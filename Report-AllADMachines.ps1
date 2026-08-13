# --- CONFIGURATION ---
$DomainController = $env:LOGONSERVER.Trim('\') 
$ExportPath = "$env:USERPROFILE\Desktop\AD_Computer_Report.csv"
$MaxThreads = 50 # Adjust this to change concurrency (higher = faster)
# ---------------------

# Import Active Directory Module
if (-not (Get-Module -Name ActiveDirectory -ErrorAction SilentlyContinue)) {
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
    } catch {
        Write-Error "The Active Directory module is not installed. Please install RSAT and try again."
        exit
    }
}

Write-Host "Connecting to Domain Controller: $DomainController..." -ForegroundColor Cyan

# === STEP 1: Print the number of results (Total Domain Count) ===
Write-Host "Querying total computer count in the domain..." -ForegroundColor Cyan
try {
    $totalCount = (Get-ADComputer -Filter * -Server $DomainController).Count
    Write-Host "Total computer accounts found in the entire domain: $totalCount" -ForegroundColor Green
} catch {
    Write-Error "Failed to query Active Directory: $_"
    exit
}

# === STEPS 2, 3, & 4: Prompt, Count OU, and Confirm ===
$computers = @()
$selectionConfirmed = $false

while (-not $selectionConfirmed) {
    # Step 2: Prompt to limit search to an OU or enter "Full"
    Write-Host ""
    $selection = Read-Host "Enter the Distinguished Name (DN) of an OU to limit the search, or enter 'Full' for the entire domain"
    
    if ($selection -eq "Full" -or [string]::IsNullOrWhiteSpace($selection)) {
        Write-Host "Retrieving all computer accounts from the entire domain (this may take a moment)..." -ForegroundColor Cyan
        try {
            $computers = Get-ADComputer -Filter '*' -Properties @('Name', 'DistinguishedName', 'Enabled', 'LastLogonDate', 'OperatingSystem', 'DNSHostName', 'Description') -Server $DomainController
            $selectionConfirmed = $true
        } catch {
            Write-Error "Failed to retrieve computers: $_"
            exit
        }
    } else {
        # Step 3: If an OU was entered, print the number of machines in that OU
        Write-Host "Querying specified OU: $selection..." -ForegroundColor Cyan
        try {
            $ouComputers = Get-ADComputer -Filter '*' -SearchBase $selection -Properties @('Name', 'DistinguishedName', 'Enabled', 'LastLogonDate', 'OperatingSystem', 'DNSHostName', 'Description') -Server $DomainController
            $ouCount = $ouComputers.Count
            Write-Host "Found $ouCount computer accounts in OU: $selection" -ForegroundColor Green
            
            # Step 4: Confirm if the report should be limited to that OU
            $confirm = Read-Host "Confirm: Do you want to limit the report to this OU? (Y/N)"
            if ($confirm -match "^[Yy](es)?$") {
                $computers = $ouComputers
                $selectionConfirmed = $true
            } else {
                Write-Host "Let's try again." -ForegroundColor Yellow
            }
        } catch {
            Write-Warning "Failed to query the specified OU. Please ensure the DN is correct (e.g., OU=Computers,DC=domain,DC=local)."
            Write-Warning "Error Details: $_"
        }
    }
}

# === Run Multi-Threaded Status Checks using Runspaces ===
Write-Host "`nBeginning multi-threaded status checks on $($computers.Count) computer accounts..." -ForegroundColor Cyan

# Map AD objects to plain PSCustomObjects so they cross the thread boundary reliably
$computerQueue = foreach ($c in $computers) {
    [PSCustomObject]@{
        Name              = $c.Name
        DNSHostName       = $c.DNSHostName
        Description       = $c.Description
        OperatingSystem   = $c.OperatingSystem
        DistinguishedName = $c.DistinguishedName
        Enabled           = $c.Enabled
        LastLogonDate     = $c.LastLogonDate
    }
}

# Define the thread execution script block
$scriptBlock = {
    param($computer)
    
    # Categorize Workstation vs Server
    $type = "Workstation"
    if ($computer.OperatingSystem -like "*Server*") {
        $type = "Server"
    }

    # Extract clean OU path from DistinguishedName
    $ou = "Unknown"
    if ($computer.DistinguishedName -match "CN=.*?,(OU=.*)") {
        $ou = $Matches[1]
    } elseif ($computer.DistinguishedName -match "CN=.*?,(CN=Computers,.*)") {
        $ou = $Matches[1]
    }

    # Determine Live Status and Last Logged On User
    $liveStatus = "Offline"
    $lastUser = "N/A"

    if ($computer.Enabled) {
        # Perform a fast ping (1 second / 1000ms timeout)
        try {
            $ping = New-Object System.Net.NetworkInformation.Ping
            $pingResult = $ping.Send($computer.DNSHostName, 1000)
            if ($pingResult.Status -eq 'Success') {
                $liveStatus = "Online"
                
                # Query the system via CIM (3-second timeout)
                try {
                    $cimOptions = New-CimSessionOption -Protocol Dcom
                    $computerSystem = Get-CimInstance -ComputerName $computer.DNSHostName -ClassName Win32_ComputerSystem -ErrorAction Stop -OperationTimeoutSec 3
                    $lastUser = $computerSystem.UserName
                    
                    if ([string]::IsNullOrEmpty($lastUser)) {
                        $lastUser = "No Active User (Online)"
                    }
                }
                catch {
                    $lastUser = "Access Denied / CIM Blocked"
                }
            }
        }
        catch {
            $liveStatus = "Unreachable (DNS/Network issue)"
        }
    } else {
        $liveStatus = "Disabled Account"
        $lastUser = "N/A (Disabled)"
    }

    # Format the Last AD Login Date
    $lastADLogin = if ($computer.LastLogonDate) { $computer.LastLogonDate.ToString("yyyy-MM-dd HH:mm:ss") } else { "Never" }

    # Return result to main pipeline
    [PSCustomObject]@{
        "Hostname"         = $computer.Name
        "DNS Hostname"     = $computer.DNSHostName
        "Description"      = $computer.Description
        "Type"             = $type
        "OperatingSystem"  = $computer.OperatingSystem
        "OU"               = $ou
        "Account Status"   = if ($computer.Enabled) { "Enabled" } else { "Disabled" }
        "Current Status"   = $liveStatus
        "Last AD Login"    = $lastADLogin
        "Last Logged User" = $lastUser
    }
}

# Create Runspace Pool for multithreading
$RunspacePool = [runspacefactory]::CreateRunspacePool(1, $MaxThreads)
$RunspacePool.Open()
$jobs = New-Object System.Collections.ArrayList

# Dispatch threads
foreach ($computer in $computerQueue) {
    $powershell = [PowerShell]::Create().AddScript($scriptBlock).AddArgument($computer)
    $powershell.RunspacePool = $RunspacePool
    
    $handle = $powershell.BeginInvoke()
    $jobs.Add([PSCustomObject]@{
        PowerShell = $powershell
        Handle     = $handle
        Harvested  = $false
    }) | Out-Null
}

# Monitor threads and gather results reactively
$results = New-Object System.Collections.ArrayList
$totalJobs = $jobs.Count
$completedCount = 0

Write-Host "Spawning $MaxThreads threads. Starting sweeps..." -ForegroundColor Yellow

while ($completedCount -lt $totalJobs) {
    foreach ($job in $jobs) {
        # Check if the job has completed and has not been harvested yet
        if ($job.Handle.IsCompleted -and -not $job.Harvested) {
            # Collect data from thread (this returns a PSDataCollection wrapper)
            $threadOutput = $job.PowerShell.EndInvoke($job.Handle)
            
            # UNPACK: Extract the actual PSCustomObject inside the collection
            foreach ($result in $threadOutput) {
                $results.Add($result) | Out-Null
                
                # Print real-time status update to terminal
                $statusColor = "DarkGray"
                if ($result."Current Status" -eq "Online") {
                    $statusColor = "Green"
                } elseif ($result."Current Status" -eq "Offline") {
                    $statusColor = "Yellow"
                }
                
                Write-Host "$($result.Hostname.PadRight(15)) | Status: $($result.'Current Status'.PadRight(8)) | User: $($result.'Last Logged User')" -ForegroundColor $statusColor
            }
            
            # Mark job as harvested so we don't process it again
            $job.Harvested = $true
            $completedCount++
            
            # Update Master Windows Progress Bar
            $percent = [math]::Round(($completedCount / $totalJobs) * 100)
            Write-Progress -Activity "Auditing Active Directory Computers" -Status "Completed: $completedCount of $totalJobs ($percent%)" -PercentComplete $percent
        }
    }
    # Small pause to prevent high CPU utilization on the monitoring loop
    Start-Sleep -Milliseconds 50
}

# Clean up Runspaces
foreach ($job in $jobs) {
    $job.PowerShell.Dispose()
}
$RunspacePool.Close()
$RunspacePool.Dispose()

# Remove active progress bar
Write-Progress -Activity "Auditing Active Directory Computers" -Completed

# Export to CSV
$results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
Write-Host "`nProcess complete! Report exported to: $ExportPath" -ForegroundColor Green

# Prompt to open GridView for immediate filtering/viewing
$title = "AD Computer Audit - $($results.Count) Records"
$results | Out-GridView -Title $title