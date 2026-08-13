// MUST use process.env.PORT for Windows App Service
const port = process.env.PORT || 3000; 

app.listen(port, () => {
    console.log(`Server is running on port ${port}`);
});