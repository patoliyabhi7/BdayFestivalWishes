const dotenv = require("dotenv");
dotenv.config({ path: "./.env" });

const app = require("./app");

app.get('/', (req,res)=>{
    res.status(200).send("Welcome!!")
})

const port = process.env.PORT || 8000;
app.listen(port, ()=>{
    console.log(`App is running on port ${port}`)
})