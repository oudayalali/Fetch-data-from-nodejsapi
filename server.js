const xlsx = require('xlsx')
const express = require('express')
const bodyParser = require('body-parser');
const app = express()
const cors = require('cors');
app.use(cors())
app.use(express.json())
app.use(bodyParser.urlencoded({ extended: false }));
const workbook = xlsx.readFile("ExcelDb.xlsx");
let worksheets = {};
for (const sheetName of workbook.SheetNames)
{
worksheets[sheetName] = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);
}


app.get('/excel', function (req, res) {
    
    var sheet_name_list = workbook.SheetNames;
    res.send(xlsx.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]))
   
    
})
app.post('/addexcel', function (req, res) {  

        worksheets.Sheet1.push({
            first_name:req.body.first_name,  
            last_name:req.body.last_name 
        }) ;

xlsx.utils.sheet_add_json(workbook.Sheets["Sheet1"], worksheets.Sheet1)
xlsx.writeFile(workbook , "ExcelDb.xlsx");
res.send(worksheets); 
})

app.listen(process.env.PORT || 8080);
//app.listen(8080);