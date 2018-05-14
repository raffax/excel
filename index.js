const express = require('express')
const bodyParser = require('body-parser');
const app = express();
app.use(express.static('public'));
app.use(bodyParser.urlencoded({ extended: true }));
//app.use(bodyParser.json())
const XlsxPopulate = require('xlsx-populate');

app.get('/xls', function(req,res) {
    var listax=req.query.lista;
    var lista=listax.split(',');
    XlsxPopulate.fromFileAsync("./wtg_checklist_it.xlsx")
        .then(workbook => {
            for(let i=0;i<lista.length;i++) {
            // Modify the workbook.
                var x=lista[i];
                var n=i+128;
                var cella='I'+n;  
                console.log(parseInt(x, 10));           
                workbook.sheet("CHECKLIST").cell(cella).value(parseInt(x, 10));   
                console.log(workbook.sheet("CHECKLIST").cell("B"+n).value())         
            // Log the value.
            }                        
            //return workbook.toFileAsync("./out.xlsx");
            
            workbook.outputAsync().then(function(aaa) {
                res.setHeader('Content-type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' );
                res.end(aaa);
            })            

        });    
})

var porta=process.env.PORT || 3000;
app.listen(porta, function () {
  console.log('servr started on port '+porta)
})