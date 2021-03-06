const express = require('express')
const bodyParser = require('body-parser');
const app = express();
app.use(express.static('public'));
app.use(bodyParser.urlencoded({ extended: true }));
//app.use(bodyParser.text({type: '*/*'}));

const XlsxPopulate = require('xlsx-populate');
var pdf = require('html-pdf');

app.get('/xls', function(req,res) {
    var listax=req.query.lista;
    var lista=listax.split(',');
    XlsxPopulate.fromFileAsync("./wtg_checklist_it.xlsx")
        .then(workbook => {
            var naz=req.query.naz+"";
            console.log("NAZIONALITA: "+naz);
            console.log("Supervisore: "+req.query.super);
            workbook.sheet("CHECKLIST").cell("C7").value(naz);           
            workbook.sheet("CHECKLIST").cell("D7").value(req.query.sito);
            workbook.sheet("CHECKLIST").cell("D9").value(req.query.manuf);   
            workbook.sheet("CHECKLIST").cell("E9").value(req.query.model);   
            workbook.sheet("CHECKLIST").cell("E7").value(req.query.posto);
            workbook.sheet("CHECKLIST").cell("F7").value(req.query.nturb);
            workbook.sheet("CHECKLIST").cell("I7").value(req.query.super);
            workbook.sheet("CHECKLIST").cell("K7").value(req.query.manut);
            workbook.sheet("CHECKLIST").cell("I9").value(req.query.turbi);
            for(let i=0;i<lista.length;i++) {
            // Modify the workbook.
                var x=lista[i];
                var n=i+128;
                var cella='I'+n;  
 //               console.log(parseInt(x, 10));           
                workbook.sheet("CHECKLIST").cell(cella).value(parseInt(x, 10));   
 //               console.log(workbook.sheet("CHECKLIST").cell("B"+n).value())         
            // Log the value.
            }                        
            //return workbook.toFileAsync("./out.xlsx");
            
            workbook.outputAsync().then(function(aaa) {
                res.setHeader('Content-type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' );
                res.end(aaa);
            })            

        });    
})

app.post('/pdf',function(req,res) {
    var html=req.body.valo;
//    console.log(html)    
    var options={
        format: "A4",
        orientation:"portrait",
        border: {
            top: "20px",
            bottom:"20px",
            left: "20px",
            right:"20px"
        }
    }
    pdf.create(html,options).toBuffer(function(err, buf){
        res.setHeader('Content-type', 'application/pdf' );
        res.end(buf);
    })
})

var porta=process.env.PORT || 3000;
app.listen(porta, function () {
  console.log('servr started on port '+porta)
})