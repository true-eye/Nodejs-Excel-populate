const express = require('express')
const bodyParser = require('body-parser')
const knex = require('knex')(require('./knexfile'))
const crypto = require('crypto')
const XlsxPopulate = require('xlsx-populate');

const app = express()
app.use(express.static(__dirname + '/public'))
app.use(bodyParser.json())

var path = require('path')

app.get('/proceso/:id', function(req, res)
{
  XlsxPopulate.fromFileAsync("./template.xlsx")
    .then(workbook => {
        // Modify the workbook.
        var ws = workbook.sheet("Hoja2");
        var result = '';
        var total = 0;
        /* --------------------------------- smoke init ----------------------------- */
        var smoke = new Array(2)
        for( var j = 0; j < 2; j++ )
        {
          smoke[j] = []
          for( var k = 0; k < 7; k++ )
            smoke[j][k] = 0;
        }
        /* --------------------------------- sport init ----------------------------- */
        var sport = new Array(3).fill(0).map(() => new Array(3).fill(0));
        knex.select('*')
        .from('result_9f')
        .where('proceso', req.params.id)
        .then((rows) => {
          var i = 2;
          for (row of rows)
          {
            index = row['id'];
            var gender = row['d1'] - 1
            /* --------------------------------------------------------- worksheet 1 --------------------------------------------------------- */
            ws.row(i).cell(1).value(row['id']);
            ws.row(i).cell(2).value(row['d1']);

            if( row['p1'] == null ) {
              ws.row(i).cell(3).value('');
              smoke[gender][0] = smoke[gender][0] + 1;
            }
            else if( row['p1'] == '0' )
            {
              ws.row(i).cell(3).value('NO');
              smoke[gender][0] = smoke[gender][0] + 1;
            }
            else
            {
              ws.row(i).cell(3).value(row['p1']);
              smoke[gender][row['p1']] = smoke[gender][row['p1']] + 1;
              smoke[gender][6] += 1;
            }

            if( row['p1'] == null || row['p1'] == 0)
              ws.row(i).cell(4).value('NO');
            else
              ws.row(i).cell(4).value('SI');

            ws.row(i).cell(5).value = '';

            if( row['p2'] == '0' )
            {
              sport[1][gender]++; sport[1][2]++; sport[2][gender]++;
              ws.row(i).cell(6).value('NO');
            }
            else
            {
              sport[0][gender]++; sport[0][2]++; sport[2][gender]++;
              ws.row(i).cell(6).value('SI');
            }

            if( row['p3'] == null )
              ws.row(i).cell(7).value('');
            else if( row['p3'] == '0' )
              ws.row(i).cell(7).value('NO');
            else
              ws.row(i).cell(7).value(row['p3']);

            if( row['p4'] == null )
              ws.row(i).cell(9).value('');
            else
              ws.row(i).cell(9).value(row['p4']);

            if( row['p5'] == null )
              ws.row(i).cell(10).value('');
            else
              ws.row(i).cell(10).value(row['p5']);
            
            if( row['p6'] == null )
              ws.row(i).cell(11).value('');
            else
              ws.row(i).cell(11).value(row['p6']);
            i++;
          }
          workbook.toFileAsync("./report.xlsx");
          res.download( __dirname + "/report.xlsx")
        })
    });
});
//______________ port listen ________________________
app.listen(7555, () => {
  console.log('Server running on http://localhost:7555')
})