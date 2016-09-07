var express = require('express');
var router = express.Router();

var sql = require("seriate");
var prodID = 5;
var PileDim = 880;

var excelbuilder = require('msexcel-builder');

// Change the config settings to match your 
// SQL Server and database

// Create a new workbook file in current working-path
  var workbook = excelbuilder.createWorkbook('./', 'Summary_List.xlsx')

  // Create a new worksheet with 10 columns and 12 rows
  var sheet1 = workbook.createSheet('sheet1', 18, 12);
  var sheet2 = workbook.createSheet('sheet2', 18, 12);

  // Fill some data
  sheet1.set(1, 1, 'Concrete Start Date');
  sheet1.set(2, 1, 'Platform Level (mRL)');
  sheet1.set(3, 1, 'Cut-off Level (mRL)');
  sheet1.set(4, 1, 'Hit Rock Level (mRL)');
  sheet1.set(5, 1, 'Toe Level (mRL)');
  sheet1.set(6, 1, 'Bored Depth PPL (m)');
  sheet1.set(7, 1, 'Pile Length (m)');
  sheet1.set(8, 1, 'Cavity (m)');
  sheet1.set(9, 1, 'Total Rock Coring (m)');
  sheet1.set(10, 1, 'Rock Socket (m)');
  sheet1.set(11, 1, 'Reinforcement Contain');
  sheet1.set(12, 1, 'Helical/Spiral');
  sheet1.set(13, 1, 'Cage Length');
  sheet1.set(14, 1, 'Theoretical');
  sheet1.set(15, 1, 'Actual');
  sheet1.set(16, 1, 'Wastage (%)');
  sheet1.set(17, 1, 'Grade');
  sheet1.set(18, 1, 'DO Number');
// end create a new worksheet

var config = {  
    "server": "bjic5w3pth.database.windows.net",
    "user": "bpldbsa@bjic5w3pth",
    "password": "Password999",
    "database": "bpldb2",
    "procName": "ups_SummaryList",
    "options" : {encrypt: true}  
};

sql.setDefaultConfig( config );

  
sql.execute( {      
        query: "execute dbo.usp_SummaryList "+prodID+","+PileDim+" "
    } ).then( function( result ) {

        var totaldata = result.length;
        
        for (var i = 0; i <totaldata;i++){
        /*var irow = parseInt(i+1);
        
        if (irow > 1){
        sheet1.set(1,irow, "test");*/
        console.log( result[i].v1 );
        //}

        
        }

         // Save it
  workbook.save(function(err){
    if (err)
      throw err;
    else
      console.log('congratulations, your workbook created');
  }); 

        
    }, function( err ) {
        console.log( "Something bad happened:", err );
    } );

//get excelbuilder
 


  



router.get('/', function(req, res) {
  res.render('index', { title: 'Express',test: 'Node JS'});
 
})

module.exports = router;
