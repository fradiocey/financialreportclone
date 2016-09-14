var express = require('express');
var router = express.Router();
var async = require("async");
var SMB2 = require('smb2');

var fs = require('fs-extra')
const MFS = require("mfs");

/*// create an SMB2 instance 
var smb2Client = new SMB2({
  share:'\\\\mymswfs01\\'
, domain:'sunway.com'
, username:'mazim'
, password:'Fradiocey10'
, debug: false,
  autoCloseTimeout: 0
});

smb2Client.mkdir('180-Public\\BPMS\\Test', function (err) {
    if (err) throw err;
    console.log('Folder created!');
});
*/

 
/*fs.copy('smb:\\\\mymswfs01\\180-Public\\BPMS\\Book1.xlsx', '//mymswfs01/180-Public/Test/BPMS/sample.xlsx', function (err) {
  if (err) return console.error(err)
  console.log("success!")
}) // copies file */


/*function copyFile(source, target, cb) {
  var cbCalled = false;

  var rd = fs.createReadStream(source);
  rd.on("error", function(err) {
    done(err);
  });
  var wr = fs.createWriteStream(target);
  wr.on("error", function(err) {
    done(err);
  });
  wr.on("close", function(ex) {
    done();
  });
  rd.pipe(wr);

  function done(err) {
    if (!cbCalled) {
      cb(err);
      cbCalled = true;
    }
  }

 
/*fs.copy('/tmp/mydir', '/tmp/mynewdir', function (err) {
  if (err) return console.error(err)
  console.log('success!')
}) // copies directory, even if it has subdirectories or files*/

var sql = require("seriate");
var prodID = 5;
var PileDim = 880;
var operation = 0;

var excelbuilder = require('msexcel-builder');
//var excelbuilder = require('msexcel-builder-colorfix-intfix');
//var excelbuilder = require('msexcel-builder');
//var excelbuilder = require('msexcel-builder-colorfix');
var workbookName = [];
var workbookID = [];
var pilediameter = [];

var totalexcel = 0;
var totalsheet = 0;
var completed = 0;

// Change the config settings to match your 
// SQL Server and database

// Create a new workbook file in current working-path
  //var workbook = excelbuilder.createWorkbook('./', 'Summary_List.xlsx')



var config = {  
    "server": "bjic5w3pth.database.windows.net",
    "user": "bpldbsa@bjic5w3pth",
    "password": "Password999",
    "database": "bpldb2",
    "procName": "ups_SummaryList",
    "options" : {encrypt: true}  
};

sql.setDefaultConfig( config );

/* template execute
sql.execute( {  
} ).then( function( result ) {
  
}, function( err ) {
        console.log( "Something bad happened:", err );
    } );
*/

  /*for (var j = 0; j < totalexcel; j++){
  var workbook1 = excelbuilder.createWorkbook('./', 'Summary List for'+workbookID[j]+' - '+workbookName[j]+'.xlsx');
  }*/

 var workbook1 = "";
//get file name
sql.execute( { 
  query: "SELECT id, ProjectCode, ProjectName FROM BplProject " 
} ).then( function( result ) {
  totalexcel = result.length;
  //for (var i=0; i < 4; i++){
  //workbookID.push(result[i].ProjectCode);
  //workbookName.push(result[i].ProjectName);
  //getPile(result[i].id,result[i].ProjectCode,result[i].ProjectName);
  //}
  
  var y = 0;
  var loopWorkbook = function(result){
    var projCode = result[y].ProjectCode
     console.log ("RUN WORKBOOK"+y) 
    getPile(result[y].id,result[y].ProjectCode,result[y].ProjectName,function(){
    
      y++
      if (y < result.length){
      loopWorkbook(result);
      }
      
    })
  }
  //start loopWorkbook
  loopWorkbook(result) 

  
}, function( err ) {
        console.log( "Something bad happened:", err );s
    } );

function getPile(id,workcode,workname,callback1){
workbook1 = excelbuilder.createWorkbook('./', 'Summary List for '+workcode+'-'+workname+'.xlsx');

console.log(workbook1)
// get pile diameter
 
sql.execute( {  
  query: "SELECT distinct pilediameter FROM bplpile where Project_Id = "+id+" order by PileDiameter asc " 
} ).then( function( result ) {
  //console.log(result)
  totalsheet = result.length;
  var x = 0;
  for (var i = 0; i < totalsheet; i++){
  pilediameter.push(result[i].pilediameter)
  }
 
  var i = 0;
  var loopSheet = function(result){
    
    getExcel(id,result[i].pilediameter,workcode,workname,function(){
      i++
      if (i < result.length){
        loopSheet(result);
         completed = 0
      }
      else{
        //console.log("Completed")
        completed = 1
         workbook1.save(function (err) {
        if (err)
        throw err;
      else
      console.log('congratulations, your workbook created');
  });  
        callback1();
      
       
      }
    })
    
  }
  // start loopSheet
  
  loopSheet(result)

  //callback();




}, function( err ) {
        console.log( "Something bad happened:", err );
    } );

}



// execute strdproc 
function getExcel(id,pileno,projCode,projName,callback){

sql.execute( {      
        query: "execute dbo.usp_SummaryList "+id+","+pileno+""
    } ).then( function( result ) {
 
  var totaldata = result.length+5;
  //for (var k=0; k < totalsheet; k++){
  
  var sheet1 = workbook1.createSheet(''+pileno+'', 26, totaldata)
  //}
    // header
  sheet1.set(2, 1, 'Summary List for '+projCode+'-'+projName+'' );
  //type
  sheet1.set(1, 3, 'Pile Details');
  sheet1.set(6, 3, 'General Details');
  sheet1.set(11, 3, 'Pile Details');
  sheet1.set(17, 3, 'Steel Cage');
  sheet1.set(21, 3, 'Concrete');
  // type center
  sheet1.align(1, 3, 'center');
  sheet1.align(6, 3, 'center');
  sheet1.align(11, 3, 'center');
  sheet1.align(17, 3, 'center');
  sheet1.align(21, 3, 'center');
  //merge table
  sheet1.merge({col:1,row:3},{col:4,row:3});
  sheet1.merge({col:6,row:3},{col:9,row:3});
  sheet1.merge({col:11,row:3},{col:15,row:3});
  sheet1.merge({col:17,row:3},{col:19,row:3});
  sheet1.merge({col:21,row:3},{col:26,row:3});
  // type border
  sheet1.border(1, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(2, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(3, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(4, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});

  sheet1.border(1, 4, {left:'medium',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(2, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(3, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(4, 4, {left:'thin',top:'thin',right:'medium',bottom:'medium'});

  sheet1.border(6, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(7, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(8, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(9, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});

  sheet1.border(6, 4, {left:'medium',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(7, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(8, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(9, 4, {left:'thin',top:'thin',right:'medium',bottom:'medium'});
  
  sheet1.border(11, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(12, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(13, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(14, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(15, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});

  sheet1.border(11, 4, {left:'medium',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(12, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(13, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(14, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(15, 4, {left:'thin',top:'thin',right:'medium',bottom:'medium'});

  sheet1.border(17, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(18, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(19, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
 
  sheet1.border(17, 4, {left:'medium',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(18, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(19, 4, {left:'thin',top:'thin',right:'medium',bottom:'medium'});

  sheet1.border(21, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(22, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(23, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(24, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(25, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  sheet1.border(26, 3, {left:'medium',top:'medium',right:'medium',bottom:'thin'});
  
  sheet1.border(21, 4, {left:'medium',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(22, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(23, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(24, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(25, 4, {left:'thin',top:'thin',right:'thin',bottom:'medium'});
  sheet1.border(26, 4, {left:'thin',top:'thin',right:'medium',bottom:'medium'});
 

  // Fill some data
  sheet1.set(1, 4, 'Pile No.');
  sheet1.set(2, 4, 'Rig Number (last rig to operate on pile)');
  sheet1.set(3, 4, 'Boring Start Date');
  sheet1.set(4, 4, 'Concrete Start Date');
  //gap
  sheet1.set(5, 4, '');
  //end gap
  sheet1.set(6, 4, 'Platform Level (mRL)');
  sheet1.set(7, 4, 'Cut-off Level (mRL)');
  sheet1.set(8, 4, 'Hit Rock Level (mRL)');
  sheet1.set(9, 4, 'Toe Level (mRL)');
   //gap
  sheet1.set(10, 4, '');
  //end gap
  sheet1.set(11, 4, 'Bored Depth PPL (m)');
  sheet1.set(12, 4, 'Pile Length (m)');
  sheet1.set(13, 4, 'Cavity (m)');
  sheet1.set(14, 4, 'Total Rock Coring (m)');
  sheet1.set(15, 4, 'Rock Socket (m)');
    //gap
  sheet1.set(16, 4, '');
  //end gap
  sheet1.set(17, 4, 'Reinforcement Content');
  sheet1.set(18, 4, 'Helical/Spiral');
  sheet1.set(19, 4, 'Cage Length');
     //gap
  sheet1.set(20, 4, '');
  //end gap
  sheet1.set(21, 4, 'Theoretical');
  sheet1.set(22, 4, 'Actual');
  sheet1.set(23, 4, 'Wastage (%)');
  sheet1.set(24, 4, 'Grade');
  sheet1.set(25, 4, 'DO Number');
  sheet1.set(26, 4, 'Concrete Volume');
// end create a new worksheet

// wrap header true
sheet1.wrap(1, 4, 'true');
sheet1.wrap(2, 4, 'true');
sheet1.wrap(3, 4, 'true');
sheet1.wrap(4, 4, 'true');
sheet1.wrap(5, 4, 'true');
sheet1.wrap(6, 4, 'true');
sheet1.wrap(7, 4, 'true');
sheet1.wrap(8, 4, 'true');
sheet1.wrap(9, 4, 'true');
sheet1.wrap(10, 4, 'true');
sheet1.wrap(11, 4, 'true');
sheet1.wrap(12, 4, 'true');
sheet1.wrap(13, 4, 'true');
sheet1.wrap(14, 4, 'true');
sheet1.wrap(15, 4, 'true');
sheet1.wrap(16, 4, 'true');
sheet1.wrap(17, 4, 'true');
sheet1.wrap(18, 4, 'true');
sheet1.wrap(19, 4, 'true');
sheet1.wrap(20, 4, 'true');
sheet1.wrap(21, 4, 'true');
sheet1.wrap(22, 4, 'true');
sheet1.wrap(23, 4, 'true');
sheet1.wrap(24, 4, 'true');
sheet1.wrap(25, 4, 'true');
sheet1.wrap(26, 4, 'true');

// header center

sheet1.align(1, 4, 'center');
sheet1.align(2, 4, 'center');
sheet1.align(3, 4, 'center');
sheet1.align(4, 4, 'center');
sheet1.align(5, 4, 'center');
sheet1.align(6, 4, 'center');
sheet1.align(7, 4, 'center');
sheet1.align(8, 4, 'center');
sheet1.align(9, 4, 'center');
sheet1.align(10, 4, 'center');
sheet1.align(11, 4, 'center');
sheet1.align(12, 4, 'center');
sheet1.align(13, 4, 'center');
sheet1.align(14, 4, 'center');
sheet1.align(15, 4, 'center');
sheet1.align(16, 4, 'center');
sheet1.align(17, 4, 'center');
sheet1.align(18, 4, 'center');
sheet1.align(19, 4, 'center');
sheet1.align(20, 4, 'center');
sheet1.align(21, 4, 'center');
sheet1.align(22, 4, 'center');
sheet1.align(23, 4, 'center');
sheet1.align(24, 4, 'center');
sheet1.align(25, 4, 'center');
sheet1.align(26, 4, 'center');

for (var i = 5; i < totaldata; i++){
//parseFloat
var Platform = result[i-5].v5;
if (typeof Platform  != "undefined" || Platform != null){
var PlatformparseFloat = parseFloat(Platform).toFixed(3)
//console.log(PlatformparseFloat)
sheet1.set(6, i, PlatformparseFloat);
}
   
sheet1.set(1, i, result[i-5].v1);
sheet1.set(2, i, result[i-5].v2);
sheet1.set(3, i, result[i-5].v3);
sheet1.set(4, i, result[i-5].v4);
sheet1.set(5, i, "");

sheet1.set(7, i, result[i-5].v6);
sheet1.set(8, i, result[i-5].v7);
sheet1.set(9, i, parseFloat(result[i-5].v8).toFixed(3));
sheet1.set(10, i, "");
sheet1.set(11, i, parseFloat(result[i-5].v9).toFixed(3));
sheet1.set(12, i, parseFloat(result[i-5].v10).toFixed(3));
sheet1.set(13, i, result[i-5].v11);
sheet1.set(14, i, parseFloat(result[i-5].v12).toFixed(1));
sheet1.set(15, i, parseFloat(result[i-5].v13).toFixed(1));
sheet1.set(16, i, "");
sheet1.set(17, i, result[i-5].v14);
sheet1.set(18, i, result[i-5].v15);
sheet1.set(19, i, result[i-5].v16);
sheet1.set(20, i, "");
sheet1.set(21, i, parseFloat(result[i-5].v17).toFixed(1));
sheet1.set(22, i, parseFloat(result[i-5].v18).toFixed(1));
sheet1.set(23, i, parseFloat(result[i-5].v19).toFixed(1));
sheet1.set(24, i, result[i-5].v20);
sheet1.set(25, i, result[i-5].v21);
sheet1.set(26, i, result[i-5].v22);


// wrap true
sheet1.wrap(1, i, 'true');
sheet1.wrap(2, i, 'true');
sheet1.wrap(3, i, 'true');
sheet1.wrap(4, i, 'true');
sheet1.wrap(5, i, 'true');
sheet1.wrap(6, i, 'true');
sheet1.wrap(7, i, 'true');
sheet1.wrap(8, i, 'true');
sheet1.wrap(9, i, 'true');
sheet1.wrap(10, i, 'true');
sheet1.wrap(11, i, 'true');
sheet1.wrap(12, i, 'true');
sheet1.wrap(13, i, 'true');
sheet1.wrap(14, i, 'true');
sheet1.wrap(15, i, 'true');
sheet1.wrap(16, i, 'true');
sheet1.wrap(17, i, 'true');
sheet1.wrap(18, i, 'true');
sheet1.wrap(19, i, 'true');
sheet1.wrap(20, i, 'true');
sheet1.wrap(21, i, 'true');
sheet1.wrap(22, i, 'true');
sheet1.wrap(23, i, 'true');
sheet1.wrap(24, i, 'true');
sheet1.wrap(25, i, 'true');
sheet1.wrap(26, i, 'true');

// border
sheet1.border(1, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(2, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(3, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(4, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
//sheet1.border(5, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(6, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(7, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(8, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(9, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
//sheet1.border(10, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(11, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(12, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(13, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(14, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(15, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
//sheet1.border(16, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(17, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(18, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(19, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
//sheet1.border(20, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(21, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(22, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(23, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(24, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(25, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});
sheet1.border(26, i, {left:'thin',top:'thin',right:'thin',bottom:'thin'});


//sheet1.numberFormat(2,1, 10); // equivalent to above
    
}

sheet1.width(1, '10');
sheet1.width(2, '10');
sheet1.width(3, '20');
sheet1.width(4, '20');
sheet1.width(5, '5');
sheet1.width(6, '10');
sheet1.width(7, '10');
sheet1.width(8, '10');
sheet1.width(9, '10');
sheet1.width(10, '5');
sheet1.width(11, '10');
sheet1.width(12, '10');
sheet1.width(13, '10');
sheet1.width(14, '10');
sheet1.width(15, '10');
sheet1.width(16, '5');
sheet1.width(17, '10');
sheet1.width(18, '10');
sheet1.width(19, '10');
sheet1.width(20, '5');
sheet1.width(21, '10');
sheet1.width(22, '10');
sheet1.width(23, '10');
sheet1.width(24, '10');
sheet1.width(25, '10');
sheet1.width(26, '10');

sheet1.merge({col:1,row:1},{col:23,row:1});
callback();

 
    
    }, function( err ) {
        console.log( "Something bad happened:", err );
    } );
//    
}//get excelbuilder
// end function 





router.get('/', function(req, res) {
  res.render('index', { title: 'Express',test: 'Node JS'});
 
})

module.exports = router;
