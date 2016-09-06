var express = require('express');
var router = express.Router();

/* GET home page. */

router.get('/', function(req, res) {
  res.render('index', { title: 'Express' });
  
  
 
});

router.all('/', function (req, res) {
console.log("Test")
});

module.exports = router;
