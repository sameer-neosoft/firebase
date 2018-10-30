const express = require("express");
const bodyParser = require("body-parser");
const app = express();
const admin = require("firebase-admin");
const XLSX = require('xlsx');
const functions = require('firebase-functions');
const serviceAccount = require('./serviceAccountKey.json');

admin.initializeApp({
	credential: admin.credential.cert(serviceAccount)
});

app.use(bodyParser.json());
app.use(
  bodyParser.urlencoded({
    extended: true
  })
);

app.get("/addDocument/:docName", function (req, res) {
  var db = admin.firestore();
  try{
    var workbook = XLSX.readFile('./files/'+req.params.docName+'.xlsx');
  }catch(e)
  {
    res.send('Something wrong with file name, please check your file name. <br> Expected URL => http://DOMAIN.COM/addDocument/{filename} <br>'+ e);
  }
  
  var sheet_name_list = workbook.SheetNames;
  var parent_exists = false;
  var data = [];
  
  // console.log(XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[1]]))
  // res.send(XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[1]]));

  sheet_name_list.forEach(function(y) {
    var worksheet = workbook.Sheets[y];
    var headers = {};
    var childHeaders = {};
    var childCol = {};
    var childValue = {};
    //console.log(xsheets);
    for(z in worksheet) {
      if(z[0] === '!') continue;
      var tt = 0;
      for (var i = 0; i < z.length; i++) {
          if (!isNaN(z[i])) {
              tt = i;
              break;
          }
      }
      console.log(worksheet[z]);
      var col = z.substring(0,tt);
      var childCol = z.substring(1,tt);
      var row = parseInt(z.substring(tt));
      var value = worksheet[z].v;

      if(row == 1 && value) 
      {
        headers[col] = value;
        if(value.includes('parent') )
        {
          parent_exists = true;
          var parent = value;
          var parentCol = col;
          var newparent = parent.split("-").pop();
          headers[col] = newparent;
        }
        continue;
      }
      if(parent_exists && row == 2)
      {
        childHeaders[col] = value;
        childCol = col;
      }
      if(!data[row]) data[row]={};
      if(childCol != col)
      {
        if(childHeaders[col] )
        {
          childValue[childHeaders[col]] = value.toString();
          data[row][headers[parentCol]] = Object.assign({},childValue);
        }else{
          data[row][headers[col]] = value.toString();
        }
      }
    }
      data.shift();
      // data.forEach(function(element) {
      // 	var docRef = db.collection(y).doc();
      // 	var setAda = docRef.set(element);
      // });
  });
  res.send(data);
});

app.listen(5500, function (err, res) {
  if (!err) console.log("Server is Running on port 5500");
});
