var express = require('express');
var router = express.Router();
const XLSX = require('xlsx');

/* GET home page. */
router.get('/', function(req, res, next) {

  const workbook = XLSX.readFile('../MOCC/public/excel/teamDetails.xlsx');
  const sheet_name_list = workbook.SheetNames;
  teamDetails=(XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]));
  console.log(teamDetails[0].TeamName);
  teamDetails[0].TeamName="Arsenal"
  console.log(teamDetails[0].TeamName);
  const excelWrite= XLSX.utils.json_to_sheet(teamDetails);
  workbook.Sheets[sheet_name_list[0]]=excelWrite;
  XLSX.writeFile( workbook, '../MOCC/public/excel/teamDetails.xlsx') 

  res.render('index', { title: 'Express' });
});

module.exports = router;
