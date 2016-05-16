var XLSX = require('xlsx');
var workbook = XLSX.readFile('input.xls');
var data = JSON.parse(require('fs').readFileSync('data.json', 'utf8'));


function Workbook() {
    if(!(this instanceof Workbook)) return new Workbook();
    this.SheetNames = [];
    this.Sheets = {};
}

var newWb = new Workbook();

var sheet_name_list = workbook.SheetNames;


sheet_name_list.forEach(function(y) { /* iterate through sheets */
    var worksheet = workbook.Sheets[y];

    newWb.SheetNames.push(y);
    newWb.Sheets[y] = worksheet;

    for (var z in worksheet) {
        /* all keys that do not begin with "!" correspond to cell addresses */
        if(z[0] !== '!') {
            console.log(y + "!" + z + "=" + JSON.stringify(worksheet[z].v));
            for(var replaceKey in data) {
                if(worksheet[z].v.match(replaceKey) !== null) {
                    worksheet[z].v = data[replaceKey];
                }
            }
        };
    }
});

XLSX.writeFile(newWb, 'output.xlsx');