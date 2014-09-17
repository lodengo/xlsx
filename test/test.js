var xlsx = require('../index');

var xls = new xlsx();
xls.loadFile('test.xlsx');

var sheet0 = xls.getSheet(0);
console.log(sheet0.cell('A1'));
sheet0.write('A1', 123);
xls.writeFile('11.xlsx');