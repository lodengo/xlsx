var xlsx = require('../index');
var assert = require('assert');
var util = require('util');
function arrays_equal(a,b) { return !(a<b || b<a); }

//load
var xls = new xlsx();
xls.loadFile('test.xlsx');
//set sheet name
assert.equal(xls.sheetsCount(), 2);
xls.setSheetName(0, 'aa');
//get sheet
var sheet0 = xls.getSheet(0);
//write sheet
sheet0.write('A1', [ 'string', 'num', 'boolean' ]);
sheet0.write('A2', [ [ 'aa', 30, true ], [ 'bb', 40, false ] ]);
sheet0.write('A4', 'cc');
//read sheet
assert.equal(sheet0.read('B2'), 30);
assert(arrays_equal(sheet0.read('A2:C2'), [ 'aa', 30, true ]));
assert.equal(sheet0.read('C2'), true);
assert(arrays_equal(sheet0.read('A2:A3'), [ 'aa', 'bb' ]));
assert.equal(sheet0.read('B3'), 40);
assert(arrays_equal(sheet0.read('A2:C3'), [ [ 'aa', 30, true ], [ 'bb', 40, false ] ]));
assert.equal(sheet0.read('A4'), 'cc');

//copy sheet with chart
var sheet2 = xls.copySheet(1, 'bb');
sheet2.write('A1', [40, 60, 20]);

var sheet3 = xls.copySheet(1, 'cc');
sheet3.write('A1', [15, 23, 36]);

//sheet count
assert.equal(xls.sheetsCount(), 4);

//hidden sheet
xls.setSheetVisible(1, false);

//write file
xls.writeFile('out.xlsx');

//express download 
//app.get('/download', function(req, res){xls.download('aa', res);});

console.log(['memoryUsage:',util.inspect(process.memoryUsage())]);
console.log(['uptime', process.uptime()]);

