var xlsx = require('../index');
var assert = require('assert');
var util = require('util');
function arrays_equal(a,b) { return !(a<b || b<a); }

//load
var xls = new xlsx();
xls.loadFile('test.xlsx');
//set sheet name
var sheetsCount = xls.sheetsCount();
xls.setSheetName(0, 'sheet新名');
//get sheet
var sheet0 = xls.getSheet(0);
//write sheet
sheet0.write('A1', [ 'string', '数字', 'boolean' ]);
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
//write more data
var datas = randomRows(1000);
sheet0.write('A5', datas, true);
//copy sheet
xls.copySheet(0, 'copy拷贝');
//write file
xls.writeFile('out.xlsx');

function randomRows(rs){
	var rows = [];
	for(var r = 0; r < rs; r++){
		var row = [];
		for(var c = 0; c < 10; c++){
			row.push(Math.floor(Math.random() * 100));
		}
		rows.push(row);
	}
	return rows;
}

console.log(['memoryUsage:',util.inspect(process.memoryUsage())]);
console.log(['uptime', process.uptime()]);
//download
function download(xls, filename, httpServerResponse) {
	httpServerResponse.writeHead(200, {
		'Content-Type' : 'application/vnd.openxmlformats',
		'Content-Disposition' : 'attachment; filename="'
				+ encodeURIComponent(filename) + '.xlsx"'
	});
	httpServerResponse.end(xls.data());
};
