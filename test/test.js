var xlsx = require('../index');
var assert = require('assert');
var util = require('util');

var xls = new xlsx();
xls.loadFile('test.xlsx');

var sheetsCount = xls.sheetsCount();
xls.setSheetName(0, 'sheet新名');

var sheet0 = xls.getSheet(0);
sheet0.write('A1', [ 'string', '数字', 'boolean' ]);
sheet0.write('A2', [ [ 'aa', 30, true ], [ 'bb', 40, false ] ]);
sheet0.write('A4', 'cc');

assert.equal(sheet0.read('B2'), 30);
assert.equal(sheet0.read('C2'), true);
assert.equal(sheet0.read('B3'), 40);
assert.equal(sheet0.read('A4'), 'cc');

//性能测试
var datas = randomRows(10000);
sheet0.write('A5', datas, true);

xls.copySheet(0, 'copy拷贝');

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
// 下载
function download(xls, filename, httpServerResponse) {
	httpServerResponse.writeHead(200, {
		'Content-Type' : 'application/vnd.openxmlformats',
		'Content-Disposition' : 'attachment; filename="'
				+ encodeURIComponent(filename) + '.xlsx"'
	});
	httpServerResponse.end(xls.data());
};
