var JSZip = require("jszip");
var ltx = require('ltx');
var fs = require('fs');
var sheet = require('./sheet');

var xlsx = module.exports = function xlsx() {
	this._fe = {};
	this._ss = {}; // sharedStrings
};

xlsx.prototype.loadData = function(data) {
	var self = this;
	var zip = new JSZip(data);
	Object.keys(zip.files).forEach(function(f) {
		var e = zip.files[f];
		if (!e.dir) {
			self._fe[f] = ltx.parse(zip.file(f).asText());
		}
	});
	// sharedStrings
	var el = this._fe['xl/sharedStrings.xml'];
	if (el) {
		var i = 0;
		el.children.forEach(function(si) {
			var t = si.getChildText('t');
			self._ss['' + i++] = t;
		});
	}
};

xlsx.prototype.loadFile = function(filename) {
	this.loadData(fs.readFileSync(filename, {}));
};

xlsx.prototype.writeFile = function(filename){
	var data = this.data();
	fs.writeFileSync(filename, data);
};

xlsx.prototype.sheetsCount = function() {
	return this._fe['xl/workbook.xml'].getChild('sheets').children.length;
};

xlsx.prototype.setSheetName = function(i, name) {
	this._fe['xl/workbook.xml'].getChild('sheets').children[i].attrs['name'] = name;
};

xlsx.prototype.getSheet = function(i) {
	var el = this._fe['xl/worksheets/sheet' + (i+1) + '.xml'];
	return new sheet(el, this._ss);
};

xlsx.prototype.copySheet = function(i, name){
	var n = this.sheetsCount() + 1;
	this._fe['[Content_Types].xml'].c('Override',{PartName:'/xl/worksheets/sheet'+n+'.xml', ContentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'});
	this._fe['xl/workbook.xml'].getChild('sheets').c('sheet', {name:name,sheetId:n,state:'visible','r:id':'sheetrId'+n});
	this._fe['xl/_rels/workbook.xml.rels'].c('Relationship', {Id:'sheetrId'+n, Type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet', Target:'worksheets/sheet'+n+'.xml'});
	this._fe['xl/worksheets/sheet'+n+'.xml'] = this._fe['xl/worksheets/sheet'+(i+1)+'.xml'].clone();
};

xlsx.prototype.data = function() {
	var zip = new JSZip();
	var self = this;
	Object.keys(self._fe).forEach(function(f) {
		zip.file(f, self._fe[f].root().toString());
	});
	return zip.generate({
		type : "nodebuffer"
	});
};
