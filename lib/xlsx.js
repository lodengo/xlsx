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
			try{
				self._fe[f] = ltx.parse(zip.file(f).asText());
			}catch(err){
				self._fe[f] = zip.file(f).asText();
			}
			
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

xlsx.prototype.download = function(filename, res){
	res.writeHead(200, {
		'Content-Type' : 'application/vnd.openxmlformats',
		'Content-Disposition' : 'attachment; filename="'
				+ encodeURIComponent(filename) + '.xlsx"'
	});
	res.end(this.data());
};

xlsx.prototype.sheetsCount = function() {
	return this._fe['xl/workbook.xml'].getChild('sheets').children.length;
};

xlsx.prototype.setSheetName = function(i, name) {
	var self = this;
	var sheet = this._fe['xl/workbook.xml'].getChild('sheets').children[i];
	var oldName = sheet.attrs['name'];
	sheet.attrs['name'] = name;
	this._sheetCharts(i).forEach(function(chart){
		self._changeChartRef(chart, oldName, name);
	});
};

xlsx.prototype.setSheetVisible = function(i, visible){
	var state = visible ? 'visible' : 'hidden'; 
	this._fe['xl/workbook.xml'].getChild('sheets').children[i].attrs['state'] = state;
};

xlsx.prototype.drawingCount = function(){
	var n = 1;
	this._fe['[Content_Types].xml'].children.forEach(function(ct){
		if('application/vnd.openxmlformats-officedocument.drawing+xml' == ct.attr('ContentType')){
			n++;
		}
	});
	return n;
};

xlsx.prototype.chartCount = function(){
	var n = 1;
	this._fe['[Content_Types].xml'].children.forEach(function(ct){
		if('application/vnd.openxmlformats-officedocument.drawingml.chart+xml' == ct.attr('ContentType')){
			n++;
		}
	});
	return n;
};

xlsx.prototype.getSheet = function(i) {
	var el = this._fe['xl/worksheets/sheet' + (i+1) + '.xml'];
	return new sheet(el, this._ss);
};

xlsx.prototype.copySheet = function(i, name){
	var fromSheet = this._fe['xl/workbook.xml'].getChild('sheets').children[i].attrs['name'];
	var n = this.sheetsCount() + 1;
	this._fe['[Content_Types].xml'].c('Override',{PartName:'/xl/worksheets/sheet'+n+'.xml', ContentType:'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'});
	this._fe['xl/workbook.xml'].getChild('sheets').c('sheet', {name:name,sheetId:n,state:'visible','r:id':'sheetrId'+n});
	this._fe['xl/_rels/workbook.xml.rels'].c('Relationship', {Id:'sheetrId'+n, Type:'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet', Target:'worksheets/sheet'+n+'.xml'});
	this._fe['xl/worksheets/sheet'+n+'.xml'] = this._fe['xl/worksheets/sheet'+(i+1)+'.xml'].clone();
	this._copySheetRels(i+1, n, fromSheet, name);
	return this.getSheet(n-1);
};

xlsx.prototype.data = function() {
	var zip = new JSZip();
	var self = this;
	Object.keys(self._fe).forEach(function(f) {		
		typeof self._fe[f] == 'string' ? zip.file(f, self._fe[f]) : zip.file(f, self._fe[f].root().toString());		
	});
	return zip.generate({
		type : "nodebuffer"
	});
};
///////////////////////////////////////////////////////////////////////
xlsx.prototype._sheetCharts = function(i){
	var charts = [];
	var self = this;
	var rels = this._fe['xl/worksheets/_rels/sheet'+(i+1)+'.xml.rels'];
	if(rels){
		rels.children.forEach(function(rel){
			var relType = rel.attrs.Type;
			relType = relType.split('/');
			relType = relType[relType.length - 1];
			if('drawing' == relType){
				var drawing = rel.attrs.Target.split('/');
				drawing = drawing[drawing.length - 1];
				var drawingRels = self._fe['xl/drawings/_rels/'+drawing+'.rels'];
				if(drawingRels){
					drawingRels.children.forEach(function(drel){
						var drelType = drel.attrs.Type;
						drelType = drelType.split('/');
						drelType = drelType[drelType.length - 1];
						if('chart' == drelType){
							var chart = drel.attrs.Target.split('/');
							chart = chart[chart.length - 1];
							charts.push('xl/charts/'+chart);
						}
					});
				}
			}
		});
	}
	return charts;
};

xlsx.prototype._copySheetRels = function(i, n, fromSheet, toSheet){	
	var self = this;
	var srcRels = this._fe['xl/worksheets/_rels/sheet'+i+'.xml.rels'];	 
	if(srcRels){
		this._fe['xl/worksheets/_rels/sheet'+n+'.xml.rels'] = srcRels.clone();
		var destRels = this._fe['xl/worksheets/_rels/sheet'+n+'.xml.rels'];
		destRels.children.forEach(function(rel){
			var relType = rel.attrs.Type;
			relType = relType.split('/');
			relType = relType[relType.length - 1];
			if('drawing' == relType){
				var src = rel.attrs.Target.split('/');
				src = src[src.length - 1];
				var dest = 'drawing'+(self.drawingCount()+1)+'.xml';
				self._copyDrawing(src, dest, fromSheet, toSheet);
				rel.attr('Target', '../drawings/'+dest);
			}
		});
	}
};

xlsx.prototype._copyDrawing = function(src, dest, fromSheet, toSheet){		
	var srcRels = 'xl/drawings/_rels/'+src+'.rels';
	var destRels = 'xl/drawings/_rels/'+dest+'.rels';
	src = 'xl/drawings/' + src;
	dest = 'xl/drawings/' + dest;
	this._fe['[Content_Types].xml'].c('Override',{PartName:'/'+dest, ContentType:'application/vnd.openxmlformats-officedocument.drawing+xml'});
	this._fe[dest] = this._fe[src].clone();	
	this._copyDrawingRels(srcRels, destRels, fromSheet, toSheet);
};

xlsx.prototype._copyDrawingRels = function(src, dest, fromSheet, toSheet){	
	var self = this;
	var srcRels = this._fe[src];
	if(srcRels){
		this._fe[dest] = srcRels.clone();
		this._fe[dest].children.forEach(function(rel){
			var relType = rel.attrs.Type;
			relType = relType.split('/');
			relType = relType[relType.length - 1];
			if('chart' == relType){
				var src = rel.attrs.Target.split('/');
				src = src[src.length - 1];
				var dest = 'chart'+(self.chartCount()+1)+'.xml';
				self._copyChart(src, dest, fromSheet, toSheet);
				rel.attr('Target', '../charts/'+dest);
			}
		});
	}
};

xlsx.prototype._copyChart = function(src, dest, fromSheet, toSheet){	
	src = 'xl/charts/' + src;
	dest = 'xl/charts/' + dest;
	this._fe['[Content_Types].xml'].c('Override',{PartName:'/'+dest, ContentType:'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'});
	this._fe[dest] = this._fe[src].clone();	
	this._changeChartRef(dest, fromSheet, toSheet);
};

xlsx.prototype._changeChartRef = function(chart, fromSheet, toSheet){
	chart = this._fe[chart];
	chart.getChildrenByFilter(function(e){
		if('c:f' == e.name){
			var t = e.text();
			t = t.replace(fromSheet, toSheet);
			e.text(t);
		}
	}, true);
};


