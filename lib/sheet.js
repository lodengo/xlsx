var ltx = require('ltx');
var utils = require('./utils');

var sheet = module.exports = function sheet(el, ss) {
	this._el = el;
	this._ss = ss;
};

sheet.prototype.read = function(ref) {
	var splt = ref.split(":");
	if(1 == splt.length){
		return this._readCell(ref);
	}else{
		return this._readRange(ref);
	}	
};

sheet.prototype.write = function(cell, v, append) {	
	var self = this;
	var cr = utils.split_cell(cell);
	if (Array.isArray(v)) {
		var sr = cr[1]; // 开始行
		var sc = utils.decode_col(cr[0]); // 开始列 int
		if(append){
			self._writeRow(sr, sc, v);
		}else{
			v.forEach(function(r) {			
				if(Array.isArray(r)){
					var c = sc;
					r.forEach(function(cv) {
						self._writeCell(sr, utils.encode_col(c++), cv);
					});
					sr++;
				}else{
					self._writeCell(sr, utils.encode_col(sc++), r);
				}			
			});
		}		
	} else {
		self._writeCell(cr[1], cr[0], v);
	}
};

sheet.prototype.dimension = function(){
	for(var i = 0; i <= this._el.children.length; i++){
		if('dimension' == this._el.children[i].name){
			return this._el.children[i].attr('ref');
		}
	}
	return '';
};
/////////////////////////////////////////////////////////////
sheet.prototype._readRange = function(range){
	range = utils.decode_range(range);
	
	if(range.s.r == range.e.r){
		return this._readRow(range.s.r, range.s.c, range.e.c);
	}else if(range.s.c == range.e.c){
		return this._readCol(range.s.c, range.s.r, range.e.r);
	}else{
		var ret = [];
		for(var r = range.s.r; r <= range.e.r; r++){
			var row = this._readRow(r, range.s.c, range.e.c);
			ret.push(row);
		}
		return ret;
	}	
};

sheet.prototype._readRow = function(r, sc, ec){
	var row = this._el.getChild('sheetData').getChildByAttr('r', '' + r);
	sc = utils.decode_col(sc);
	ec = utils.decode_col(ec);
	var ret = [];
	for(var i = sc; i <= ec; i++){
		if(row){
			var cell = row.getChildByAttr('r', utils.encode_col(i)+r); 			
			ret.push(this._cellv(cell));
		}else{
			ret.push('');
		}		
	}
	return ret;	
};

sheet.prototype._readCol = function(c, sr, er){	
	var ret = [];
	var sd = this._el.getChild('sheetData');
	for(var i = sr; i <= er; i++){
		var row = sd.getChildByAttr('r', '' + i);
		var cell = row.getChildByAttr('r', c+i); 		
		ret.push(this._cellv(cell));
	}
	return ret;
};

sheet.prototype._readCell = function(cell) {
	var cr = utils.split_cell(cell);
	var r = this._el.getChild('sheetData').getChildByAttr('r', '' + cr[1]);
	var v = '';	
	if (r) {
		var c = r.getChildByAttr('r', cell);		
		v = this._cellv(c);
	}	
	return v;
};

sheet.prototype._cellv = function(c){
	if(!c) return '';
	var v = '';	
	var t = c.attrs['t'];
	var v = c.getChildText('v');
	switch (t) {
	case 'n':
		v = parseFloat(v);
		break;
	case 'b':
		v = v == 1 ? true : false;
		break;
	case 's':
		v = this._ss[v];
		break;
	}
	return v;
};

sheet.prototype._writeRow = function(sr, sc, rows){
	sr = parseInt(sr);
	var sd = this._el.getChild('sheetData');
	var srow = sd.c('row', {r : ''+sr});
	var self = this;
	rows.forEach(function(r, ri) {		
		if(Array.isArray(r)){
			var wr = ri == 0 ? srow : sd.c('row',{r:''+(sr+ri)});
			var c = sc;
			r.forEach(function(rcv) {
				var cv = self._v2cell(rcv);
				var cr = '' + utils.encode_col(c++) + (sr+ri);	
				
				wr.c('c', {
					r : cr,
					s : cv.s,
					t : cv.t
				}).c('v').t(cv.v);				
			});			
		}else{
			var cv = self._v2cell(r);
			var cr = '' + utils.encode_col(sc++) + sr;	
			srow.c('c', {
				r : cr,
				s : cv.s,
				t : cv.t
			}).c('v').t(cv.v);
		}			
	});
};

sheet.prototype._writeCell = function(ri, c, v) {
	var sd = this._el.getChild('sheetData');
	var r = sd.getChildByAttr('r', '' + ri);
	var cr = '' + c + ri;	
	var cv = this._v2cell(v);
	if (r) {
		var c = r.getChildByAttr('r', cr);
		if (c) {
			c.attr('t', cv.t);
			var rcv = c.getChild('v');
			rcv ? rcv.text(cv.v) : c.c('v').t(cv.v);
		} else {
			r.c('c', {
				r : cr,
				s : cv.s,
				t : cv.t
			}).c('v').t(cv.v);
		}
	} else {		
		sd.c('row', {
			r : ''+ri
		}).c('c', {
			r : cr,
			s : cv.s,
			t : cv.t
		}).c('v').t(cv.v);
	}
};

sheet.prototype._v2cell =function(v){
	var cv = {
		v:v,
		s: 0,
		t:'str'
	};
	var vt = typeof v;
	switch(vt){
		case 'string': break;
		case 'boolean': cv.t = 'b'; cv.v = v ? 1 : 0;  break;
		case 'number': cv.t = 'n'; break;
		default: break;
	}
	return cv;
};
