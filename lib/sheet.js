var ltx = require('ltx');

var sheet = module.exports = function sheet(el, ss) {
	this._el = el;
	this._ss = ss;
};

sheet.prototype.read = function(cell) {
	var cr = cell.replace(/(\$?[A-Z]*)(\$?\d*)/, "$1,$2").split(",");
	var r = this._el.getChild('sheetData').getChildByAttr('r', '' + cr[1]);
	var v = '';	
	if (r) {
		var c = r.getChildByAttr('r', cell);		
		if (c) {
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
		}
	}
	
	return v;
};

sheet.prototype.write = function(cell, v, fast) {
	fast = fast || false;
	var self = this;
	var cr = cell.replace(/(\$?[A-Z]*)(\$?\d*)/, "$1,$2").split(",");
	if (Array.isArray(v)) {
		var sr = cr[1]; // 开始行
		var sc = decode_col(cr[0]); // 开始列 int
		if(fast){
			self._writeRow(sr, sc, v);
		}else{
			v.forEach(function(r) {			
				if(Array.isArray(r)){
					var c = sc;
					r.forEach(function(cv) {
						self._writeCell(sr, encode_col(c++), cv);
					});
					sr++;
				}else{
					self._writeCell(sr, encode_col(sc++), r);
				}			
			});
		}		
	} else {
		self._writeCell(cr[1], cr[0], v);
	}
};

sheet.prototype._writeRow = function(sr, sc, rows){
	sr = parseInt(sr);
	var sd = this._el.getChild('sheetData');
	var srow = sd.c('row', {
		r : ''+sr
	});
	
	rows.forEach(function(r, ri) {		
		if(Array.isArray(r)){
			var wr = ri == 0 ? srow : sd.c('row',{r:''+(sr+ri)});
			var c = sc;
			r.forEach(function(rcv) {
				var cv = cellv(rcv);
				var cr = '' + encode_col(c++) + (sr+ri);	
				
				wr.c('c', {
					r : cr,
					s : cv.s,
					t : cv.t
				}).c('v').t(cv.v);				
			});			
		}else{
			var cv = cellv(r);
			var cr = '' + encode_col(sc++) + sr;	
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
	var cv = cellv(v);
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
//////////////////////////////////////////////////
function decode_col(colstr) {
	var c = colstr.replace(/^\$([A-Z])/,"$1"), d = 0, i = 0;
	for (; i !== c.length; ++i)
		d = 26 * d + c.charCodeAt(i) - 64;
	return d - 1;
};

function encode_col(col) {
	var s = "";
	for (++col; col; col = Math.floor((col - 1) / 26))
		s = String.fromCharCode(((col - 1) % 26) + 65) + s;
	return s;
};

function cellv(v){
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
