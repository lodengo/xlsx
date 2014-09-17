var ltx = require('ltx');

var sheet = module.exports = function sheet(el, ss) {
	this._el = el;
	this._ss = ss;
};

sheet.prototype.cell = function(cell){
	var cr = cell.replace(/(\$?[A-Z]*)(\$?\d*)/,"$1,$2").split(",");
	var r = this._el.getChild('sheetData').getChildByAttr('r', ''+cr[1]);	
	if(r){
		var c = r.getChildByAttr('r', cell);
		if(c){
			var t = c.attrs['t'];
			var v = c.getChildText('v');
			switch(t){
				case 'n': v = parseFloat(v); break;
				case 's': v = this._ss[v]; break;
			}
			return v;
		}		
	}
	return '';
};

sheet.prototype.write = function(cell, v){
	var cr = cell.replace(/(\$?[A-Z]*)(\$?\d*)/,"$1,$2").split(",");
	var r = this._el.getChild('sheetData').getChildByAttr('r', ''+cr[1]);	
	if(r){
		var c = r.getChildByAttr('r', cell);
		if(c){
			c.attr('t', 'str');
			c.getChild('v').text(v);
		}else{
			r.c('c', {r:cell, s:'0', t:'str'}).c('v').t(v);
		}
	}else{
		var sd = this._el.getChild('sheetData');
		sd.c('row', {r:cr[1]}).c('c', {r:cell, s:'0', t:'str'}).c('v').t(v);
	}
};
