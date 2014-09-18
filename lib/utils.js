function decode_col(colstr) {
	var c = colstr.replace(/^\$([A-Z])/, "$1"), d = 0, i = 0;
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

function split_cell(cstr) {
	return cstr.replace(/(\$?[A-Z]*)(\$?\d*)/, "$1,$2").split(",");
};

function decode_range(range) {
	var x = range.split(":").map(function(cell) {
		var splt = split_cell(cell);
		return {
			c : splt[0],
			r : parseInt(splt[1])
		};
	});
	return {
		s : x[0],
		e : x[x.length - 1]
	};
}

module.exports = {
	encode_col : encode_col,
	decode_col : decode_col,
	split_cell : split_cell,
	decode_range : decode_range
}