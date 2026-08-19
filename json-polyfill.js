/* json-polyfill.js ?? ?? JScript ?? JSON.parse?? parse?ES3 ???
 * ??????MSScriptControl ? JScript ? ES3 ??????? JSON ???
 * ???????? ASCII ???? BOM???? AddCode ?????????
 * ?? stringify ????????????? json2.js?MIT??
 */
(function () {
    if (typeof JSON === 'object') { return; }   /* ???? JSON ??? */
    JSON = {};
    var at = 0, ch = ' ', text = '';

    function error(m) { throw { name: 'SyntaxError', message: m, at: at }; }
    function next(c) {
        if (c && c !== ch) { error("Expected '" + c + "' instead of '" + ch + "'"); }
        ch = text.charAt(at); at += 1; return ch;
    }
    function white() { while (ch && ch <= ' ') { next(); } }
    function number() {
        var n = '';
        if (ch === '-') { n = '-'; next(); }
        while (ch >= '0' && ch <= '9') { n += ch; next(); }
        if (ch === '.') { n += '.'; next(); while (ch >= '0' && ch <= '9') { n += ch; next(); } }
        if (ch === 'e' || ch === 'E') {
            n += ch; next();
            if (ch === '+' || ch === '-') { n += ch; next(); }
            while (ch >= '0' && ch <= '9') { n += ch; next(); }
        }
        var v = +n;
        if (!isFinite(v)) { error("Bad number"); }
        return v;
    }
    function string() {
        var s = '', i, hex, uffff;
        if (ch === '"') {
            next('"');
            while (ch) {
                if (ch === '"') { next(); return s; }
                if (ch === '\\') {
                    next();
                    switch (ch) {
                        case 'n': s += '\n'; next(); break;
                        case 't': s += '\t'; next(); break;
                        case 'r': s += '\r'; next(); break;
                        case 'b': s += '\b'; next(); break;
                        case 'f': s += '\f'; next(); break;
                        case 'u':
                            uffff = 0;
                            for (i = 0; i < 4; i += 1) {
                                hex = parseInt(next(), 16);
                                if (!isFinite(hex)) { error("Bad \\u escape"); }
                                uffff = uffff * 16 + hex;
                            }
                            s += String.fromCharCode(uffff);
            next(); // pre-read next char, same as other escape branches
                            break;
                        case '"': case '/': case '\\':
                            s += ch; next(); break;
                        default: error("Bad escape");
                    }
                } else {
                    if (ch === '\n' || ch === '\r') { error("Bad control in string"); }
                    s += ch; next();
                }
            }
        }
        error("Bad string");
    }
    function word() {
        switch (ch) {
            case 't': next('t'); next('r'); next('u'); next('e'); return true;
            case 'f': next('f'); next('a'); next('l'); next('s'); next('e'); return false;
            case 'n': next('n'); next('u'); next('l'); next('l'); return null;
        }
        error("Unexpected '" + ch + "'");
    }
    function array() {
        var a = [];
        next('['); white();
        if (ch === ']') { next(']'); return a; }
        while (ch) {
            a.push(value()); white();
            if (ch === ']') { next(']'); return a; }
            next(','); white();
        }
        error("Bad array");
    }
    function object() {
        var o = {}, k;
        next('{'); white();
        if (ch === '}') { next('}'); return o; }
        while (ch) {
            k = string(); white(); next(':'); o[k] = value(); white();
            if (ch === '}') { next('}'); return o; }
            next(','); white();
        }
        error("Bad object");
    }
    function value() {
        white();
        switch (ch) {
            case '{': return object();
            case '[': return array();
            case '"': return string();
            case '-': return number();
            default: return (ch >= '0' && ch <= '9') ? number() : word();
        }
    }

    JSON.parse = function (source) {
        text = String(source);
        at = 0; ch = ' ';
        white();
        if (ch === '') { error("Empty JSON text"); }
        var result = value();
        white();
        if (ch) { error("Trailing characters"); }
        return result;
    };
})();
