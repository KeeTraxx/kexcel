var Util = require("./Util");
var Saveable = (function () {
    function Saveable(path) {
        this.path = path;
    }
    Saveable.prototype.save = function () {
        return Util.saveXML(this.xml, this.path);
    };
    return Saveable;
})();
module.exports = Saveable;
//# sourceMappingURL=Saveable.js.map