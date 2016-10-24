module.exports = /******/
function(modules) {
    // webpackBootstrap
    /******/
    // The module cache
    /******/
    var installedModules = {};
    /******/
    /******/
    // The require function
    /******/
    function __webpack_require__(moduleId) {
        /******/
        /******/
        // Check if module is in cache
        /******/
        if (installedModules[moduleId]) /******/
        return installedModules[moduleId].exports;
        /******/
        /******/
        // Create a new module (and put it into the cache)
        /******/
        var module = installedModules[moduleId] = {
            /******/
            exports: {},
            /******/
            id: moduleId,
            /******/
            loaded: false
        };
        /******/
        /******/
        // Execute the module function
        /******/
        modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
        /******/
        /******/
        // Flag the module as loaded
        /******/
        module.loaded = true;
        /******/
        /******/
        // Return the exports of the module
        /******/
        return module.exports;
    }
    /******/
    /******/
    /******/
    // expose the modules object (__webpack_modules__)
    /******/
    __webpack_require__.m = modules;
    /******/
    /******/
    // expose the module cache
    /******/
    __webpack_require__.c = installedModules;
    /******/
    /******/
    // __webpack_public_path__
    /******/
    __webpack_require__.p = "";
    /******/
    /******/
    // Load entry module and return exports
    /******/
    return __webpack_require__(0);
}([ /* 0 */
/***/
function(module, exports, __webpack_require__) {
    /* WEBPACK VAR INJECTION */
    (function(__dirname) {
        "use strict";
        var fs = __webpack_require__(1);
        var Promise = __webpack_require__(2);
        var path = __webpack_require__(3);
        var _ = __webpack_require__(5);
        var rimraf = __webpack_require__(6);
        var unzip = __webpack_require__(7);
        var mkTempDir = Promise.promisify(__webpack_require__(8).mkdir);
        var fstream = __webpack_require__(9);
        var archiver = __webpack_require__(10);
        var Sheet_1 = __webpack_require__(11);
        var XMLFile_1 = __webpack_require__(15);
        var SharedStrings_1 = __webpack_require__(16);
        /**
	 * Main class for KExcel. Use .new() and .open(file | stream) to open an .xlsx file.
	 */
        var Workbook = function() {
            function Workbook(input) {
                /**
	         * Dictionary which holds pointers to files.
	         * @type {{}}
	         */
                this.files = {};
                /**
	         * The array of sheets in this workbook.
	         * @type {Array}
	         */
                this.sheets = [];
                if (typeof input == "string") {
                    this.filename = input;
                    this.source = fs.createReadStream(input);
                } else {
                    this.source = input;
                }
            }
            /**
	     * Creates an empty workbook
	     * @returns {Promise<Workbook>}
	     */
            Workbook.new = function() {
                return Workbook.open(path.join(__dirname, "..", "templates", "empty.xlsx"));
            };
            /**
	     * Opens an existing .xlsx file
	     * @param input
	     * @returns {Promise<Workbook>}
	     */
            Workbook.open = function(input) {
                var workbook = new Workbook(input);
                return workbook.init();
            };
            Workbook.prototype.init = function() {
                var _this = this;
                return this.extract().then(function() {
                    var p = Workbook.autoload.map(function(filepath) {
                        var xmlfile = new XMLFile_1.XMLFile(path.join(_this.tempDir, filepath));
                        _this.files[filepath] = xmlfile;
                        return xmlfile.load();
                    });
                    return Promise.all(p);
                }).then(function() {
                    _this.emptySheet = new XMLFile_1.XMLFile(path.join(__dirname, "..", "templates", "emptysheet.xml"));
                    return _this.emptySheet.load();
                }).then(function() {
                    return Promise.all([ _this.initSharedStrings(), _this.initSheets() ]);
                }).thenReturn(this);
            };
            Workbook.prototype.initSharedStrings = function() {
                this.sharedStrings = new SharedStrings_1.SharedStrings(path.join(this.tempDir, "xl", "sharedStrings.xml"), this);
                return this.sharedStrings.load();
            };
            Workbook.prototype.initSheets = function() {
                var _this = this;
                var wbxml = this.getXML("xl/workbook.xml");
                var relxml = this.getXML("xl/_rels/workbook.xml.rels");
                var p = _.map(wbxml.workbook.sheets[0].sheet, function(sheetXml) {
                    var r = _.find(relxml.Relationships.Relationship, function(rel) {
                        return rel.$.Id == sheetXml.$["r:id"];
                    });
                    var sheet = new Sheet_1.Sheet(_this, sheetXml, r);
                    _this.files[sheet.filename] = sheet;
                    _this.sheets.push(sheet);
                    return sheet.load();
                });
                return Promise.all(p);
            };
            Workbook.prototype.extract = function() {
                var _this = this;
                if (this.filename && !fs.existsSync(this.filename)) return Promise.reject(this.filename + " not found.");
                return mkTempDir("xlsx").then(function(tempDir) {
                    _this.tempDir = tempDir;
                    return new Promise(function(resolve, reject) {
                        var parser = unzip.Parse();
                        var writer = fstream.Writer(tempDir);
                        var outstream = _this.source.pipe(parser).pipe(writer);
                        outstream.on("close", function() {
                            resolve(_this);
                        });
                        parser.on("error", function(error) {
                            reject(error);
                        });
                    });
                });
            };
            /**
	     * Returns the xml object for the requested file
	     * @param filePath
	     * @returns {any}
	     */
            Workbook.prototype.getXML = function(filePath) {
                return this.files[filePath].xml;
            };
            /**
	     * Creates a new sheet in the workbook
	     * @param name Optionally set the name of the sheet
	     * @returns {Sheet}
	     */
            Workbook.prototype.createSheet = function(name) {
                var sheet = new Sheet_1.Sheet(this);
                sheet.create();
                this.sheets.push(sheet);
                this.files[sheet.filename] = sheet;
                if (name != undefined) {
                    sheet.setName(name);
                }
                return sheet;
            };
            Workbook.prototype.getSheet = function(input) {
                if (typeof input == "number") return this.sheets[input];
                return _.find(this.sheets, function(sheet) {
                    return sheet.getName() == input;
                });
            };
            Workbook.prototype.pipe = function(destination, options) {
                var _this = this;
                var archive = archiver("zip");
                Promise.all(_.map(this.files, function(file) {
                    return file.save();
                })).then(function() {
                    return _this.sharedStrings.save();
                }).then(function() {
                    archive.on("finish", function() {
                        // Async version somehow doesn't work?
                        /*rimraf(this.tempDir, function(error){
	                 console.log('errr');
	                 console.log(error);
	                 });*/
                        rimraf.sync(_this.tempDir);
                    });
                    archive.pipe(destination, options);
                    archive.bulk([ {
                        expand: true,
                        cwd: _this.tempDir,
                        src: [ "**", "_rels/.rels" ],
                        data: {
                            date: new Date()
                        }
                    } ]);
                    archive.finalize();
                });
                return archive;
            };
            /**
	     * These files are automatically loaded into files[]
	     * @type {string[]}
	     */
            Workbook.autoload = [ "xl/workbook.xml", "xl/_rels/workbook.xml.rels", "[Content_Types].xml" ];
            return Workbook;
        }();
        exports.Workbook = Workbook;
        module.exports = {
            Workbook: Workbook,
            open: Workbook.open,
            "new": Workbook.new
        };
    }).call(exports, "ts");
}, /* 1 */
/***/
function(module, exports) {
    module.exports = require("fs");
}, /* 2 */
/***/
function(module, exports) {
    module.exports = require("bluebird");
}, /* 3 */
/***/
function(module, exports, __webpack_require__) {
    /* WEBPACK VAR INJECTION */
    (function(process) {
        // Copyright Joyent, Inc. and other Node contributors.
        //
        // Permission is hereby granted, free of charge, to any person obtaining a
        // copy of this software and associated documentation files (the
        // "Software"), to deal in the Software without restriction, including
        // without limitation the rights to use, copy, modify, merge, publish,
        // distribute, sublicense, and/or sell copies of the Software, and to permit
        // persons to whom the Software is furnished to do so, subject to the
        // following conditions:
        //
        // The above copyright notice and this permission notice shall be included
        // in all copies or substantial portions of the Software.
        //
        // THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
        // OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
        // MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN
        // NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
        // DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
        // OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE
        // USE OR OTHER DEALINGS IN THE SOFTWARE.
        // resolves . and .. elements in a path array with directory names there
        // must be no slashes, empty elements, or device names (c:\) in the array
        // (so also no leading and trailing slashes - it does not distinguish
        // relative and absolute paths)
        function normalizeArray(parts, allowAboveRoot) {
            // if the path tries to go above the root, `up` ends up > 0
            var up = 0;
            for (var i = parts.length - 1; i >= 0; i--) {
                var last = parts[i];
                if (last === ".") {
                    parts.splice(i, 1);
                } else if (last === "..") {
                    parts.splice(i, 1);
                    up++;
                } else if (up) {
                    parts.splice(i, 1);
                    up--;
                }
            }
            // if the path is allowed to go above the root, restore leading ..s
            if (allowAboveRoot) {
                for (;up--; up) {
                    parts.unshift("..");
                }
            }
            return parts;
        }
        // Split a filename into [root, dir, basename, ext], unix version
        // 'root' is just a slash, or nothing.
        var splitPathRe = /^(\/?|)([\s\S]*?)((?:\.{1,2}|[^\/]+?|)(\.[^.\/]*|))(?:[\/]*)$/;
        var splitPath = function(filename) {
            return splitPathRe.exec(filename).slice(1);
        };
        // path.resolve([from ...], to)
        // posix version
        exports.resolve = function() {
            var resolvedPath = "", resolvedAbsolute = false;
            for (var i = arguments.length - 1; i >= -1 && !resolvedAbsolute; i--) {
                var path = i >= 0 ? arguments[i] : process.cwd();
                // Skip empty and invalid entries
                if (typeof path !== "string") {
                    throw new TypeError("Arguments to path.resolve must be strings");
                } else if (!path) {
                    continue;
                }
                resolvedPath = path + "/" + resolvedPath;
                resolvedAbsolute = path.charAt(0) === "/";
            }
            // At this point the path should be resolved to a full absolute path, but
            // handle relative paths to be safe (might happen when process.cwd() fails)
            // Normalize the path
            resolvedPath = normalizeArray(filter(resolvedPath.split("/"), function(p) {
                return !!p;
            }), !resolvedAbsolute).join("/");
            return (resolvedAbsolute ? "/" : "") + resolvedPath || ".";
        };
        // path.normalize(path)
        // posix version
        exports.normalize = function(path) {
            var isAbsolute = exports.isAbsolute(path), trailingSlash = substr(path, -1) === "/";
            // Normalize the path
            path = normalizeArray(filter(path.split("/"), function(p) {
                return !!p;
            }), !isAbsolute).join("/");
            if (!path && !isAbsolute) {
                path = ".";
            }
            if (path && trailingSlash) {
                path += "/";
            }
            return (isAbsolute ? "/" : "") + path;
        };
        // posix version
        exports.isAbsolute = function(path) {
            return path.charAt(0) === "/";
        };
        // posix version
        exports.join = function() {
            var paths = Array.prototype.slice.call(arguments, 0);
            return exports.normalize(filter(paths, function(p, index) {
                if (typeof p !== "string") {
                    throw new TypeError("Arguments to path.join must be strings");
                }
                return p;
            }).join("/"));
        };
        // path.relative(from, to)
        // posix version
        exports.relative = function(from, to) {
            from = exports.resolve(from).substr(1);
            to = exports.resolve(to).substr(1);
            function trim(arr) {
                var start = 0;
                for (;start < arr.length; start++) {
                    if (arr[start] !== "") break;
                }
                var end = arr.length - 1;
                for (;end >= 0; end--) {
                    if (arr[end] !== "") break;
                }
                if (start > end) return [];
                return arr.slice(start, end - start + 1);
            }
            var fromParts = trim(from.split("/"));
            var toParts = trim(to.split("/"));
            var length = Math.min(fromParts.length, toParts.length);
            var samePartsLength = length;
            for (var i = 0; i < length; i++) {
                if (fromParts[i] !== toParts[i]) {
                    samePartsLength = i;
                    break;
                }
            }
            var outputParts = [];
            for (var i = samePartsLength; i < fromParts.length; i++) {
                outputParts.push("..");
            }
            outputParts = outputParts.concat(toParts.slice(samePartsLength));
            return outputParts.join("/");
        };
        exports.sep = "/";
        exports.delimiter = ":";
        exports.dirname = function(path) {
            var result = splitPath(path), root = result[0], dir = result[1];
            if (!root && !dir) {
                // No dirname whatsoever
                return ".";
            }
            if (dir) {
                // It has a dirname, strip trailing slash
                dir = dir.substr(0, dir.length - 1);
            }
            return root + dir;
        };
        exports.basename = function(path, ext) {
            var f = splitPath(path)[2];
            // TODO: make this comparison case-insensitive on windows?
            if (ext && f.substr(-1 * ext.length) === ext) {
                f = f.substr(0, f.length - ext.length);
            }
            return f;
        };
        exports.extname = function(path) {
            return splitPath(path)[3];
        };
        function filter(xs, f) {
            if (xs.filter) return xs.filter(f);
            var res = [];
            for (var i = 0; i < xs.length; i++) {
                if (f(xs[i], i, xs)) res.push(xs[i]);
            }
            return res;
        }
        // String.prototype.substr - negative index don't work in IE8
        var substr = "ab".substr(-1) === "b" ? function(str, start, len) {
            return str.substr(start, len);
        } : function(str, start, len) {
            if (start < 0) start = str.length + start;
            return str.substr(start, len);
        };
    }).call(exports, __webpack_require__(4));
}, /* 4 */
/***/
function(module, exports) {
    // shim for using process in browser
    var process = module.exports = {};
    // cached from whatever global is present so that test runners that stub it
    // don't break things.  But we need to wrap it in a try catch in case it is
    // wrapped in strict mode code which doesn't define any globals.  It's inside a
    // function because try/catches deoptimize in certain engines.
    var cachedSetTimeout;
    var cachedClearTimeout;
    function defaultSetTimout() {
        throw new Error("setTimeout has not been defined");
    }
    function defaultClearTimeout() {
        throw new Error("clearTimeout has not been defined");
    }
    (function() {
        try {
            if (typeof setTimeout === "function") {
                cachedSetTimeout = setTimeout;
            } else {
                cachedSetTimeout = defaultSetTimout;
            }
        } catch (e) {
            cachedSetTimeout = defaultSetTimout;
        }
        try {
            if (typeof clearTimeout === "function") {
                cachedClearTimeout = clearTimeout;
            } else {
                cachedClearTimeout = defaultClearTimeout;
            }
        } catch (e) {
            cachedClearTimeout = defaultClearTimeout;
        }
    })();
    function runTimeout(fun) {
        if (cachedSetTimeout === setTimeout) {
            //normal enviroments in sane situations
            return setTimeout(fun, 0);
        }
        // if setTimeout wasn't available but was latter defined
        if ((cachedSetTimeout === defaultSetTimout || !cachedSetTimeout) && setTimeout) {
            cachedSetTimeout = setTimeout;
            return setTimeout(fun, 0);
        }
        try {
            // when when somebody has screwed with setTimeout but no I.E. maddness
            return cachedSetTimeout(fun, 0);
        } catch (e) {
            try {
                // When we are in I.E. but the script has been evaled so I.E. doesn't trust the global object when called normally
                return cachedSetTimeout.call(null, fun, 0);
            } catch (e) {
                // same as above but when it's a version of I.E. that must have the global object for 'this', hopfully our context correct otherwise it will throw a global error
                return cachedSetTimeout.call(this, fun, 0);
            }
        }
    }
    function runClearTimeout(marker) {
        if (cachedClearTimeout === clearTimeout) {
            //normal enviroments in sane situations
            return clearTimeout(marker);
        }
        // if clearTimeout wasn't available but was latter defined
        if ((cachedClearTimeout === defaultClearTimeout || !cachedClearTimeout) && clearTimeout) {
            cachedClearTimeout = clearTimeout;
            return clearTimeout(marker);
        }
        try {
            // when when somebody has screwed with setTimeout but no I.E. maddness
            return cachedClearTimeout(marker);
        } catch (e) {
            try {
                // When we are in I.E. but the script has been evaled so I.E. doesn't  trust the global object when called normally
                return cachedClearTimeout.call(null, marker);
            } catch (e) {
                // same as above but when it's a version of I.E. that must have the global object for 'this', hopfully our context correct otherwise it will throw a global error.
                // Some versions of I.E. have different rules for clearTimeout vs setTimeout
                return cachedClearTimeout.call(this, marker);
            }
        }
    }
    var queue = [];
    var draining = false;
    var currentQueue;
    var queueIndex = -1;
    function cleanUpNextTick() {
        if (!draining || !currentQueue) {
            return;
        }
        draining = false;
        if (currentQueue.length) {
            queue = currentQueue.concat(queue);
        } else {
            queueIndex = -1;
        }
        if (queue.length) {
            drainQueue();
        }
    }
    function drainQueue() {
        if (draining) {
            return;
        }
        var timeout = runTimeout(cleanUpNextTick);
        draining = true;
        var len = queue.length;
        while (len) {
            currentQueue = queue;
            queue = [];
            while (++queueIndex < len) {
                if (currentQueue) {
                    currentQueue[queueIndex].run();
                }
            }
            queueIndex = -1;
            len = queue.length;
        }
        currentQueue = null;
        draining = false;
        runClearTimeout(timeout);
    }
    process.nextTick = function(fun) {
        var args = new Array(arguments.length - 1);
        if (arguments.length > 1) {
            for (var i = 1; i < arguments.length; i++) {
                args[i - 1] = arguments[i];
            }
        }
        queue.push(new Item(fun, args));
        if (queue.length === 1 && !draining) {
            runTimeout(drainQueue);
        }
    };
    // v8 likes predictible objects
    function Item(fun, array) {
        this.fun = fun;
        this.array = array;
    }
    Item.prototype.run = function() {
        this.fun.apply(null, this.array);
    };
    process.title = "browser";
    process.browser = true;
    process.env = {};
    process.argv = [];
    process.version = "";
    // empty string to avoid regexp issues
    process.versions = {};
    function noop() {}
    process.on = noop;
    process.addListener = noop;
    process.once = noop;
    process.off = noop;
    process.removeListener = noop;
    process.removeAllListeners = noop;
    process.emit = noop;
    process.binding = function(name) {
        throw new Error("process.binding is not supported");
    };
    process.cwd = function() {
        return "/";
    };
    process.chdir = function(dir) {
        throw new Error("process.chdir is not supported");
    };
    process.umask = function() {
        return 0;
    };
}, /* 5 */
/***/
function(module, exports) {
    module.exports = require("lodash");
}, /* 6 */
/***/
function(module, exports) {
    module.exports = require("rimraf");
}, /* 7 */
/***/
function(module, exports) {
    module.exports = require("unzip");
}, /* 8 */
/***/
function(module, exports) {
    module.exports = require("temp");
}, /* 9 */
/***/
function(module, exports) {
    module.exports = require("fstream");
}, /* 10 */
/***/
function(module, exports) {
    module.exports = require("archiver");
}, /* 11 */
/***/
function(module, exports, __webpack_require__) {
    "use strict";
    var __extends = this && this.__extends || function(d, b) {
        for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
        function __() {
            this.constructor = d;
        }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
    var Promise = __webpack_require__(2);
    var _ = __webpack_require__(5);
    var Util = __webpack_require__(12);
    var path = __webpack_require__(3);
    var Saveable_1 = __webpack_require__(14);
    var Sheet = function(_super) {
        __extends(Sheet, _super);
        function Sheet(workbook, workbookXml, relationshipXml) {
            _super.call(this, null);
            this.workbook = workbook;
            this.workbookXml = workbookXml;
            this.relationshipXml = relationshipXml;
            if (this.workbookXml) {
                this.filename = this.relationshipXml.$.Target;
                this.path = path.join(this.workbook.tempDir, "xl", this.filename);
                this.id = this.relationshipXml.$.Id;
            }
        }
        Sheet.prototype.load = function() {
            var _this = this;
            return this.xml ? Promise.resolve(this.xml) : Util.loadXML(this.path).then(function(xml) {
                return _this.xml = xml;
            });
        };
        Sheet.prototype.getName = function() {
            return this.workbookXml.$.name;
        };
        Sheet.prototype.setName = function(name) {
            this.workbookXml.$.name = name;
        };
        Sheet.prototype.create = function() {
            this.addRelationship();
            this.addContentType();
            this.addToWorkbook();
            this.xml = _.cloneDeep(this.workbook.emptySheet.xml);
        };
        Sheet.prototype.copyFrom = function(sheet) {
            this.xml = _.cloneDeep(sheet.xml);
            // delete selections if any
            delete this.xml.worksheet.sheetViews;
        };
        Sheet.prototype.setCellValue = function(rownum_or_ref, colnum, cellvalue, copyCellStyle) {
            var cell = this.getCell(rownum_or_ref, colnum);
            var value = typeof colnum == "string" ? colnum : cellvalue;
            var from = typeof colnum == "string" ? cellvalue : copyCellStyle;
            if (cellvalue === undefined || cellvalue === null) {
                var matches = cell.$.r.match(Sheet.refRegex);
                var rownum = parseInt(matches[2]);
                // delete cell
                var row = this.getRowXML(rownum);
                row.c.splice(row.c.indexOf(cell), 1);
            } else {
                this.setValue(cell, value);
                if (from !== undefined) {
                    cell.$.s = this.getCell(from).$.s;
                }
            }
        };
        Sheet.prototype.setValue = function(cell, cellvalue) {
            if (typeof cellvalue == "number") {
                // number
                cell.v = [ cellvalue ];
                delete cell.f;
            } else if (cellvalue[0] == "=") {
                // function
                cell.f = [ cellvalue.substr(1).replace(/;/g, ",") ];
            } else {
                // assume string
                cell.v = [ this.workbook.sharedStrings.getIndex(cellvalue) ];
                cell.$.t = "s";
                // reset cell type
                delete cell.$.s;
            }
            return cell;
        };
        Sheet.prototype.getCellValue = function(r, colnum) {
            var cell = this.getCell(r, colnum);
            if (cell.$.t == "s") {
                // Sharedstring
                return this.workbook.sharedStrings.getString(cell.v[0]);
            } else if (cell.f && cell.v) {
                var value = cell.v[0].hasOwnProperty("_") ? cell.v[0]._ : cell.v[0];
                return value != "" ? value : undefined;
            } else if (cell.f) {
                return "=" + cell.f[0];
            } else {
                return cell.hasOwnProperty("v") ? cell.v[0] : undefined;
            }
        };
        Sheet.prototype.getCellFunction = function(r, colnum) {
            var cell = this.getCell(r, colnum);
            if (cell === undefined || cell === null || !cell.f) return undefined;
            var func = cell.f[0].hasOwnProperty("_") ? cell.f[0]._ : cell.f[0];
            return "=" + func;
        };
        Sheet.prototype.getCell = function(rownum_or_ref, colnum) {
            var rownum;
            var cellId;
            if (typeof rownum_or_ref == "string") {
                var matches = rownum_or_ref.match(Sheet.refRegex);
                rownum = parseInt(matches[2]);
                //colnum = Sheet.excelColumnToInt(matches[1]);
                cellId = rownum_or_ref;
            } else if (typeof rownum_or_ref == "number") {
                rownum = rownum_or_ref;
                //colnum = colnum;
                cellId = Sheet.intToExcelColumn(colnum) + rownum;
            } else {
                return rownum_or_ref;
            }
            var row = this.getRowXML(rownum);
            var cell = _.find(row.c, function(cell) {
                return cell.$.r == cellId;
            });
            if (cell === undefined) {
                cell = {
                    $: {
                        r: cellId
                    }
                };
                row.c = row.c || [];
                row.c.push(cell);
                row.c.sort(function(a, b) {
                    return Sheet.excelColumnToInt(a.$.r) - Sheet.excelColumnToInt(b.$.r);
                });
            }
            return cell;
        };
        Sheet.prototype.getRow = function(r) {
            var _this = this;
            var row = r;
            if (typeof r == "number") {
                row = this.getRowXML(r);
            }
            if (!row.c) return undefined;
            var result = [];
            row.c.forEach(function(cell) {
                result[Sheet.excelColumnToInt(cell.$.r) - 1] = _this.getCellValue(cell);
            });
            return result;
        };
        Sheet.prototype.setRow = function(r, values) {
            var _this = this;
            var row = r;
            if (typeof r == "number") {
                row = this.getRowXML(r);
            }
            var rownum = row.$.r;
            row.c = _.compact(values.map(function(value, index) {
                if (!value) return undefined;
                var cellId = Sheet.intToExcelColumn(index + 1) + rownum;
                return _this.setValue({
                    $: {
                        r: cellId
                    }
                }, value);
            }));
        };
        Sheet.prototype.appendRow = function(values) {
            var row = this.getRowXML(this.getLastRowNumber() + 1);
            this.setRow(row, values);
            return row.$.r;
        };
        Sheet.prototype.getLastRowNumber = function() {
            if (this.xml.worksheet.sheetData && this.xml.worksheet.sheetData[0].row) {
                // Remove empty rows
                this.xml.worksheet.sheetData[0].row = this.xml.worksheet.sheetData[0].row.filter(function(row) {
                    return row.c && row.c.length > 0;
                });
                return +_.last(this.xml.worksheet.sheetData[0].row).$.r || 0;
            } else {
                return 0;
            }
        };
        Sheet.prototype.getRowXML = function(rownum) {
            if (!this.xml.worksheet.sheetData) {
                this.xml.worksheet.sheetData = [];
            }
            if (!this.xml.worksheet.sheetData[0]) {
                this.xml.worksheet.sheetData[0] = {
                    row: []
                };
            }
            var rows = this.xml.worksheet.sheetData[0].row;
            var row = _.find(rows, function(r) {
                return r.$.r == rownum;
            });
            if (!row) {
                row = {
                    $: {
                        r: rownum
                    }
                };
                rows.push(row);
                rows.sort(function(row1, row2) {
                    return row1.$.r - row2.$.r;
                });
            }
            return row;
        };
        Sheet.prototype.toJSON = function() {
            var _this = this;
            var keys = this.getRow(1);
            var rows = this.xml.worksheet.sheetData[0].row.slice(1);
            return rows.map(function(row) {
                return _.zipObject(keys, _this.getRow(row));
            });
        };
        Sheet.prototype.addRelationship = function() {
            var relationships = this.workbook.getXML("xl/_rels/workbook.xml.rels");
            this.id = "rId" + (relationships.Relationships.Relationship.length + 1);
            this.filename = "worksheets/kexcel_" + this.id + ".xml";
            this.relationshipXml = {
                $: {
                    Id: this.id,
                    Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
                    Target: this.filename
                }
            };
            relationships.Relationships.Relationship.push(this.relationshipXml);
        };
        Sheet.prototype.addContentType = function() {
            var contentTypes = this.workbook.getXML("[Content_Types].xml");
            this.path = path.join(this.workbook.tempDir, "xl", "worksheets", "kexcel_" + this.id + ".xml");
            contentTypes.Types.Override.push({
                $: {
                    PartName: "/xl/" + this.filename,
                    ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
                }
            });
        };
        Sheet.prototype.addToWorkbook = function() {
            var wbxml = this.workbook.getXML("xl/workbook.xml");
            var sheets = wbxml.workbook.sheets[0].sheet;
            this.workbookXml = {
                $: {
                    name: "Sheet" + (sheets.length + 1),
                    sheetId: sheets.length + 1,
                    "r:id": this.id
                }
            };
            sheets.push(this.workbookXml);
        };
        Sheet.intToExcelColumn = function(col) {
            var result = "";
            var mod;
            while (col > 0) {
                mod = (col - 1) % 26;
                result = String.fromCharCode(65 + mod) + result;
                col = Math.floor((col - mod) / 26);
            }
            return result;
        };
        Sheet.excelColumnToInt = function(ref) {
            var number = 0;
            var pow = 1;
            for (var i = ref.length - 1; i >= 0; i--) {
                var c = ref.charCodeAt(i) - 64;
                if (c > 0 && c < 27) {
                    number += c * pow;
                    pow *= 26;
                }
            }
            return number;
        };
        Sheet.refRegex = /^([A-Z]+)(\d+)$/i;
        return Sheet;
    }(Saveable_1.Saveable);
    exports.Sheet = Sheet;
}, /* 12 */
/***/
function(module, exports, __webpack_require__) {
    "use strict";
    var fs = __webpack_require__(1);
    var xml2js = __webpack_require__(13);
    var Promise = __webpack_require__(2);
    var parseString = Promise.promisify(xml2js.parseString);
    var readFile = Promise.promisify(fs.readFile);
    var writeFile = Promise.promisify(fs.writeFile);
    var builder = new xml2js.Builder();
    function parseXML(input) {
        return parseString(input);
    }
    exports.parseXML = parseXML;
    function loadXML(path) {
        return readFile(path).then(function(buffer) {
            return parseString(buffer.toString());
        });
    }
    exports.loadXML = loadXML;
    function saveXML(xmlobj, path) {
        var contents = builder.buildObject(xmlobj);
        return writeFile(path, contents).thenReturn(path);
    }
    exports.saveXML = saveXML;
}, /* 13 */
/***/
function(module, exports) {
    module.exports = require("xml2js");
}, /* 14 */
/***/
function(module, exports, __webpack_require__) {
    "use strict";
    var Util = __webpack_require__(12);
    var Saveable = function() {
        function Saveable(path) {
            this.path = path;
        }
        Saveable.prototype.save = function() {
            return Util.saveXML(this.xml, this.path);
        };
        return Saveable;
    }();
    exports.Saveable = Saveable;
}, /* 15 */
/***/
function(module, exports, __webpack_require__) {
    "use strict";
    var __extends = this && this.__extends || function(d, b) {
        for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
        function __() {
            this.constructor = d;
        }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
    var Promise = __webpack_require__(2);
    var Util = __webpack_require__(12);
    var Saveable_1 = __webpack_require__(14);
    var XMLFile = function(_super) {
        __extends(XMLFile, _super);
        function XMLFile(path) {
            _super.call(this, path);
            this.path = path;
        }
        XMLFile.prototype.load = function() {
            var _this = this;
            return this.xml ? Promise.resolve(this.xml) : Util.loadXML(this.path).then(function(xml) {
                return _this.xml = xml;
            });
        };
        return XMLFile;
    }(Saveable_1.Saveable);
    exports.XMLFile = XMLFile;
}, /* 16 */
/***/
function(module, exports, __webpack_require__) {
    "use strict";
    var __extends = this && this.__extends || function(d, b) {
        for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
        function __() {
            this.constructor = d;
        }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
    var Promise = __webpack_require__(2);
    var _ = __webpack_require__(5);
    var Util = __webpack_require__(12);
    var Saveable_1 = __webpack_require__(14);
    var SharedStrings = function(_super) {
        __extends(SharedStrings, _super);
        function SharedStrings(path, workbook) {
            _super.call(this, path);
            this.path = path;
            this.workbook = workbook;
            this.cache = {};
        }
        SharedStrings.prototype.load = function() {
            var _this = this;
            return this.xml ? Promise.resolve(this.xml) : Util.loadXML(this.path).then(function(xml) {
                return _this.xml = xml;
            }).catch(function() {
                return Promise.all([ _this.addRelationship(), _this.addContentType() ]).then(function() {
                    return Util.parseXML('<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"><si/></sst>');
                }).then(function(xml) {
                    _this.xml = xml;
                });
            }).finally(function() {
                _this.xml.sst.si.forEach(function(si, index) {
                    if (si.t) {
                        _this.cache[si.t[0]] = index;
                    } else {}
                });
                return _this.xml;
            });
        };
        SharedStrings.prototype.getIndex = function(s) {
            return this.cache[s] || this.storeString(s);
        };
        SharedStrings.prototype.getString = function(n) {
            var sxml = this.xml.sst.si[n];
            if (!sxml) return undefined;
            return sxml.hasOwnProperty("t") ? sxml.t[0] : _.compact(sxml.r.map(function(d) {
                return d.t[0];
            })).join("");
        };
        SharedStrings.prototype.storeString = function(s) {
            var index = this.xml.sst.si.push({
                t: [ s ]
            }) - 1;
            this.cache[s] = index;
            return index;
        };
        SharedStrings.prototype.addRelationship = function() {
            var relationships = this.workbook.getXML("xl/_rels/workbook.xml.rels");
            relationships.Relationships.Relationship.push({
                $: {
                    Id: "rId1ss",
                    Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
                    Target: "sharedStrings.xml"
                }
            });
        };
        SharedStrings.prototype.addContentType = function() {
            var contentTypes = this.workbook.getXML("[Content_Types].xml");
            contentTypes.Types.Override.push({
                $: {
                    PartName: "/xl/sharedStrings.xml",
                    ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
                }
            });
        };
        return SharedStrings;
    }(Saveable_1.Saveable);
    exports.SharedStrings = SharedStrings;
} ]);
//# sourceMappingURL=kexcel.bundle.map