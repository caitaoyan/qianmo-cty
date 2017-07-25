/******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// identity function for calling harmony imports with the correct context
/******/ 	__webpack_require__.i = function(value) { return value; };
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, {
/******/ 				configurable: false,
/******/ 				enumerable: true,
/******/ 				get: getter
/******/ 			});
/******/ 		}
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "dist/";
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = 7);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ (function(module, exports) {

/**
 * 配置文件，主要是个各类属性默认值
 * 整张表命名为：WorkSpace
 * 表格上半部分叫做：Tool
 * 表格叫做：Sheet
 *
 */

/**
 * 整张表-WorkSpace的属性默认值
 *
 */
var WSConfig = {
    isInit: true
}

/**
 * 表头-Tool的属性默认值
 *
 */
var ToolConfig = {
    height: '12%',
    width: '100%',

    cmdCodeMap: {
        borderAll: '&#xe228', borderBottom: '&#xe229', borderClear: '&#xe22a',
        borderLeft: '&#xe22e', borderOuter: '&#xe22f', borderRight: '&#xe230',
        borderTop: "&#xe232", mergeCell: '&#xe252', splitCell: '&#xe0b6',
        copy: '&#xe14d', cut: '&#xe14e', paste: '&#xe14f', find: '&#xe881',
        font: '&#xe167', alignCenter: '&#xe234', alignLeft: '&#xe236', alignRight: '&#xe237',
        bold: '&#xe238', italic: '&#xe23f', fillColor: '&#xe23a', textColor: '&#xe23c',
        brush: '&#xe243', underline: '&#xe249', sigma: '&#xe24a', func: '&#xe940',
        preview: '&#xe8f4', sort: '&#xe164', wrapText: '&#xe25b', undo: '&#xe967', redo: '&#xe968',
        cleanText: '清除文字', cleanStyle: '清除格式', cleanAll: '清除文字和格式', multiLine: '多行输入',
        addCol: '插列', addRow: '插列', menu: '&#xe3c7', close: '&#xe5cd'
    },
    styleHtml: ['font', 'bold', 'italic', 'textColor', 'fillColor', 'align', 'border', 'brush', 'cleanText',
        'cleanStyle', 'cleanAll'],
    toolHtml: ['copy', 'paste', 'cut', 'sort', 'find', 'wrapText', 'func', 'sigma', 'mergeCell', 'splitCell',
        'multiLine', 'addCol','addRow'],
    dataHtml: ['数据源', '数据集', 'Init'],
    defaultHtml: ['undo', 'redo', 'preview'],
    borderHtml: ['borderAll', 'borderBottom', 'borderTop', 'borderClear', 'borderLeft', 'borderOuter', 'borderRight'],
    alignHtml: ['alignLeft', 'alignCenter', 'alignRight'],
    previewHtml: ['Edit', 'Down'],
    fontBorderOption: {
        borderLeft: 'left',
        borderTop: 'top',
        borderRight: 'right',
        borderBottom: 'bottom',
        borderClear: 'clear',
        borderOuter: 'outer',
        borderAll: 'all'
    },
    alignOption: {
        alignLeft: 'left',
        alignCenter: 'center',
        alignRight: 'right'
    },

    menu: ['样式', '工具', '数据']
}

/**
 * 侧边栏-sliderBar的常量
 *
 * */
var SlideBarConfig = {
    toggleOpenIcon: 'open',
    toggleCloseIcon: 'close',

    iconHtmlMap: {
		open: '&#xe3c7',
        close: '&#xe5cd'
    }
}

/**
 * 表单-Cell的属性默认值
 *
 */
var CellConfig = {
    height: 20,
    width: 100
}

/**
 * 表单-Sheet的属性默认值
 *
 */
var SheetConfig = {
    headWidth: 30,
    headHeight: 25,

    rowNum: 100,
    colNum: 25
}

var CellPropConfig = {
    id: 'cellProp'
}

var InputConfig = {
    id: 'input'
}

var ClipBoardConfig = {
    id: 'ta',
    value: ''
}
var keyboardTables = {

    specialKeysCommon: {
        8: "[backspace]", 9: "[tab]", 13: "[enter]", 25: "[tab]", 27: "[esc]", 33: "[pgup]", 34: "[pgdn]",
        35: "[end]", 36: "[home]", 37: "[aleft]", 38: "[aup]", 39: "[aright]", 40: "[adown]", 45: "[ins]",
        46: "[del]", 113: "[f2]"
    },

    specialKeysIE: {
        8: "[backspace]", 9: "[tab]", 13: "[enter]", 25: "[tab]", 27: "[esc]", 33: "[pgup]", 34: "[pgdn]",
        35: "[end]", 36: "[home]", 37: "[aleft]", 38: "[aup]", 39: "[aright]", 40: "[adown]", 45: "[ins]",
        46: "[del]", 113: "[f2]"
    },

    controlKeysIE: {
        67: "[ctrl-c]",
        83: "[ctrl-s]",
        86: "[ctrl-v]",
        88: "[ctrl-x]",
        90: "[ctrl-z]"
    },

    specialKeysOpera: {
        8: "[backspace]", 9: "[tab]", 13: "[enter]", 25: "[tab]", 27: "[esc]", 33: "[pgup]", 34: "[pgdn]",
        35: "[end]", 36: "[home]", 37: "[aleft]", 38: "[aup]", 39: "[aright]", 40: "[adown]",
        45: "[ins]", // issues with releases before 9.5 - same as "-" ("-" changed in 9.5)
        46: "[del]", // issues with releases before 9.5 - same as "." ("." changed in 9.5)
        113: "[f2]"
    },

    controlKeysOpera: {
        67: "[ctrl-c]",
        83: "[ctrl-s]",
        86: "[ctrl-v]",
        88: "[ctrl-x]",
        90: "[ctrl-z]"
    },

    specialKeysSafari: {
        8: "[backspace]", 9: "[tab]", 13: "[enter]", 25: "[tab]", 27: "[esc]", 63232: "[aup]", 63233: "[adown]",
        63234: "[aleft]", 63235: "[aright]", 63272: "[del]", 63273: "[home]", 63275: "[end]", 63276: "[pgup]",
        63277: "[pgdn]", 63237: "[f2]"
    },

    controlKeysSafari: {
        99: "[ctrl-c]",
        115: "[ctrl-s]",
        118: "[ctrl-v]",
        120: "[ctrl-x]",
        122: "[ctrl-z]"
    },

    ignoreKeysSafari: {
        63236: "[f1]", 63238: "[f3]", 63239: "[f4]", 63240: "[f5]", 63241: "[f6]", 63242: "[f7]",
        63243: "[f8]", 63244: "[f9]", 63245: "[f10]", 63246: "[f11]", 63247: "[f12]", 63289: "[numlock]"
    },

    specialKeysFirefox: {
        8: "[backspace]", 9: "[tab]", 13: "[enter]", 25: "[tab]", 27: "[esc]", 33: "[pgup]", 34: "[pgdn]",
        35: "[end]", 36: "[home]", 37: "[aleft]", 38: "[aup]", 39: "[aright]", 40: "[adown]", 45: "[ins]",
        46: "[del]", 113: "[f2]"
    },

    controlKeysFirefox: {
        99: "[ctrl-c]",
        115: "[ctrl-s]",
        118: "[ctrl-v]",
        120: "[ctrl-x]",
        122: "[ctrl-z]"
    },

    ignoreKeysFirefox: {
        16: "[shift]", 17: "[ctrl]", 18: "[alt]", 20: "[capslock]", 19: "[pause]", 44: "[printscreen]",
        91: "[windows]", 92: "[windows]", 112: "[f1]", 114: "[f3]", 115: "[f4]", 116: "[f5]",
        117: "[f6]", 118: "[f7]", 119: "[f8]", 120: "[f9]", 121: "[f10]", 122: "[f11]", 123: "[f12]",
        144: "[numlock]", 145: "[scrolllock]", 224: "[cmd]"
    }
}
/**
 * 公式类型
 * @type {[*]}
 */
// var function_classlist= ["all", "stat", "lookup", "datetime", "financial", "test", "math", "text"]
/**
 * 输出模块
 *
 */
module.exports.WSConfig = WSConfig
module.exports.ToolConfig = ToolConfig
module.exports.SheetConfig = SheetConfig
module.exports.CellConfig = CellConfig
module.exports.CellPropConfig = CellPropConfig
module.exports.InputConfig = InputConfig
module.exports.ClipBoardConfig = ClipBoardConfig
module.exports.SlideBarConfig = SlideBarConfig


/***/ }),
/* 1 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * 整体工作管理对象
 * @constructor
 */
var config = __webpack_require__(0)
var WSManager = function () {
}

WSManager.prototype.init = function (parentNode) {

	//实例化初始化Tool对象
	var ToolModule = __webpack_require__(12)
	var Tool = ToolModule.Tool
	var tool = new Tool()

	//实例化初始化UndoStack对象
	var UndoStackModule = __webpack_require__(15)
	var UndoStack = UndoStackModule.UndoStack
	var undoStack = new UndoStack()

	//实例化初始化Sheet对象
	var SheetModule = __webpack_require__(9)
	var Sheet = SheetModule.Sheet
	var sheet = new Sheet(undoStack)

	//实例化初始化SliderBar对象
	// var SliderBarModule = require('SliderBar')

	//实例化初始化Workspace对象
	var WorkspaceModule = __webpack_require__(17)
	var Workspace = WorkspaceModule.Workspace
	this.workspace = new Workspace(tool, sheet)

	//实例化初始化WSRender对象
	var WSRenderModule = __webpack_require__(16)
	var WSRender = WSRenderModule.WSRender
	var wsRender = new WSRender(this, parentNode)

	wsRender.init()

	if (config.WSConfig.isInit && parentNode.getAttribute("url")) {
		document.getElementById("Init").onclick()
		config.WSConfig.isInit = false
	}
}

module.exports.WSManager = WSManager

/***/ }),
/* 2 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * Created by Ian on 17/6/8.
 */
/**
 * 样式常量类
 *
 * 存储各种默认样式
 * 以及后续要添加的主题样式
 *
 * 注意，大多数宽与高，是成员属性，写在config里面
 */

var config = __webpack_require__(0)
var WSConfig = config.WSConfig
var SheetConfig = config.SheetConfig
var ToolConfig = config.ToolConfig

var WSStyle =
	'position: absolute;' +
	'left: 0;' +
	'top: 0;' +
	'width:100%;' +
	'height:100%;'

var SheetDivStyle =
	'position: relative;' +
	'top: 0;' +
	'left: 0;' +
	'width: 100%;' +
	'height: 87%'

var SheetTableStyle = 'user-select: none;'

var SheetTableDivStyle =
	'display: inline-block;' +
	'height:100%;' +
	'width: 100%;' +
	'overflow: auto;' +
	'border: none;' +
	'transition: all 0.5s ease'

var SliderBarStyle =
	'display: inline-block;' +
	'position: relative;' +
	'right :0;' +
	'top: 0;' +
	'width: 0;' +
	'height: 100%;' +
	'background-color: gray;' +
	'transition: all 0.5s ease;'

var ToggleDivCloseLeft = 'left: 95%;'
var ToggleDivOpenLeft = 'left: 76%;'

var ToggleDivStyle =
	'position: fixed;' +
	'bottom: 20px;' +
	'background-color: gray;' +
	'width: 40px;' +
	'height: 40px;' +
	'border-radius: 20px;' +
	'color: white;' +
	'line-height: 40px;' +
	'text-align: center;' +
	'z-index: 100;' +
	'user-select: none;' +
	'cursor: pointer;' +
	'transition: all 0.5s ease;'

var ToggleDivHoverStyle =
	'position: fixed;' +
	'bottom: 20px;' +
	'background-color: white;' +
	'width: 40px;' +
	'height: 40px;' +
	'border-radius: 20px;' +
	'color: gray;' +
	'line-height: 40px;' +
	'text-align: center;' +
	'z-index: 100;' +
	'user-select: none;' +
	'cursor: pointer;' +
	'transition: all 0.5s ease;'

var CellPropDivStyle =
	'position: absolute;' +
	'z-index: 100;' +
	'background-color: #fff;' +
	'border: 1px solid black;'

var InputStyle =
	'display: none;' +
	'position: absolute'

var ClipBoardStyle =
	'display:none;' +
	'position:absolute;' +
	'height:1px;' +
	'width:1px;' +
	'opacity:0;' +
	'filter:alpha(opacity=0);'

var SpanStyle =
	''

var ToolStyle = ''

var MenuDivStyle =
	'height:37%; ' +
	'width: 100%;' +
	'line-height: 30px;' +
	'min-height: 28px'

var MenuBoxStyle =
	'display: inline-block; ' +
	'height: 100%; ' +
	'letter-spacing: 7px; ' +
	'padding-left: 7px; ' +
	'cursor: pointer; ' +
	'transition: all 0.1s'+
	'background-color: transparent; ' +
	'border-bottom: none;'

var MenuBoxSelectedStyle =
	'display: inline-block; ' +
	'height: 100%; ' +
	'letter-spacing: 10px; ' +
	'padding-left: 10px; ' +
	'cursor: pointer; ' +
	'transition: all 0.1s; ' +
	'background-color: #ddd; ' +
	'border-bottom: solid 2px grey;'

var ButtonBoxStyle =
	'height: 63%;' +
	'width: 100%;' +
	'background-color: #ddd;' +
	'line-height: 50px' +
	'min-height: 50px;'

var ButtonDivStyle =
	'display: inline-block;' +
	'height: 50px;' +
	'cursor: pointer;' +
	'transition: all 0.1s;' +
	'color: #333;' +
	'line-height: 50px;' +
	'margin-left: 20px;' +
	'font-size: 20px'

var ButtonDivSelectedStyle =
	'display: inline-block;' +
	'height: 50px;' +
	'cursor: pointer;' +
	'transition: all 0.1s;' +
	'color: #fff;' +
	'line-height: 50px;' +
	'margin-left: 20px;' +
	'font-size: 20px'

var CellStyle = ''

var ColorDivStyle =
	'display: inline; ' +
	'margin-left: 10px; ' +
	'padding-left: 15px; ' +
	'cursor: pointer; ' +
	'border: none; ' +
	'width: 15px; ' +
	'height: 15px; ' +
	'background-color: #fff;'

var backgroundColorDivStyle =
	'display: inline; ' +
	'margin-left: 10px; ' +
	'padding-left: 15px; ' +
	'cursor: pointer; ' +
	'border: none; ' +
	'width: 15px; ' +
	'height: 15px; ' +
	'background-color: #fff;'

var ColorSelectDivStyle =
	'display: none;' +
	'position: absolute;' +
	'z-index: 100;' +
	'background-color: #fff;' +
	'border: 1px solid black;' +
	'width: 106px;'

module.exports.WSStyle = WSStyle

module.exports.SheetDivStyle = SheetDivStyle

module.exports.SheetTableStyle = SheetTableStyle

module.exports.SheetTableDivStyle = SheetTableDivStyle

module.exports.ToggleDivOpenLeft = ToggleDivOpenLeft
module.exports.ToggleDivCloseLeft = ToggleDivCloseLeft
module.exports.ToggleDivStyle =  ToggleDivStyle
module.exports.ToggleDivHoverStyle = ToggleDivHoverStyle

module.exports.CellPropDivStyle = CellPropDivStyle

module.exports.InputStyle = InputStyle

module.exports.ToolStyle = ToolStyle

module.exports.MeunDivStyle = MenuDivStyle

module.exports.MeunBoxStyle = MenuBoxStyle

module.exports.MeunBoxSelectedStyle = MenuBoxSelectedStyle

module.exports.CellStyle = CellStyle

module.exports.ClipBoardStyle = ClipBoardStyle

module.exports.ButtonBoxStyle = ButtonBoxStyle

module.exports.ButtonDivStyle = ButtonDivStyle

module.exports.ButtonDivSelectedStyle = ButtonDivSelectedStyle

module.exports.SpanStyle = SpanStyle

module.exports.ColorDivStyle = ColorDivStyle

module.exports.backgroundColorDivStyle = backgroundColorDivStyle

module.exports.ColorSelectDivStyle = ColorSelectDivStyle

module.exports.SliderBarStyle = SliderBarStyle

/***/ }),
/* 3 */
/***/ (function(module, exports, __webpack_require__) {

//鼠标和键盘事件

// var UtilModule=require('Util')
// var Util=UtilModule.Util
// var util=new Util()

var SheetRenderModule = __webpack_require__(5)
var SheetRender = SheetRenderModule.SheetRender

var SheetEventHandler = function (sheet) {

    this.sheet = sheet
    this.sheetRender = new SheetRender(sheet)
}
/**
 * 鼠标按下事件
 * @param element
 */
SheetEventHandler.prototype.mouseDown = function (elementid) {
    //var cellPropDiv=document.getElementById("cellProp")
    //cellPropDiv.style.display='none'
    //cellPropDiv.innerHTML=''
    if (!this.sheet.isEditing) {
        //if (!this.sheet.firstCellid) {
        this.sheet.firstCellid = elementid
        this.sheet.lastCellid = this.sheet.firstCellid
    }

        this.setCellBackgroundColor('#fff')
        if (elementid) {
            //setCellProp(element,this.sheet.cells[element.id])
            this.sheet.firstCellid = elementid
        }
    document.getElementById('dragBar').style.display = 'none'
    this.sheet.isMouseDown = true
}
SheetEventHandler.prototype.dragBarMouseDown = function () {
    this.setCellBackgroundColor('#fff')
    this.sheet.firstCellid = this.sheet.lastCellid
    this.sheet.isDraging = true
}
/**
 * 鼠标移动事件
 * @param element
 */
SheetEventHandler.prototype.mouseMove = function (elementid) {
    this.setCellBackgroundColor('#fff')
    if (elementid) {
        //var cellPropDiv=document.getElementById("cellProp")
        //cellPropDiv.style.display='none'
        //cellPropDiv.innerHTML=''
        this.sheet.lastCellid = elementid
        this.sheet.editCells = this.sheet.getColAndRow()
        if(this.sheet.isDraging) this.setCellBackgroundColor('#f69')
        else this.setCellBackgroundColor('#69f')
    }
}
/**
 * 鼠标抬起事件
 * @param element
 */
SheetEventHandler.prototype.mouseUp = function (element) {
    if(this.sheet.isDraging){
        if(!this.sheet.cells[this.sheet.firstCellid]){
            this.sheet.setAttr('content','')
        }else{
            this.sheet.setAttr('content',this.sheet.cells[this.sheet.firstCellid].content)
        }

        this.sheet.render()
    }
    if (this.sheet.isMouseDown || this.sheet.isDraging) {
        if (element.id) {
            var dragBar = document.getElementById('dragBar')
            dragBar.style.left = getLeft(element) +  element.offsetWidth -8+ 'px'
            dragBar.style.top = getTop(element) +element.offsetHeight-8+ 'px'
            dragBar.style.display = 'block'
            this.sheet.lastCellid = element.id
            this.sheet.editCells = this.sheet.getColAndRow()
            this.setCellBackgroundColor('#69f')

        }
    }
    this.sheet.isMouseDown = false
    this.sheet.isDraging = false
}
SheetEventHandler.prototype.dragBarMouseUp = function () {
    this.setCellBackgroundColor('#69f')
    this.sheet.isMouseDown = false
    this.sheet.isDraging = false
}
/**
 * 鼠标双击事件
 * @param element
 */
SheetEventHandler.prototype.dblclick = function (element) {
    if (element.id) {
        var input = document.getElementById('input')
        input.style.display = 'block'
        input.style.backgroundColor = '#efe'
        input.style.left = getLeft(element) + 'px'
        input.style.top = getTop(element) + 'px'
        input.style.height = element.offsetHeight + 'px'
        input.style.width = element.offsetWidth + 'px'
        if (this.sheet.cells[element.id]) {
            input.value = this.sheet.cells[element.id].content
        }
        else {
            input.value = ''
        }
        input.focus()
        this.sheet.isEditing = true
        this.sheet.isMouseDown = false
    }
}
SheetEventHandler.prototype.inputBlur = function () {
    this.sheet.isEditing = false
    var input = document.getElementById('input')
    input.blur()
    input.style.display = 'none'
    //var ele = this.sheet.lastCell
    if (input.value) {
        this.sheet.setAttr('content', input.value)
        this.sheet.render()
        // var sp = document.getElementById('sp')
        // sp.value = input.value
        // sp.style.font = ele.firstChild.style.font
        // sp.style.fontSize = ele.firstChild.style.fontSize
        // this.sheet.cells[ele.id].viewWidth = sp.offsetWidth
        // if(this.sheet.cells[ele.id].autoLF){
        //     ele.width = this.sheet.cells[ele.id].width
        // }else{
        //     ele.width = this.sheet.cells[ele.id].viewWidth
        // }
    }
    this.sheet.isEditing = false
}
SheetEventHandler.prototype.inputFocus = function () {
    this.sheet.isEditing = true
    var input = document.getElementById('input')
    input.focus()
}
// //鼠标失去焦点事件
// SheetEventHandler.prototype.mouseBlur=function(element){
//     element.style.color='transparent'
//     element.style.background='transparent'
//     if(config.WSConfig.isKeyDown){
//         var value=element.value
//         if(element.nextSibling.innerHTML==''&&value!=''){
//             var command = 'set '+ element.id + ' ' + value
//             var type = 'set'
//             this.sheet.UndoStack.pushCommand(command,type)
//         }
//         element.nextSibling.innerHTML=value
//     }
//     config.WSConfig.isKeyDown=true
//     element.parentNode.style.backgroundColor = 'transparent'
// }
/**
 * 设置单元格背景颜色
 * @param backgroundColor
 */
SheetEventHandler.prototype.setCellBackgroundColor = function (backgroundColor) {
    var cells = this.sheet.editCells
    if (cells != null) {
        for (var i = cells.firstCellCol; i <= cells.lastCellCol; i++) {
            for (var j = cells.firstCellRow; j <= cells.lastCellRow; j++) {
                var ele = document.getElementById(String.fromCharCode(i) + '_' + j)
                if (ele) {
                    if (backgroundColor == '#fff'
                        && this.sheet.cells[ele.id] != undefined
                        && this.sheet.cells[ele.id].backgroundColor != undefined) {
                        ele.style.backgroundColor = this.sheet.cells[ele.id].backgroundColor
                    } else {
                        ele.style.backgroundColor = backgroundColor
                    }
                }
            }
        }
    }
}

/**
 * 键盘按下事件
 * @param element
 * @param event
 */
SheetEventHandler.prototype.keyDown = function (event) {
    if (!this.sheet.isMultiLineEditing) {
        var col = this.sheet.editCells.lastCellCol//element.id.split('_')[0].charCodeAt(0)
        var row = this.sheet.editCells.lastCellRow//parseInt(element.id.split('_')[1])
        //console.log(event.which)
        switch (event.which) {
            case 37://左键
                do {
                    col = col - 1
                } while (this.sheet.cells[String.fromCharCode(col) + '_' + row] && !this.sheet.cells[String.fromCharCode(col) + '_' + row].show)
                if (!this.sheet.isEditing && col > 64) {
                    var cellid = String.fromCharCode(col) + '_' + row

                    this.mouseDown(cellid)
                    this.mouseUp(document.getElementById(cellid))

                }
                return false
            case 38://上键
                do {
                    row = row - 1
                } while (this.sheet.cells[String.fromCharCode(col) + '_' + row] && !this.sheet.cells[String.fromCharCode(col) + '_' + row].show)
                var cellid = String.fromCharCode(col) + '_' + row

                if (this.sheet.isEditing) {
                    this.sheet.isEditing = false
                    document.getElementById('input').blur()
                }
                if(row > 0){
                    this.mouseDown(cellid)
                    this.mouseUp(document.getElementById(cellid))
                }
                return false
            case 39://右键
                do {
                    col = col + 1
                } while (this.sheet.cells[String.fromCharCode(col) + '_' + row] && !this.sheet.cells[String.fromCharCode(col) + '_' + row].show)
                if (!this.sheet.isEditing) {
                    var cellid = String.fromCharCode(col) + '_' + row
                    if (document.getElementById(cellid)) {
                        this.mouseDown(cellid)
                        this.mouseUp(document.getElementById(cellid))
                    }
                }
                return false
            case 40://下键
                do {
                    row = row + 1
                } while (this.sheet.cells[String.fromCharCode(col) + '_' + row] && !this.sheet.cells[String.fromCharCode(col) + '_' + row].show)
                var cellid = String.fromCharCode(col) + '_' + row
                if (document.getElementById(cellid)) {
                    if (this.sheet.isEditing) {
                        this.sheet.isEditing = false
                        document.getElementById('input').blur()
                    }
                    this.mouseDown(cellid)
                    this.mouseUp(document.getElementById(cellid))
                }
                return false
            case 17://ctrl
                break
            case 18://alt
                break
            case 67://ctrl+c
                var ta = document.getElementById('ta')
                var text = ''
                //var firstCell=this.sheet.firstCell
                //var lastCell=this.sheet.lastCell
                var cells = this.sheet.getColAndRow()
                if (cells != null) {
                    var rowCount = 0
                    for (var i = cells.firstCellRow; i <= cells.lastCellRow; i++) {
                        var colCount = 0
                        for (var j = cells.firstCellCol; j <= cells.lastCellCol; j++) {
                            var eleOld = document.getElementById(
                                String.fromCharCode(j)
                                + '_' + i)
                            text += eleOld.firstChild.innerHTML + '\t'
                            colCount++
                        }
                        text = text.substr(0, text.length - 1)
                        text += '\n'
                        rowCount++
                    }
                    text = text.substr(0, text.length - 1)
                }
                ta.value = text
                ta.style.display = 'block'
                ta.focus()
                ta.select()

                window.setTimeout(function () {
                    ta.style.display = 'none'
                    var cell = document.getElementById(col + '_' + row)
                    //cell.focus()
                }, 200)

                break
            case 86://ctrl+v
                var ta = document.getElementById('ta')
                ta.style.display = 'block'
                ta.focus()
                ta.select()
                var sheet = this.sheet
                window.setTimeout(function () {
                    ta.blur()
                    var v = ta.value
                    ta.style.display = 'none'
                    v = v.split('\n')
                    for (var i = 0; i < v.length; i++) {
                        var vv = v[i].split('\t')
                        for (var j = 0; j < vv.length; j++) {
                            var cvCol = col + j
                            var cvRow = row + i
                            var cells = {}
                            cells.firstCellCol = cvCol
                            cells.firstCellRow = cvRow
                            cells.lastCellCol = cvCol
                            cells.lastCellRow = cvRow
                            sheet.setAttr('content', vv[j], cells)
                        }
                    }
                    var cell = document.getElementById(String.fromCharCode(col) + '_' + row)
                    cell.focus()
                    sheet.render()
                }, 200)

                this.sheet.isKeyDown = false
                break
            case 88://ctrl+x
                var ta = document.getElementById('ta')
                var text = ''
                //var firstCell=this.sheet.firstCell
                //var lastCell=this.sheet.lastCell
                var cells = this.sheet.getColAndRow()
                if (cells != null) {
                    var rowCount = 0
                    for (var i = cells.firstCellRow; i <= cells.lastCellRow; i++) {
                        var colCount = 0
                        for (var j = cells.firstCellCol; j <= cells.lastCellCol; j++) {
                            var eleOld = document.getElementById(
                                String.fromCharCode(j)
                                + '_' + i)
                            text += eleOld.firstChild.innerHTML + '\t'
                            colCount++
                        }
                        text = text.substr(0, text.length - 1)
                        text += '\n'
                        rowCount++
                    }
                    text = text.substr(0, text.length - 1)
                }
                ta.value = text
                ta.style.display = 'block'
                ta.focus()
                ta.select()

                window.setTimeout(function () {
                    ta.style.display = 'none'
                    var cell = document.getElementById(String.fromCharCode(col) + '_' + row)
                    cell.focus()
                }, 200)

                this.sheet.setAttr('content', '')
                this.sheet.render()
                break
            case 8://backspace
                this.sheet.setAttr('content', '')
                this.sheet.render()
                break
            case 46://delete
                var sheet = this.sheet
                sheet.cellAttrs.forEach(function (attr) {
                    sheet.setAttr(attr[0], attr[1])
                })
                sheet.render()
                break
            default:
                if (!this.sheet.isEditing) {
                    if (this.sheet.lastCellid) this.dblclick(document.getElementById(this.sheet.lastCellid))
                }
        }
    }
}

SheetEventHandler.prototype.multiLineBlur = function (text) {
    this.sheet.setAttr('content', text)
    this.sheet.setAttr('wrapText', true)
    this.sheet.render()
    this.sheet.isMultiLineEditing = false
}
// SheetEventHandler.prototype.setCellProp=function(element,e){
//     if(e){
//         var cellPropDiv=document.getElementById("cellProp")
//         cellPropDiv.style.display='block'
//         var left=element.offsetLeft
//         var top=element.offsetTop
//         cellPropDiv.style.top=top+element.offsetHeight+26+'px'
//         cellPropDiv.style.left=left+element.offsetWidth+9+'px'
//         var cellPropTable=document.createElement('table')
//         cellPropDiv.appendChild(cellPropTable)
//         var tr=document.createElement('tr')
//         cellPropTable.appendChild(tr)
//         var th=document.createElement('th')
//         th.innerHTML='属性名称'
//         tr.appendChild(th)
//         th=document.createElement('th')
//         th.innerHTML='属性值'
//         tr.appendChild(th)
//         for (var key in e) {
//             if(e[key]!=''){
//                 var tr=document.createElement('tr')
//                 cellPropTable.appendChild(tr)
//                 var td=document.createElement('td')
//                 td.innerHTML=key
//                 tr.appendChild(td)
//                 var td=document.createElement('td')
//                 td.innerHTML=e[key]
//                 tr.appendChild(td)
//             }
//         }
//         this.addProp(element,cellPropTable)
//     }
// }
// SheetEventHandler.prototype.addProp=function(ele,cellPropTable){
//     var SheetEventHandler=this
//     var tr=document.createElement('tr')
//     cellPropTable.appendChild(tr)
//     var td=document.createElement('td')
//     var inputTd1=document.createElement('input')
//     inputTd1.setAttribute('list','prop')
//     td.appendChild(inputTd1)
//     var datalist=document.createElement('datalist')
//     datalist.id='prop'
//     td.appendChild(datalist)
//     var option=document.createElement('option')
//     option.value='foregroundColor'
//     datalist.appendChild(option)
//     option=document.createElement('option')
//     option.value='bold'
//     datalist.appendChild(option)
//     option=document.createElement('option')
//     option.value='backgroundColor'
//     datalist.appendChild(option)
//     option=document.createElement('option')
//     option.value='italic'
//     datalist.appendChild(option)
//     option=document.createElement('option')
//     option.value='fontSize'
//     datalist.appendChild(option)
//     tr.appendChild(td)
//     var td=document.createElement('td')
//     var inputTd2=document.createElement('input')
//     td.appendChild(inputTd2)
//     tr.appendChild(td)
//     inputTd1.onclick=function(){
//         SheetEventHandler.firstCell=this
//     }
//     inputTd2.onclick=function(){
//         SheetEventHandler.firstCell=this
//     }
//     inputTd1.onblur=function(){
//         SheetEventHandler.firstCell=ele
//         var input2=this.parentNode.nextSibling.firstChild
//         if(inputTd1.value!=null &&inputTd1.value !=''
//             &&input2.value!=null &&input2.value !=''){
//             cellRender.renderCell(ele.id,inputTd1.value,input2.value)
//             if(!this.parentNode.parentNode.nextSibling){
//                 SheetEventHandler.addProp(ele,cellPropTable)
//             }
//         }
//
//     }
//     inputTd2.onblur=function(){
//         SheetEventHandler.firstCell=ele
//         var input1=this.parentNode.previousSibling.firstChild
//         if(input1.value!=null &&input1.value !=''
//             &&inputTd2.value!=null &&inputTd2.value !=''){
//             cellRender.renderCell(ele.id,input1.value,inputTd2.value)
//             if(!this.parentNode.parentNode.nextSibling){
//                 SheetEventHandler.addProp(ele,cellPropTable)
//             }
//         }
//     }
// }

//获取元素的纵坐标
function getTop(e) {
    var offset = e.offsetTop;
    if (e.offsetParent != null) offset += getTop(e.offsetParent);
    return offset;
}

//获取元素的横坐标
function getLeft(e) {
    var offset = e.offsetLeft
    if (e.offsetParent != null) offset += getLeft(e.offsetParent)
    return offset
}

module.exports.SheetEventHandler = SheetEventHandler

/***/ }),
/* 4 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * 单元格对象
 */

var configModule = __webpack_require__(0)

var cellConfig = configModule.CellConfig

var Cell = function (coord) {

	//成员属性
	//设置默认值
	//this.autoLF = false

	//this.viewWidth = 0

	this.height = cellConfig.height
	this.width = cellConfig.width

	this.foregroundColor = ''
	this.backgroundColor = ''

	this.font = ''
	this.fontSize = ''
	// // this.fontWeight = false
	// // this.fontStyle = false

	this.topFrame = false
	this.bottomFrame = false
	this.leftFrame = false
	this.rightFrame = false

	this.alignment = 'left'
	this.wrapText = false

	this.coord = coord

	this.content = ''
	this.value = null

	this.show = true
	this.area = [1,1]

	this.bold = false
	this.italic = false

	//this.colSpan = 1
	//this.rowSpan = 1
}

module.exports.Cell = Cell

/***/ }),
/* 5 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * 渲染表格
 */

var config = __webpack_require__(0)
var style = __webpack_require__(2)
var SliderBarRender = __webpack_require__(11).SliderBarRender
// config.CellConfig
var SheetRender = function (sheet) {
	this.sheet=sheet
}

SheetRender.prototype.init = function (sheetDiv) {

	sheetDiv.style = style.SheetDivStyle

	var sheetTableDiv = document.createElement('div')
	sheetTableDiv.style = style.SheetTableDivStyle

	var sheetTable = document.createElement('table')
    sheetTable.style = style.SheetTableStyle
	sheetTable.cellSpacing = '10px'

	var sliderBarDiv = document.createElement('div')
	var sliderBarRender = new SliderBarRender(this.sheet, sheetDiv)
	sliderBarRender.init(sliderBarDiv, sheetTableDiv)

    sheetTableDiv.appendChild(sheetTable)

	sheetDiv.appendChild(sheetTableDiv)
	sheetDiv.appendChild(sliderBarDiv)

    // var cellPropDiv = document.createElement("div")
    // cellPropDiv.id = config.CellPropConfig.id
	//

    // cellPropDiv.style = style.CellPropDivStyle

    // sheetDiv.appendChild(cellPropDiv)

	//表格的具体渲染
	renderSheet(this.sheet, sheetTable)
}

SheetRender.prototype.reRender = function() {
	while(this.sheet.undostackCache !== []){

	}
}


/**
 * 构建表格的数据模型
 * 一个存放表头的二维数组
 * @param rowNum
 * @param colNum
 */
function createHeader(rowNum, colNum) {
	var gridHeader =  []
	for(var a=0;a<rowNum+1;a++){
        gridHeader[a] = []

		for(var b=0;b<colNum+1;b++){
            gridHeader[a][b] = null
		}
	}
	var startString = "A"

	for(var i=1; i<colNum+1; i++){
        gridHeader[0][i] = String.fromCharCode(startString.charCodeAt(0) + i - 1)
	}

	for(var j=1; j<rowNum+1; j++){
        gridHeader[j][0] = j
	}

	return gridHeader
}
/**
 * 表格渲染
 * @param sheet
 * @param sheetTable
 */
function renderSheet(sheet, sheetTable) {
	var gridHeader = createHeader(sheet.rowNum,  sheet.colNum)

	var rowNum = gridHeader.length
	var colNum = gridHeader[0].length

	var rowHead = document.createElement('tr')
	sheetTable.appendChild(rowHead)
	//渲染第一行表头
	for(var i=0;i<colNum;i++){
		var rowHeadTH  = document.createElement('th')
		// rowHeadTH.style.height = '20px'
		rowHeadTH.id = gridHeader[0][i]+'_0'
		rowHead.appendChild(rowHeadTH)

		var rowHeadDiv = document.createElement('div')
		rowHeadTH.appendChild(rowHeadDiv)
		// 第一行第一列特殊处理
		if(i === 0){
            //rowHeadTH.style.width = '29px'
		}else{
			rowHeadDiv.innerHTML = gridHeader[0][i]
            rowHeadDiv.style.width = config.CellConfig.width + 'px'
		}
	}

    var SheetEventBinderModule = __webpack_require__(10)
    var SheetEventBinder=SheetEventBinderModule.SheetEventBinder
    var sheetEventBinder=new SheetEventBinder(sheet)

	//渲染除第一行外单元格
	for(var j = 1;j<rowNum;j++){
		var rowTR = document.createElement('tr')
		sheetTable.appendChild(rowTR)

		for(var k=0; k<colNum;k++){
			//表格第一列特殊处理
			if(k === 0){
				var rowTH = document.createElement('th')
				rowTH.id = '@_' + gridHeader[j][0]
				rowTR.appendChild(rowTH)
				var rowDiv = document.createElement('div')
				rowTH.appendChild(rowDiv)
                rowDiv.innerHTML = gridHeader[j][k]
			}else {
				var rowTD = document.createElement('td')
                rowTD.id = gridHeader[0][k]+"_"+j
                rowTR.appendChild(rowTD)
                if(!config.WSConfig.isPreview){
                    sheetEventBinder.initRowTD(rowTD)
                }
				var rowDiv = document.createElement('div')
				rowTD.className = "noWrap"
                rowTD.appendChild(rowDiv)
			}
		}
	}
	var input = document.createElement('input')
	input.style = style.InputStyle
	input.id = config.InputConfig.id
    if(!config.WSConfig.isPreview){
        sheetEventBinder.initInput(input)
    }

	sheetTable.appendChild(input)

    var clipBoard = document.createElement('textarea') // used for ctrl-c/ctrl-v where an invisible text area is needed
	clipBoard.style = style.ClipBoardStyle
	clipBoard.value = config.ClipBoardConfig.value;
	clipBoard.id = config.ClipBoardConfig.id;
    sheetTable.appendChild(clipBoard)

    // var span = document.createElement('span') // 用于获得字符串的显示长度
    // span.style = 'visibility: hidden;white-space: nowrap;filter:alpha(opacity=0);'
    // span.value = '';
    // span.id = 'sp'
    // sheetTable.appendChild(span)

    var multiLine = document.createElement('textarea')
    multiLine.style = 'position: absolute;left: 30px;top: 20px;display:none;'
	multiLine.style.left = screen.width/2 -150+ 'px'
    multiLine.style.top = screen.height/2 -300+ 'px'
    multiLine.style.width = 400 + 'px'
    multiLine.style.height = 400 + 'px'
    multiLine.value = '';
    multiLine.id = 'multiLine1'
    if(!config.WSConfig.isPreview){
        sheetEventBinder.initMultiLine(multiLine)
    }
	sheetTable.appendChild(multiLine)
    var dragBar = document.createElement('div')
    dragBar.style = 'position: absolute;'
    dragBar.style.backgroundColor = 'yellow'
    dragBar.style.width = 8 + 'px'
    dragBar.style.height = 8 + 'px'
    dragBar.style.display = 'none'
	dragBar.id = 'dragBar'
	dragBar.webkitUserDrag = 'true'
    if(!config.WSConfig.isPreview){
        sheetEventBinder.initDragBar(dragBar)
    }
	sheetTable.appendChild(dragBar)
}



module.exports.SheetRender = SheetRender

/***/ }),
/* 6 */
/***/ (function(module, exports, __webpack_require__) {

/**
 工具栏绑定事件
 */

var ToolEventHandlerModule = __webpack_require__(13)
var ToolEventHandler = ToolEventHandlerModule.ToolEventHandler
var toolEventHandler = null

var SheetEventHandlerModule = __webpack_require__(3)
var SheetEventHandler = SheetEventHandlerModule.SheetEventHandler
var sheetEventHandler = null


var ToolEventBinder = function (sheet) {
    this.sheet = sheet
    sheetEventHandler = new SheetEventHandler(sheet)
    this.sheet.colorOrBackgroudcolor = null
    toolEventHandler = new ToolEventHandler(sheet)
}
ToolEventBinder.prototype.buttonClick = function (action) {
    toolEventHandler.buttonClick(action)
}
ToolEventBinder.prototype.initFontFamily = function (fontFamilySelect) {
    var toolEventBinder = this
    fontFamilySelect.onchange = function () {
        toolEventBinder.sheet.setAttr('font', fontFamilySelect.value)
        toolEventBinder.sheet.render()
        fontFamilySelect.value = 'fontFamily'
    }
}
ToolEventBinder.prototype.initFontSize = function (fontSizeSelect) {
    var toolEventBinder = this
    fontSizeSelect.onchange = function () {
        //todo
        //switch
        toolEventBinder.sheet.setAttr('fontSize', fontSizeSelect.value)
        toolEventBinder.sheet.render()
        fontSizeSelect.value = 'fontSize'
    }
}
ToolEventBinder.prototype.initFontWeight = function (fontWeightDiv) {
    var toolEventBinder = this
    fontWeightDiv.onclick = function () {
        var lastCellId = toolEventBinder.sheet.lastCellid
        if (!toolEventBinder.sheet.cells[lastCellId]) {
            toolEventBinder.sheet.setAttr('bold', true)
        }
        toolEventBinder.sheet.setAttr('bold', !toolEventBinder.sheet.cells[lastCellId].bold)
        toolEventBinder.sheet.render()
    }
}
ToolEventBinder.prototype.initFontStyle = function (fontStyleDiv) {
    var toolEventBinder = this
    fontStyleDiv.onclick = function () {
        var lastCellId = toolEventBinder.sheet.lastCellid
        if (!toolEventBinder.sheet.cells[lastCellId]) {
            toolEventBinder.sheet.setAttr('italic', true)
        }
        toolEventBinder.sheet.setAttr('italic', !toolEventBinder.sheet.cells[lastCellId].italic)
        toolEventBinder.sheet.render()
    }
}
ToolEventBinder.prototype.initColor = function (colorDiv) {
    var toolEventBinder = this
    colorDiv.onclick = function () {
        var colorSelectDiv = document.getElementById('colorSelect')
        if (colorSelectDiv.style.display === 'none' || toolEventBinder.fontType === 'backgroundColor') {
            toolEventBinder.colorOrBackgroudcolor = 'foregroundColor'
            colorSelectDiv.style.display = 'block'
            var left = this.offsetLeft
            var top = this.offsetTop
            colorSelectDiv.style.top = top + 18 + 'px'
            colorSelectDiv.style.left = left + 'px'
        } else {
            colorSelectDiv.style.display = 'none'
        }

    }
}
ToolEventBinder.prototype.initBackgroundColor = function (backgroundColorDiv) {
    var toolEventBinder = this
    backgroundColorDiv.onclick = function () {
        var colorSelectDiv = document.getElementById('colorSelect')
        if (colorSelectDiv.style.display === 'none' || toolEventBinder.fontType === 'foregroundColor') {
            toolEventBinder.colorOrBackgroudcolor = 'backgroundColor'
            colorSelectDiv.style.display = 'block'
            var left = this.offsetLeft
            var top = this.offsetTop
            colorSelectDiv.style.top = top + 18 + 'px'
            colorSelectDiv.style.left = left + 'px'
        } else {
            colorSelectDiv.style.display = 'none'
        }
    }
}
ToolEventBinder.prototype.initColorSelect = function (td) {
    var toolEventBinder = this
    td.onclick = function () {
        document.getElementById('colorSelect').style.display = 'none'
        toolEventBinder.sheet.setAttr(toolEventBinder.colorOrBackgroudcolor, this.style.backgroundColor)
        toolEventBinder.sheet.render()

    }
}
ToolEventBinder.prototype.initMerge = function (mergeDiv, value) {
    var toolEventBinder = this
    mergeDiv.onclick = function () {
        sheetEventHandler.setCellBackgroundColor('#fff')
        toolEventBinder.sheet.setAttr('merge', value)
        toolEventBinder.sheet.render()
        sheetEventHandler.mouseDown(toolEventBinder.sheet.firstCellid)
        sheetEventHandler.mouseUp(document.getElementById(toolEventBinder.sheet.lastCellid))
    }
}

ToolEventBinder.prototype.initFontBorder = function (value) {
    var toolEventBinder = this
    switch (value) {
        case 'left':
            toolEventBinder.sheet.setAttr('leftFrame', true)

            break
        case 'top':
            toolEventBinder.sheet.setAttr('topFrame', true)
            break
        case 'right':
            toolEventBinder.sheet.setAttr('rightFrame', true)
            break
        case 'bottom':
            toolEventBinder.sheet.setAttr('bottomFrame', true)
            break
        case 'clear':
            toolEventBinder.sheet.setAttr('allFrameEx', false)
            break
        case 'outer':
            toolEventBinder.sheet.setAttr('leftFrame', true)
            toolEventBinder.sheet.setAttr('topFrame', true)
            toolEventBinder.sheet.setAttr('rightFrame', true)
            toolEventBinder.sheet.setAttr('bottomFrame', true)
            break
        case 'all':
            toolEventBinder.sheet.setAttr('allFrameEx', true)
            break
    }
    toolEventBinder.sheet.render()
}

ToolEventBinder.prototype.initTextAlign = function (value) {
    var toolEventBinder = this
    toolEventBinder.sheet.setAttr('alignment', value)
    toolEventBinder.sheet.render()
}
ToolEventBinder.prototype.initFileInput = function (fileInput) {
    fileInput.onchange = function () {
        var ajax
        if (window.XMLHttpRequest) {
            ajax = new XMLHttpRequest()
        } else if (window.ActiveXObject) {
            ajax = new window.ActiveXObject()
        } else {
            alert("请升级至最新版本的浏览器")
        }
        if (ajax !== null) {
            ajax.open("GET", "json.json", true)
            ajax.send(null)
            ajax.onreadystatechange = function () {
                if (ajax.readyState === 4 && ajax.status === 200) {
                    var CellList = JSON.parse(ajax.responseText)
                    CellList.forEach(function (e) {
                        for (var key in e) {
                            cellRender.renderCell(e['cellName'], key, e[key])
                        }
                    })
                }
            }
        }
    }
}
module.exports.ToolEventBinder = ToolEventBinder

/***/ }),
/* 7 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * Created by Ian on 17/6/5.
 */

var WSManagerModule = __webpack_require__(1)
var WSManager = WSManagerModule.WSManager
var wsManager=new WSManager()

var parentNode = document.getElementById('QianMoApp')

var myUrl=GetQueryString("url")
if(myUrl!=null && myUrl.toString().length>1){
    parentNode.setAttribute('url',myUrl)
}
function GetQueryString(name){
    var reg=new RegExp("(^|&)"+name+"=([^&]*)(&|$)")
    var r=window.location.search.substring(1).match(reg)
    if(r!=null){
        return unescape(r[2])
    }
    return null
}

wsManager.init(parentNode)

/***/ }),
/* 8 */
/***/ (function(module, exports, __webpack_require__) {

var config = __webpack_require__(0)

// var UtilModule=require('Util')
// var Util=UtilModule.Util
// var util=new Util()

var CellModule = __webpack_require__(4)
var Cell = CellModule.Cell

var CellRender = function (sheet) {
    this.sheet = sheet
}
/**
 * 渲染单元格
 * @param id
 * @param cmd
 * @param value
 */
CellRender.prototype.renderCell = function (id, cmd, value) {
    // var sheet=this.sheet
    // if(id && id.indexOf('_')==-1){
    //     var index=id.search(/\d/)
    //     id=id.substring(0,index)+'_'+id.substring(index,id.length)
    // }
    var ele = document.getElementById(id)

    // if(id && !sheet.cells[id]){
    //     sheet.cells[id] = new Cell(id)
    // }
    if (ele) {
        switch (cmd) {
            //合并范围
            case "colSpan":
                if(value>0) ele.colSpan = value
                break
            case 'rowSpan':
                if(value>0)ele.rowSpan = value
                break
            case "show":
                if (value) ele.style.display = 'table-cell'
                else {
                    ele.style.display = 'none'
                }
                break
            //字体颜色
            case "foregroundColor":
                ele.style.color = value


                break
            //背景色
            case "backgroundColor":

                ele.style.backgroundColor = value

                break
            //文本内容
            case "content":
                if (value === null) {
                    value = ''
                }
                ele.firstChild.innerHTML = value
                break
            //字体
            case "font":
                if (value == 'Default') {
                    value = ''
                }
                ele.style.fontFamily = value
                break
            //字体大小
            case "fontSize":
                if (value == 'Default' || value == '-----') {
                    value = ''
                }
                ele.style.fontSize = value
                //this.sheet.cells[id].fontSize=value
                break
            //字体类型
            //case "format":
            //    this.sheet.cells[id].format=value
            //    break
            //加粗
            case "bold":
                if (value) ele.style.fontWeight = 'bold'
                else ele.style.fontWeight = ''
                break

            //斜体
            case "italic":
                if (value) ele.style.fontStyle = 'italic'
                else ele.style.fontStyle = ''
                break
            //
            //对齐方向
            case "alignment":

                ele.style.textAlign = value
                ele.firstChild.style.textAlign = value

                break
            //换行
            case "wrapText":
                if (value) ele.firstChild.className = 'wrap'
                else ele.firstChild.className = 'noWrap'
                break
            //边框
            case "bottomFrame":
                var newId = id.split('_')[0] + '_' + (parseInt(id.split('_')[1])+1)
                if (value ) {
                    ele.style.borderBottom = '1px solid #000'
                } else if(!this.sheet.cells[newId] || !this.sheet.cells[newId].topFrame){
                    ele.style.borderBottom = ''
                    console.log(newId)
                }
                // if (value) ele.style.borderBottom = '2px solid #000'
                // else ele.style.borderBottom = ''
                break
            case "leftFrame":
                var newId = String.fromCharCode(id.split('_')[0].charCodeAt(0) - 1) + '_' + id.split('_')[1]
                if (value) {
                    document.getElementById(newId).style.borderRight = '1px solid #000'
                } else if(!this.sheet.cells[newId] || !this.sheet.cells[newId].rightFrame)document.getElementById(newId).style.borderRight = ''
                // if (value) ele.style.borderLeft = '2px solid #000'
                // else ele.style.borderLeft = ''
                break
            case "rightFrame":
                var newId = String.fromCharCode(id.split('_')[0].charCodeAt(0) + 1) + '_' + id.split('_')[1]
                if (value) {
                    ele.style.borderRight = '1px solid #000'
                } else if(!this.sheet.cells[newId] || !this.sheet.cells[newId].leftFrame) ele.style.borderRight = ''
                // if (value) ele.style.borderRight = '2px solid #000'
                // else ele.style.borderRight = ''
                break
            case "topFrame":
                var newId = id.split('_')[0] + '_' + (parseInt(id.split('_')[1])-1)
                if (value) {
                    document.getElementById(newId).style.borderBottom = '1px solid #000'
                } else if(!this.sheet.cells[newId] || !this.sheet.cells[newId].bottomFrame) document.getElementById(newId).style.borderBottom = ''
                // if (value) ele.style.borderTop = '2px solid #000'
                // else ele.style.borderTop = ''
                break
            //宽高
            case 'width':
                if (value) {
                    var newId = id.split('_')[0]+'_0'
                    document.getElementById(newId).firstChild.style.width = parseInt(value) + 'px'
                }
                break
            case 'height':
                if (value){
                    var newId = '@_'+id.split('_')[1]
                    document.getElementById(newId).firstChild.style.height = parseInt(value) + 'px'
                }
                break
            case 'area':
                console.log('!!!')
                break
            // case "formula":
            //     this.sheet.cells[id].formula=value
            //     break
        }
    }
}

module.exports.CellRender = CellRender

/***/ }),
/* 9 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * 表格对象
 */


var configModule = __webpack_require__(0)
var CellModule = __webpack_require__(4)
var Cell = CellModule.Cell
var sheetConfig = configModule.SheetConfig
// var UtilModule = require('Util')
// var Util = UtilModule.Util
// var util = new Util()
var CellRenderModule = __webpack_require__(8)
var CellRender = CellRenderModule.CellRender
//var cellRender = new CellRender
var Sheet = function (UndoStack) {
    this.cellRender = new CellRender(this)
    this.height = sheetConfig.height
    this.width = sheetConfig.width

    this.headWidth = sheetConfig.headWidth
    this.headHeight = sheetConfig.headHeight


    this.rowNum = sheetConfig.rowNum
    this.colNum = sheetConfig.colNum

    this.maxCol = 65
    this.maxRow = 1

    this.cells = {}
    this.firstCopiedCell = null
    this.copiedCells = {}
    // todo cellAttrs
    this.cellAttrs = [['backgroundColor', ''], ['foregroundColor', ''], ['bold', false], ['italic', false],
        ['content', ''], ['fontSize', ''], ['alignment', 'left'], ['font', ''], ['show', true],
        ['colSpan', 1], ['rowSpan', 1]]
    //['bottomFrame', false], ['leftFrame', false], ['topFrame', false], ['rightFrame', false]
    this.UndoStack = UndoStack
    this.undostackCache = []

    this.isMouseDown = false
    this.isKeyDown = false
    this.isEditing = false
    this.isMultiLineEditing = false
    this.isDraging = false

    //this.isFont = false
    this.isPreview = false

    this.firstCellid = 'A_1'
    this.lastCellid = 'A_1'
    this.editCells = this.getColAndRow()
}

Sheet.prototype.addCell = function (id) {
    if (!this.cells[id]) {
        this.cells[id] = new Cell(id)
        if (id.split('_')[0].charCodeAt(0) > this.maxCol) this.maxCol = id.split('_')[0].charCodeAt(0)
        if (parseInt(id.split('_')[1]) > this.maxRow) this.maxRow = parseInt(id.split('_')[1])
    }
}

Sheet.prototype.setAttr = function (attr, value, cells) {
    if (!cells) var cells = this.editCells
    if (cells != null) {
        switch (attr) {
            //缩进
            case 'indentation':
                if (value > 0) {
                    var html = '';
                    for (var i = 0; i < value; i++) {
                        html += '&nbsp;'
                    }
                    var i = cells.firstCellCol
                    var j = cells.firstCellRow
                    var id = String.fromCharCode(i) + '_' + j
                    if (this.cells[id]) {
                        if (this.cells[id]['content']) {
                            if (this.cells[id].show) {
                                html = html + this.cells[id]['content']
                                this.undostackCache.push([id, 'content', html, this.cells[id]['content']])
                                this.cells[id]['content'] = html
                            }
                        }
                    }
                }
                break
            case 'cellname':
            case 'area':
                break
            case 'merge':
                if (value) {
                    for (var i = cells.firstCellCol; i <= cells.lastCellCol; i++) {
                        for (var j = cells.firstCellRow; j <= cells.lastCellRow; j++) {
                            var id = String.fromCharCode(i) + '_' + j
                            this.addCell(id)
                            if (this.cells[id].show) {
                                if (i === cells.firstCellCol && j === cells.firstCellRow) {
                                    var colSpan = cells.lastCellCol - cells.firstCellCol + 1
                                    var rowSpan = cells.lastCellRow - cells.firstCellRow + 1
                                    if (colSpan !== this.cells[id].colSpan || rowSpan !== this.cells[id].rowSpan) {
                                        this.undostackCache.push([id, 'colSpan', colSpan, this.cells[id].colSpan])
                                        this.undostackCache.push([id, 'rowSpan', rowSpan, this.cells[id].rowSpan])
                                        //this.undostackCache.push([id, 'area', [colSpan, rowSpan], [this.cells[id].colSpan, this.cells[id].rowSpan]])
                                        this.cells[id].colSpan = colSpan
                                        this.cells[id].rowSpan = rowSpan
                                    }
                                }
                                else {
                                    this.cells[id].show = false
                                    this.undostackCache.push([id, 'show', false, true])
                                    this.undostackCache.push([id, 'colSpan', cells.firstCellCol - i + 1, this.cells[id].colSpan])
                                    this.undostackCache.push([id, 'rowSpan', cells.firstCellRow - j + 1, this.cells[id].rowSpan])
                                    this.cells[id].colSpan = cells.firstCellCol - i + 1
                                    this.cells[id].rowSpan = cells.firstCellRow - j + 1
                                }
                            }
                        }
                    }
                    this.lastCellid = this.firstCellid
                }
                else {
                    this.firstCellid = String.fromCharCode(cells.firstCellCol) + '_' + cells.firstCellRow
                    this.lastCellid = String.fromCharCode(cells.lastCellCol) + '_' + cells.lastCellRow
                    for (var i = cells.firstCellCol; i <= cells.lastCellCol; i++) {
                        for (var j = cells.firstCellRow; j <= cells.lastCellRow; j++) {
                            var id = String.fromCharCode(i) + '_' + j
                            if (!this.cells[id]) {
                                continue
                            }
                            if (this.cells[id].colSpan != 1 || this.cells[id].rowSpan != 1) {
                                this.undostackCache.push([id, 'colSpan', 1, this.cells[id].colSpan])
                                this.undostackCache.push([id, 'rowSpan', 1, this.cells[id].rowSpan])
                                this.cells[id].colSpan = 1
                                this.cells[id].rowSpan = 1
                            }
                            if (!this.cells[id].show) {
                                this.undostackCache.push([id, 'show', true, false])
                                this.cells[id].show = true
                            }
                        }
                    }
                }
                break
            case 'leftFrame':
            case 'topFrame':
            case 'rightFrame':
            case 'bottomFrame':
                switch (attr) {
                    case 'leftFrame':
                        var mini = cells.firstCellCol
                        var minj = cells.firstCellRow
                        var maxi = cells.firstCellCol
                        var maxj = cells.lastCellRow
                        break
                    case 'topFrame':
                        var mini = cells.firstCellCol
                        var minj = cells.firstCellRow
                        var maxi = cells.lastCellCol
                        var maxj = cells.firstCellRow
                        break
                    case 'rightFrame':
                        var mini = cells.lastCellCol
                        var minj = cells.firstCellRow
                        var maxi = cells.lastCellCol
                        var maxj = cells.lastCellRow
                        break
                    case 'bottomFrame':
                        var mini = cells.firstCellCol
                        var minj = cells.lastCellRow
                        var maxi = cells.lastCellCol
                        var maxj = cells.lastCellRow
                        break

                }
                for (var i = mini; i <= maxi; i++) {
                    for (var j = minj; j <= maxj; j++) {
                        var id = String.fromCharCode(i) + '_' + j
                        this.addCell(id)
                        if (this.cells[id][attr] != value) {
                            this.cells[id][attr] = value
                            this.undostackCache.push([id, attr, value, !value])
                        }
                    }
                }
                break
            case 'allFrameEx':
                for (var i = cells.firstCellCol; i <= cells.lastCellCol; i++) {
                    for (var j = cells.firstCellRow; j <= cells.lastCellRow; j++) {
                        var id = String.fromCharCode(i) + '_' + j
                        this.addCell(id)
                        // var env = this
                        //
                        // if (env.cells[id].rightFrame != value && i !== cells.firstCellCol - 1) {
                        //     env.undostackCache.push([id, 'topFrame', value, !value])
                        //     env.cells[id].rightFrame = value
                        // }
                        // if (env.cells[id].bottomFrame != value && j !== cells.firstCellRow - 1) {
                        //     env.undostackCache.push([id, 'rightFrame', value, !value])
                        //     env.cells[id].bottomFrame = value
                        // }
                        var env = this

                        var a = ['leftFrame', 'rightFrame', 'topFrame', 'bottomFrame']

                        a.forEach(function (frame) {
                            if (env.cells[id][frame] != value) {
                                env.undostackCache.push([id, frame, value, !value])
                                env.cells[id][frame] = value
                            }
                        })

                    }
                }
                if (value == false) {
                    if (cells.firstCellCol > 65) {
                        var firstCellid = String.fromCharCode(cells.firstCellCol - 1)
                            + '_' + cells.firstCellRow
                        var lastCellid = String.fromCharCode(cells.firstCellCol - 1)
                            + '_' + cells.lastCellRow
                        var newCell = this.getColAndRow(firstCellid, lastCellid)
                        this.setAttr('rightFrame', false, newCell)
                        delete  newCell
                    }
                    if (cells.firstCellRow > 1) {

                        var firstCellid = String.fromCharCode(cells.firstCellCol)
                            + '_' + (cells.firstCellRow - 1)
                        var lastCellid = String.fromCharCode(cells.lastCellCol)
                            + '_' + (cells.firstCellRow - 1)
                        var newCell = this.getColAndRow(firstCellid, lastCellid)
                        this.setAttr('bottomFrame', false, newCell)
                        delete  newCell
                    }
                    if (document.getElementById(String.fromCharCode(cells.lastCellCol)
                            + '_' + (cells.lastCellRow + 1))) {
                        var firstCellid = String.fromCharCode(cells.firstCellCol)
                            + '_' + (cells.lastCellRow + 1)
                        var lastCellid = String.fromCharCode(cells.lastCellCol)
                            + '_' + (cells.lastCellRow + 1)
                        var newCell = this.getColAndRow(firstCellid, lastCellid)
                        this.setAttr('topFrame', false, newCell)
                        delete  newCell
                    }
                    if (document.getElementById(String.fromCharCode(cells.lastCellCol + 1)
                            + '_' + (cells.lastCellRow))) {
                        var firstCellid = String.fromCharCode(cells.lastCellCol + 1)
                            + '_' + (cells.firstCellRow)
                        var lastCellid = String.fromCharCode(cells.lastCellCol + 1)
                            + '_' + (cells.lastCellRow )
                        var newCell = this.getColAndRow(firstCellid, lastCellid)
                        this.setAttr('leftFrame', false, newCell)
                        delete  newCell
                    }
                }
                break
            case 'allFrame':
                for (var i = cells.firstCellCol; i <= cells.lastCellCol; i++) {
                    for (var j = cells.firstCellRow; j <= cells.lastCellRow; j++) {
                        var id = String.fromCharCode(i) + '_' + j
                        this.addCell(id)
                        var env = this

                        var a = ['leftFrame', 'rightFrame', 'topFrame', 'bottomFrame']

                        a.forEach(function (frame) {
                            if (env.cells[id][frame] != value) {
                                env.undostackCache.push([id, frame, value, !value])
                                env.cells[id][frame] = value
                            }
                        })

                    }
                }
                break
            default:
                for (var i = cells.firstCellCol; i <= cells.lastCellCol; i++) {
                    for (var j = cells.firstCellRow; j <= cells.lastCellRow; j++) {
                        var id = String.fromCharCode(i) + '_' + j
                        this.addCell(id)

                            if (this.cells[id][attr] !== value) {
                                this.undostackCache.push([id, attr, value, this.cells[id][attr]])
                                this.cells[id][attr] = value
                            }

                    }
                }
                break
        }
    }
}
//将value赋值给attr属性，赋值范围由sheet.firstcell和sheet.lastcell指定

Sheet.prototype.sheet2json = function () {
//todo
}

Sheet.prototype.json2sheet = function () {
//todo
}

Sheet.prototype.render = function () {
    var sheet = this
    this.undostackCache.forEach(function (cache) {
        sheet.cellRender.renderCell(cache[0], cache[1], cache[2])
    })
    this.UndoStack.pushStep(this.undostackCache)
    this.undostackCache = []
    //遍历sheet.undostackCache,将sheet.undostackCache压入sheet.undostack
}
Sheet.prototype.renderCell = function (id, cmd, value) {
    this.cellRender.renderCell(id, cmd, value)
}

Sheet.prototype.undo = function () {
    var undoSteps = this.UndoStack.unDo()
    if (undoSteps) {
        var sheet = this
        undoSteps.forEach(function (step) {
            sheet.cells[step[0]][step[1]] = step[3]
            sheet.cellRender.renderCell(step[0], step[1], step[3])
        })
    }
}

Sheet.prototype.redo = function () {
    var undoSteps = this.UndoStack.reDo()
    if (undoSteps) {
        var sheet = this
        undoSteps.forEach(function (step) {
            sheet.cells[step[0]][step[1]] = step[2]
            sheet.cellRender.renderCell(step[0], step[1], step[2])
        })
    }

}

Sheet.prototype.getColAndRow = function (firstCellid, lastCellid) {
    if (firstCellid == null) {
        firstCellid = this.firstCellid
        lastCellid = this.lastCellid
    }
    //if (firstCell && firstCell.id && lastCell && lastCell.id) {
    var result = {}
    result.firstCellCol = Math.min(firstCellid.split('_')[0].charCodeAt(0), lastCellid.split('_')[0].charCodeAt(0))
    result.firstCellRow = Math.min(parseInt(firstCellid.split('_')[1]), parseInt(lastCellid.split('_')[1]))
    result.lastCellCol = Math.max(lastCellid.split('_')[0].charCodeAt(0), firstCellid.split('_')[0].charCodeAt(0))
    result.lastCellRow = Math.max(parseInt(lastCellid.split('_')[1]), parseInt(firstCellid.split('_')[1]))
    //}
    var sheet = this
    return getRectangle(result, sheet)
}

getRectangle = function (cells, sheet) {
    //console.log(cells)
    var flag = false
    var minCol = cells.firstCellCol
    var minRow = cells.firstCellRow
    var maxCol = cells.lastCellCol
    var maxRow = cells.lastCellRow

    for (var i = cells.firstCellCol; i <= cells.lastCellCol; i++) {
        for (var j = cells.firstCellRow; j <= cells.lastCellRow; j++) {
            var id = String.fromCharCode(i) + '_' + j
            if (!sheet.cells[id]) {
                continue

            }
            if (!sheet.cells[id].show) {
                if ((i + sheet.cells[id].colSpan - 1) < minCol) {
                    flag = true
                    minCol = (i + sheet.cells[id].colSpan - 1)

                }
                if ((j + sheet.cells[id].rowSpan - 1) < minRow) {
                    flag = true
                    minRow = (j + sheet.cells[id].rowSpan - 1)
                }
            }
            else {
                if (sheet.cells[id].colSpan - 1 + i > maxCol) {
                    flag = true
                    maxCol = sheet.cells[id].colSpan - 1 + i
                }
                if (sheet.cells[id].rowSpan - 1 + j > maxRow) {
                    flag = true
                    maxRow = sheet.cells[id].rowSpan - 1 + j
                }
            }

        }
    }
    cells.firstCellCol = minCol
    cells.firstCellRow = minRow
    cells.lastCellCol = maxCol
    cells.lastCellRow = maxRow
    if (flag) return getRectangle(cells, sheet)
    else return cells
}


module.exports.Sheet = Sheet

/***/ }),
/* 10 */
/***/ (function(module, exports, __webpack_require__) {

//事件绑定对象
var SheetEventHandlerModule = __webpack_require__(3)
var SheetEventHandler = SheetEventHandlerModule.SheetEventHandler
var sheetEventHandler = null

var SheetEventBinder = function (sheet) {
    this.sheet = sheet
    sheetEventHandler = new SheetEventHandler(sheet)
    window.onmouseup = function () {
        sheet.isMouseDown = false
    }
    document.onkeydown = function (event) {
        sheetEventHandler.keyDown(event)

        if (!sheet.isMultiLineEditing && !sheet.isEditing) return false
    }
}

SheetEventBinder.prototype.initRowTD = function (rowTD) {
    var sheet = this.sheet
    rowTD.onmousedown = function () {
        sheetEventHandler.mouseDown(rowTD.id)
    }
    rowTD.onmouseup = function () {
        sheetEventHandler.mouseUp(rowTD)
    }
    rowTD.onmousemove = function () {
        if (sheet.isMouseDown || sheet.isDraging) {
            sheetEventHandler.mouseMove(rowTD.id)
        }
    }
    rowTD.ondblclick = function () {
        sheetEventHandler.dblclick(this)
    }
}
SheetEventBinder.prototype.initInput = function (input) {
    input.onblur = function () {
        sheetEventHandler.inputBlur()
    }
    input.onfocus = function () {
        sheetEventHandler.inputFocus()
    }
}
SheetEventBinder.prototype.initMultiLine = function (multiLine) {
    multiLine.onblur = function () {
        sheetEventHandler.multiLineBlur(multiLine.value)
        multiLine.value = ''
        multiLine.style.display = 'none'
    }
}
SheetEventBinder.prototype.initDragBar = function (dragBar) {
    dragBar.onmousedown = function () {
        sheetEventHandler.dragBarMouseDown()
    }
    dragBar.onmouseup = function () {
        sheetEventHandler.dragBarMouseUp()
    }
}
module.exports.SheetEventBinder = SheetEventBinder

/***/ }),
/* 11 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * Created by Ian on 17/7/22.
 */

var style  = __webpack_require__(2)
var config = __webpack_require__(0)

var SliderBarRender = function (sheet, sheetDiv) {
	this.sheet = sheet
	this.sheetDiv = sheetDiv
	this.isOpen = false

	// alert(this.sheetDiv)
}

SliderBarRender.prototype.init = function (sliderBarDiv, sheetTableDiv) {

	var toggleDiv = document.createElement('div')

	// sliderBarDiv.style = style.SliderBarStyle

	// if(this.isOpen){
	// 	sheetTableDiv.style.width = '75%'
	// 	sliderBarDiv.style.width = '25%'
	// }else{
	// 	sheetTableDiv.style.width = '100%'
	// 	sliderBarDiv.style.width = '0'
	// }

	// console.log(this.sheet)
	renderBar(sliderBarDiv, sheetTableDiv)
	renderToggle(this.sheetDiv, toggleDiv)

	var isOpen = this.isOpen
	var sheetDiv = this.sheetDiv

	toggleDiv.onclick = function () {
		isOpen = !isOpen

		renderBar(sliderBarDiv, sheetTableDiv, isOpen)
		renderToggle(sheetDiv, toggleDiv, isOpen)
	}
}

function renderToggle(sheetDiv, toggleDiv, isOpen) {
	var iconHtmlMap = config.SlideBarConfig.iconHtmlMap
	var toggleOpenIcon = config.SlideBarConfig.toggleOpenIcon
	var toggleCloseIcon = config.SlideBarConfig.toggleCloseIcon

	// var toggleDiv = document.createElement('div')

	if(isOpen){

		toggleDiv.innerHTML = iconHtmlMap[toggleCloseIcon]
		toggleDiv.style = style.ToggleDivOpenLeft + style.ToggleDivStyle

		toggleDiv.onmouseover = function () {
			toggleDiv.style = style.ToggleDivOpenLeft + style.ToggleDivHoverStyle
		}

		toggleDiv.onmouseout = function () {
			toggleDiv.style = style.ToggleDivOpenLeft + style.ToggleDivStyle
			// toggleDiv.style.left = '76%'
		}
	}else{

		toggleDiv.innerHTML = iconHtmlMap[toggleOpenIcon]
		toggleDiv.style = style.ToggleDivCloseLeft + style.ToggleDivStyle

		toggleDiv.onmouseover = function () {
			toggleDiv.style = style.ToggleDivCloseLeft + style.ToggleDivHoverStyle
		}

		toggleDiv.onmouseout = function () {
			toggleDiv.style = style.ToggleDivCloseLeft + style.ToggleDivStyle
		}
	}

	// console.log(sheet)
	sheetDiv.appendChild(toggleDiv)
}

function renderBar(sliderBarDiv, sheetTableDiv, isOpen) {

	sliderBarDiv.style = style.SliderBarStyle

	if(isOpen){
		sheetTableDiv.style.width = '75%'
		sliderBarDiv.style.width = '25%'
	}else{
		sheetTableDiv.style.width = '100%'
		sliderBarDiv.style.width = '0'
	}
}


module.exports.SliderBarRender = SliderBarRender

/***/ }),
/* 12 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * 工具栏对象
 */
var configModule = __webpack_require__(0)

var ToolConfig = configModule.ToolConfig

var Tool = function () {
    this.width = ToolConfig.width
    this.height = ToolConfig.height
}

module.exports.Tool = Tool

/***/ }),
/* 13 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * 工具栏事件
 */
var config = __webpack_require__(0)


// var CellRenderModule = require('CellRender')
// var CellRender = CellRenderModule.CellRender
// var cellRender = null

var SheetEventHandlerModule = __webpack_require__(3)
var SheetEventHandler = SheetEventHandlerModule.SheetEventHandler
var sheetEventHandler = null

var ToolEventHandler = function (sheet) {
    this.sheet = sheet
    // cellRender = new CellRender(sheet)
    sheetEventHandler = new SheetEventHandler(sheet)
}
ToolEventHandler.prototype.buttonClick = function (action, value) {
    var sheet = this.sheet
    switch (action) {
        case "wrapText":
            var lastCellId = this.sheet.lastCellid
            if (!sheet.cells[lastCellId]) {
                sheet.setAttr('wrapText', true)
            }
            else {
                sheet.setAttr('wrapText', !this.sheet.cells[lastCellId].wrapText)
            }
            sheet.render()
            break
        case "preview":
            var e = this.sheet.cells
            if (JSON.stringify(e) !== "{}") {
                var needEditCells = []
                for (var key in e) {
                    if (e[key]["content"] != '' && e[key]["content"].indexOf('=') == 0) {
                        var formula = {
                            'coord': key,
                            'value': e[key]["content"]
                        }
                        needEditCells.push(formula)
                        this.sheet.cells[key]["formula"] = e[key]["content"]
                    }
                }
                if (needEditCells.length > 0) {
                    var ajax
                    if (window.XMLHttpRequest) {
                        ajax = new XMLHttpRequest()
                    } else if (window.ActiveXObject) {
                        ajax = new window.ActiveXObject()
                    } else {
                        alert("请升级至最新版本的浏览器")
                    }
                    if (ajax !== null) {
                        ajax.open("POST", "http://localhost:8088/qianmo-service/changeContent", true)
                        needEditCells = JSON.stringify(needEditCells)
                        ajax.send(needEditCells)
                        ajax.onreadystatechange = function () {
                            if (ajax.readyState == 4 && ajax.status == 200) {
                                var CellList = JSON.parse(ajax.responseText)
                                for (var key in CellList) {
                                    var value = CellList[key].value.substring(1)
                                    if (value.indexOf("+") != -1) {
                                        var result = 0
                                        var values = value.split("+")
                                        for (var v in values) {
                                            result += parseInt(values[v])
                                        }
                                        value = result
                                    }
                                    e[CellList[key].coord]["content"] = value
                                }
                                config.WSConfig.isPreview = true
                                var WSManagerModule = __webpack_require__(1)
                                var WSManager = WSManagerModule.WSManager
                                var newWsManager = new WSManager()
                                var parentNode = document.getElementById('QianMoApp2')
                                // parentNode.firstChild.remove()
                                newWsManager.init(parentNode)
                                for (var key in e) {
                                    for (var key1 in e[key]) {
                                        if (key1 !== "formula") {
                                            sheet.renderCell(key,key1, e[key][key1])
                                        }

                                    }
                                }
                            }
                        }
                    }
                }else{
                    config.WSConfig.isPreview = true
                    var WSManagerModule = __webpack_require__(1)
                    var WSManager = WSManagerModule.WSManager
                    var newWsManager = new WSManager()
                    var parentNode = document.getElementById('QianMoApp2')
                    // parentNode.firstChild.remove()
                    newWsManager.init(parentNode)
                    for (var key in e) {
                        for (var key1 in e[key]) {
                            if (key1 !== "formula") {
                                sheet.renderCell(key,key1, e[key][key1])
                            }
                        }
                    }
                }
                document.getElementById('QianMoApp').style.display='none'
                document.getElementById('QianMoApp2').style.display='inline'
            } else {
                alert("无内容，不允许预览！")
            }

            console.log(this.sheet.cells)
            break
        case "Edit":
            document.getElementById('QianMoApp2').style.display='none'
            document.getElementById('QianMoApp').style.display='inline'
            console.log(this.sheet.cells)
            config.WSConfig.isPreview = false
            var WSManagerModule = __webpack_require__(1)
            var WSManager = WSManagerModule.WSManager
            var newWsManager = new WSManager()
            var parentNode = document.getElementById('QianMoApp')
            parentNode.firstChild.remove()
            newWsManager.init(parentNode)
            var e = sheet.cells
            console.log(e)
            for (var key in e) {
                var cells=sheet.getColAndRow(key,key)
                for (var key1 in e[key]) {
                    if (key1 === "formula" && e[key][key1] !== "") {
                        // cellRender.renderCell(key, "content", e[key][key1])
                        sheet.setAttr('content', e[key][key1],cells)
                        sheet.render()
                    } else {
                        // cellRender.renderCell(key, key1, e[key][key1])
                        sheet.setAttr(key1, e[key][key1],cells)
                        sheet.render()
                    }

                }
            }
            break
        case "Down":
            var e = this.sheet.cells
            console.log(e)
            var a = {}
            var CellList = []
            for (var key in e) {
                e[key]["area"] = e[key]["area"].split("_")[0] + e[key]["area"].split("_")[1]
                    + e[key]["area"].split("_")[2]
                e[key]["area"] = e[key]["area"].split(":")[0] + "-" + e[key]["area"].split(":")[1]
                var addCell = {
                    "cellName": key.split("_")[0] + key.split("_")[1],
                    "area": e[key]["area"],
                    "content": e[key]["content"] === undefined ? "" : e[key]["content"],
                    "format": e[key]["format"] === undefined ? "" : e[key]["format"],
                    "font": e[key]["font"] === undefined ? "" : e[key]["font"],
                    "fontSize": e[key]["fontSize"] === undefined ? "" : e[key]["fontSize"],
                    "foregroundColor": e[key]["foregroundColor"] === undefined ? "" : e[key]["foregroundColor"],
                    "backgroundColor": e[key]["backgroundColor"] === undefined ? "" : e[key]["backgroundColor"],
                    "formula": e[key]["formula"] === undefined ? "" : e[key]["formula"],
                    "leftFrame": e[key]["leftFrame"] === undefined ? "" : e[key]["leftFrame"],
                    "topFrame": e[key]["topFrame"] === undefined ? "" : e[key]["topFrame"],
                    "rightFrame": e[key]["rightFrame"] === undefined ? "" : e[key]["rightFrame"],
                    "bottomFrame": e[key]["bottomFrame"] === undefined ? "" : e[key]["bottomFrame"],
                    "indentation": e[key]["indentation"] === undefined ? "" : e[key]["indentation"],
                    "alignment": e[key]["alignment"] === undefined ? "" : e[key]["alignment"],
                    "bold": e[key]["bold"] === undefined ? "" : e[key]["bold"],
                    "italic": e[key]["italic"] === undefined ? "" : e[key]["italic"]
                }

                CellList.push(addCell)
            }
            var json= JSON.stringify(CellList)
            var ajax
            if (window.XMLHttpRequest) {
                ajax = new XMLHttpRequest()
            } else if (window.ActiveXObject) {
                ajax = new window.ActiveXObject()
            } else {
                alert("请升级至最新版本的浏览器")
            }
            if(ajax !=null){
                ajax.open("POST","http://localhost:8088/qianmo-service/excelDownload",true)
                ajax.onload=function(){
                    if(ajax.status==200){
                        var filename = "";
                        console.log(ajax.getResponseHeader)
                        var disposition = ajax.getResponseHeader('Content-Disposition');
                        if (disposition && disposition.indexOf('attachment') !== -1) {
                            var filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
                            var matches = filenameRegex.exec(disposition);
                            if (matches != null && matches[1]) filename = matches[1].replace(/['"]/g, '');
                        }
                        var type = ajax.getResponseHeader('Content-Type');
                        var blob = new Blob([this.response], { type: type });
                        if (typeof window.navigator.msSaveBlob !== 'undefined') {
                            // IE workaround for "HTML7007: One or more blob URLs were revoked by closing the blob for which they were created. These URLs will no longer resolve as the data backing the URL has been freed."
                            window.navigator.msSaveBlob(blob, filename);
                        } else {
                            var URL = window.URL || window.webkitURL;
                            var downloadUrl = URL.createObjectURL(blob);
                            console.log(downloadUrl)
                            if (filename) {
                                // use HTML5 a[download] attribute to specify filename
                                var a = document.createElement("a");
                                // safari doesn't support this yet
                                if (typeof a.download === 'undefined') {
                                    window.location = downloadUrl;
                                } else {
                                    a.href = downloadUrl;
                                    a.download = filename;
                                    console.log(a)
                                    document.body.appendChild(a);
                                    a.click();
                                }
                            } else {
                                window.location = downloadUrl;
                            }

                            setTimeout(function () { URL.revokeObjectURL(downloadUrl); }, 100); // cleanup
                        }
                    }
                    document.getElementById('statu').innerHTML='SUCCESS'
                }
                ajax.responseType = 'blob'
                ajax.send(json)
            }
            break
        case  "Init":
            var parentNode = document.getElementById("QianMoApp")
            var param = parentNode.getAttribute("url")
            var url = ''
            var method="GET"
            if(!param&& param!=''){
                url='json.json'
                param=null
            }else{
                method="POST"
                url='http://localhost:8088/qianmo-service/getContentJson'
                param=param.substring(param.lastIndexOf("\\")+1)
            }
            var ajax
            if (window.XMLHttpRequest) {
                ajax = new XMLHttpRequest()
            } else if (window.ActiveXObject) {
                ajax = new window.ActiveXObject()
            } else {
                alert("请升级至最新版本的浏览器")
            }
            if (ajax !== null) {
                ajax.open(method, url, true)
                ajax.send(param)
                ajax.onreadystatechange = function () {
                    if (ajax.readyState === 4 && ajax.status === 200) {
                        var CellList = JSON.parse(ajax.responseText)
                        console.log(CellList)
                        CellList.forEach(function (e) {

                            var area = e['area']
                            var firstCell = e['cellName'].split('_')

                            var cells = {}
                            cells.firstCellCol = firstCell[0].charCodeAt(0)
                            cells.firstCellRow = parseInt(firstCell[1])
                            var areaSubstring = area.split('_')
                            cells.lastCellCol = areaSubstring[2].charCodeAt(0)
                            cells.lastCellRow = parseInt(areaSubstring[3])
                            if (cells.firstCellCol !== cells.lastCellCol || cells.firstCellRow !== cells.lastCellRow) {
                                sheet.setAttr('merge', true, cells)
                            }


                            cells.lastCellCol = firstCell[0].charCodeAt(0)
                            cells.lastCellRow = parseInt(firstCell[1])


                            for (var key in e) {
                                switch (e[key]) {
                                    case 'false':
                                        var value = false
                                        break
                                    case 'true':
                                        var value = true
                                        break
                                    default:
                                        var value = e[key]
                                }
                                if ((key === 'bottomFrame' || key === 'topFrame' || key === 'leftFrame' ||
                                        key === 'rightFrame') && value === false) continue

                                sheet.setAttr(key, value, cells)
                            }
                        })

                        sheet.render()
                    }
                }
            }
            break
        case 'copy':
            cells = sheet.editCells
            sheet.firstCopiedCell = String.fromCharCode(cells.firstCellCol) + '_' + cells.firstCellRow
            sheet.copiedCells = []
            for (var i = cells.firstCellCol; i <= cells.lastCellCol; i++) {
                for (var j = cells.firstCellRow; j <= cells.lastCellRow; j++) {
                    var id = String.fromCharCode(i) + '_' + j
                    var newCell = {}
                    newCell.coord = id
                    sheet.cellAttrs.forEach(function (attr) {
                        if (!sheet.cells[id]) {
                            newCell[attr[0]] = attr[1]
                        }
                        else {
                            newCell[attr[0]] = sheet.cells[id][attr[0]]
                        }
                    })
                    if (!sheet.cells[id]) {
                        newCell.bottomFrame = false
                        newCell.topFrame = false
                        newCell.leftFrame = false
                        newCell.rightFrame = false
                    }
                    else {
                        newCell.bottomFrame = sheet.cells[id].bottomFrame
                        newCell.topFrame = sheet.cells[id].topFrame
                        newCell.leftFrame = sheet.cells[id].leftFrame
                        newCell.rightFrame = sheet.cells[id].rightFrame
                    }
                    sheet.copiedCells.push(newCell)
                }
            }
            console.log('copy done!')
            break
        case 'cut':
            cells = sheet.editCells
            sheet.firstCopiedCell = String.fromCharCode(cells.firstCellCol) + '_' + cells.firstCellRow
            sheet.copiedCells = []
            for (var i = cells.firstCellCol; i <= cells.lastCellCol; i++) {
                for (var j = cells.firstCellRow; j <= cells.lastCellRow; j++) {
                    var id = String.fromCharCode(i) + '_' + j
                    var newCell = {}
                    newCell.coord = id
                    sheet.cellAttrs.forEach(function (attr) {
                        if (!sheet.cells[id]) {
                            newCell[attr[0]] = attr[1]
                        }
                        else {
                            newCell[attr[0]] = sheet.cells[id][attr[0]]
                        }
                    })
                    if (!sheet.cells[id]) {
                        newCell.bottomFrame = false
                        newCell.topFrame = false
                        newCell.leftFrame = false
                        newCell.rightFrame = false
                    }
                    else {
                        newCell.bottomFrame = sheet.cells[id].bottomFrame
                        newCell.topFrame = sheet.cells[id].topFrame
                        newCell.leftFrame = sheet.cells[id].leftFrame
                        newCell.rightFrame = sheet.cells[id].rightFrame
                    }
                    sheet.copiedCells.push(newCell)
                }
            }
            sheet.cellAttrs.forEach(function (attr) {
                sheet.setAttr(attr[0], attr[1], cells)
            })
            sheet.setAttr('allFrame', false, cells)
            sheet.render()
            sheet.lastCellid = sheet.firstCellid

            console.log('cut done!')
            break
        case 'paste':
            var borders = [['bottomFrame', false], ['leftFrame', false], ['topFrame', false], ['rightFrame', false]]
            if (sheet.firstCopiedCell) {
                cells = sheet.editCells
                sheet.copiedCells.forEach(function (cell) {
                    var newCell = {}
                    var firstCopiedCellCol = sheet.firstCopiedCell.split('_')[0].charCodeAt(0)
                    var firstCopiedCellRow = parseInt(sheet.firstCopiedCell.split('_')[1])
                    cellCol = cell.coord.split('_')[0].charCodeAt(0)
                    cellRow = parseInt(cell.coord.split('_')[1])
                    newCell.firstCellCol = cells.firstCellCol + cellCol - firstCopiedCellCol
                    newCell.firstCellRow = cells.firstCellRow + cellRow - firstCopiedCellRow
                    newCell.lastCellCol = newCell.firstCellCol
                    newCell.lastCellRow = newCell.firstCellRow
                    sheet.cellAttrs.forEach(function (attr) {
                        sheet.setAttr(attr[0], cell[attr[0]], newCell)
                    })
                    borders.forEach(function (frame) {
                        sheet.setAttr(frame[0], cell[frame[0]], newCell)
                    })
                    delete newCell
                })
                sheet.render()
                console.log('paste done!')
                console.log(sheet.cells)
            } else {
                console.log('nothing to paste')
            }
            break
        case 'font':
            sheet.setAttr('font', value)
            sheet.render()
            break
        case 'fontSize':
            sheet.setAttr('fontSize', value)
            sheet.render()
            break
        case 'bold':
            var lastCellId = sheet.lastCell.id
            if (!sheet.cells[lastCellId]) {
                sheet.setAttr('bold', true)
            }
            sheet.setAttr('bold', !sheet.cells[lastCellId].bold)
            sheet.render()
            break
        case 'italic':
            var lastCellId = sheet.lastCell.id
            if (!sheet.cells[lastCellId]) {
                sheet.setAttr('italic', true)
            }
            sheet.setAttr('italic', !this.sheet.cells[lastCellId].italic)
            sheet.render()
            break
        case 'merge':
            this.setCellBackgroundColor('#fff')
            this.sheet.setAttr('merge', value)
            this.sheet.render()
            this.mouseDown(sheet.firstCellid)
            console.log(sheet.firstCellid)
            this.mouseUp(document.getElementById(sheet.lastCellid))
            break
        // case 'border':
        // switch(value){
        // 	case 'left':
        // 		sheet.setAttr('leftFrame',false)
        //             console.log('a')
        // 		break
        // 	case 'top':
        // 		sheet.setAttr('topFrame',false)
        // 		break
        // 	case 'right':
        // 		sheet.setAttr('rightFrame',false)
        // 		break
        // 	case 'bottom':
        // 		sheet.setAttr('bottomFrame',false)
        // 		break
        // 	case 'clear':
        // 		sheet.setAttr('allFrame',false)
        // 		break
        // 	case 'outer':
        // 		sheet.setAttr('leftFrame',true)
        // 		sheet.setAttr('topFrame',true)
        // 		sheet.setAttr('rightFrame',true)
        // 		sheet.setAttr('bottomFrame',true)
        // 		break
        // 	case 'all':
        // 		sheet.setAttr('allFrame',true)
        // 		break
        //
        // 	//fontBorderSelect.value='border'
        // }
        // sheet.render()
        //     break
        case 'alignment':
            sheet.setAttr('alignment', value)
            sheet.render()
            break
        case 'undo':
            sheet.undo()
            break
        case 'redo':
            sheet.redo()
            break
        // case 'sort':
        //     break
        case 'multiLine':
            document.getElementById('dragBar').style.display = 'none'
            sheet.isMultiLineEditing = true
            var ml = document.getElementById('multiLine1')
            ml.style.display = 'block'
            if (sheet.cells[sheet.lastCellid]) ml.value = sheet.cells[sheet.lastCellid].content
            else ml.value = ''
            ml.focus()
            break
        case 'cleanText':
            sheet.setAttr('content', '')
            sheet.render()
            break
        case 'cleanStyle':
            sheet.cellAttrs.forEach(function (attr) {
                if (attr[0] !== 'content') sheet.setAttr(attr[0], attr[1])
            })
            sheet.render()
            break
        case 'cleanAll':
            sheet.cellAttrs.forEach(function (attr) {
                sheet.setAttr(attr[0], attr[1])
            })
            sheet.setAttr('allFrameEx', false)
            sheet.render()
            break
        case 'addCol':
        case 'deleteCol':
            cells = sheet.getColAndRow(sheet.firstCellid.split('_')[0]+'_1',String.fromCharCode(sheet.maxCol)+'_'+sheet.maxRow)
            //sheet.firstCopiedCell = String.fromCharCode(cells.firstCellCol) + '_' + cells.firstCellRow
            var sheetcopiedCells = []
            for (var i = cells.firstCellCol; i <= cells.lastCellCol; i++) {
                for (var j = cells.firstCellRow; j <= cells.lastCellRow; j++) {
                    var id = String.fromCharCode(i) + '_' + j
                    var newCell = {}
                    newCell.coord = id
                    sheet.cellAttrs.forEach(function (attr) {
                        if (!sheet.cells[id]) {
                            newCell[attr[0]] = attr[1]
                        }
                        else {
                            newCell[attr[0]] = sheet.cells[id][attr[0]]
                        }
                    })
                    if (!sheet.cells[id]) {
                        newCell.bottomFrame = false
                        newCell.topFrame = false
                        newCell.leftFrame = false
                        newCell.rightFrame = false
                    }
                    else {
                        newCell.bottomFrame = sheet.cells[id].bottomFrame
                        newCell.topFrame = sheet.cells[id].topFrame
                        newCell.leftFrame = sheet.cells[id].leftFrame
                        newCell.rightFrame = sheet.cells[id].rightFrame
                    }
                    sheetcopiedCells.push(newCell)
                }
            }
            sheet.cellAttrs.forEach(function (attr) {
                sheet.setAttr(attr[0], attr[1], cells)
            })
            sheet.setAttr('allFrame', false, cells)

            var borders = [['bottomFrame', false], ['leftFrame', false], ['topFrame', false], ['rightFrame', false]]
            if (sheetcopiedCells) {
                if(action === 'deleteCol'){
                    var firstCopiedCellCol = cells.firstCellCol+1
                    var firstCopiedCellRow = cells.firstCellRow
                }else{
                    var firstCopiedCellCol = cells.firstCellCol-1
                    var firstCopiedCellRow = cells.firstCellRow
                }
                sheetcopiedCells.forEach(function (cell) {
                    var newCell = {}
                    cellCol = cell.coord.split('_')[0].charCodeAt(0)
                    cellRow = parseInt(cell.coord.split('_')[1])
                    newCell.firstCellCol = cells.firstCellCol + cellCol - firstCopiedCellCol
                    newCell.firstCellRow = cells.firstCellRow + cellRow - firstCopiedCellRow
                    newCell.lastCellCol = newCell.firstCellCol
                    newCell.lastCellRow = newCell.firstCellRow
                    sheet.cellAttrs.forEach(function (attr) {
                        sheet.setAttr(attr[0], cell[attr[0]], newCell)
                    })
                    borders.forEach(function (frame) {
                        sheet.setAttr(frame[0], cell[frame[0]], newCell)
                    })
                    delete newCell
                })
                sheet.render()
                console.log('addCol')
            }
            break
        case 'addRow':
        case 'deleteRow':
            cells = sheet.getColAndRow('A_'+sheet.firstCellid.split('_')[1],String.fromCharCode(sheet.maxCol)+'_'+sheet.maxRow)
            //sheet.firstCopiedCell = String.fromCharCode(cells.firstCellCol) + '_' + cells.firstCellRow
            var sheetcopiedCells = []
            for (var i = cells.firstCellCol; i <= cells.lastCellCol; i++) {
                for (var j = cells.firstCellRow; j <= cells.lastCellRow; j++) {
                    var id = String.fromCharCode(i) + '_' + j
                    var newCell = {}
                    newCell.coord = id
                    sheet.cellAttrs.forEach(function (attr) {
                        if (!sheet.cells[id]) {
                            newCell[attr[0]] = attr[1]
                        }
                        else {
                            newCell[attr[0]] = sheet.cells[id][attr[0]]
                        }
                    })
                    if (!sheet.cells[id]) {
                        newCell.bottomFrame = false
                        newCell.topFrame = false
                        newCell.leftFrame = false
                        newCell.rightFrame = false
                    }
                    else {
                        newCell.bottomFrame = sheet.cells[id].bottomFrame
                        newCell.topFrame = sheet.cells[id].topFrame
                        newCell.leftFrame = sheet.cells[id].leftFrame
                        newCell.rightFrame = sheet.cells[id].rightFrame
                    }
                    sheetcopiedCells.push(newCell)
                }
            }
            sheet.cellAttrs.forEach(function (attr) {
                sheet.setAttr(attr[0], attr[1], cells)
            })
            sheet.setAttr('allFrame', false, cells)

            var borders = [['bottomFrame', false], ['leftFrame', false], ['topFrame', false], ['rightFrame', false]]
            if (sheetcopiedCells) {
                if(action === 'deleteRow'){
                    var firstCopiedCellCol = cells.firstCellCol
                    var firstCopiedCellRow = cells.firstCellRow+1
                }else{
                    var firstCopiedCellCol = cells.firstCellCol
                    var firstCopiedCellRow = cells.firstCellRow-1
                }
                sheetcopiedCells.forEach(function (cell) {
                    var newCell = {}
                    cellCol = cell.coord.split('_')[0].charCodeAt(0)
                    cellRow = parseInt(cell.coord.split('_')[1])
                    newCell.firstCellCol = cells.firstCellCol + cellCol - firstCopiedCellCol
                    newCell.firstCellRow = cells.firstCellRow + cellRow - firstCopiedCellRow
                    newCell.lastCellCol = newCell.firstCellCol
                    newCell.lastCellRow = newCell.firstCellRow
                    sheet.cellAttrs.forEach(function (attr) {
                        sheet.setAttr(attr[0], cell[attr[0]], newCell)
                    })
                    borders.forEach(function (frame) {
                        sheet.setAttr(frame[0], cell[frame[0]], newCell)
                    })
                    delete newCell
                })
                sheet.render()
            }
            break
        default:
            alert('该功能未实现，请期待')
            break
    }
    if(action !== ('multiLine' || 'copy')){
        sheetEventHandler.mouseDown(sheet.firstCellid)
        sheetEventHandler.mouseUp(document.getElementById(sheet.lastCellid))
    }
}


module.exports.ToolEventHandler = ToolEventHandler

/***/ }),
/* 14 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * 工具栏渲染
 */
var config = __webpack_require__(0)
var style = __webpack_require__(2)
var ToolRender = function (sheet) {
	this.sheet = sheet
}
ToolRender.prototype.init = function (toolDiv, width, height) {
	var ToolEventBinderModule = __webpack_require__(6)
	var ToolEventBinder = ToolEventBinderModule.ToolEventBinder
	// var toolEventBinder = new ToolEventBinder(this.sheet)

	toolDiv.style.width = width
	toolDiv.style.height = height

	renderMenu(toolDiv, this.sheet)

}

function renderMenu(toolDiv, sheet) {
	var menuDiv = document.createElement('div')
	var buttonBox = document.createElement('div')

	buttonBox.style = style.ButtonBoxStyle

	toolDiv.appendChild(menuDiv)
	toolDiv.appendChild(buttonBox)

	var toolMenu = config.ToolConfig.menu

	var styleBox = document.createElement('div')
	styleBox.innerHTML = toolMenu[0]
	var toolBox = document.createElement('div')
	toolBox.innerHTML = toolMenu[1]
	var dataBox = document.createElement('div')
	dataBox.innerHTML = toolMenu[2]

	var boxes = {
		styleBox: {node: styleBox, selected: false},
		toolBox: {node: toolBox, selected: false},
		dataBox: {node: dataBox, selected: false}
	}

	menuDiv.style = style.MeunDivStyle

	changeSelected(boxes, 'styleBox')
	renderTool('styleBox', buttonBox, sheet)

	styleBox.onclick = function () {
		changeSelected(boxes, 'styleBox')
		renderTool('styleBox', buttonBox, sheet)
	}
	toolBox.onclick = function () {
		changeSelected(boxes, 'toolBox')
		renderTool('toolBox', buttonBox, sheet)
	}
	dataBox.onclick = function () {
		changeSelected(boxes, 'dataBox')
		renderTool('dataBox', buttonBox, sheet)
	}

	for (var box in boxes) {
		menuDiv.appendChild(boxes[box].node)
	}

}

function renderTool(menu, buttonBox, sheet) {

	removeAllChild(buttonBox)

	var ToolEventBinderModule = __webpack_require__(6)
	var ToolEventBinder = ToolEventBinderModule.ToolEventBinder
	var toolEventBinder = new ToolEventBinder(sheet)

	var cmdMap = config.ToolConfig.cmdCodeMap
	var toolHtml = config.ToolConfig.toolHtml
	var styleHtml = config.ToolConfig.styleHtml
	var dataHtml = config.ToolConfig.dataHtml
	var defaultHtml = config.ToolConfig.defaultHtml

	if (config.WSConfig.isPreview) {
		var previewHtml = config.ToolConfig.previewHtml

		previewHtml.forEach(function (innerhtml) {

			var buttonDiv = document.createElement('div')
			buttonDiv.id = innerhtml
			buttonDiv.style = style.ButtonDivStyle
			buttonDiv.innerHTML = innerhtml

			buttonDiv.onclick = function () {
				toolEventBinder.buttonClick(innerhtml)
			}
			buttonBox.appendChild(buttonDiv)
		})

		return
	}

	defaultHtml.forEach(function (innerhtml) {

		var buttonDiv = document.createElement('div')
		buttonDiv.id = innerhtml
		buttonDiv.style = style.ButtonDivStyle
		buttonDiv.innerHTML = cmdMap[innerhtml]

		buttonDiv.onclick = function () {
			toolEventBinder.buttonClick(innerhtml)
		}
		buttonBox.appendChild(buttonDiv)
	})

	if (menu === 'styleBox') {

		for (var i = 0; i < styleHtml.length; i++) {

			if (styleHtml[i] === 'font') {
				renderFont(toolEventBinder, buttonBox)
				continue
			}

			if (styleHtml[i] === 'align') {
				renderAlign(toolEventBinder, buttonBox)
				continue
			}

			if (styleHtml[i] === 'border') {
				renderBorder(toolEventBinder, buttonBox)
				continue
			}

			var styleButtonDiv = document.createElement('div')
			styleButtonDiv.id = styleHtml[i]
			styleButtonDiv.style = style.ButtonDivStyle
			styleButtonDiv.innerHTML = cmdMap[styleHtml[i]]

			buttonBox.appendChild(styleButtonDiv)

			if (styleHtml[i] === 'bold') {
				toolEventBinder.initFontWeight(styleButtonDiv)
				continue
			}
			if (styleHtml[i] === 'italic') {
				toolEventBinder.initFontStyle(styleButtonDiv)
				continue
			}
			if (styleHtml[i] === 'textColor') {
				renderFontColor(toolEventBinder, styleButtonDiv)
				renderColorSelect(toolEventBinder, buttonBox)
				continue
			}
			if (styleHtml[i] === 'fillColor') {
				renderBackgroundColor(toolEventBinder, styleButtonDiv)
				renderColorSelect(toolEventBinder, buttonBox)
				continue
			}


			styleButtonDiv.onclick = function (e) {
				toolEventBinder.buttonClick(e.currentTarget.id)
			}
		}
	} else if (menu === 'toolBox') {

		for (var m = 0; m < toolHtml.length; m++) {

			var toolButtonDiv = document.createElement('div')
			toolButtonDiv.id = toolHtml[m]
			toolButtonDiv.style = style.ButtonDivStyle
			toolButtonDiv.innerHTML = cmdMap[toolHtml[m]]

			buttonBox.appendChild(toolButtonDiv)

			if (toolHtml[m] === 'mergeCell') {
				toolEventBinder.initMerge(toolButtonDiv, true)
				continue
			}

			if (toolHtml[m] === 'splitCell') {
				toolEventBinder.initMerge(toolButtonDiv, false)
				continue
			}

			toolButtonDiv.onclick = function (e) {
				toolEventBinder.buttonClick(e.currentTarget.id)
			}
		}
	} else {

		for (var n = 0; n < dataHtml.length; n++) {

			var dataButtonDiv = document.createElement('div')
			dataButtonDiv.id = dataHtml[n]
			dataButtonDiv.style = style.ButtonDivStyle
			dataButtonDiv.style.fontSize = '15px'
			dataButtonDiv.innerHTML = dataHtml[n]

			dataButtonDiv.onclick = function (e) {
				console.log(e.currentTarget.id)
				toolEventBinder.buttonClick(e.currentTarget.id)
			}
			buttonBox.appendChild(dataButtonDiv)
		}
	}

}

function changeSelected(boxes, selectedNode) {

	for (var box in boxes) {
		boxes[box].selected = false
	}
	boxes[selectedNode].selected = true

	for (var b in boxes) {
		if (boxes[b].selected) {
			boxes[b].node.style = style.MeunBoxSelectedStyle
		} else {
			boxes[b].node.style = style.MeunBoxStyle
		}
	}
}

function renderFont(toolEventBinder, buttonBox) {
	//字体
	var fontFamilySelect = document.createElement('select')
	fontFamilySelect.style.marginLeft = '10px'
	var fontFamilyOption = ['fontFamily', 'Default', 'Custom', 'Verdana', 'Arial', 'Courier']
	fontFamilyOption.forEach(function (o) {
		var option = document.createElement('option')
		option.innerHTML = o
		fontFamilySelect.appendChild(option)
	})
	toolEventBinder.initFontFamily(fontFamilySelect)
	buttonBox.appendChild(fontFamilySelect)

	//字体大小
	var fontSizeSelect = document.createElement('select')
	fontSizeSelect.style.marginLeft = '10px'
	var fontSizeOption = ['fontSize', 'Default', 'X-Small', 'Small', 'Medium', 'Large',
		'X-Large', '6pt', '7pt', '8pt', '9pt', '10pt', '11pt', '12pt', '14pt', '16pt', '18pt',
		'20pt', '22pt', '24pt', '28pt', '36pt', '48pt', '72pt']
	fontSizeOption.forEach(function (o) {
		var option = document.createElement('option')
		option.innerHTML = o
		fontSizeSelect.appendChild(option)
	})
	toolEventBinder.initFontSize(fontSizeSelect)
	buttonBox.appendChild(fontSizeSelect)
}

function renderBorder(toolEventBinder, buttonBox) {
	var borderHtml = config.ToolConfig.borderHtml
	var cmdMap = config.ToolConfig.cmdCodeMap
	var borderButton = []
	var i

	for (i = 0; i < borderHtml.length; i++) {
		var borderButtonDiv = document.createElement('div')
		borderButtonDiv.id = borderHtml[i]
		borderButtonDiv.style = style.ButtonDivStyle
		borderButtonDiv.innerHTML = cmdMap[borderHtml[i]]

		buttonBox.appendChild(borderButtonDiv)
		borderButton.push(borderButtonDiv)
	}

	for (i = 0; i < borderButton.length; i++) {
		borderButton[i].onclick = function (e) {
			for (var j = 0; j < borderButton.length; j++) {
				borderButton[j].style = style.ButtonDivStyle
			}

			var fontBorderOption = config.ToolConfig.fontBorderOption
			toolEventBinder.initFontBorder(fontBorderOption[e.currentTarget.id])

			e.currentTarget.style = style.ButtonDivSelectedStyle
		}
	}
}

function renderAlign(toolEventBinder, buttonBox) {
	var alignHtml = config.ToolConfig.alignHtml
	var cmdMap = config.ToolConfig.cmdCodeMap
	var alignButton = []
	var i

	for (i = 0; i < alignHtml.length; i++) {
		var alignButtonDiv = document.createElement('div')
		alignButtonDiv.id = alignHtml[i]
		alignButtonDiv.style = style.ButtonDivStyle
		alignButtonDiv.innerHTML = cmdMap[alignHtml[i]]

		buttonBox.appendChild(alignButtonDiv)
		alignButton.push(alignButtonDiv)
	}

	for (i = 0; i < alignButton.length; i++) {
		alignButton[i].onclick = function (e) {
			for (var j = 0; j < alignButton.length; j++) {
				alignButton[j].style = style.ButtonDivStyle
			}

			var alignOption = config.ToolConfig.alignOption
			toolEventBinder.initTextAlign(alignOption[e.currentTarget.id])

			e.currentTarget.style = style.ButtonDivSelectedStyle
		}
	}
}

function renderFontColor(toolEventBinder, styleButtonDiv) {
	//color,字体颜色
	// var colorDiv = document.createElement('div')
	// colorDiv.style = style.ColorDivStyle
	// buttonBox.appendChild(colorDiv)
	toolEventBinder.initColor(styleButtonDiv)
}

function renderBackgroundColor(toolEventBinder, styleButtonDiv) {

	//backgroundColor背景颜色
	// var backgroundColorDiv = document.createElement('div')
	// backgroundColorDiv.style = style.backgroundColorDivStyle
	// buttonBox.appendChild(backgroundColorDiv)
	toolEventBinder.initBackgroundColor(styleButtonDiv)
}

function renderColorSelect(toolEventBinder, buttonBox) {
	var colorSelectDiv = document.createElement("div")
	colorSelectDiv.id = 'colorSelect'
	colorSelectDiv.style = style.ColorSelectDivStyle
	// colorSelectDiv.style.display = 'none'
	// colorSelectDiv.style.position = "absolute"
	// colorSelectDiv.style.zIndex = 100
	// colorSelectDiv.style.backgroundColor = "#FFF"
	// colorSelectDiv.style.border = "1px solid black"
	// colorSelectDiv.style.width = '106px'
	buttonBox.appendChild(colorSelectDiv)

	var mainDiv = document.createElement("div")
	// mainDiv.style.padding = "3px"
	// mainDiv.style.backgroundColor = "#CCC"
	colorSelectDiv.appendChild(mainDiv)

	var mainTable = document.createElement("table")
	mainTable.cellSpacing = 0
	mainTable.cellPadding = 0
	mainTable.style.width = "100px"
	mainDiv.appendChild(mainTable)

	var mainTbody = document.createElement("tbody")
	mainTable.appendChild(mainTbody)
	var steps = [0, 68, 153, 204, 255]
	var commonRgb = ["400", "310", "420", "440", "442", "340", "040", "042", "032", "044", "024", "004", "204", "314", "402", "414"]

	for (var row = 0; row < 16; row++) {
		var tr = document.createElement("tr")
		for (var col = 0; col < 5; col++) {
			var td = document.createElement("td")
			toolEventBinder.initColorSelect(td)
			td.style.fontSize = "1px"
			td.innerHTML = "&nbsp;"
			td.style.height = "10px"
			td.style.lineHeight = '12px'
			if (col <= 1) {
				td.style.width = "17px"
				td.style.borderRight = "3px solid white"
			}
			else {
				td.style.width = "20px"
				td.style.backgroundRepeat = "no-repeat"
			}
			if (col === 0) {
				var x = commonRgb[row]
				td.style.backgroundColor = "rgb(" + steps[x.charAt(0) - 0] + "," + steps[x.charAt(1) - 0] + "," + steps[x.charAt(2) - 0] + ")"
			}
			if (col === 1) {
				td.style.backgroundColor = makeRGB(17 * (15 - row), 17 * (15 - row), 17 * (15 - row))
			}
			if (col === 2) {
				td.style.backgroundColor = makeRGB(17 * (15 - row), 0, 0)
			}
			if (col === 3) {
				td.style.backgroundColor = makeRGB(0, 17 * (15 - row), 0)
			}
			if (col === 4) {
				td.style.backgroundColor = makeRGB(0, 0, 17 * (15 - row))
			}
			tr.appendChild(td)
		}
		mainTbody.appendChild(tr)
	}
	mainTable.appendChild(mainTbody)
}

function makeRGB(r, g, b) {
	return "rgb(" + (r > 0 ? r : 0) + "," + (g > 0 ? g : 0) + "," + (b > 0 ? b : 0) + ")"
}

function removeAllChild(node) {

	while (node.hasChildNodes()) {
		node.removeChild(node.firstChild)
	}
}

module.exports.ToolRender = ToolRender

/***/ }),
/* 15 */
/***/ (function(module, exports) {

/**
 *  堆栈对象
 */
var UndoStack = function () {

    this.tos = -1
    //this.maxRedo = 0
    this.maxUndo = 50
    this.stack = []

}
UndoStack.prototype.pushStep = function (stack) {
    while(this.stack.length > this.tos+1) this.stack.pop()
    this.stack.push(stack)
    if(this.stack.length > this.maxUndo) this.stack.shift()
    else this.tos++
}
UndoStack.prototype.reDo = function () {
    if (this.tos < this.stack.length-1){
        this.tos ++
        var result = this.stack[this.tos]
    }
    else result = null
    return result
}
UndoStack.prototype.unDo = function () {
    if (this.tos > -1){
        var result = this.stack[this.tos]
        this.tos --
    }
    else result = null
    return result
}
module.exports.UndoStack = UndoStack

/***/ }),
/* 16 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * 工作区域渲染
 * @param wsManager
 * @param parentNode
 * @constructor
 */


var WSRender = function (wsManager, parentNode) {
	this.manager = wsManager
	this.parNode = parentNode
}

WSRender.prototype.init = function () {
	var ws = this.manager.workspace

	//workspace
	var WSDiv = document.createElement('div')
	// WSDiv.style.width = ws.width + 'px'
	// WSDiv.style.height = ws.height + 'px'
	WSDiv.style = __webpack_require__(2).WSStyle
	this.parNode.appendChild(WSDiv)

	//tool
	var ToolRenderModule = __webpack_require__(14)
	var ToolRender = ToolRenderModule.ToolRender
	var toolRender = new ToolRender(ws.Sheet)
	var tool = ws.Tool
	var toolDiv = document.createElement('div')
	toolRender.init(toolDiv, tool.width, tool.height)
	WSDiv.appendChild(toolDiv)

	//sheet
	var SheetRenderModule = __webpack_require__(5)
	var SheetRender = SheetRenderModule.SheetRender
	var sheetRender = new SheetRender(ws.Sheet)
	var sheet = ws.Sheet
	var sheetDiv = document.createElement('div')
	sheetRender.init(sheetDiv)
	WSDiv.appendChild(sheetDiv)

	//slideBar
	// var SliderBarRenderModule = require('SliderBarRender')
	// var SliderBarRender = SliderBarRenderModule.SliderBarRender
	// var sliderBarRender = new SliderBarRender(ws.Sheet)
	// var sliderBar = ws.SliderBar
	// var sliderBarDiv = document.createElement('div')
	// sliderBarRender.init(sliderBarDiv)
	// WSDiv.appendChild(sliderBarDiv)

}
module.exports.WSRender = WSRender

/***/ }),
/* 17 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * 工作区对象，包括工具栏，表格等
 * @type {*}
 */
var config= __webpack_require__(0)

var Workspace = function (Tool, Sheet, SliderBar) {

    this.width = config.WSConfig.width
    this.height = config.WSConfig.height

    this.Tool =Tool
    this.Sheet= Sheet
    this.SliderBar = SliderBar
}

module.exports.Workspace = Workspace

/***/ })
/******/ ]);
//# sourceMappingURL=app.js.map