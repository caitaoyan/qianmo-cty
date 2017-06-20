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
/******/ 	return __webpack_require__(__webpack_require__.s = 3);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ (function(module, exports) {

/**
 * 配置文件，主要是个各类属性默认值
 * 整张表命名为：sheet
 * 表格上半部分叫做：sheetTool
 * 表格叫做：sheetGrid
 *
 */

/**
 * 整张表-Sheet的属性默认值
 *
 */
var WSConfig = {
	height: 600,
	width: 1030,
	isMouseDown:false,
	isKeyDown:true,
	isEditing:false,
	editCell:null,
	isFont:false
}

/**
 * 表头-SheetTool的属性默认值
 *
 */
var ToolConfig = {
	height: 40,
	width: WSConfig.width
}

/**
 * 单元格-Cell属性默认值
 *
 */
var CellConfig = {
	height: 20,
	width: 100
}

/**
 * 表单-SheetGrid的属性默认值
 *
 */
var SheetConfig = {
	height: WSConfig.height - ToolConfig.height,
	width: WSConfig.width,
	headWidth: 30,
	headHeight: 25,

	rowNum: (WSConfig.height - ToolConfig.height-30)/CellConfig.height,
	colNum: (WSConfig.width-30)/CellConfig.width
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
 * 输出模块
 *
 */
module.exports.WSConfig = WSConfig
module.exports.ToolConfig = ToolConfig
module.exports.SheetConfig = SheetConfig
module.exports.CellConfig = CellConfig


/***/ }),
/* 1 */
/***/ (function(module, exports) {

var Util=function(){

}
/**
 * 将第一个单元格和最后一个单元格进行排序
 * @param firstCell
 * @param lastCell
 * @returns {*}
 */
Util.prototype.getColAndRow=function(firstCell,lastCell){
    var result=null
    if(firstCell&&firstCell.id&&lastCell&&lastCell.id) {
        result=new Object()
        result.firstCellCol = firstCell.id.split('_')[0].charCodeAt(0)
        result.firstCellRow = parseInt(firstCell.id.split('_')[1])
        result.lastCellCol = lastCell.id.split('_')[0].charCodeAt(0)
        result.lastCellRow = parseInt(lastCell.id.split('_')[1])
        if( result.firstCellCol> result.lastCellCol){
            var newFirstCellCol=result.firstCellCol
            result.firstCellCol=result.lastCellCol
            result.lastCellCol=newFirstCellCol
        }
        if( result.firstCellRow> result.lastCellRow){
            var newFirstCellRow=result.firstCellRow
            result.firstCellRow=result.lastCellRow
            result.lastCellRow=newFirstCellRow
        }

    }
    return result
}

module.exports.Util=Util

/***/ }),
/* 2 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * 整体工作管理对象
 * @constructor
 */
var WSManager = function () {}

WSManager.prototype.init = function (parentNode) {

    //实例化初始化Tool对象
    var ToolModule = __webpack_require__(9)
    var Tool=ToolModule.Tool
    var tool=new Tool()

    //实例化初始化UndoStack对象
    var UndoStackModule =__webpack_require__(13)
    var UndoStack = UndoStackModule.UndoStack
    var undoStack = new UndoStack()

    //实例化初始化Sheet对象
    var SheetModule = __webpack_require__(5)
    var Sheet =SheetModule.Sheet
    var sheet=new Sheet(undoStack)

    //实例化初始化Workspace对象
    var WorkspaceModule = __webpack_require__(15)
    var Workspace = WorkspaceModule.Workspace
    this.workspace=new Workspace(tool,sheet)

    //实例化初始化WSRender对象
    var WSRenderModule = __webpack_require__(14)
    var WSRender = WSRenderModule.WSRender
    var wsRender=new WSRender(this,parentNode)

    wsRender.init()
}

module.exports.WSManager = WSManager

/***/ }),
/* 3 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * Created by Ian on 17/6/5.
 */

var WSManagerModule = __webpack_require__(2)
var WSManager = WSManagerModule.WSManager
var wsManager=new WSManager()

var parentNode = document.getElementById('QianMoApp')

wsManager.init(parentNode)


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
	this.height = cellConfig.height
	this.width = cellConfig.width

	this.text = null
	this.coord = coord
}

module.exports.Cell = Cell

/***/ }),
/* 5 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * 表格对象
 */


var configModule = __webpack_require__(0)

var sheetConfig = configModule.SheetConfig

var Sheet = function (UndoStack) {
    this.height = sheetConfig.height
    this.width = sheetConfig.width

    this.headWidth = sheetConfig.headWidth
    this.headHeight = sheetConfig.headHeight

    this.rowNum = sheetConfig.rowNum
    this.colNum = sheetConfig.colNum

    this.cells = {}
    this.UndoStack = UndoStack
}

//Sheet.prototype.addCell = function(newcell) {return this.cells[newcell.coord]=newcell}

module.exports.Sheet = Sheet

/***/ }),
/* 6 */
/***/ (function(module, exports, __webpack_require__) {

//事件绑定对象
var config = __webpack_require__(0)
var SheetEventHandlerModule=__webpack_require__(7)
var SheetEventHandler = SheetEventHandlerModule.SheetEventHandler
var sheetEventHandler = null

var SheetEventBinder = function (sheet) {
    this.sheet=sheet
    sheetEventHandler=new SheetEventHandler(sheet)
    window.onmousedown=function(){
        document.getElementById('colorSelect').style.display='none'
    }
    window.onmouseup = function(){
        config.WSConfig.isMouseDown=false
    }
    document.onkeydown = function (event) {

        if(sheetEventHandler.firstCell){
            sheetEventHandler.keyDown(sheetEventHandler.firstCell,event)
        }
    }
}

SheetEventBinder.prototype.initRowTD = function(rowTD){

    rowTD.onmousedown = function(){
        sheetEventHandler.mouseDown(this)
    }
    rowTD.onmouseup = function(){
        sheetEventHandler.mouseUp(this)
    }
    rowTD.onmousemove = function(){
        if(config.WSConfig.isMouseDown) {
            sheetEventHandler.mouseMove(this)
        }
    }
    rowTD.ondblclick = function() {
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
module.exports.SheetEventBinder=SheetEventBinder

/***/ }),
/* 7 */
/***/ (function(module, exports, __webpack_require__) {

//鼠标和键盘事件
var config = __webpack_require__(0)
var CellModule = __webpack_require__(4)
var Cell=CellModule.Cell
var UtilModule=__webpack_require__(1)
var Util=UtilModule.Util
var util=new Util()

var SheetEventHandler = function (sheet) {
    this.firstCell = null
    this.lastCell = null
    this.sheet=sheet
}
/**
 * 鼠标按下事件
 * @param element
 */
SheetEventHandler.prototype.mouseDown = function(element){
    this.setCellBackgroundColor('#fff')
    if(element.id){
        this.firstCell = element
        config.WSConfig.isMouseDown=true
    }
}
/**
 * 鼠标移动事件
 * @param element
 */
SheetEventHandler.prototype.mouseMove = function(element){
    this.setCellBackgroundColor('#fff')
    if(element.id){
        this.lastCell=element
        this.setCellBackgroundColor('#69f')
    }
}
/**
 * 鼠标抬起事件
 * @param element
 */
SheetEventHandler.prototype.mouseUp = function(element){
    if(element.id){
        this.lastCell = element
        this.sheet.range=this.firstCell.id+':'+this.lastCell.id
        this.setCellBackgroundColor('#69f')
        config.editCell = element
    }
    config.WSConfig.isMouseDown=false
}
/**
 * 鼠标双击事件
 * @param element
 */
SheetEventHandler.prototype.dblclick = function(element){
    if(element.id){
        config.editCell = element
        var input = document.getElementById('input')
        input.style.display = 'block'
        input.style.backgroundColor = '#efe'
        input.style.left = getLeft(element) + 'px'
        input.style.top = getTop(element) + 'px'
        input.style.height = element.offsetHeight + 'px'
        input.style.width = element.offsetWidth + 'px'
        if(this.sheet.cells[element.id]){
            input.value = this.sheet.cells[element.id].text
        }
        else{
            input.value = ''
        }
        input.focus()
    }
    config.WSConfig.isMouseDown=false
}
SheetEventHandler.prototype.inputBlur = function () {
    config.isEditing = false
    var input = document.getElementById('input')
    input.blur()
    input.style.display = 'none'
    var ele = config.editCell
    if(input.value){
        if(!this.sheet.cells[ele.id]) {
            this.sheet.cells[ele.id] = new Cell(ele)
        }
        this.sheet.cells[ele.id].text = input.value
        ele.firstChild.innerHTML = input.value
    }
}
SheetEventHandler.prototype.inputFocus = function () {
    config.isEditing = true
    var input = document.getElementById('input')
    input.focus()
}
//鼠标失去焦点事件
SheetEventHandler.prototype.mouseBlur=function(element){
    element.style.color='transparent'
    element.style.background='transparent'
    if(config.WSConfig.isKeyDown){
        var value=element.value
        if(element.nextSibling.innerHTML==''&&value!=''){
            var command = 'set '+ element.id + ' ' + value
            var type = 'set'
            this.sheet.UndoStack.pushCommand(command,type)
        }
        element.nextSibling.innerHTML=value
    }
    config.WSConfig.isKeyDown=true
    element.parentNode.style.backgroundColor = 'transparent'
}
/**
 * 设置单元格背景颜色
 * @param backgroundColor
 */
SheetEventHandler.prototype.setCellBackgroundColor=function(backgroundColor){
    var cells=util.getColAndRow(this.firstCell,this.lastCell)
    if(cells!=null) {
        for(var i=cells.firstCellCol;i<=cells.lastCellCol;i++){
            for(var j=cells.firstCellRow;j<=cells.lastCellRow;j++){
                var ele = document.getElementById(String.fromCharCode(i) + '_' + j)
                if(backgroundColor=='#fff'&&ele.backBackgroundColor!=undefined){
                    ele.style.backgroundColor = ele.backBackgroundColor
                }else{
                    ele.style.backgroundColor = backgroundColor
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
SheetEventHandler.prototype.keyDown = function(element,event){
    var col=element.id.split('_')[0]
    var row=element.id.split('_')[1]

    switch (event.which) {
        case 37://左键
            if(!config.isEditing){
                col=String.fromCharCode(col.charCodeAt(0) - 1)
                var cell = document.getElementById( col +'_'+ row )
                if (cell) {
                    this.mouseDown(cell)
                    this.mouseUp(cell)
                }
            }
            break
        case 38://上键
            var cell = document.getElementById( col +'_'+ (parseInt(row)-1) )
            if (cell) {
                if(config.isEditing){
                    config.isEditing = false
                    document.getElementById('input').blur()
                }
                this.mouseDown(cell)
                this.mouseUp(cell)
            }
            break
        case 39://右键
            if(!config.isEditing){
                col=String.fromCharCode(col.charCodeAt(0) + 1)
                var cell = document.getElementById( col +'_'+ row)
                if (cell) {
                    this.mouseDown(cell)
                    this.mouseUp(cell)
                }
            }
            break
        case 40://下键
            var cell = document.getElementById( col +'_'+ (parseInt(row)+1) )
            if (cell) {
                if(config.isEditing){
                    config.isEditing = false
                    document.getElementById('input').blur()
                }
                this.mouseDown(cell)
                this.mouseUp(cell)
            }
            break
        case 17://ctrl
            break
        case 18://alt
            break
        case 67://ctrl+c
            this.sheet.copiedfrom=this.sheet.range
            var command = 'copy ' + this.sheet.copiedfrom + ' formulas'
            var type = 'copy'
            this.sheet.UndoStack.pushCommand(command,type)
            ta = document.getElementById('ta')
            var text = ''
            var firstCell=this.sheet.copiedfrom.split(':')[0]
            var lastCell=this.sheet.copiedfrom.split(':')[1]
            var cells=util.getColAndRow(document.getElementById(firstCell),
                document.getElementById(lastCell))
            if(cells!=null) {
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
                    text=text.substr(0,text.length-1)
                    text += '\n'
                    rowCount++
                }
                text=text.substr(0,text.length-1)
            }
            ta.value = text
            ta.style.display = 'block'
            ta.focus()
            ta.select()

            window.setTimeout(function() {
                ta.style.display = 'none'
                var cell = document.getElementById(col + '_' + row)
                cell.focus()
            },200)

            break
        case 86://ctrl+v
            ta = document.getElementById('ta')
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
                        var cvCol = String.fromCharCode(col.charCodeAt(0) + j)
                        var cvRow = parseInt(row) + i
                        var cell = document.getElementById(cvCol + '_' + cvRow)
                        if(cell){
                            cell.firstChild.innerHTML = vv[j]
                            if(!sheet.cells[cell.id]) {
                                sheet.cells[cell.id] = new Cell(cell)
                            }
                            sheet.cells[cell.id].text = vv[j]
                            //还需要给cell height和width
                        }

                    }
                }
                var cell = document.getElementById(col + '_' + row)
                cell.focus()
            }, 200)

            var command = 'past ' + element.id + ' formulas'
            var type = 'past'
            this.sheet.UndoStack.pushCommand(command,type)
            config.WSConfig.isKeyDown=false
            break
        default:
            if(!config.isEditing){
                if(config.editCell) this.dblclick(config.editCell)
            }
    }
}

//获取元素的纵坐标
function getTop(e){
    var offset=e.offsetTop;
    if(e.offsetParent!=null) offset+=getTop(e.offsetParent);
    return offset;
}
//获取元素的横坐标
function getLeft(e){
    var offset=e.offsetLeft
    if(e.offsetParent!=null) offset+=getLeft(e.offsetParent)
    return offset
}
module.exports.SheetEventHandler = SheetEventHandler

/***/ }),
/* 8 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * 渲染表格
 */

var config = __webpack_require__(0)

var SheetRender = function (sheet) {
	this.sheet=sheet
}

SheetRender.prototype.init = function (sheetDiv) {

	var sheetTable = document.createElement('table')
    sheetTable.style.userSelect = 'none'
    sheetDiv.appendChild(sheetTable)
	sheetTable.style.width = this.sheet.width + 'px'
	sheetTable.style.height = this.sheet.height + 'px'
	//表格的具体渲染
	renderSheet(this.sheet, sheetTable)
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
	var cellHeight = config.CellConfig.height
	var cellWidth = config.CellConfig.width
	var	headWidth = config.SheetConfig.headWidth
	var	headHeight = config.SheetConfig.headHeight

	var rowNum = gridHeader.length
	var colNum = gridHeader[0].length


	var rowHead = document.createElement('tr')
	sheetTable.appendChild(rowHead)
	//渲染第一行表头
	for(var i=0;i<colNum;i++){
		var rowHeadTH  = document.createElement('th')
		rowHead.appendChild(rowHeadTH)

		var rowHeadDiv = document.createElement('div')
		rowHeadTH.appendChild(rowHeadDiv)
		//第一行第一列特殊处理
		if(i === 0){
			rowHeadTH.style.minWidth = headWidth + 'px'
			rowHeadTH.style.minHeight = headHeight + 'px'
			rowHeadDiv.style.minWidth = headWidth-2 + 'px'
			rowHeadDiv.style.minHeight = headHeight-2 + 'px'
		}else{
			rowHeadTH.style.minHeight = headHeight + 'px'
			rowHeadTH.style.minWidth = cellWidth + 'px'
			rowHeadDiv.style.minHeight = headHeight-2 + 'px'
			rowHeadDiv.style.minWidth = cellWidth-2 + 'px'

			rowHeadDiv.innerHTML = gridHeader[0][i]
		}

	}

    var SheetEventBinderModule = __webpack_require__(6)
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
				rowTR.appendChild(rowTH)
                rowTH.style.minWidth = headWidth + 'px'
                rowTH.style.minHeight = cellHeight + 'px'

				var rowDiv = document.createElement('div')
				rowTH.appendChild(rowDiv)
                rowDiv.style.minWidth = headWidth-2 + 'px'
                rowDiv.style.minHeight = cellHeight-2 + 'px'
                rowDiv.innerHTML = gridHeader[j][k]
			}else {
				var rowTD = document.createElement('td')
                rowTD.id = gridHeader[0][k]+"_"+j
                rowTR.appendChild(rowTD)
                rowTD.style.minWidth = cellWidth + 'px'
                rowTD.style.minHeight = cellHeight + 'px'
                rowTD.style.display = 'table-cell'
                sheetEventBinder.initRowTD(rowTD)
				var rowDiv = document.createElement('div')
                rowTD.appendChild(rowDiv)
                rowDiv.style.minWidth = cellWidth-2 + 'px'
                rowDiv.style.minHeight = cellHeight-2 + 'px'
                rowDiv.style.textAlign = 'right'
			}
		}
	}
	var input = document.createElement('input')
	input.style.display = 'none'
    input.style.position = 'absolute'
	input.id = 'input'
    sheetEventBinder.initInput(input)
	sheetTable.appendChild(input)
    var ta = document.createElement('textarea') // used for ctrl-c/ctrl-v where an invisible text area is needed
    ta.style = 'display:none;position:absolute;height:1px;width:1px;opacity:0;filter:alpha(opacity=0);'
    ta.value = '';
    ta.id = 'ta'
    sheetTable.appendChild(ta)
    //sheetEventBinder.init()
}

module.exports.SheetRender = SheetRender

/***/ }),
/* 9 */
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
/* 10 */
/***/ (function(module, exports, __webpack_require__) {

/**
 工具栏绑定事件
 */
var config = __webpack_require__(0)
var ToolEventHandlerModule=__webpack_require__(11)
var ToolEventHandler = ToolEventHandlerModule.ToolEventHandler
var toolEventHandler=null

var ToolEventBinder=function(sheet){
    this.sheet=sheet
    toolEventHandler=new ToolEventHandler(sheet)
}
ToolEventBinder.prototype.initFontFamily=function(fontFamilySelect){
    var toolEventBinder=this
    fontFamilySelect.onchange=function(){
        var fontFamily=fontFamilySelect.value
        toolEventBinder.fontValue=fontFamily
        toolEventBinder.fontType='family'
        toolEventHandler.setFont(toolEventBinder)
    }
}

ToolEventBinder.prototype.initFontSize=function(fontSizeSelect){
    var toolEventBinder=this
    fontSizeSelect.onchange=function(){
        var fontSize=fontSizeSelect.value
        toolEventBinder.fontValue=fontSize
        toolEventBinder.fontType='size'
        toolEventHandler.setFont(toolEventBinder)
    }
}

ToolEventBinder.prototype.initFontWeight=function(fontWeightDiv){
    var toolEventBinder=this
    fontWeightDiv.onclick=function(){
        config.WSConfig.isFont=false
        toolEventBinder.fontValue='bold'
        toolEventBinder.fontType='weight'
        toolEventHandler.setFont(toolEventBinder)
        if(config.WSConfig.isFont){
            toolEventHandler.setFont(toolEventBinder)
        }
    }
}

ToolEventBinder.prototype.initFontStyle=function(fontStyleDiv){
    var toolEventBinder=this
    fontStyleDiv.onclick=function(){
        config.WSConfig.isFont=false
        toolEventBinder.fontValue='italic'
        toolEventBinder.fontType='style'
        toolEventHandler.setFont(toolEventBinder)
        if(config.WSConfig.isFont){
            toolEventHandler.setFont(toolEventBinder)
        }
    }
}

ToolEventBinder.prototype.initColor=function(colorDiv){
    var toolEventBinder=this
    colorDiv.onclick=function(){
        var colorSelectDiv=document.getElementById('colorSelect')
        if(colorSelectDiv.style.display=='none'||toolEventBinder.fontType=='backgroundColor'){
            toolEventBinder.fontType='color'
            colorSelectDiv.style.display='block'
            var left=this.offsetLeft
            var top=this.offsetTop
            colorSelectDiv.style.top=top+18+'px'
            colorSelectDiv.style.left=left+'px'
        }else{
            colorSelectDiv.style.display='none'
        }

    }
}

ToolEventBinder.prototype.initBackgroundColor=function(backgroundColorDiv){
    var toolEventBinder=this
    backgroundColorDiv.onclick=function(){
        var colorSelectDiv=document.getElementById('colorSelect')
        if(colorSelectDiv.style.display=='none'||toolEventBinder.fontType=='color'){
            toolEventBinder.fontType='backgroundColor'
            colorSelectDiv.style.display='block'
            var left=this.offsetLeft
            var top=this.offsetTop
            colorSelectDiv.style.top=top+18+'px'
            colorSelectDiv.style.left=left+'px'
        }else{
            colorSelectDiv.style.display='none'
        }

    }
}

ToolEventBinder.prototype.initColorSelect=function(td){
    var toolEventBinder=this
    td.onclick=function(){
        toolEventBinder.fontValue=this.style.backgroundColor
        document.getElementById('colorSelect').style.display='none'
        toolEventHandler.setFont(toolEventBinder)
    }
}
module.exports.ToolEventBinder=ToolEventBinder

/***/ }),
/* 11 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * 工具栏事件
 */
var config = __webpack_require__(0)
var UtilModule=__webpack_require__(1)
var Util=UtilModule.Util
var util=new Util()

var ToolEventHandler=function(sheet){
    this.sheet=sheet
}

ToolEventHandler.prototype.setFont=function(toolEventBinder){
    var sheet=this.sheet
    var fontType=toolEventBinder.fontType
    var fontValue=toolEventBinder.fontValue
    if(sheet.range){
        var firstCell=sheet.range.split(':')[0]
        var lastCell=sheet.range.split(':')[1]
        var cells=util.getColAndRow(document.getElementById(firstCell),
            document.getElementById(lastCell))
        if(cells!=null) {
            for(var i=cells.firstCellCol;i<=cells.lastCellCol;i++){
                for(var j=cells.firstCellRow;j<=cells.lastCellRow;j++){
                    var ele = document.getElementById(String.fromCharCode(i) + '_' + j).firstChild
                    if(fontType=='family'){
                        setFontFamily(ele,fontValue)
                    }else if(fontType=='size'){
                        setFontSize(ele,fontValue)
                    }else if(fontType=='weight'){
                        setFontWeight(ele,fontValue)
                    }else if(fontType=='style'){
                        setFontStyle(ele,fontValue)
                    }else if(fontType=='color'){
                        setFontColor(ele,fontValue)
                    }else if(fontType=='backgroundColor'){
                        setFontBackgroundColor(ele,fontValue)
                    }
                }
            }
        }
    }
}
function setFontFamily(ele,fontFamily) {
    if(fontFamily=='Default'){
        fontFamily=''
    }
    ele.style.fontFamily=fontFamily
}
function setFontSize(ele,fontSize){
    if(fontSize=='Default'||fontSize=='-----'){
        fontSize=''
    }
    ele.style.fontSize=fontSize
}
function setFontWeight(ele,fontWeight){
    if(!config.WSConfig.isFont){
        if(ele.style.fontWeight!=''){
            fontWeight=''
        }else{
            config.WSConfig.isFont=true
        }
    }

    ele.style.fontWeight=fontWeight
}
function setFontStyle(ele,fontStyle){
    if(!config.WSConfig.isFont){
        if(ele.style.fontStyle!=''){
            fontStyle=''
        }else{
            config.WSConfig.isFont=true
        }
    }

    ele.style.fontStyle=fontStyle
}

function setFontColor(ele,fontColor){
    ele.backBackgroundColor=fontColor
    ele.parentNode.style.color=fontColor
}
function setFontBackgroundColor(ele,fontBackgroundColor){
    ele.parentNode.backBackgroundColor=fontBackgroundColor
    ele.parentNode.style.backgroundColor=fontBackgroundColor
}
module.exports.ToolEventHandler = ToolEventHandler

/***/ }),
/* 12 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * 工具栏渲染
 */

var ToolRender=function(sheet){
    this.sheet=sheet
}
ToolRender.prototype.init=function(toolDiv,width,height){
    toolDiv.style.width = width + 'px'
    toolDiv.style.height = height + 'px'
    toolDiv.style.backgroundColor = '#aaaaaa'
    //sheetToolDiv.innerHTML = "&#xe900;&#xe901;&#xe14d;&#xe14e;&#xe14f;"
    var html = ["&#xe900","&#xe901","&#xe14d","&#xe14e","&#xe14f"]
    //渲染工具栏具体button按钮
    html.forEach(function(innerhtml){
        var buttonDiv = document.createElement('div')
        buttonDiv.style.display = "inline"
        buttonDiv.innerHTML=innerhtml

        buttonDiv.onclick = function(){
            // eventHandle.buttonClick(innerhtml)
        }
        toolDiv.appendChild(buttonDiv)
    })

    var ToolEventBinderModule = __webpack_require__(10)
    var ToolEventBinder=ToolEventBinderModule.ToolEventBinder
    var toolEventBinder=new ToolEventBinder(this.sheet)

    renderFont(toolEventBinder,toolDiv)
    //color,字体颜色
    var colorDiv=document.createElement('div')
    colorDiv.style.display='inline'
    colorDiv.style.marginLeft='10px'
    colorDiv.style.paddingLeft='15px'
    colorDiv.style.cursor='pointer'
    colorDiv.style.border='1px solid black'
    colorDiv.style.width='15px'
    colorDiv.style.height='15px'
    colorDiv.style.backgroundColor='#fff'
    toolDiv.appendChild(colorDiv)
    toolEventBinder.initColor(colorDiv)

    //backgroundColor背景颜色
    var backgroundColorDiv=document.createElement('div')
    backgroundColorDiv.style.display='inline'
    backgroundColorDiv.style.marginLeft='10px'
    backgroundColorDiv.style.paddingLeft='15px'
    backgroundColorDiv.style.cursor='pointer'
    backgroundColorDiv.style.border='1px solid black'
    backgroundColorDiv.style.width='15px'
    backgroundColorDiv.style.height='15px'
    backgroundColorDiv.style.backgroundColor='#fff'
    toolDiv.appendChild(backgroundColorDiv)
    toolEventBinder.initBackgroundColor(backgroundColorDiv)
    renderColorSelect(toolEventBinder,toolDiv)
}

function renderFont(toolEventBinder,toolDiv){
    //字体
    var fontFamilySelect=document.createElement('select')
    var fontFamilyOption=['-----','Default','Custom','Verdana','Arial','Courier']
    fontFamilyOption.forEach(function(o){
        var option=document.createElement('option')
        option.innerHTML=o
        fontFamilySelect.appendChild(option)
    })
    toolEventBinder.initFontFamily(fontFamilySelect)
    toolDiv.appendChild(fontFamilySelect)

    //字体大小
    var fontSizeSelect=document.createElement('select')
    var fontSizeOption=['-----','Default','X-Small','Small','Medium','Large',
        'X-Large','6pt','7pt','8pt','9pt','10pt','11pt','12pt','14pt','16pt','18pt',
        '20pt','22pt','24pt','28pt','36pt','48pt','72pt']
    fontSizeOption.forEach(function(o){
        var option=document.createElement('option')
        option.innerHTML=o
        fontSizeSelect.appendChild(option)
    })
    toolEventBinder.initFontSize(fontSizeSelect)
    toolDiv.appendChild(fontSizeSelect)

    //加粗
    var fontWeightDiv=document.createElement('div')
    fontWeightDiv.innerHTML='B'
    fontWeightDiv.style.display='inline'
    fontWeightDiv.style.paddingLeft='10px'
    fontWeightDiv.style.cursor='pointer'
    fontWeightDiv.style.fontWeight='bold'
    toolEventBinder.initFontWeight(fontWeightDiv)
    toolDiv.appendChild(fontWeightDiv)

    //斜体
    var fontStyleDiv=document.createElement('div')
    fontStyleDiv.innerHTML='I'
    fontStyleDiv.style.display='inline'
    fontStyleDiv.style.paddingLeft='10px'
    fontStyleDiv.style.cursor='pointer'
    fontStyleDiv.style.fontStyle='italic'
    fontStyleDiv.style.fontFamily='Verdana'
    toolEventBinder.initFontStyle(fontStyleDiv)
    toolDiv.appendChild(fontStyleDiv)
}
function renderColorSelect(toolEventBinder,toolDiv){
    var colorSelectDiv=document.createElement("div")
    colorSelectDiv.id='colorSelect'
    colorSelectDiv.style.display='none'
    colorSelectDiv.style.position = "absolute"
    colorSelectDiv.style.zIndex = 100
    colorSelectDiv.style.backgroundColor = "#FFF"
    colorSelectDiv.style.border = "1px solid black"
    colorSelectDiv.style.width='106px'
    toolDiv.appendChild(colorSelectDiv)
    //
    // var titleTable=document.createElement('table')
    // titleTable.cellspacing='0'
    // titleTable.cellpadding='0'
    // titleTable.style.borderBottom='1px solid black'
    // colorSelectDiv.appendChild(titleTable)
    //
    // var td=document.createElement('td')
    // td.style.fontSize='10px'
    // td.style.cursor='default'
    // td.style.width='100%'
    // td.style.backgroundColor='#999'
    // td.style.color='#fff'
    // titleTable.appendChild(td)
    //
    // td=document.createElement('td')
    // td.style.fontSize='10px'
    // td.style.cursor='default'
    // td.style.color='#666'
    // titleTable.appendChild(td)

    var mainDiv = document.createElement("div")
    mainDiv.style.padding = "3px"
    mainDiv.style.backgroundColor = "#CCC"
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

    for (var row=0; row<16; row++) {
        var tr = document.createElement("tr")
        for (var col=0; col<5; col++) {
            var td = document.createElement("td")
            toolEventBinder.initColorSelect(td)
            td.style.fontSize = "1px"
            td.innerHTML = "&nbsp;"
            td.style.height = "10px"
            if (col<=1) {
                td.style.width = "17px"
                td.style.borderRight = "3px solid white"
            }
            else {
                td.style.width = "20px"
                td.style.backgroundRepeat = "no-repeat"
            }
            if(col==0){
                var x = commonRgb[row]
                td.style.backgroundColor="rgb("+steps[x.charAt(0)-0]+","+steps[x.charAt(1)-0]+","+steps[x.charAt(2)-0]+")"
            }
            if(col==1){
                td.style.backgroundColor=makeRGB(17*(15-row),17*(15-row),17*(15-row))
            }
            if(col==2){
                td.style.backgroundColor=makeRGB(17*(15-row),0,0)
            }
            if(col==3){
                td.style.backgroundColor=makeRGB(0,17*(15-row),0)
            }
            if(col==4){
                td.style.backgroundColor=makeRGB(0,0,17*(15-row))
            }
            tr.appendChild(td)
        }
        mainTbody.appendChild(tr)
    }
    mainTable.appendChild(mainTbody)
}

function makeRGB(r, g, b) {
    return "rgb("+(r>0?r:0)+","+(g>0?g:0)+","+(b>0?b:0)+")"
}
module.exports.ToolRender = ToolRender

/***/ }),
/* 13 */
/***/ (function(module, exports) {

/**
 *  堆栈对象
 */
var UndoStack = function () {

    this.tos=-1
    this.maxRedo=0
    this.maxUndo=50
    this.stack=[]

}
UndoStack.prototype.pushCommand = function(command,type){
    var stack={
        command : command,
        type: type
    }
    this.stack.push(stack)
    this.tos++
}
UndoStack.prototype.reDo = function(){

}
UndoStack.prototype.unDo = function(){

}
module.exports.UndoStack = UndoStack

/***/ }),
/* 14 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * 工作区域渲染
 * @param wsManager
 * @param parentNode
 * @constructor
 */
var WSRender=function(wsManager,parentNode){
    this.manager=wsManager
    this.parNode=parentNode
}

WSRender.prototype.init=function () {
    var ws=this.manager.workspace

    //workspace
    var WSDiv = document.createElement('div')
    WSDiv.style.width = ws.width + 'px'
    WSDiv.style.height = ws.height + 'px'
    this.parNode.appendChild(WSDiv)

    //tool
    var ToolRenderModule = __webpack_require__(12)
    var ToolRender = ToolRenderModule.ToolRender
    var toolRender=new ToolRender(ws.Sheet)
    var tool = ws.Tool
    var toolDiv = document.createElement('div')
    toolRender.init(toolDiv,tool.width,tool.height)
    WSDiv.appendChild(toolDiv)

    //sheet
    var SheetRenderModule= __webpack_require__(8)
    var SheetRender= SheetRenderModule.SheetRender
    var sheetRender=new SheetRender(ws.Sheet)
    var sheet=ws.Sheet
    var sheetDiv=document.createElement('div')
    sheetRender.init(sheetDiv)
    WSDiv.appendChild(sheetDiv)
}
module.exports.WSRender = WSRender

/***/ }),
/* 15 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * 工作区对象，包括工具栏，表格等
 * @type {*}
 */
var config= __webpack_require__(0)

var Workspace = function (Tool, Sheet) {

    this.width = config.WSConfig.width
    this.height = config.WSConfig.height

    this.Tool =Tool
    this.Sheet= Sheet
}

module.exports.Workspace = Workspace

/***/ })
/******/ ]);
//# sourceMappingURL=app.js.map