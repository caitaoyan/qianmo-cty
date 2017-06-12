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
var SheetConfig = {
	height: 600,
	width: 1030,
	isMouseDown:false,
	isKeyDown:true
}

/**
 * 表头-SheetTool的属性默认值
 *
 */
var SheetToolConfig = {
	height: 40,
	width: SheetConfig.width
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
var SheetGridConfig = {
	height: SheetConfig.height - SheetToolConfig.height,
	width: SheetConfig.width,
	headWidth: 30,
	headHeight: 25,

	rowNum: (SheetConfig.height - SheetToolConfig.height-30)/CellConfig.height,
	colNum: (SheetConfig.width-30)/CellConfig.width
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
module.exports.SheetConfig = SheetConfig
module.exports.SheetToolConfig = SheetToolConfig
module.exports.SheetGridConfig = SheetGridConfig
module.exports.CellConfig = CellConfig


/***/ }),
/* 1 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * Created by Ian on 17/6/7.
 */

var configModule = __webpack_require__(0)

var cellConfig = configModule.CellConfig

var Cell = function (coord) {

	//成员属性
	//设置默认值
	this.height = cellConfig.height
	this.width = cellConfig.width

	this.coord = coord
}

module.exports.Cell = Cell

/***/ }),
/* 2 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * Created by Ian on 17/6/7.
 */


var Controller = function () {}

Controller.prototype.init = function (parentNode) {

	//实例化初始化需要的对象

	var sheetGridModule = __webpack_require__(7)
	var sheetToolModule = __webpack_require__(8)
	var sheetModule = __webpack_require__(6)
	var renderModule = __webpack_require__(5)
	var undoStackModule = __webpack_require__(9)

	var SheetGrid = sheetGridModule.SheetGrid
	var SheetTool = sheetToolModule.SheetTool
	var Sheet = sheetModule.Sheet
	var Render = renderModule.Render
	var UndoStack = undoStackModule.UndoStack

	this.sheetGrid = new SheetGrid()
	this.sheetTool = new SheetTool()
	this.undoStack = new UndoStack()
	this.sheet = new Sheet(this.sheetTool, this.sheetGrid,this.undoStack)
	this.render = new Render()

	this.render.renderSheet(this.sheet, parentNode)
}

module.exports.Controller = Controller

/***/ }),
/* 3 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * Created by Ian on 17/6/5.
 */

var controllerModule = __webpack_require__(2)
var Controller = controllerModule.Controller

var parentNode = document.getElementById('QianMoApp')

var controller = new Controller()
controller.init(parentNode)

/***/ }),
/* 4 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * Created by ASUS on 2017/6/8.
 */
//
var config = __webpack_require__(0)
var CellModule = __webpack_require__(1)
var Cell=CellModule.Cell
var EventHandle = function () {
    this.firstCell = null
    this.lastCell = null
}

EventHandle.prototype.buttonClick = function(str){
    alert(str)
}
EventHandle.prototype.mouseDown = function(element){
    this.setCellBackgroundColor('transparent')
    if(element.firstChild.id){
        this.firstCell = element.firstChild
        config.SheetConfig.isMouseDown=true
    }
}
EventHandle.prototype.mouseMove = function(element){
    this.setCellBackgroundColor('transparent')
    if(element.firstChild.id){
        this.lastCell=element.firstChild
        this.setCellBackgroundColor('#69f')
    }
}
EventHandle.prototype.mouseUp = function(element,sheet){
    if(element.firstChild.id){
        this.lastCell = element.firstChild
        sheet.range=this.firstCell.id+':'+this.lastCell.id
        this.setCellBackgroundColor('#69f')
    }
}
EventHandle.prototype.setCellBackgroundColor=function(backgroundColor){
    var cells=this.getColAndRow(this.firstCell,this.lastCell)
    if(cells!=null) {
        for(var i=cells.firstCellCol;i<=cells.lastCellCol;i++){
            for(var j=cells.firstCellRow;j<=cells.lastCellRow;j++){
                var ele = document.getElementById(String.fromCharCode(i) + '_' + j)
                ele.parentNode.style.backgroundColor = backgroundColor
            }
        }
    }
}
EventHandle.prototype.getColAndRow=function(firstCell,lastCell){
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
EventHandle.prototype.keyDown = function(element,event,sheet){
    var col=element.id.split('_')[0]
    var row=element.id.split('_')[1]

    switch (event.which) {
        case 37://左键
            col=String.fromCharCode(col.charCodeAt(0) - 1)
            var cell = document.getElementById( col +'_'+ row )
            if (cell) {cell.focus()}
            break
        case 38://上键
            var cell = document.getElementById( col +'_'+ (parseInt(row)-1) )
            if (cell) {cell.focus()}
            break
        case 39://右键
            col=String.fromCharCode(col.charCodeAt(0) + 1)
            var cell = document.getElementById( col +'_'+ row)
            if (cell) {cell.focus()}
            break
        case 40://下键
            var cell = document.getElementById( col +'_'+ (parseInt(row)+1) )
            if (cell) {cell.focus()}
            break
        case 17://ctrl
            break
        case 18://alt
            break
        case 67://ctrl+c
             sheet.copiedfrom=sheet.range
            var command = 'copy ' + sheet.copiedfrom + ' formulas'
            var type = 'copy'
            sheet.UndoStack.setStack(command,type)
            break
        case 86://ctrl+v
            var firstCell=sheet.copiedfrom.split(':')[0]
            var lastCell=sheet.copiedfrom.split(':')[1]
            var cells=this.getColAndRow(document.getElementById(firstCell),
                document.getElementById(lastCell))
            if(cells!=null) {
                var colCount=0
                for(var i=cells.firstCellCol;i<=cells.lastCellCol;i++){
                    var rowCount=0
                    for(var j=cells.firstCellRow;j<=cells.lastCellRow;j++){
                        var eleOld = document.getElementById(
                            String.fromCharCode(i)
                            + '_' + j)
                        var eleNew = document.getElementById(
                            String.fromCharCode(col.charCodeAt(0)+colCount)
                            + '_' + (parseInt(row)+rowCount))
                        if(eleNew){
                            eleNew.nextSibling.innerHTML=eleOld.nextSibling.innerHTML
                            var cell=new Cell()
                            cell.value=eleOld.nextSibling.innerHTML
                            sheet.SheetGrid.cells[eleNew.id]=cell
                        }
                        rowCount++

                    }
                    colCount++
                }
                var command = 'past ' + element.id + ' formulas'
                var type = 'past'
                sheet.UndoStack.setStack(command,type)
            }
            config.SheetConfig.isKeyDown=false
            break
        default:
            var cell=new Cell()
            cell.value=element.value
            sheet.SheetGrid.cells[element.id]=cell
            element.style.color='#111'
            element.style.background='#efe'
            break
    }
    console.log(sheet);
}


module.exports.EventHandle = EventHandle

/***/ }),
/* 5 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * Created by Ian on 17/6/5.
 */

/**
 * SheetView Class
 *  负责
 *  @渲染表格
 *  @改变样式
 *
 *  总之，跟dom的直接交互，全由SheetView完成
 */

/**
 * 这段代码还有很多地方需要重构
 * 包括各种样式
 * 希望不要动
 */
var config = __webpack_require__(0)
var EventHandleModule = __webpack_require__(4)
var EventHandle = EventHandleModule.EventHandle
var eventHandle = new EventHandle()
var CellModule = __webpack_require__(1)
var Cell=CellModule.Cell

var Render = function () {}

Render.prototype.renderSheet = function (sheet, parentNode) {
	var sheetDiv = document.createElement('div')

	sheetDiv.style.width = sheet.width + 'px'
	sheetDiv.style.height = sheet.height + 'px'

	parentNode.appendChild(sheetDiv)

	var sheetTool = sheet.SheetTool
	// console.log(sheet)
	var sheetToolDiv = document.createElement('div')

	sheetToolDiv.style.width = sheetTool.width + 'px'
	sheetToolDiv.style.height = sheetTool.height + 'px'
	sheetToolDiv.style.backgroundColor = '#aaaaaa'
	//sheetToolDiv.innerHTML = "&#xe900;&#xe901;&#xe14d;&#xe14e;&#xe14f;"
    var html = ["&#xe900","&#xe901","&#xe14d","&#xe14e","&#xe14f"]

	html.forEach(function(innerhtml){
	    var buttonDiv = document.createElement('div')
		buttonDiv.style.display = "inline"
        buttonDiv.innerHTML=innerhtml

        buttonDiv.onclick = function(){
            eventHandle.buttonClick(innerhtml)
        }
	    sheetToolDiv.appendChild(buttonDiv)
    })


	sheetDiv.appendChild(sheetToolDiv)

	var sheetGrid = sheet.SheetGrid
	var sheetGridTable = document.createElement('table')
	sheetDiv.appendChild(sheetGridTable)
	sheetGridTable.style.width = sheetGrid.width + 'px'
	sheetGridTable.style.height = sheetGrid.height + 'px'

	fillGrid(sheet, sheetGridTable)

}

/**
 * 构建表格的数据模型
 * 一个二维数组
 *
 * 感觉这个应该放在controller里面
 * 但是先这样写
 * 你们不要改
 *
 * 表头存放表头显示的值
 * 其他为空，表示要放input
 *
 * @param rowNum
 * @param colNum
 */
function createModel(rowNum, colNum) {
	var sheetGridModel =  []
	for(var a=0;a<rowNum+1;a++){
		sheetGridModel[a] = []

		for(var b=0;b<colNum+1;b++){
			sheetGridModel[a][b] = null
		}
	}
	var startString = "A"

	for(var i=1; i<colNum+1; i++){
		sheetGridModel[0][i] = String.fromCharCode(startString.charCodeAt(0) + i - 1)
	}

	for(var j=1; j<rowNum+1; j++){
		sheetGridModel[j][0] = j
	}

	return sheetGridModel
}

function fillGrid(sheet, sheetGridTable) {
	var sheetGridModel = createModel(sheet.SheetGrid.rowNum,  sheet.SheetGrid.colNum)
	var cellHeight = config.CellConfig.height
	var cellWidth = config.CellConfig.width
	var	headWidth = config.SheetGridConfig.headWidth
	var	headHeight = config.SheetGridConfig.headHeight

	var rowNum = sheetGridModel.length
	var colNum = sheetGridModel[0].length


	var rowHead = document.createElement('tr')
	sheetGridTable.appendChild(rowHead)

	for(var i=0;i<colNum;i++){
		var rowHeadTH  = document.createElement('th')
		rowHead.appendChild(rowHeadTH)

		var rowHeadDiv = document.createElement('div')
		rowHeadTH.appendChild(rowHeadDiv)

		if(i === 0){
			rowHeadTH.style.width = headWidth + 'px'
			rowHeadTH.style.height = headHeight + 'px'
			rowHeadDiv.style.width = headWidth-2 + 'px'
			rowHeadDiv.style.height = headHeight-2 + 'px'
		}else{
			rowHeadTH.style.height = headHeight + 'px'
			rowHeadTH.style.width = cellWidth + 'px'
			rowHeadDiv.style.height = headHeight-2 + 'px'
			rowHeadDiv.style.width = cellWidth-2 + 'px'

			rowHeadDiv.innerHTML = sheetGridModel[0][i]
		}

	}

	for(var j = 1;j<rowNum;j++){
		var rowTR = document.createElement('tr')
		sheetGridTable.appendChild(rowTR)

		for(var k=0; k<colNum;k++){
			if(k === 0){
				var rowTH = document.createElement('th')
				rowTR.appendChild(rowTH)
				var rowDiv = document.createElement('div')
				rowTH.appendChild(rowDiv)

				rowTH.style.width = headWidth + 'px'
				rowTH.style.height = cellHeight + 'px'
				rowDiv.style.width = headWidth-2 + 'px'
				rowDiv.style.height = cellHeight-2 + 'px'
				rowDiv.innerHTML = sheetGridModel[j][k]
			}else {
				var rowTD = document.createElement('td')
                rowTD.style.display = 'table-cell'
				rowTR.appendChild(rowTD)

				var rowInput = document.createElement('input')
                rowInput.id = sheetGridModel[0][k]+"_"+j
                rowInput.style.position = 'absolute'
                rowInput.style.color = 'transparent'
                rowInput.style.backgroundColor = 'transparent'
				rowTD.onmousedown = function(){
                    eventHandle.mouseDown(this)
                }
                window.onmousedown=function(){
					eventHandle.setCellBackgroundColor('transparent')
				}
				rowTD.onmouseup = function(){
                    eventHandle.mouseUp(this,sheet)
                }
				window.onmouseup = function(){
					config.SheetConfig.isMouseDown=false
				}
				rowTD.onmousemove = function(){
					if(config.SheetConfig.isMouseDown) {
						eventHandle.mouseMove(this)
					}
				}
                rowInput.ondblclick = function () {
                    this.style.color='#111'
                    this.style.background='#efe'
                }
				rowInput.onfocus = function () {
					this.parentNode.style.backgroundColor = '#69f'
				}
                rowInput.onblur=function () {
                    this.style.color='transparent'
                    this.style.background='transparent'
					if(config.SheetConfig.isKeyDown){
                    	var value=this.value
                    	if(this.nextSibling.innerHTML==''&&value!=''){
                    		var command = 'set '+ this.id + ' ' + value
							var type = 'set'
							sheet.UndoStack.setStack(command,type)
						}
						this.nextSibling.innerHTML=value
					}
					this.parentNode.style.backgroundColor = 'transparent'
                }
				rowInput.onkeydown=function (event) {
					eventHandle.keyDown(this,event,sheet)
				}
                rowTD.appendChild(rowInput)

                var rowDiv = document.createElement('div')
                rowDiv.style.textAlign = 'right'
                rowTD.appendChild(rowDiv)

				rowTD.style.width = cellWidth + 'px'
				rowTD.style.height = cellHeight + 'px'
				rowInput.style.width = cellWidth-2 + 'px'
				rowInput.style.height = cellHeight-2 + 'px'
                rowDiv.style.width = cellWidth-2 + 'px'
                rowDiv.style.height = cellHeight-2 + 'px'

			}
		}
	}
}
module.exports.Render = Render

/***/ }),
/* 6 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * Created by Ian on 17/6/7.
 */

var configModule = __webpack_require__(0)

var sheetConfig = configModule.SheetConfig

var Sheet = function (SheetTool, SheetGrid,UndoStack) {

	this.width = sheetConfig.width
	this.height = sheetConfig.height

	this.SheetTool = SheetTool
	this.SheetGrid = SheetGrid
	this.UndoStack = UndoStack
}

module.exports.Sheet = Sheet

/***/ }),
/* 7 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * Created by Ian on 17/6/7.
 */

/**
 * 表格对象
 */


var configModule = __webpack_require__(0)

var sheetGridConfig = configModule.SheetGridConfig

var SheetGrid = function () {
	this.height = sheetGridConfig.height
	this.width = sheetGridConfig.width

	this.headWidth = sheetGridConfig.headWidth
	this.headHeight = sheetGridConfig.headHeight

	this.rowNum = sheetGridConfig.rowNum
	this.colNum = sheetGridConfig.colNum

	this.cells = {}
}

module.exports.SheetGrid = SheetGrid

/***/ }),
/* 8 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * Created by Ian on 17/6/7.
 */

var configModule = __webpack_require__(0)

var sheetToolConfig = configModule.SheetToolConfig

var SheetTool = function () {
	this.width = sheetToolConfig.width
	this.height = sheetToolConfig.height
}

module.exports.SheetTool = SheetTool

/***/ }),
/* 9 */
/***/ (function(module, exports) {

/**
 * Created by cty on 17/6/12.
 */


var UndoStack = function () {

    this.tos=-1
    this.maxRedo=0
    this.maxUndo=50
    this.stack=[]

}
UndoStack.prototype.setStack=function(command,type){
    var stack={
        command : command,
        type: type
    }
    this.stack.push(stack)
    this.tos++
}
UndoStack.prototype.reDo=function(){

}
UndoStack.prototype.unDo=function(){

}
module.exports.UndoStack = UndoStack

/***/ })
/******/ ]);
//# sourceMappingURL=app.js.map