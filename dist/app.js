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
/******/ 	return __webpack_require__(__webpack_require__.s = 5);
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
	height: 4400,
	width: 2030,
	isMouseDown:false,
	isKeyDown:true,
	isEditing:false,
	editCell:null,
	isFont:false,
	isPreview:false
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
 * 公式类型
 * @type {[*]}
 */
var function_classlist= ["all", "stat", "lookup", "datetime", "financial", "test", "math", "text"]
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
/***/ (function(module, exports, __webpack_require__) {

/**
 * 整体工作管理对象
 * @constructor
 */
var WSManager = function () {}

WSManager.prototype.init = function (parentNode) {

    //实例化初始化Tool对象
    var ToolModule = __webpack_require__(10)
    var Tool=ToolModule.Tool
    var tool=new Tool()

    //实例化初始化UndoStack对象
    var UndoStackModule =__webpack_require__(14)
    var UndoStack = UndoStackModule.UndoStack
    var undoStack = new UndoStack()

    //实例化初始化Sheet对象
    var SheetModule = __webpack_require__(6)
    var Sheet =SheetModule.Sheet
    var sheet=new Sheet(undoStack)

    //实例化初始化Workspace对象
    var WorkspaceModule = __webpack_require__(16)
    var Workspace = WorkspaceModule.Workspace
    this.workspace=new Workspace(tool,sheet)

    //实例化初始化WSRender对象
    var WSRenderModule = __webpack_require__(15)
    var WSRender = WSRenderModule.WSRender
    var wsRender=new WSRender(this,parentNode)

    wsRender.init()

    if(parentNode.getAttribute("url")){
        document.getElementById("Init").onclick()
    }
}

module.exports.WSManager = WSManager

/***/ }),
/* 2 */
/***/ (function(module, exports, __webpack_require__) {

var config = __webpack_require__(0)

var UtilModule=__webpack_require__(3)
var Util=UtilModule.Util
var util=new Util()

var CellModule = __webpack_require__(4)
var Cell=CellModule.Cell

var CellRender = function (sheet) {
    this.sheet=sheet
}
/**
 * 渲染单元格
 * @param id
 * @param cmd
 * @param value
 */
CellRender.prototype.renderCell=function (id,cmd,value) {
    var sheet=this.sheet
    if(id && id.indexOf('_')==-1){
        var index=id.search(/\d/)
        id=id.substring(0,index)+'_'+id.substring(index,id.length)
    }
    var  ele = document.getElementById(id)
    if(id && !sheet.cells[id]){
        sheet.cells[id] = new Cell(id)
    }
    if(id==null||ele){
        switch (cmd) {
            //合并范围
            case "area":
                if(value.indexOf(':')==-1){
                    var firstCell=value.split('-')[0]
                    var index=firstCell.search(/\d/)
                    firstCell=firstCell.substring(0,index)+'_'+firstCell.substring(index,firstCell.length)
                    var lastCell=value.split('-')[1]
                    var index=lastCell.search(/\d/)
                    lastCell=lastCell.substring(0,index)+'_'+lastCell.substring(index,lastCell.length)
                    sheet.range=firstCell+':'+lastCell
                    value=firstCell+':'+lastCell
                }
                this.sheet.cells[id].area=value
                var firstCell=value.split(':')[0]
                var lastCell=value.split(':')[1]
                var cells=util.getColAndRow(document.getElementById(firstCell),
                    document.getElementById(lastCell))
                this.mergeCell(firstCell,lastCell,cells,this.sheet)
                break
            //前景色
            case "foregroundColor":
                if(value!='#000000'){
                    ele.style.color=value
                    this.sheet.cells[id].foregroundColor=value
                }

                break
            //背景色
            case "backgroundColor":
                if(value!='#000000'){
                    ele.style.backgroundColor=value
                    this.sheet.cells[id].backgroundColor=value
                }
                break
            //文本内容
            case "content":
                if(value==null){
                    value=''
                }
                ele.firstChild.innerHTML=value
                this.sheet.cells[id].content=value
                break
            //字体
            case "font":
                if(value=='Default'){
                    value=''
                }
                ele.style.fontFamily=value
                this.sheet.cells[id].font=value
                break
            //字体大小
            case "fontSize":
                if(value=='Default'||value=='-----'){
                    value=''
                }
                ele.style.fontSize=value
                this.sheet.cells[id].fontSize=value
                break
            //字体类型
            case "format":
                this.sheet.cells[id].format=value
                break
            //加粗
            case "bold":
                if(value=='加粗'){
                    value='bold'
                }
                if(!config.WSConfig.isFont){
                    if(ele.style.fontWeight!=''){
                        value=''
                    }else{
                        config.WSConfig.isFont=true
                    }
                }
                ele.style.fontWeight=value
                this.sheet.cells[id].bold=value
                break
            //斜体
            case "italic":
                if(!config.WSConfig.isFont){
                    if(ele.style.fontStyle!=''){
                        value=''
                    }else{
                        config.WSConfig.isFont=true
                    }
                }

                ele.style.fontStyle=value
                this.sheet.cells[id].italic=value
                break
            //对齐方向
            case "alignment":
                if(value=='左对齐'){
                    value='left'
                }else if(value=='居中对齐'){
                    value='center'
                }else if(value=='右对齐'){
                    value='right'
                }
                ele.style.textAlign=value
                this.sheet.cells[id].alignment=value
                break
            //边框
            case "bottomFrame":
                if(value=='有'||value=='bottom'){
                    value='bottom_bottomFrame'
                }
            case "leftFrame":
                if(value=='有'||value=='left'){
                    value='left_leftFrame'
                }
            case "rightFrame":
                if(value=='有'||value=='right'){
                    value='right_rightFrame'
                }
            case "topFrame":
                if(value=='有'||value=='top'){
                    value='top_topFrame'

                }
            case "frame":
                // '-----','左边框','上边框','右边框','下边框','外侧边框','无边框','全边框'
                if(value=='左边框'){
                    value='left_frame'
                }else if(value=='上边框'){
                    value='top_frame'
                }else if(value=='右边框'){
                    value='right_frame'
                }else if(value=='下边框'){
                    value='bottom_frame'
                }else if(value=='外侧边框'){
                    value='out_frame'
                }else if(value=='无边框'){
                    value='none_frame'
                }else if(value=='全边框'){
                    value='all_frame'
                }
                if(value!='无'){
                    var firstCell=sheet.range.split(':')[0]
                    var lastCell=sheet.range.split(':')[1]
                    var cells=util.getColAndRow(document.getElementById(firstCell),
                        document.getElementById(lastCell))
                    this.setFontBorder(cells,value.split('_')[0],value.split('_')[1])
                }
                break
            //缩进
            case "indentation":
                var html='';
                for(var i=0;i<value;i++){
                    html+='&nbsp;'
                }
                ele.firstChild.innerHTML=html+this.sheet.cells[id].content
                this.sheet.cells[id].indentation=value
                break
        }
    }
}
/**
 * 合并单元格
 * @param firstCell
 * @param lastCell
 * @param cells
 * @param sheet
 */
CellRender.prototype.mergeCell=function(firstCell,lastCell,cells,sheet){
    var newFirstCell=document.getElementById(String.fromCharCode(cells.firstCellCol)+"_"+cells.firstCellRow)
    var col=0
    var row=0
    if(firstCell!=lastCell){
        for(var i=cells.firstCellCol;i<=cells.lastCellCol;i++) {
            var newRow=0
            for (var j = cells.firstCellRow; j <= cells.lastCellRow; j++) {
                var ele=document.getElementById(String.fromCharCode(i) + '_' + j)
                if(ele){
                    if(ele.style.display!='none'){
                        newRow=newRow+ele.rowSpan
                    }
                }
            }
            if(newRow>row){
                row=newRow
            }
        }
        for(var i=cells.firstCellRow;i<=cells.lastCellRow;i++) {
            var newCol=0
            for (var j = cells.firstCellCol; j <= cells.lastCellCol; j++) {
                var ele=document.getElementById(String.fromCharCode(j) + '_' + i)
                if(ele){
                    if(ele.style.display!='none'){
                        newCol=newCol+ele.colSpan
                    }
                }
            }
            if(newCol>col){
                col=newCol
            }
        }
        var newLastCell=String.fromCharCode(cells.firstCellCol+col-1) + '_' + (cells.firstCellRow+row-1)
        cells=util.getColAndRow(newFirstCell,
            document.getElementById(newLastCell))
        var colSpanNum=0
        var rowSpanNum=0
        for(var i=cells.firstCellCol;i<=cells.lastCellCol;i++) {
            var newRowSpanNum=0
            for (var j = cells.firstCellRow; j <= cells.lastCellRow; j++) {
                    var ele=document.getElementById(String.fromCharCode(i) + '_' + j)
                    if(ele){
                        newRowSpanNum++
                        if(i!=cells.firstCellCol||j!=cells.firstCellRow){
                            ele .style.display='none'
                        }
                }
            }
            if(newRowSpanNum>=rowSpanNum){
                rowSpanNum=newRowSpanNum
            }
            colSpanNum++
        }
        newFirstCell.rowSpan=rowSpanNum
        newFirstCell.colSpan=colSpanNum
        if(!this.sheet.cells[newFirstCell.id]){
            this.sheet.cells[newFirstCell.id] = new Cell(newFirstCell.id)
        }
        this.sheet.cells[newFirstCell.id].rowSpan=rowSpanNum
        this.sheet.cells[newFirstCell.id].colSpan=colSpanNum
        sheet.range=newFirstCell.id+":"+newFirstCell.id
    }else{
        var colNum=newFirstCell.colSpan
        var rowNum=newFirstCell.rowSpan
        var flag=0;
        var last=null;
        for(var i=0;i<colNum;i++) {
            for (var j = 0; j <rowNum; j++) {
                var id=String.fromCharCode(cells.firstCellCol+i) + '_' + (cells.firstCellRow+j)
                var ele=document.getElementById(id)
                if(ele){
                    last=ele
                    ele.colSpan=1
                    ele.rowSpan=1
                    ele .style.display='table-cell'
                    if(!this.sheet.cells[ele.id]){
                        this.sheet.cells[ele.id] = new Cell(ele.id)
                    }
                    this.sheet.cells[ele.id].rowSpan=ele.rowSpan
                    this.sheet.cells[ele.id].colSpan=ele.colSpan
                    if(flag==0){
                        sheet.range=ele.id+":"
                        flag++
                    }

                }
            }
        }
        sheet.range+=last.id
    }
}
/**
 * 设置边框
 * @param cells
 * @param fontValue
 */
CellRender.prototype.setFontBorder=function(cells,fontValue,fontType){
    // '-----','左边框','上边框','右边框','下边框','外侧边框','无边框','全边框'
        for(var i=cells.firstCellCol;i<=cells.lastCellCol;i++) {
        for (var j = cells.firstCellRow; j <= cells.lastCellRow; j++) {
            var ele = document.getElementById(String.fromCharCode(i) + '_' + j)
            if(!this.sheet.cells[ele.id]){
                this.sheet.cells[ele.id] = new Cell(ele.id)
            }
            this.sheet.cells[ele.id][fontType]=fontValue
            if(fontValue=='none'){
                ele.style.border='1px solid #ccc'
                var colSpan=ele.colSpan
                for(var z=0;z<colSpan;z++) {
                    ele = document.getElementById(String.fromCharCode(i + z) + '_' + (j - 1))
                    if(ele){
                        ele.style.borderBottom = '1px solid #ccc'
                    }
                }
                ele = document.getElementById(String.fromCharCode(i-1) + '_' + j)
                if(ele){
                    ele.style.borderRight='1px solid #ccc'
                }
            }else if(fontValue=='all'){
                ele.style.border='1px solid #000'
                var colSpan=ele.colSpan
                for(var z=0;z<colSpan;z++){
                    ele = document.getElementById(String.fromCharCode(i+z) + '_' + (j-1))
                    if(ele){
                        ele.style.borderBottom='1px solid #000'
                    }
                }
                ele = document.getElementById(String.fromCharCode(i-1) + '_' + j)
                if(ele){
                    ele.style.borderRight='1px solid #000'
                }


            }else{
                if(i==cells.firstCellCol&&(fontValue=='left'||fontValue=='out')){
                    var eleNew = document.getElementById(String.fromCharCode(i-1) + '_' + j)
                    if(eleNew){
                        eleNew.style.borderRight='1px solid #000'
                    }

                }
                if(j == cells.firstCellRow&&(fontValue=='top'||fontValue=='out')){
                    var colSpan=ele.colSpan
                    for(var z=0;z<colSpan;z++){
                        var eleNew = document.getElementById(String.fromCharCode(i+z) + '_' + (j-1))
                        if(eleNew){
                            eleNew.style.borderBottom='1px solid #000'
                        }

                    }
                }
                if(i==cells.lastCellCol&&(fontValue=='right'||fontValue=='out')){
                    ele.style.borderRight='1px solid #000'
                }
                if(j == cells.lastCellRow&&(fontValue=='bottom'||fontValue=='out')){
                    ele.style.borderBottom='1px solid #000'
                }
            }
        }
    }
}


module.exports.CellRender = CellRender

/***/ }),
/* 3 */
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
	this.autoLF = false
	this.viewWidth = 0
	this.height = cellConfig.height
	this.width = cellConfig.width

	this.coord = coord
}

module.exports.Cell = Cell

/***/ }),
/* 5 */
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
/* 6 */
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
/* 7 */
/***/ (function(module, exports, __webpack_require__) {

//事件绑定对象
var config = __webpack_require__(0)
var SheetEventHandlerModule=__webpack_require__(8)
var SheetEventHandler = SheetEventHandlerModule.SheetEventHandler
var sheetEventHandler = null

var SheetEventBinder = function (sheet) {
    this.sheet=sheet
    sheetEventHandler=new SheetEventHandler(sheet)
    window.onmouseup = function(){
        config.WSConfig.isMouseDown=false
    }
    document.onkeydown = function (event) {
        if(sheetEventHandler.firstCell.id){
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
/* 8 */
/***/ (function(module, exports, __webpack_require__) {

//鼠标和键盘事件
var config = __webpack_require__(0)
var CellModule = __webpack_require__(4)
var Cell=CellModule.Cell
var UtilModule=__webpack_require__(3)
var Util=UtilModule.Util
var util=new Util()

var CellRenderModule=__webpack_require__(2)
var CellRender=CellRenderModule.CellRender
var cellRender=null

var SheetEventHandler = function (sheet) {
    this.firstCell = null
    this.lastCell = null
    this.sheet=sheet
    cellRender=new CellRender(sheet)
}
/**
 * 鼠标按下事件
 * @param element
 */
SheetEventHandler.prototype.mouseDown = function(element){
    var cellPropDiv=document.getElementById("cellProp")
    cellPropDiv.style.display='none'
    cellPropDiv.innerHTML=''
    this.setCellBackgroundColor('#fff')
    if(element.id){
        this.setCellProp(element,this.sheet.cells[element.id])
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
        var cellPropDiv=document.getElementById("cellProp")
        cellPropDiv.style.display='none'
        cellPropDiv.innerHTML=''
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
        if(this.firstCell){
            this.sheet.range=this.firstCell.id+':'+this.lastCell.id
            this.setCellBackgroundColor('#69f')
        }
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
            input.value = this.sheet.cells[element.id].content
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
            this.sheet.cells[ele.id] = new Cell(ele.id)
        }
        this.sheet.cells[ele.id].content = input.value
        ele.firstChild.innerHTML = input.value
        var sp = document.getElementById('sp')
        sp.value = input.value
        sp.style.font = ele.firstChild.style.font
        sp.style.fontSize = ele.firstChild.style.fontSize
        this.sheet.cells[ele.id].viewWidth = sp.offsetWidth
        if(this.sheet.cells[ele.id].autoLF){
            ele.width = this.sheet.cells[ele.id].width
        }else{
            ele.width = this.sheet.cells[ele.id].viewWidth
        }
    }else{
        ele.firstChild.innerHTML = ''
    }
}
SheetEventHandler.prototype.inputFocus = function () {
    config.isEditing = true
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
SheetEventHandler.prototype.setCellBackgroundColor=function(backgroundColor){
    var cells=util.getColAndRow(this.firstCell,this.lastCell)
    if(cells!=null) {
        for(var i=cells.firstCellCol;i<=cells.lastCellCol;i++){
            for(var j=cells.firstCellRow;j<=cells.lastCellRow;j++){
                var ele = document.getElementById(String.fromCharCode(i) + '_' + j)
                if(ele){
                    if(backgroundColor=='#fff'
                        &&this.sheet.cells[ele.id]!=undefined
                        &&this.sheet.cells[ele.id].backgroundColor!=undefined){
                        ele.style.backgroundColor = this.sheet.cells[ele.id].backgroundColor
                    }else{
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
            var ta = document.getElementById('ta')
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
                        var cvCol = String.fromCharCode(col.charCodeAt(0) + j)
                        var cvRow = parseInt(row) + i
                        var cell = document.getElementById(cvCol + '_' + cvRow)
                        if(cell){
                            cell.firstChild.innerHTML = vv[j]
                            if(!sheet.cells[cell.id]) {
                                sheet.cells[cell.id] = new Cell(cell)
                            }
                            sheet.cells[cell.id].content = vv[j]
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
SheetEventHandler.prototype.setCellProp=function(element,e){
    if(e){
        var cellPropDiv=document.getElementById("cellProp")
        cellPropDiv.style.display='block'
        var left=element.offsetLeft
        var top=element.offsetTop
        cellPropDiv.style.top=top+element.offsetHeight+26+'px'
        cellPropDiv.style.left=left+element.offsetWidth+9+'px'
        var cellPropTable=document.createElement('table')
        cellPropDiv.appendChild(cellPropTable)
        var tr=document.createElement('tr')
        cellPropTable.appendChild(tr)
        var th=document.createElement('th')
        th.innerHTML='属性名称'
        tr.appendChild(th)
        th=document.createElement('th')
        th.innerHTML='属性值'
        tr.appendChild(th)
        for (var key in e) {
            if(e[key]!=''){
                var tr=document.createElement('tr')
                cellPropTable.appendChild(tr)
                var td=document.createElement('td')
                td.innerHTML=key
                tr.appendChild(td)
                var td=document.createElement('td')
                td.innerHTML=e[key]
                tr.appendChild(td)
            }
        }
        this.addProp(element,cellPropTable)
    }
}
SheetEventHandler.prototype.addProp=function(ele,cellPropTable){
    var SheetEventHandler=this
    var tr=document.createElement('tr')
    cellPropTable.appendChild(tr)
    var td=document.createElement('td')
    var inputTd1=document.createElement('input')
    inputTd1.setAttribute('list','prop')
    td.appendChild(inputTd1)
    var datalist=document.createElement('datalist')
    datalist.id='prop'
    td.appendChild(datalist)
    var option=document.createElement('option')
    option.value='foregroundColor'
    datalist.appendChild(option)
    option=document.createElement('option')
    option.value='bold'
    datalist.appendChild(option)
    option=document.createElement('option')
    option.value='backgroundColor'
    datalist.appendChild(option)
    option=document.createElement('option')
    option.value='italic'
    datalist.appendChild(option)
    option=document.createElement('option')
    option.value='fontSize'
    datalist.appendChild(option)
    tr.appendChild(td)
    var td=document.createElement('td')
    var inputTd2=document.createElement('input')
    td.appendChild(inputTd2)
    tr.appendChild(td)
    inputTd1.onclick=function(){
        SheetEventHandler.firstCell=this
    }
    inputTd2.onclick=function(){
        SheetEventHandler.firstCell=this
    }
    inputTd1.onblur=function(){
        SheetEventHandler.firstCell=ele
        var input2=this.parentNode.nextSibling.firstChild
        if(inputTd1.value!=null &&inputTd1.value !=''
            &&input2.value!=null &&input2.value !=''){
            cellRender.renderCell(ele.id,inputTd1.value,input2.value)
            if(!this.parentNode.parentNode.nextSibling){
                SheetEventHandler.addProp(ele,cellPropTable)
            }
        }

    }
    inputTd2.onblur=function(){
        SheetEventHandler.firstCell=ele
        var input1=this.parentNode.previousSibling.firstChild
        if(input1.value!=null &&input1.value !=''
            &&inputTd2.value!=null &&inputTd2.value !=''){
            cellRender.renderCell(ele.id,input1.value,inputTd2.value)
            if(!this.parentNode.parentNode.nextSibling){
                SheetEventHandler.addProp(ele,cellPropTable)
            }
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
/* 9 */
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
    var cellPropDiv=document.createElement("div")
    cellPropDiv.id='cellProp'
    // cellPropDiv.style.display='none'
    cellPropDiv.style.position = "absolute"
    cellPropDiv.style.zIndex = 100
    cellPropDiv.style.backgroundColor = "#FFF"
    cellPropDiv.style.border = "1px solid black"
    sheetDiv.appendChild(cellPropDiv)

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

		}else{
			rowHeadDiv.innerHTML = '&nbsp;&nbsp;&nbsp;&nbsp;'+gridHeader[0][i]+'&nbsp;&nbsp;&nbsp;&nbsp;'
		}
	}

    var SheetEventBinderModule = __webpack_require__(7)
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
                rowTD.appendChild(rowDiv)
			}
		}
	}
	var input = document.createElement('input')
	input.style.display = 'none'
    input.style.position = 'absolute'
	input.id = 'input'
    if(!config.WSConfig.isPreview){
        sheetEventBinder.initInput(input)
    }

	sheetTable.appendChild(input)
    var ta = document.createElement('textarea') // used for ctrl-c/ctrl-v where an invisible text area is needed
    ta.style = 'display:none;position:absolute;height:1px;width:1px;opacity:0;filter:alpha(opacity=0);'
    ta.value = '';
    ta.id = 'ta'
    sheetTable.appendChild(ta)
    var span = document.createElement('span') // 用于获得字符串的显示长度
    span.style = 'visibility: hidden;white-space: nowrap;filter:alpha(opacity=0);'
    span.value = '';
    span.id = 'sp'
    sheetTable.appendChild(span)


}


module.exports.SheetRender = SheetRender

/***/ }),
/* 10 */
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
/* 11 */
/***/ (function(module, exports, __webpack_require__) {

/**
 工具栏绑定事件
 */
var config = __webpack_require__(0)
var ToolEventHandlerModule=__webpack_require__(12)
var ToolEventHandler = ToolEventHandlerModule.ToolEventHandler
var toolEventHandler=null

var CellRenderModule=__webpack_require__(2)
var CellRender=CellRenderModule.CellRender
var cellRender=null

var ToolEventBinder=function(sheet){
    this.sheet=sheet
    toolEventHandler=new ToolEventHandler(sheet)
    cellRender=new CellRender(sheet)
}
ToolEventBinder.prototype.buttonClick = function (buttonDiv) {
    toolEventHandler.buttonClick(buttonDiv.innerHTML)
}
ToolEventBinder.prototype.initFontFamily=function(fontFamilySelect){
    var toolEventBinder=this
    fontFamilySelect.onchange=function(){
        var fontFamily=fontFamilySelect.value
        toolEventBinder.fontValue=fontFamily
        toolEventBinder.fontType='font'
        toolEventHandler.setFont(toolEventBinder)
        fontFamilySelect.value='-----'
    }
}
ToolEventBinder.prototype.initFontSize=function(fontSizeSelect){
    var toolEventBinder=this
    fontSizeSelect.onchange=function(){
        var fontSize=fontSizeSelect.value
        toolEventBinder.fontValue=fontSize
        toolEventBinder.fontType='fontSize'
        toolEventHandler.setFont(toolEventBinder)
        fontSizeSelect.value='-----'
    }
}
ToolEventBinder.prototype.initFontWeight=function(fontWeightDiv){
    var toolEventBinder=this
    fontWeightDiv.onclick=function(){
        config.WSConfig.isFont=false
        toolEventBinder.fontValue='bold'
        toolEventBinder.fontType='bold'
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
        toolEventBinder.fontType='italic'
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
            toolEventBinder.fontType='foregroundColor'
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
        if(colorSelectDiv.style.display=='none'||toolEventBinder.fontType=='foregroundColor'){
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
ToolEventBinder.prototype.initMerge=function(mergeDiv){
    var toolEventBinder=this
    mergeDiv.onclick=function(){
        config.WSConfig.isFont=false
        // toolEventBinder.fontValue='italic'
        toolEventBinder.fontType='merge'
        toolEventHandler.setFont(toolEventBinder)
    }
}
ToolEventBinder.prototype.initFontBorder=function(fontBorderSelect){
    var toolEventBinder=this
    fontBorderSelect.onchange=function(){
        var fontBorder=fontBorderSelect.value
        toolEventBinder.fontValue=fontBorder
        toolEventBinder.fontType='border'
        toolEventHandler.setFont(toolEventBinder)
        fontBorderSelect.value='-----'
    }
}
ToolEventBinder.prototype.initTextAlign=function(textAlignSelect){
    var toolEventBinder=this
    textAlignSelect.onchange=function(){
        var textAlign=textAlignSelect.value
        toolEventBinder.fontValue=textAlign
        toolEventBinder.fontType='alignment'
        toolEventHandler.setFont(toolEventBinder)
        textAlignSelect.value='-----'
    }
}
ToolEventBinder.prototype.initFileInput=function(fileInput){
    fileInput.onchange=function(){
        var ajax
        if(window.XMLHttpRequest){
            ajax = new XMLHttpRequest()
        }else if(window.ActiveXObject){
            ajax = new window.ActiveXObject()
        }else{
            alert("请升级至最新版本的浏览器")
        }
        if(ajax !=null){
            ajax.open("GET","json.json",true)
            ajax.send(null)
            ajax.onreadystatechange=function(){
                if(ajax.readyState==4&&ajax.status==200){
                    var CellList = JSON.parse(ajax.responseText)
                    CellList.forEach(function(e){
                        for (var key in e) {
                            cellRender.renderCell(e['cellName'],key,e[key])
                        }
                    })
                }
            }
        }
    }
}
module.exports.ToolEventBinder=ToolEventBinder

/***/ }),
/* 12 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * 工具栏事件
 */
var config = __webpack_require__(0)

var UtilModule=__webpack_require__(3)
var Util=UtilModule.Util
var util=new Util()

var CellRenderModule = __webpack_require__(2)
var CellRender=CellRenderModule.CellRender
var cellRender=null

var ToolEventHandler = function(sheet){
    this.sheet=sheet
    cellRender=new CellRender(sheet)
}
ToolEventHandler.prototype.buttonClick = function(action){
    console.log(action)
    switch(action){
        case "autoLF":
            var sheet = this.sheet
            if(sheet.range){
                console.log('a')
                var firstCell=sheet.range.split(':')[0]
                var lastCell=sheet.range.split(':')[1]
                var cells=util.getColAndRow(document.getElementById(firstCell),
                    document.getElementById(lastCell))
                if(cells!=null){
                    for(var i=cells.firstCellCol;i<=cells.lastCellCol;i++) {
                        for (var j = cells.firstCellRow; j <= cells.lastCellRow; j++) {
                            var ele = document.getElementById(String.fromCharCode(i) + '_' + j)
                            if(sheet.cells[ele.id]){
                                sheet.cells[ele.id].autoLF = !sheet.cells[ele.id].autoLF
                                if(sheet.cells[ele.id].autoLF){
                                    ele.width = sheet.cells[ele.id].width
                                    console.log(ele.width)
                                }
                                else{
                                    ele.width = sheet.cells[ele.id].viewWidth
                                    console.log(ele.width)
                                }
                            }
                        }
                    }
                }
            }
            break
        case "Preview":
            var flag=false
            var e=this.sheet.cells
            for (var key in e) {
                flag=true
                break
            }
            if(flag){
                var needEditCells={}
                for (var key in e) {
                    if(e[key]["content"]!=''&&e[key]["content"].indexOf('=')==0){
                        needEditCells[key]=e[key]["content"]
                    }
                }
                config.WSConfig.isPreview=true
                var WSManagerModule = __webpack_require__(1)
                var WSManager = WSManagerModule.WSManager
                var newWsManager=new WSManager()
                var parentNode = document.getElementById('QianMoApp')
                parentNode.firstChild.remove()
                newWsManager.init(parentNode)
                for (var key in e) {
                    for(var key1 in e[key]){
                        cellRender.renderCell(key,key1,e[key][key1])
                    }
                }
            }else{
                alert("无内容，不允许预览！")
            }

            break
        case "Edit":
            config.WSConfig.isPreview=false
            var WSManagerModule = __webpack_require__(1)
            var WSManager = WSManagerModule.WSManager
            var newWsManager=new WSManager()
            var parentNode = document.getElementById('QianMoApp')
            parentNode.firstChild.remove()
            newWsManager.init(parentNode)
            var e=this.sheet.cells
            for (var key in e) {
                for(var key1 in e[key]){
                    cellRender.renderCell(key,key1,e[key][key1])
                }
            }
            break
        case "Down":
            var ajax
            if(window.XMLHttpRequest){
                ajax = new XMLHttpRequest()
            }else if(window.ActiveXObject){
                ajax = new window.ActiveXObject()
            }else{
                alert("请升级至最新版本的浏览器")
            }
            if(ajax !=null){
                ajax.open("POST","json.json",true)
                ajax.send(null)
                ajax.onreadystatechange=function(){
                    if(ajax.readyState==4&&ajax.status==200){
                        var CellList = JSON.parse(ajax.responseText)
                        CellList.forEach(function(e){
                            for (var key in e) {
                                cellRender.renderCell(e['cellName'],key,e[key])
                            }
                        })
                    }
                }
            }
            break
        case  "Init":
            var parentNode=document.getElementById("QianMoApp")
            var ajax
            if(window.XMLHttpRequest){
                ajax = new XMLHttpRequest()
            }else if(window.ActiveXObject){
                ajax = new window.ActiveXObject()
            }else{
                alert("请升级至最新版本的浏览器")
            }
            if(ajax !=null){
                ajax.open("GET",parentNode.getAttribute("url"),true)
                ajax.send(null)
                ajax.onreadystatechange=function(){
                    if(ajax.readyState==4&&ajax.status==200){
                        var CellList = JSON.parse(ajax.responseText)
                        CellList.forEach(function(e){
                            for (var key in e) {
                                cellRender.renderCell(e['cellName'],key,e[key])
                            }
                        })
                    }
                }
            }
            break
        default:
            break
    }
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
            if(fontType=='merge'){
                cellRender.renderCell(null,'area',sheet.range)
            }else if(fontType=='border'){
                cellRender.renderCell(null,'frame',fontValue)
            }else{
                for(var i=cells.firstCellCol;i<=cells.lastCellCol;i++){
                    for(var j=cells.firstCellRow;j<=cells.lastCellRow;j++){
                        var id = String.fromCharCode(i) + '_' + j
                        cellRender.renderCell(id,fontType,fontValue)
                    }
                }
            }
        }
    }
}



module.exports.ToolEventHandler = ToolEventHandler

/***/ }),
/* 13 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * 工具栏渲染
 */
var config = __webpack_require__(0)
var ToolRender=function(sheet){
    this.sheet=sheet
}
ToolRender.prototype.init=function(toolDiv,width,height){
    var ToolEventBinderModule = __webpack_require__(11)
    var ToolEventBinder=ToolEventBinderModule.ToolEventBinder
    var toolEventBinder=new ToolEventBinder(this.sheet)

    toolDiv.style.width = width + 'px'
    toolDiv.style.height = height + 'px'
    toolDiv.style.backgroundColor = '#aaaaaa'
    var html = []
    if(config.WSConfig.isPreview){
        html = ["Edit","Down"]
    }else{

        //sheetToolDiv.innerHTML = "&#xe900;&#xe901;&#xe14d;&#xe14e;&#xe14f;"
        html = ["&#xe900","&#xe901","&#xe14d","&#xe14e","&#xe14f","autoLF","Preview","Init"]
    }
//渲染工具栏具体button按钮
    html.forEach(function(innerhtml){
        var buttonDiv = document.createElement('div')
        buttonDiv.id=innerhtml
        buttonDiv.style.display = "inline"
        buttonDiv.innerHTML=innerhtml
        buttonDiv.style.cursor='pointer'
        buttonDiv.style.marginLeft='10px'

        buttonDiv.onclick = function(){
            toolEventBinder.buttonClick(this)
        }
        toolDiv.appendChild(buttonDiv)
    })
    if(!config.WSConfig.isPreview){
        renderFont(toolEventBinder,toolDiv)
        renderColorAndBackgroundColor(toolEventBinder,toolDiv)
        renderColorSelect(toolEventBinder,toolDiv)

        var fileInput=document.createElement('input')
        fileInput.type='file'
        fileInput.style.marginLeft='10px'
        toolEventBinder.initFileInput(fileInput)
        toolDiv.appendChild(fileInput)
    }


}

function renderFont(toolEventBinder,toolDiv){
    //字体
    var fontFamilySelect=document.createElement('select')
    fontFamilySelect.style.marginLeft='10px'
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
    fontSizeSelect.style.marginLeft='10px'
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
    fontWeightDiv.style.marginLeft='10px'
    fontWeightDiv.style.cursor='pointer'
    fontWeightDiv.style.fontWeight='bold'
    toolEventBinder.initFontWeight(fontWeightDiv)
    toolDiv.appendChild(fontWeightDiv)

    //斜体
    var fontStyleDiv=document.createElement('div')
    fontStyleDiv.innerHTML='I'
    fontStyleDiv.style.display='inline'
    fontStyleDiv.style.marginLeft='10px'
    fontStyleDiv.style.cursor='pointer'
    fontStyleDiv.style.fontStyle='italic'
    fontStyleDiv.style.fontFamily='Verdana'
    toolEventBinder.initFontStyle(fontStyleDiv)
    toolDiv.appendChild(fontStyleDiv)

    //合并单元格
    var mergeDiv=document.createElement('div')
    mergeDiv.style.display = "inline"
    mergeDiv.innerHTML='&#xe157'
    mergeDiv.style.marginLeft='10px'
    mergeDiv.style.cursor='pointer'
    toolEventBinder.initMerge(mergeDiv)
    toolDiv.appendChild(mergeDiv)

    //单元格边框
    var fontBorderSelect=document.createElement('select')
    fontBorderSelect.style.marginLeft='10px'
    var fontBorderOption=['-----','左边框','上边框','右边框','下边框','无边框','外侧边框','全边框']
    fontBorderOption.forEach(function(o){
        var option=document.createElement('option')
        option.innerHTML=o
        fontBorderSelect.appendChild(option)
    })
    toolEventBinder.initFontBorder(fontBorderSelect)
    toolDiv.appendChild(fontBorderSelect)

    //字体位置
    var textAlignSelect=document.createElement('select')
    textAlignSelect.style.marginLeft='10px'
    var textAlignOption=['-----','左对齐','居中对齐','右对齐']
    textAlignOption.forEach(function(o){
        var option=document.createElement('option')
        option.innerHTML=o
        textAlignSelect.appendChild(option)
    })
    toolEventBinder.initTextAlign(textAlignSelect)
    toolDiv.appendChild(textAlignSelect)
}
function renderColorAndBackgroundColor(toolEventBinder,toolDiv){
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
/* 14 */
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
/* 15 */
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
    var ToolRenderModule = __webpack_require__(13)
    var ToolRender = ToolRenderModule.ToolRender
    var toolRender=new ToolRender(ws.Sheet)
    var tool = ws.Tool
    var toolDiv = document.createElement('div')
    toolRender.init(toolDiv,tool.width,tool.height)
    WSDiv.appendChild(toolDiv)

    //sheet
    var SheetRenderModule= __webpack_require__(9)
    var SheetRender= SheetRenderModule.SheetRender
    var sheetRender=new SheetRender(ws.Sheet)
    var sheet=ws.Sheet
    var sheetDiv=document.createElement('div')
    sheetRender.init(sheetDiv)
    WSDiv.appendChild(sheetDiv)
}
module.exports.WSRender = WSRender

/***/ }),
/* 16 */
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