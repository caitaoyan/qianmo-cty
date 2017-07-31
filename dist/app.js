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
/******/ 	return __webpack_require__(__webpack_require__.s = 9);
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
        addCol: '插列', addRow: '插列', menu: '&#xe3c7', close: '&#xe5cd',Init:'Init'
    },
    styleHtml: ['font', 'bold', 'italic', 'textColor', 'fillColor', 'align', 'border', 'brush', 'cleanText',
        'cleanStyle', 'cleanAll'],
    toolHtml: ['copy', 'paste', 'cut', 'sort', 'find', 'wrapText', 'func', 'sigma', 'mergeCell', 'splitCell',
        'multiLine', 'addCol','addRow'],
    dataHtml: ['数据源', '数据集'],
    defaultHtml: ['undo', 'redo', 'preview', 'Init'],
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
	arrowIcon: 'arrow',

    id: 'sliderBarDiv',

    iconHtmlMap: {
		open: '&#xe3c7',
        close: '&#xe5cd',
        arrow: '&#xe313'
    },

    sliderPaneTitle: ['单元格属性','表格属性', '公式','数据源','数据集'],
    sliderPaneId: ['cellAttrPane', 'sheetAttrPane','funcPane', 'dataSrcPane', 'dataTarPane'],
    sliderPaneConfig:{
        cellAttr: {
            title: '单元格属性',
            paneId: 'cellAttrPane',
            titleId: 'cellAttrTitle',
            arrowId: 'cellAttrArrow',
            contentId: 'cellAttrContent'
        },
		sheetAttr: {
			title: '表格属性',
			paneId: 'sheetAttrPane',
			titleId: 'sheetAttrTitle',
			arrowId: 'sheetAttrArrow',
			contentId: 'sheetAttrContent'
		},
		func: {
			title: '公式',
			paneId: 'funcPane',
			titleId: 'funcTitle',
			arrowId: 'funcArrow',
			contentId: 'funcContent'
		},
		dataSrc: {
			title: '数据源',
			paneId: 'dataSrcPane',
			titleId: 'dataSrcTitle',
			arrowId: 'dataSrcArrow',
			contentId: 'dataSrcContent'
		},
		dataTar: {
			title: '数据集',
			paneId: 'dataTarPane',
			titleId: 'dataTarTitle',
			arrowId: 'dataTarArrow',
			contentId: 'dataTarContent'
		}
    },

    toggleId: 'toggleDiv'
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

    rowNum: 200,
    colNum: 25
}

var SheetTableDivConfig = {
    id: 'sheetTableDiv'
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
module.exports.SheetTableDivConfig = SheetTableDivConfig


/***/ }),
/* 1 */
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
	'height:100%;' +
	'overflow: hidden'

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
	'position: absolute;' +
	'right :0;' +
	'top: 0;' +
	'width: 0;' +
	'height: 100%;' +
	'background-color: gray;' +
	'transition: all 0.5s ease;' +
	'overflow: auto;'

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

var PaneStyle =
	'padding: 10px;' +
	'top: 0;' +
	'left: 0;'

var PaneTitleStyle =
	'border-bottom: solid 1px rgba(255,255,255,0.3);' +
	'color: rgba(255,255,255,0.9);' +
	'margin-bottom: 10px;' +
	'font-size: 0.9em;' +
	'user-select: none;' +
	'cursor: pointer;'

var PaneContentCloseStyle =
	'height: 0;' +
	'width: 100%;' +
	'min-height: 0px;' +
	'background-color: rgba(255,255,255,0.1);' +
	'border-radius: 8px;' +
	'transition: all 0.5s ease;'

var PaneContentOpenStyle =
	'width: 100%;' +
	'height: auto;' +
	'min-height: 200px;' +
	'background-color: rgba(255,255,255,0.1);' +
	'border-radius: 8px;' +
	'transition: all 0.5s ease;'

var ArrowDownStyle =
	'display: inline-block;' +
	'right: 0;' +
	'transition: all 0.5s ease;'

var ArrowUpStyle =
	'display: inline-block;' +
	'right: 0;' +
	'transform: rotate(180deg);' +
	'transition: all 0.5s ease'

var SliderTableStyle =
	'width: 100%;' +
	'border: none;'

var SliderOddTrStyle =
	'border: none;' +
	'background-color: rgba(0,0,0,0.1);'

var SliderEvenTrStyle=
	'border: none;' +
	'background-color: rgba(0,0,0,0);'

var  SliderTdStyle =
	'width: 50%;' +
	'border: none;' +
	'padding-left: 5px;' +
	'color: whiteSmoke;' +
	'font-size: 0.9em;' +
	'line-height: 1.8em'

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

module.exports.PaneStyle = PaneStyle
module.exports.PaneTitleStyle = PaneTitleStyle

module.exports.PaneContentCloseStyle = PaneContentCloseStyle
module.exports.PaneContentOpenStyle = PaneContentOpenStyle

module.exports.ArrowDownStyle = ArrowDownStyle
module.exports.ArrowUpStyle = ArrowUpStyle

module.exports.SliderTableStyle = SliderTableStyle
module.exports.SliderOddTrStyle = SliderOddTrStyle
module.exports.SliderEvenTrStyle = SliderEvenTrStyle
module.exports.SliderTdStyle = SliderTdStyle

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
/* 2 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * Created by Ian on 17/7/22.
 */

var style  = __webpack_require__(1)
var config = __webpack_require__(0)

var SliderBarRender = function (sheet, sheetDiv) {
	this.sheet = sheet
	this.sheetDiv = sheetDiv

	// alert(this.sheetDiv)
}

SliderBarRender.prototype.init = function (sliderBarDiv, sheetTableDiv) {

	var toggleDiv = document.createElement('div')
	toggleDiv.id = config.SlideBarConfig.toggleId
	toggleDiv.isOpen = false
	this.sheetDiv.appendChild(toggleDiv)

	renderBar(sliderBarDiv, sheetTableDiv, toggleDiv)
	renderToggle(toggleDiv)

	var sliderPaneConfig = config.SlideBarConfig.sliderPaneConfig
	var sliderPanes = {}

	for(var paneConfig in sliderPaneConfig){
		var paneDiv = document.createElement('div')
		paneDiv.style = style.PaneStyle

		paneDiv.id = sliderPaneConfig[paneConfig].paneId
		// paneDiv.config = paneConfig
		sliderPanes[paneConfig] = paneDiv

		sliderBarDiv.appendChild(paneDiv)
	}
	renderPanes(sliderPanes)

	// var isOpen = this.isOpen
	// var sheetDiv = this.sheetDiv

	toggleDiv.onclick = function () {
		toggleDiv.isOpen = !toggleDiv.isOpen

		renderBar(sliderBarDiv, sheetTableDiv, toggleDiv)
		renderToggle(toggleDiv)
	}
}

function renderToggle(toggleDiv) {

	var iconHtmlMap = config.SlideBarConfig.iconHtmlMap
	var toggleOpenIcon = config.SlideBarConfig.toggleOpenIcon
	var toggleCloseIcon = config.SlideBarConfig.toggleCloseIcon

	var isOpen = toggleDiv.isOpen
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
	// sheetDiv.appendChild(toggleDiv)

	setTimeout(function () {
		toggleDiv.style.opacity = '0.5'
	}, 5000)
}

function renderBar(sliderBarDiv, sheetTableDiv, toggleDiv) {

	var isOpen = toggleDiv.isOpen

	sliderBarDiv.style = style.SliderBarStyle

	if(isOpen){
		sheetTableDiv.style.width = '75%'
		sliderBarDiv.style.width = '25%'
	}else{
		sheetTableDiv.style.width = '100%'
		sliderBarDiv.style.width = '0'
	}
}

function renderPanes(sliderPanes) {

	var sliderPaneConfig = config.SlideBarConfig.sliderPaneConfig

	for(var key in sliderPanes) {
		var isOpen = false
		var paneDiv = sliderPanes[key]
		var paneConfig = key

		var paneTitleDiv = document.createElement('div')
		paneTitleDiv.innerHTML = sliderPaneConfig[paneConfig].title
		paneTitleDiv.id = sliderPaneConfig[paneConfig].titleId
		paneTitleDiv.style = style.PaneTitleStyle
		paneTitleDiv.isOpen = isOpen

		var arrowDiv = document.createElement('div')
		arrowDiv.innerHTML = config.SlideBarConfig.iconHtmlMap[config.SlideBarConfig.arrowIcon]
		arrowDiv.id = sliderPaneConfig[paneConfig].arrowId
		paneTitleDiv.appendChild(arrowDiv)

		var paneContentDiv = document.createElement('div')
		paneContentDiv.id = sliderPaneConfig[paneConfig].contentId

		renderPane(arrowDiv, paneContentDiv, isOpen)

		//第一次写闭包！！！
		paneTitleDiv.onclick = function (arrowDiv, paneContentDiv, paneTitleDiv) {
			return function () {
				paneTitleDiv.isOpen = !paneTitleDiv.isOpen
				renderPane(arrowDiv, paneContentDiv, paneTitleDiv)
			}
		}(arrowDiv, paneContentDiv, paneTitleDiv)
		// paneContentDiv.style = style.PaneContentOpenStyle


		paneDiv.appendChild(paneTitleDiv)
		paneDiv.appendChild(paneContentDiv)

		// if(paneDiv.id === 'cellAttrPane'){
		//
		// }else if(paneDiv.id === 'sheetAttrPane'){
		//
		// }else if(paneDiv.id === 'funcPane'){
		//
		// }else if(paneDiv.id === 'dataSrcPane'){
		//
		// }else{
		//
		// }
	}
}

function renderPane(arrowDiv, paneContentDiv, paneTitleDiv) {

	var isOpen = paneTitleDiv.isOpen

	if(isOpen){
		paneContentDiv.style = style.PaneContentOpenStyle
		arrowDiv.style = style.ArrowUpStyle

		if(paneContentDiv.firstChild){
			setTimeout(function () {paneContentDiv.firstChild.style.display = 'table'}, 200)
		}
	}else{
		paneContentDiv.style = style.PaneContentCloseStyle
		arrowDiv.style = style.ArrowDownStyle

		if(paneContentDiv.firstChild){paneContentDiv.firstChild.style.display = 'none'}
	}
}

// SliderBarRender.prototype
//
// function addFunc() {
//
// }

module.exports.SliderBarRender = SliderBarRender

module.exports.renderBar = renderBar

module.exports.renderPane = renderPane

module.exports.renderToggle = renderToggle

/***/ }),
/* 3 */
/***/ (function(module, exports, __webpack_require__) {

//鼠标和键盘事件

// var UtilModule=require('Util')
// var Util=UtilModule.Util
// var util=new Util()

var config = __webpack_require__(0)

var SheetRenderModule = __webpack_require__(4)
var SheetRender = SheetRenderModule.SheetRender
var SliderBarHandler = __webpack_require__(14).SliderBarHandler

var Formula = __webpack_require__(6).Formula
var formula = new Formula()

var SheetEventHandler = function (sheet) {

    this.sheet = sheet
    this.sheetRender = new SheetRender(sheet)
	this.sliderBarHandler = new SliderBarHandler(sheet)
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
        if (this.sheet.isDraging) this.setCellBackgroundColor('#f69')
        else this.setCellBackgroundColor('#69f')
    }
}
/**
 * 鼠标抬起事件
 * @param element
 */
SheetEventHandler.prototype.mouseUp = function (element) {
    if (this.sheet.isDraging) {
        if (!this.sheet.cells[this.sheet.firstCellid]) {
            this.sheet.setAttr('content', '')
        } else {
            this.sheet.setAttr('content', this.sheet.cells[this.sheet.firstCellid].content)
        }

        this.sheet.render()
    }
    if (this.sheet.isMouseDown || this.sheet.isDraging) {
        if (element.id) {
            var dragBar = document.getElementById('dragBar')
            dragBar.style.left = element.offsetLeft +element.offsetWidth-document.getElementById(config.SheetTableDivConfig.id).scrollLeft - 8 + 'px'
            dragBar.style.top = element.offsetTop +element.offsetHeight -document.getElementById(config.SheetTableDivConfig.id).scrollTop  - 8 + 'px'
            dragBar.style.display = 'block'
            this.sheet.lastCellid = element.id
            this.sheet.editCells = this.sheet.getColAndRow()
            this.setCellBackgroundColor('#69f')

			this.sliderBarHandler.autoOpen()
			this.sliderBarHandler.addCellAttr()

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
        input.style.left = element.offsetLeft -document.getElementById(config.SheetTableDivConfig.id).scrollLeft + 'px'
        input.style.top = element.offsetTop -document.getElementById(config.SheetTableDivConfig.id).scrollTop+ 'px'
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
        if(input.value[0] === '='){
            var cc = formula.ParseFormulaIntoTokens(input.value.substring(1))
            var dd = formula.evaluate_parsed_formula(cc)
            this.sheet.setAttr('content', dd.value)
        }
        else this.sheet.setAttr('content', input.value)
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
        console.log(event.which)
        switch (event.which) {
            case 17:
                this.sheet.isCtrlDown = true
                return
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
                if (row > 0) {
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
            case 18://alt
                break
            case 67://ctrl+c
                if(this.sheet.isCtrlDown){
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
                }
                break
            case 86://ctrl+v
                if(this.sheet.isCtrlDown){
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

                }
                this.sheet.isKeyDown = false
                break
            case 88://ctrl+x
                if(this.sheet.isCtrlDown){
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
                }
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

SheetEventHandler.prototype.formulaButonClick = function () {
    var fname = document.getElementById('funcNameSelect').value
    this.sheet.setAttr('content','='+fname+'(')
    this.sheet.render()
}

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
 * 渲染表格
 */

var config = __webpack_require__(0)
var style = __webpack_require__(1)
var SliderBarRender = __webpack_require__(2).SliderBarRender
var FuncListRender = __webpack_require__(11).FuncListRender
var funcListRender = new FuncListRender()
// config.CellConfig
var SheetRender = function (sheet) {
	this.sheet=sheet
}

SheetRender.prototype.init = function (sheetDiv) {

	sheetDiv.style = style.SheetDivStyle

	var sheetTableDiv = document.createElement('div')
	sheetTableDiv.style = style.SheetTableDivStyle
	sheetTableDiv.id = config.SheetTableDivConfig.id

	var sheetTable = document.createElement('table')
    sheetTable.style = style.SheetTableStyle
	sheetTable.cellSpacing = '10px'

	var sliderBarDiv = document.createElement('div')
	sliderBarDiv.id = config.SlideBarConfig.id
	var sliderBarRender = new SliderBarRender(this.sheet, sheetDiv)
	sliderBarRender.init(sliderBarDiv, sheetTableDiv)

	//sheetTable.cellSpacing = '10px'//
	sheetTable.id = 'sheetTable'
    sheetDiv.appendChild(sheetTable)

    sheetTableDiv.appendChild(sheetTable)

	sheetDiv.appendChild(sheetTableDiv)
	sheetDiv.appendChild(sliderBarDiv)

    // var cellPropDiv = document.createElement("div")
    // cellPropDiv.id = config.CellPropConfig.id
	//

    // cellPropDiv.style = style.CellPropDivStyle

    // sheetDiv.appendChild(cellPropDiv)

	//表格的具体渲染
	this.renderSheet(this.sheet, sheetTable)
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
SheetRender.prototype.renderSheet = function(sheet, sheetTable) {
    for (var child = sheetTable.firstChild; child != null; child = sheetTable.firstChild) {
        sheetTable.removeChild(child);
    }
    var gridHeader = createHeader(sheet.rowNum, sheet.colNum)

    var rowNum = gridHeader.length
    var colNum = gridHeader[0].length

    var rowHead = document.createElement('tr')
    sheetTable.appendChild(rowHead)
    //渲染第一行表头
    for (var i = 0; i < colNum; i++) {
        var rowHeadTH = document.createElement('th')
        // rowHeadTH.style.height = '20px'
        rowHeadTH.id = gridHeader[0][i] + '_0'
        rowHead.appendChild(rowHeadTH)

        var rowHeadDiv = document.createElement('div')
        rowHeadTH.appendChild(rowHeadDiv)
        // 第一行第一列特殊处理
        if (i === 0) {
            //rowHeadTH.style.width = '29px'
        } else {
            rowHeadDiv.innerHTML = gridHeader[0][i]
            rowHeadDiv.style.width = config.CellConfig.width + 'px'
        }
    }

    var SheetEventBinderModule = __webpack_require__(13)
    var SheetEventBinder = SheetEventBinderModule.SheetEventBinder
    var sheetEventBinder = new SheetEventBinder(sheet)

    //渲染除第一行外单元格
    for (var j = 1; j < rowNum; j++) {
        var rowTR = document.createElement('tr')
        sheetTable.appendChild(rowTR)

        for (var k = 0; k < colNum; k++) {
            //表格第一列特殊处理
            if (k === 0) {
                var rowTH = document.createElement('th')
                rowTH.id = '@_' + gridHeader[j][0]
                rowTR.appendChild(rowTH)
                var rowDiv = document.createElement('div')
                rowDiv.style.height = 25 + 'px'
                rowTH.appendChild(rowDiv)
                rowDiv.innerHTML = gridHeader[j][k]
            } else {
                var rowTD = document.createElement('td')
                rowTD.id = gridHeader[0][k] + "_" + j
                rowTR.appendChild(rowTD)
                if (!sheet.isPreview) {
                    sheetEventBinder.initRowTD(rowTD)
                }
                var rowDiv = document.createElement('div')
                rowTD.className = "noWrap"
                rowTD.appendChild(rowDiv)
            }
        }
    }
    if (!sheet.isPreview) {
        sheetTable.style.borderCollapse = 'collapse'
        sheetTable.style.border = '1px solid #ccc'
        var input = document.createElement('input')
        input.style = style.InputStyle
        input.id = config.InputConfig.id

        sheetEventBinder.initInput(input)


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
        multiLine.style.left = screen.width / 2 - 150 + 'px'
        multiLine.style.top = screen.height / 2 - 300 + 'px'
        multiLine.style.width = 400 + 'px'
        multiLine.style.height = 400 + 'px'
        multiLine.value = '';
        multiLine.id = 'multiLine1'
        if (!config.WSConfig.isPreview) {
            sheetEventBinder.initMultiLine(multiLine)
        }
        sheetTable.appendChild(multiLine)

        var funcListDiv = document.createElement('div')
        funcListRender.init( funcListDiv)
        var formulaButton = document.createElement('input')
        formulaButton.type = 'button'
        formulaButton.value = 'Paste'
        sheetEventBinder.initFormulaButton(formulaButton)
        funcListDiv.appendChild(formulaButton)
        sheetTable.appendChild(funcListDiv)
        funcListDiv.id = 'funcListDiv'
        funcListDiv.onblur = function () {
            this.style.display = 'none'
        }
        funcListDiv.style = 'position: absolute;left: 30px;top: 20px;display:none'
        funcListDiv.style.left = screen.width / 2 - 250 + 'px'
        funcListDiv.style.top = screen.height / 2 - 380 + 'px'

        var dragBar = document.createElement('div')
        dragBar.style = 'position: absolute;'
        dragBar.style.backgroundColor = 'yellow'
        dragBar.style.width = 8 + 'px'
        dragBar.style.height = 8 + 'px'
        dragBar.style.display = 'none'
        dragBar.id = 'dragBar'
        dragBar.webkitUserDrag = 'true'
        if (!config.WSConfig.isPreview) {
            sheetEventBinder.initDragBar(dragBar)
        }
        sheetTable.appendChild(dragBar)
    }

}

module.exports.SheetRender = SheetRender

/***/ }),
/* 5 */
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
	// this.value = null

	this.show = true

	this.bold = false
	this.italic = false

	this.colSpan = 1
	this.rowSpan = 1
}

module.exports.Cell = Cell

/***/ }),
/* 6 */
/***/ (function(module, exports) {

var Formula = function () {

    //
// Formula constants for parsing:
//

    this.ParseState = {
        num: 1,
        alpha: 2,
        coord: 3,
        string: 4,
        stringquote: 5,
        numexp1: 6,
        numexp2: 7,
        alphanumeric: 8,
        specialvalue: 9
    };

    this.TokenType = {num: 1, coord: 2, op: 3, name: 4, error: 5, string: 6, space: 7};

    this.CharClass = {
        num: 1,
        numstart: 2,
        op: 3,
        eof: 4,
        alpha: 5,
        incoord: 6,
        error: 7,
        quote: 8,
        space: 9,
        specialstart: 10
    };

    this.CharClassTable = {
        " ": 9,
        "!": 3,
        '"': 8,
        "#": 10,
        "$": 6,
        "%": 3,
        "&": 3,
        "(": 3,
        ")": 3,
        "*": 3,
        "+": 3,
        ",": 3,
        "-": 3,
        ".": 2,
        "/": 3,
        "0": 1,
        "1": 1,
        "2": 1,
        "3": 1,
        "4": 1,
        "5": 1,
        "6": 1,
        "7": 1,
        "8": 1,
        "9": 1,
        ":": 3,
        "<": 3,
        "=": 3,
        ">": 3,
        "A": 5,
        "B": 5,
        "C": 5,
        "D": 5,
        "E": 5,
        "F": 5,
        "G": 5,
        "H": 5,
        "I": 5,
        "J": 5,
        "K": 5,
        "L": 5,
        "M": 5,
        "N": 5,
        "O": 5,
        "P": 5,
        "Q": 5,
        "R": 5,
        "S": 5,
        "T": 5,
        "U": 5,
        "V": 5,
        "W": 5,
        "X": 5,
        "Y": 5,
        "Z": 5,
        "^": 3,
        "_": 5,
        "a": 5,
        "b": 5,
        "c": 5,
        "d": 5,
        "e": 5,
        "f": 5,
        "g": 5,
        "h": 5,
        "i": 5,
        "j": 5,
        "k": 5,
        "l": 5,
        "m": 5,
        "n": 5,
        "o": 5,
        "p": 5,
        "q": 5,
        "r": 5,
        "s": 5,
        "t": 5,
        "u": 5,
        "v": 5,
        "w": 5,
        "x": 5,
        "y": 5,
        "z": 5
    };

    this.UpperCaseTable = {
        "a": "A",
        "b": "B",
        "c": "C",
        "d": "D",
        "e": "E",
        "f": "F",
        "g": "G",
        "h": "H",
        "i": "I",
        "j": "J",
        "k": "K",
        "l": "L",
        "m": "M",
        "n": "N",
        "o": "O",
        "p": "P",
        "q": "Q",
        "r": "R",
        "s": "S",
        "t": "T",
        "u": "U",
        "v": "V",
        "w": "W",
        "x": "X",
        "y": "Y",
        "z": "Z"
    }

    this.SpecialConstants = { // names that turn into constants for name lookup
        "#NULL!": "0,e#NULL!", "#NUM!": "0,e#NUM!", "#DIV/0!": "0,e#DIV/0!", "#VALUE!": "0,e#VALUE!",
        "#REF!": "0,e#REF!", "#NAME?": "0,e#NAME?"
    };


    // Operator Precedence table
    //
    // 1- !, 2- : ,, 3- M P, 4- %, 5- ^, 6- * /, 7- + -, 8- &, 9- < > = G(>=) L(<=) N(<>),
    // Negative value means Right Associative

    this.TokenPrecedence = {
        "!": 1,
        ":": 2, ",": 2,
        "M": -3, "P": -3,
        "%": 4,
        "^": 5,
        "*": 6, "/": 6,
        "+": 7, "-": 7,
        "&": 8,
        "<": 9, ">": 9, "G": 9, "L": 9, "N": 9
    };

    // Convert one-char token text to input text:

    this.TokenOpExpansion = {'G': '>=', 'L': '<=', 'M': '-', 'N': '<>', 'P': '+'};

    //
    // Information about the resulting value types when doing operations on values (used by LookupResultType)
    //
    // Each object entry is an object with specific types with result type info as follows:
    //
    //    'type1a': '|type2a:resulta|type2b:resultb|...
    //    Type of t* or n* matches any of those types not listed
    //    Results may be a type or the numbers 1 or 2 specifying to return type1 or type2


    this.TypeLookupTable = {
        unaryminus: {'n*': '|n*:1|', 'e*': '|e*:1|', 't*': '|t*:e#VALUE!|', 'b': '|b:n|'},
        unaryplus: {'n*': '|n*:1|', 'e*': '|e*:1|', 't*': '|t*:e#VALUE!|', 'b': '|b:n|'},
        unarypercent: {'n*': '|n:n%|n*:n|', 'e*': '|e*:1|', 't*': '|t*:e#VALUE!|', 'b': '|b:n|'},
        plus: {
            'n%': '|n%:n%|nd:n|nt:n|ndt:n|n$:n|n:n|n*:n|b:n|e*:2|t*:e#VALUE!|',
            'nd': '|n%:n|nd:nd|nt:ndt|ndt:ndt|n$:n|n:nd|n*:n|b:n|e*:2|t*:e#VALUE!|',
            'nt': '|n%:n|nd:ndt|nt:nt|ndt:ndt|n$:n|n:nt|n*:n|b:n|e*:2|t*:e#VALUE!|',
            'ndt': '|n%:n|nd:ndt|nt:ndt|ndt:ndt|n$:n|n:ndt|n*:n|b:n|e*:2|t*:e#VALUE!|',
            'n$': '|n%:n|nd:n|nt:n|ndt:n|n$:n$|n:n$|n*:n|b:n|e*:2|t*:e#VALUE!|',
            'nl': '|n%:n|nd:n|nt:n|ndt:n|n$:n|n:n|n*:n|b:n|e*:2|t*:e#VALUE!|',
            'n': '|n%:n|nd:nd|nt:nt|ndt:ndt|n$:n$|n:n|n*:n|b:n|e*:2|t*:e#VALUE!|',
            'b': '|n%:n%|nd:nd|nt:nt|ndt:ndt|n$:n$|n:n|n*:n|b:n|e*:2|t*:e#VALUE!|',
            't*': '|n*:e#VALUE!|t*:e#VALUE!|b:e#VALUE!|e*:2|',
            'e*': '|e*:1|n*:1|t*:1|b:1|'
        },
        concat: {
            't': '|t:t|th:th|tw:tw|tl:t|tr:tr|t*:2|e*:2|',
            'th': '|t:th|th:th|tw:t|tl:th|tr:t|t*:t|e*:2|',
            'tw': '|t:tw|th:t|tw:tw|tl:tw|tr:tw|t*:t|e*:2|',
            'tl': '|t:tl|th:th|tw:tw|tl:tl|tr:tr|t*:t|e*:2|',
            't*': '|t*:t|e*:2|',
            'e*': '|e*:1|n*:1|t*:1|'
        },
        oneargnumeric: {'n*': '|n*:n|', 'e*': '|e*:1|', 't*': '|t*:e#VALUE!|', 'b': '|b:n|'},
        twoargnumeric: {
            'n*': '|n*:n|t*:e#VALUE!|e*:2|',
            'e*': '|e*:1|n*:1|t*:1|',
            't*': '|t*:e#VALUE!|n*:e#VALUE!|e*:2|'
        },
        propagateerror: {'n*': '|n*:2|e*:2|', 'e*': '|e*:2|', 't*': '|t*:2|e*:2|', 'b': '|b:2|e*:2|'}
    };
    this.scc = {
        s_parseerrexponent: "Improperly formed number exponent",
        s_parseerrchar: "Unexpected character in formula",
        s_parseerrstring: "Improperly formed string",
        s_parseerrspecialvalue: "Improperly formed special value",
        s_parseerrtwoops: "Error in formula (two operators inappropriately in a row)",
        s_parseerrmissingopenparen: "Missing open parenthesis in list with comma(s). ",
        s_parseerrcloseparennoopen: "Closing parenthesis without open parenthesis. ",
        s_parseerrmissingcloseparen: "Missing close parenthesis. ",
        s_parseerrmissingoperand: "Missing operand. ",
        s_parseerrerrorinformula: "Error in formula.",
        s_calcerrerrorvalueinformula: "Error value in formula",
        s_parseerrerrorinformulabadval: "Error in formula resulting in bad value",
        s_formularangeresult: "Formula results in range value:",
        s_calcerrnumericnan: "Formula results in an bad numeric value",
        s_calcerrnumericoverflow: "Numeric overflow",
        s_sheetunavailable: "Sheet unavailable:", // when FindSheetInCache returns null
        s_calcerrcellrefmissing: "Cell reference missing when expected.",
        s_calcerrsheetnamemissing: "Sheet name missing when expected.",
        s_circularnameref: "Circular name reference to name",
        s_calcerrunknownname: "Unknown name",
        s_calcerrincorrectargstofunction: "Incorrect arguments to function",
        s_sheetfuncunknownfunction: "Unknown function",
        s_sheetfunclnarg: "LN argument must be greater than 0",
        s_sheetfunclog10arg: "LOG10 argument must be greater than 0",
        s_sheetfunclogsecondarg: "LOG second argument must be numeric greater than 0",
        s_sheetfunclogfirstarg: "LOG first argument must be greater than 0",
        s_sheetfuncroundsecondarg: "ROUND second argument must be numeric",
        s_sheetfuncddblife: "DDB life must be greater than 1",
        s_sheetfuncslnlife: "SLN life must be greater than 1",
    }


    this.FunctionClassList = {
        //'statistics': [],
        'lookup': [],
        'datetime': [],
        'financial': [],
        'test': [],
        'math': [],
        'text': [],
        'stat': []
    }


    this.funcParem = {
        'CHOOSE': 'choose',
        'COLUMNS': 'range',
        'ROWS': 'range',
        'INDEX': 'index',
        'HLOOKUP': 'hlookup',
        'MATCH': 'match',
        'VLOOKUP': 'vlookup',
        'TODAY': '',
        'HOUR': 'v',
        'MINUTE': 'v',
        'SECOND': 'v',
        'DAY': 'v',
        'MONTH': 'v',
        'YEAR': 'v',
        'WEEKDAY': 'weekday',
        'TIME': 'hms',
        'DATE': 'date',
        'DDB': 'ddb',
        'SLN': 'csl',
        'SYD': 'cslp',
        'FV': 'fv',
        'NPER': 'nper',
        'PMT': 'pmt',
        'PV': 'pv',
        'RATE': 'rate',
        'NPV': 'npv',
        'IRR': 'irr',
        'AND': 'vn',
        'OR': 'vn',
        'NOT': 'v',
        'FALSE': '',
        'NA': '',
        'NOW': '',
        'PI': '',
        'TRUE': '',
        'T': 'v',
        'VALUE': 'v',
        'ISBLANK': 'v',
        'ISERR': 'v',
        'ISERROR': 'v',
        'ISLOGICAL': 'v',
        'ISNA': 'v',
        'ISNONTEXT': 'v',
        'ISNUMBER': 'v',
        'ISTEXT': 'v',
        'IF': 'iffunc',
        'ABS': 'v',
        'ACOS': 'v',
        'ASIN': 'v',
        'ATAN': 'v',
        'COS': 'v',
        'DEGREES': 'v',
        'EVEN': 'v',
        'EXP': 'v',
        'FACT': 'v',
        'INT': 'v',
        'LN': 'v',
        'LOG10': 'v',
        'ODD': 'v',
        'RADIANS': 'v',
        'SIN': 'v',
        'SQRT': 'v',
        'TAN': 'v',
        'ATAN2': 'xy',
        'MOD': '',
        'POWER': '',
        'TRUNC': 'valpre',
        'LOG': 'log',
        'ROUND': 'vp',
        'N': 'v',
        'FIND': 'find',
        'LEFT': 'tc',
        'LEN': 'txt',
        'LOWER': 'txt',
        'MID': 'mid',
        'PROPER': 'v',
        'REPLACE': 'replace',
        'REPT': 'tc',
        'RIGHT': 'tc',
        'SUBSTITUTE': 'subs',
        'TRIM': 'v',
        'UPPER': 'v',
        'EXACT': '',
        'COUNTIF': 'rangec',
        'SUMIF': 'sumif',
        'DAVERAGE': 'dfunc',
        'DCOUNT': 'dfunc',
        'DCOUNTA': 'dfunc',
        'DGET': 'dfunc',
        'DMAX': 'dfunc',
        'DMIN': 'dfunc',
        'DPRODUCT': 'dfunc',
        'DSTDEV': 'dfunc',
        'DSTDEVP': 'dfunc',
        'DSUM': 'dfunc',
        'DVAR': 'dfunc',
        'DVARP': 'dfunc',
        'AVERAGE': 'vn',
        'COUNT': 'vn',
        'COUNTA': 'vn',
        'COUNTBLANK': 'vn',
        'MAX': 'vn',
        'MIN': 'vn',
        'PRODUCT': 'vn',
        'STDEV': 'vn',
        'STDEVP': 'vn',
        'SUM': 'vn',
        'VAR': 'vn',
        'VARP': 'vn',


    }

    this.funcInfo = {
        s_fdef_ABS: 'Absolute value function. ',
        s_fdef_ACOS: 'Trigonometric arccosine function. ',
        s_fdef_AND: 'True if all arguments are true. ',
        s_fdef_ASIN: 'Trigonometric arcsine function. ',
        s_fdef_ATAN: 'Trigonometric arctan function. ',
        s_fdef_ATAN2: 'Trigonometric arc tangent function (result is in radians). ',
        s_fdef_AVERAGE: 'Averages the values. ',
        s_fdef_CHOOSE: 'Returns the value specified by the index. The values may be ranges of cells. ',
        s_fdef_COLUMNS: 'Returns the number of columns in the range. ',
        s_fdef_COS: 'Trigonometric cosine function (value is in radians). ',
        s_fdef_COUNT: 'Counts the number of numeric values, not blank, text, or error. ',
        s_fdef_COUNTA: 'Counts the number of non-blank values. ',
        s_fdef_COUNTBLANK: 'Counts the number of blank values. (Note: "" is not blank.) ',
        s_fdef_COUNTIF: 'Counts the number of number of cells in the range that meet the criteria. The criteria may be a value ("x", 15, 1+3) or a test (>25). ',
        s_fdef_DATE: 'Returns the appropriate date value given numbers for year, month, and day. For example: DATE(2006,2,1) for February 1, 2006. Note: In this program, day "1" is December 31, 1899 and the year 1900 is not a leap year. Some programs use January 1, 1900, as day "1" and treat 1900 as a leap year. In both cases, though, dates on or after March 1, 1900, are the same. ',
        s_fdef_DAVERAGE: 'Averages the values in the specified field in records that meet the criteria. ',
        s_fdef_DAY: 'Returns the day of month for a date value. ',
        s_fdef_DCOUNT: 'Counts the number of numeric values, not blank, text, or error, in the specified field in records that meet the criteria. ',
        s_fdef_DCOUNTA: 'Counts the number of non-blank values in the specified field in records that meet the criteria. ',
        s_fdef_DDB: 'Returns the amount of depreciation at the given period of time (the default factor is 2 for double-declining balance).   ',
        s_fdef_DEGREES: 'Converts value in radians into degrees. ',
        s_fdef_DGET: 'Returns the value of the specified field in the single record that meets the criteria. ',
        s_fdef_DMAX: 'Returns the maximum of the numeric values in the specified field in records that meet the criteria. ',
        s_fdef_DMIN: 'Returns the maximum of the numeric values in the specified field in records that meet the criteria. ',
        s_fdef_DPRODUCT: 'Returns the result of multiplying the numeric values in the specified field in records that meet the criteria. ',
        s_fdef_DSTDEV: 'Returns the sample standard deviation of the numeric values in the specified field in records that meet the criteria. ',
        s_fdef_DSTDEVP: 'Returns the standard deviation of the numeric values in the specified field in records that meet the criteria. ',
        s_fdef_DSUM: 'Returns the sum of the numeric values in the specified field in records that meet the criteria. ',
        s_fdef_DVAR: 'Returns the sample variance of the numeric values in the specified field in records that meet the criteria. ',
        s_fdef_DVARP: 'Returns the variance of the numeric values in the specified field in records that meet the criteria. ',
        s_fdef_EVEN: 'Rounds the value up in magnitude to the nearest even integer. ',
        s_fdef_EXACT: 'Returns "true" if the values are exactly the same, including case, type, etc. ',
        s_fdef_EXP: 'Returns e raised to the value power. ',
        s_fdef_FACT: 'Returns factorial of the value. ',
        s_fdef_FALSE: 'Returns the logical value "false". ',
        s_fdef_FIND: 'Returns the starting position within string2 of the first occurrence of string1 at or after "start". If start is omitted, 1 is assumed. ',
        s_fdef_FV: 'Returns the future value of repeated payments of money invested at the given rate for the specified number of periods, with optional present value (default 0) and payment type (default 0 = at end of period, 1 = beginning of period). ',
        s_fdef_HLOOKUP: 'Look for the matching value for the given value in the range and return the corresponding value in the cell specified by the row offset. If rangelookup is 1 (the default) and not 0, match if within numeric brackets (match<=value) instead of exact match. ',
        s_fdef_HOUR: 'Returns the hour portion of a time or date/time value. ',
        s_fdef_IF: 'Results in true-value if logical-expression is TRUE or non-zero, otherwise results in false-value. ',
        s_fdef_INDEX: 'Returns a cell or range reference for the specified row and column in the range. If range is 1-dimensional, then only one of rownum or colnum are needed. If range is 2-dimensional and rownum or colnum are zero, a reference to the range of just the specified column or row is returned. You can use the returned reference value in a range, e.g., sum(A1:INDEX(A2:A10,4)). ',
        s_fdef_INT: 'Returns the value rounded down to the nearest integer (towards -infinity). ',
        s_fdef_IRR: 'Returns the interest rate at which the cash flows in the range have a net present value of zero. Uses an iterative process that will return #NUM! error if it does not converge. There may be more than one possible solution. Providing the optional guess value may help in certain situations where it does not converge or finds an inappropriate solution (the default guess is 10%). ',
        s_fdef_ISBLANK: 'Returns "true" if the value is a reference to a blank cell. ',
        s_fdef_ISERR: 'Returns "true" if the value is of type "Error" but not "NA". ',
        s_fdef_ISERROR: 'Returns "true" if the value is of type "Error". ',
        s_fdef_ISLOGICAL: 'Returns "true" if the value is of type "Logical" (true/false). ',
        s_fdef_ISNA: 'Returns "true" if the value is the error type "NA". ',
        s_fdef_ISNONTEXT: 'Returns "true" if the value is not of type "Text". ',
        s_fdef_ISNUMBER: 'Returns "true" if the value is of type "Number" (including logical values). ',
        s_fdef_ISTEXT: 'Returns "true" if the value is of type "Text". ',
        s_fdef_LEFT: 'Returns the specified number of characters from the text value. If count is omitted, 1 is assumed. ',
        s_fdef_LEN: 'Returns the number of characters in the text value. ',
        s_fdef_LN: 'Returns the natural logarithm of the value. ',
        s_fdef_LOG: 'Returns the logarithm of the value using the specified base. ',
        s_fdef_LOG10: 'Returns the base 10 logarithm of the value. ',
        s_fdef_LOWER: 'Returns the text value with all uppercase characters converted to lowercase. ',
        s_fdef_MATCH: 'Look for the matching value for the given value in the range and return position (the first is 1) in that range. If rangelookup is 1 (the default) and not 0, match if within numeric brackets (match<=value) instead of exact match. If rangelookup is -1, act like 1 but the bracket is match>=value. ',
        s_fdef_MAX: 'Returns the maximum of the numeric values. ',
        s_fdef_MID: 'Returns the specified number of characters from the text value starting from the specified position. ',
        s_fdef_MIN: 'Returns the minimum of the numeric values. ',
        s_fdef_MINUTE: 'Returns the minute portion of a time or date/time value. ',
        s_fdef_MOD: 'Returns the remainder of the first value divided by the second. ',
        s_fdef_MONTH: 'Returns the month part of a date value. ',
        s_fdef_N: 'Returns the value if it is a numeric value otherwise an error. ',
        s_fdef_NA: 'Returns the #N/A error value which propagates through most operations. ',
        s_fdef_NOT: 'Returns FALSE if value is true, and TRUE if it is false. ',
        s_fdef_NOW: 'Returns the current date/time. ',
        s_fdef_NPER: 'Returns the number of periods at which payments invested each period at the given rate with optional future value (default 0) and payment type (default 0 = at end of period, 1 = beginning of period) has the given present value. ',
        s_fdef_NPV: 'Returns the net present value of cash flows (which may be individual values and/or ranges) at the given rate. The flows are positive if income, negative if paid out, and are assumed at the end of each period. ',
        s_fdef_ODD: 'Rounds the value up in magnitude to the nearest odd integer. ',
        s_fdef_OR: 'True if any argument is true ',
        s_fdef_PI: 'The value 3.1415926... ',
        s_fdef_PMT: 'Returns the amount of each payment that must be invested at the given rate for the specified number of periods to have the specified present value, with optional future value (default 0) and payment type (default 0 = at end of period, 1 = beginning of period). ',
        s_fdef_POWER: 'Returns the first value raised to the second value power. ',
        s_fdef_PRODUCT: 'Returns the result of multiplying the numeric values. ',
        s_fdef_PROPER: 'Returns the text value with the first letter of each word converted to uppercase and the others to lowercase. ',
        s_fdef_PV: 'Returns the present value of the given number of payments each invested at the given rate, with optional future value (default 0) and payment type (default 0 = at end of period, 1 = beginning of period). ',
        s_fdef_RADIANS: 'Converts value in degrees into radians. ',
        s_fdef_RATE: 'Returns the rate at which the given number of payments each invested at the given rate has the specified present value, with optional future value (default 0) and payment type (default 0 = at end of period, 1 = beginning of period). Uses an iterative process that will return #NUM! error if it does not converge. There may be more than one possible solution. Providing the optional guess value may help in certain situations where it does not converge or finds an inappropriate solution (the default guess is 10%). ',
        s_fdef_REPLACE: 'Returns text1 with the specified number of characters starting from the specified position replaced by text2. ',
        s_fdef_REPT: 'Returns the text repeated the specified number of times. ',
        s_fdef_RIGHT: 'Returns the specified number of characters from the text value starting from the end. If count is omitted, 1 is assumed. ',
        s_fdef_ROUND: 'Rounds the value to the specified number of decimal places. If precision is negative, then round to powers of 10. The default precision is 0 (round to integer). ',
        s_fdef_ROWS: 'Returns the number of rows in the range. ',
        s_fdef_SECOND: 'Returns the second portion of a time or date/time value (truncated to an integer). ',
        s_fdef_SIN: 'Trigonometric sine function (value is in radians) ',
        s_fdef_SLN: 'Returns the amount of depreciation at each period of time using the straight-line method. ',
        s_fdef_SQRT: 'Square root of the value ',
        s_fdef_STDEV: 'Returns the sample standard deviation of the numeric values. ',
        s_fdef_STDEVP: 'Returns the standard deviation of the numeric values. ',
        s_fdef_SUBSTITUTE: 'Returns text1 with the all occurrences of oldtext replaced by newtext. If "occurrence" is present, then only that occurrence is replaced. ',
        s_fdef_SUM: 'Adds the numeric values. The values to the sum function may be ranges in the form similar to A1:B5. ',
        s_fdef_SUMIF: 'Sums the numeric values of cells in the range that meet the criteria. The criteria may be a value ("x", 15, 1+3) or a test (>25). If range2 is present, then range1 is tested and the corresponding range2 value is summed. ',
        s_fdef_SYD: 'Depreciation by Sum of Year\'s Digits method. ',
        s_fdef_T: 'Returns the text value or else a null string. ',
        s_fdef_TAN: 'Trigonometric tangent function (value is in radians) ',
        s_fdef_TIME: 'Returns the time value given the specified hour, minute, and second. ',
        s_fdef_TODAY: 'Returns the current date (an integer). Note: In this program, day "1" is December 31, 1899 and the year 1900 is not a leap year. Some programs use January 1, 1900, as day "1" and treat 1900 as a leap year. In both cases, though, dates on or after March 1, 1900, are the same. ',
        s_fdef_TRIM: 'Returns the text value with leading, trailing, and repeated spaces removed. ',
        s_fdef_TRUE: 'Returns the logical value "true". ',
        s_fdef_TRUNC: 'Truncates the value to the specified number of decimal places. If precision is negative, truncate to powers of 10. ',
        s_fdef_UPPER: 'Returns the text value with all lowercase characters converted to uppercase. ',
        s_fdef_VALUE: 'Converts the specified text value into a numeric value. Various forms that look like numbers (including digits followed by %, forms that look like dates, etc.) are handled. This may not handle all of the forms accepted by other spreadsheets and may be locale dependent. ',
        s_fdef_VAR: 'Returns the sample variance of the numeric values. ',
        s_fdef_VARP: 'Returns the variance of the numeric values. ',
        s_fdef_VLOOKUP: 'Look for the matching value for the given value in the range and return the corresponding value in the cell specified by the column offset. If rangelookup is 1 (the default) and not 0, match if within numeric brackets (match>=value) instead of exact match. ',
        s_fdef_WEEKDAY: 'Returns the day of week specified by the date value. If type is 1 (the default), Sunday is day and Saturday is day 7. If type is 2, Monday is day 1 and Sunday is day 7. If type is 3, Monday is day 0 and Sunday is day 6. ',
        s_fdef_YEAR: 'Returns the year part of a date value. ',

        s_farg_v: "value",
        s_farg_vn: "value1, value2, ...",
        s_farg_xy: "valueX, valueY",
        s_farg_choose: "index, value1, value2, ...",
        s_farg_range: "range",
        s_farg_rangec: "range, criteria",
        s_farg_date: "year, month, day",
        s_farg_dfunc: "databaserange, fieldname, criteriarange",
        s_farg_ddb: "cost, salvage, lifetime, period [, factor]",
        s_farg_find: "string1, string2 [, start]",
        s_farg_fv: "rate, n, payment, [pv, [paytype]]",
        s_farg_hlookup: "value, range, row, [rangelookup]",
        s_farg_iffunc: "logical-expression, true-value, false-value",
        s_farg_index: "range, rownum, colnum",
        s_farg_irr: "range, [guess]",
        s_farg_tc: "text, count",
        s_farg_log: "value, base",
        s_farg_match: "value, range, [rangelookup]",
        s_farg_mid: "text, start, length",
        s_farg_nper: "rate, payment, pv, [fv, [paytype]]",
        s_farg_npv: "rate, value1, value2, ...",
        s_farg_pmt: "rate, n, pv, [fv, [paytype]]",
        s_farg_pv: "rate, n, payment, [fv, [paytype]]",
        s_farg_rate: "n, payment, pv, [fv, [paytype, [guess]]]",
        s_farg_replace: "text1, start, length, text2",
        s_farg_vp: "value, [precision]",
        s_farg_valpre: "value, precision",
        s_farg_csl: "cost, salvage, lifetime",
        s_farg_cslp: "cost, salvage, lifetime, period",
        s_farg_subs: "text1, oldtext, newtext [, occurrence]",
        s_farg_sumif: "range1, criteria [, range2]",
        s_farg_hms: "hour, minute, second",
        s_farg_txt: "text",
        s_farg_vlookup: "value, range, col, [rangelookup]",
        s_farg_weekday: "date, [type]",
        s_farg_dt: "date",
        s_farg_: ''
    }


/////////////////////////
//
// FUNCTION DEFINITIONS
//
// The standard function definitions follow.
//
// Note that some need SocialCalc.DetermineValueType to be defined.
//

    /*
    #
    # AVERAGE(v1,c1:c2,...)
    # COUNT(v1,c1:c2,...)
    # COUNTA(v1,c1:c2,...)
    # COUNTBLANK(v1,c1:c2,...)
    # MAX(v1,c1:c2,...)
    # MIN(v1,c1:c2,...)
    # PRODUCT(v1,c1:c2,...)
    # STDEV(v1,c1:c2,...)
    # STDEVP(v1,c1:c2,...)
    # SUM(v1,c1:c2,...)
    # VAR(v1,c1:c2,...)
    # VARP(v1,c1:c2,...)
    #
    # Calculate all of these and then return the desired one (overhead is in accessing not calculating)
    # If this routine is changed, check the dseries_functions, too.
    #
    */

    //var SocialCalc = {}
    //this.SocialCalcFormula = {}
    var SocialCalcFormula = this
    SocialCalcFormula.PushOperand = function(operand, t, v) {

        operand.push({type: t, value: v});

    }
    var FunctionClassList = this.FunctionClassList

    SocialCalcFormula.SeriesFunctions = function (fname, operand, foperand) {

        var value1, t, v1;

        var scf = SocialCalcFormula;
        var operand_value_and_type = scf.OperandValueAndType;
        var lookup_result_type = scf.LookupResultType;
        var typelookupplus = scf.TypeLookupTable.plus;

        var PushOperand = function (t, v) {
            operand.push({type: t, value: v});
        };

        var sum = 0;
        var resulttypesum = "";
        var count = 0;
        var counta = 0;
        var countblank = 0;
        var product = 1;
        var maxval;
        var minval;
        var mk, sk, mk1, sk1; // For variance, etc.: M sub k, k-1, and S sub k-1
                              // as per Knuth "The Art of Computer Programming" Vol. 2 3rd edition, page 232

        while (foperand.length > 0) {
            value1 = operand_value_and_type(foperand);
            t = value1.type.charAt(0);
            if (t == "n") count += 1;
            if (t != "b") counta += 1;
            if (t == "b") countblank += 1;

            if (t == "n") {
                v1 = value1.value - 0; // get it as a number
                sum += v1;
                product *= v1;
                maxval = (maxval != undefined) ? (v1 > maxval ? v1 : maxval) : v1;
                minval = (minval != undefined) ? (v1 < minval ? v1 : minval) : v1;
                if (count == 1) { // initialize with first values for variance used in STDEV, VAR, etc.
                    mk1 = v1;
                    sk1 = 0;
                }
                else { // Accumulate S sub 1 through n as per Knuth noted above
                    mk = mk1 + (v1 - mk1) / count;
                    sk = sk1 + (v1 - mk1) * (v1 - mk);
                    sk1 = sk;
                    mk1 = mk;
                }
                resulttypesum = lookup_result_type(value1.type, resulttypesum || value1.type, typelookupplus);
            }
            else if (t == "e" && resulttypesum.charAt(0) != "e") {
                resulttypesum = value1.type;
            }
        }

        resulttypesum = resulttypesum || "n";

        switch (fname) {
            case "SUM":
                PushOperand(resulttypesum, sum);
                break;

            case "PRODUCT": // may handle cases with text differently than some other spreadsheets
                PushOperand(resulttypesum, product);
                break;

            case "MIN":
                PushOperand(resulttypesum, minval || 0);
                break;

            case "MAX":
                PushOperand(resulttypesum, maxval || 0);
                break;

            case "COUNT":
                PushOperand("n", count);
                break;

            case "COUNTA":
                PushOperand("n", counta);
                break;

            case "COUNTBLANK":
                PushOperand("n", countblank);
                break;

            case "AVERAGE":
                if (count > 0) {
                    PushOperand(resulttypesum, sum / count);
                }
                else {
                    PushOperand("e#DIV/0!", 0);
                }
                break;

            case "STDEV":
                if (count > 1) {
                    PushOperand(resulttypesum, Math.sqrt(sk / (count - 1))); // sk is never negative according to Knuth
                }
                else {
                    PushOperand("e#DIV/0!", 0);
                }
                break;

            case "STDEVP":
                if (count > 1) {
                    PushOperand(resulttypesum, Math.sqrt(sk / count));
                }
                else {
                    PushOperand("e#DIV/0!", 0);
                }
                break;

            case "VAR":
                if (count > 1) {
                    PushOperand(resulttypesum, sk / (count - 1));
                }
                else {
                    PushOperand("e#DIV/0!", 0);
                }
                break;

            case "VARP":
                if (count > 1) {
                    PushOperand(resulttypesum, sk / count);
                }
                else {
                    PushOperand("e#DIV/0!", 0);
                }
                break;
        }

        return null;

    }

// Add to function list
    FunctionClassList['stat']["AVERAGE"] = [SocialCalcFormula.SeriesFunctions, -1, "vn", null, "stat"];
    FunctionClassList['stat']["COUNT"] = [SocialCalcFormula.SeriesFunctions, -1, "vn", null, "stat"];
    FunctionClassList['stat']["COUNTA"] = [SocialCalcFormula.SeriesFunctions, -1, "vn", null, "stat"];
    FunctionClassList['stat']["COUNTBLANK"] = [SocialCalcFormula.SeriesFunctions, -1, "vn", null, "stat"];
    FunctionClassList['stat']["MAX"] = [SocialCalcFormula.SeriesFunctions, -1, "vn", null, "stat"];
    FunctionClassList['stat']["MIN"] = [SocialCalcFormula.SeriesFunctions, -1, "vn", null, "stat"];
    FunctionClassList['stat']["PRODUCT"] = [SocialCalcFormula.SeriesFunctions, -1, "vn", null, "stat"];
    FunctionClassList['stat']["STDEV"] = [SocialCalcFormula.SeriesFunctions, -1, "vn", null, "stat"];
    FunctionClassList['stat']["STDEVP"] = [SocialCalcFormula.SeriesFunctions, -1, "vn", null, "stat"];
    FunctionClassList['stat']["SUM"] = [SocialCalcFormula.SeriesFunctions, -1, "vn", null, "stat"];
    FunctionClassList['stat']["VAR"] = [SocialCalcFormula.SeriesFunctions, -1, "vn", null, "stat"];
    FunctionClassList['stat']["VARP"] = [SocialCalcFormula.SeriesFunctions, -1, "vn", null, "stat"];

    /*
    #
    # DAVERAGE(databaserange, fieldname, criteriarange)
    # DCOUNT(databaserange, fieldname, criteriarange)
    # DCOUNTA(databaserange, fieldname, criteriarange)
    # DGET(databaserange, fieldname, criteriarange)
    # DMAX(databaserange, fieldname, criteriarange)
    # DMIN(databaserange, fieldname, criteriarange)
    # DPRODUCT(databaserange, fieldname, criteriarange)
    # DSTDEV(databaserange, fieldname, criteriarange)
    # DSTDEVP(databaserange, fieldname, criteriarange)
    # DSUM(databaserange, fieldname, criteriarange)
    # DVAR(databaserange, fieldname, criteriarange)
    # DVARP(databaserange, fieldname, criteriarange)
    #
    # Calculate all of these and then return the desired one (overhead is in accessing not calculating)
    # If this routine is changed, check the series_functions, too.
    #
    */

    SocialCalcFormula.DSeriesFunctions = function (fname, operand, foperand) {

        var value1, tostype, cr, dbrange, fieldname, criteriarange, dbinfo, criteriainfo;
        var fieldasnum, targetcol, i, j, k, cell, criteriafieldnums;
        var testok, criteriacr, criteria, testcol, testcr;
        var t;

        var scf = SocialCalcFormula;
        var operand_value_and_type = scf.OperandValueAndType;
        var lookup_result_type = scf.LookupResultType;
        var typelookupplus = scf.TypeLookupTable.plus;

        var PushOperand = function (t, v) {
            operand.push({type: t, value: v});
        };

        var value1 = {};

        var sum = 0;
        var resulttypesum = "";
        var count = 0;
        var counta = 0;
        var countblank = 0;
        var product = 1;
        var maxval;
        var minval;
        var mk, sk, mk1, sk1; // For variance, etc.: M sub k, k-1, and S sub k-1
                              // as per Knuth "The Art of Computer Programming" Vol. 2 3rd edition, page 232

        dbrange = scf.TopOfStackValueAndType( foperand); // get a range
        fieldname = scf.OperandValueAndType( foperand); // get a value
        criteriarange = scf.TopOfStackValueAndType( foperand); // get a range

        if (dbrange.type != "range" || criteriarange.type != "range") {
            return scf.FunctionArgsError(fname, operand);
        }

        dbinfo = scf.DecodeRangeParts(dbrange.value);
        criteriainfo = scf.DecodeRangeParts(criteriarange.value);

        fieldasnum = scf.FieldToColnum(dbinfo.sheetdata, dbinfo.col1num, dbinfo.ncols, dbinfo.row1num, fieldname.value, fieldname.type);
        if (fieldasnum <= 0) {
            PushOperand("e#VALUE!", 0);
            return;
        }

        targetcol = dbinfo.col1num + fieldasnum - 1;
        criteriafieldnums = [];

        for (i = 0; i < criteriainfo.ncols; i++) { // get criteria field colnums
            cell = criteriainfo.sheetdata.GetAssuredCell(SocialCalc.crToCoord(criteriainfo.col1num + i, criteriainfo.row1num));
            criterianum = scf.FieldToColnum(dbinfo.sheetdata, dbinfo.col1num, dbinfo.ncols, dbinfo.row1num, cell.datavalue, cell.valuetype);
            if (criterianum <= 0) {
                PushOperand("e#VALUE!", 0);
                return;
            }
            criteriafieldnums.push(dbinfo.col1num + criterianum - 1);
        }

        for (i = 1; i < dbinfo.nrows; i++) { // go through each row of the database
            testok = false;
            CRITERIAROW:
                for (j = 1; j < criteriainfo.nrows; j++) { // go through each criteria row
                    for (k = 0; k < criteriainfo.ncols; k++) { // look at each column
                        criteriacr = SocialCalc.crToCoord(criteriainfo.col1num + k, criteriainfo.row1num + j); // where criteria is
                        cell = criteriainfo.sheetdata.GetAssuredCell(criteriacr);
                        criteria = cell.datavalue;
                        if (typeof criteria == "string" && criteria.length == 0) continue; // blank items are OK
                        testcol = criteriafieldnums[k];
                        testcr = SocialCalc.crToCoord(testcol, dbinfo.row1num + i); // cell to check
                        cell = criteriainfo.sheetdata.GetAssuredCell(testcr);
                        if (!scf.TestCriteria(cell.datavalue, cell.valuetype || "b", criteria)) {
                            continue CRITERIAROW; // does not meet criteria - check next row
                        }
                    }
                    testok = true; // met all the criteria
                    break CRITERIAROW;
                }
            if (!testok) {
                continue;
            }

            cr = SocialCalc.crToCoord(targetcol, dbinfo.row1num + i); // get cell of this row to do the function on
            cell = dbinfo.sheetdata.GetAssuredCell(cr);

            value1.value = cell.datavalue;
            value1.type = cell.valuetype;
            t = value1.type.charAt(0);
            if (t == "n") count += 1;
            if (t != "b") counta += 1;
            if (t == "b") countblank += 1;

            if (t == "n") {
                v1 = value1.value - 0; // get it as a number
                sum += v1;
                product *= v1;
                maxval = (maxval != undefined) ? (v1 > maxval ? v1 : maxval) : v1;
                minval = (minval != undefined) ? (v1 < minval ? v1 : minval) : v1;
                if (count == 1) { // initialize with first values for variance used in STDEV, VAR, etc.
                    mk1 = v1;
                    sk1 = 0;
                }
                else { // Accumulate S sub 1 through n as per Knuth noted above
                    mk = mk1 + (v1 - mk1) / count;
                    sk = sk1 + (v1 - mk1) * (v1 - mk);
                    sk1 = sk;
                    mk1 = mk;
                }
                resulttypesum = lookup_result_type(value1.type, resulttypesum || value1.type, typelookupplus);
            }
            else if (t == "e" && resulttypesum.charAt(0) != "e") {
                resulttypesum = value1.type;
            }
        }

        resulttypesum = resulttypesum || "n";

        switch (fname) {
            case "DSUM":
                PushOperand(resulttypesum, sum);
                break;

            case "DPRODUCT": // may handle cases with text differently than some other spreadsheets
                PushOperand(resulttypesum, product);
                break;

            case "DMIN":
                PushOperand(resulttypesum, minval || 0);
                break;

            case "DMAX":
                PushOperand(resulttypesum, maxval || 0);
                break;

            case "DCOUNT":
                PushOperand("n", count);
                break;

            case "DCOUNTA":
                PushOperand("n", counta);
                break;

            case "DAVERAGE":
                if (count > 0) {
                    PushOperand(resulttypesum, sum / count);
                }
                else {
                    PushOperand("e#DIV/0!", 0);
                }
                break;

            case "DSTDEV":
                if (count > 1) {
                    PushOperand(resulttypesum, Math.sqrt(sk / (count - 1))); // sk is never negative according to Knuth
                }
                else {
                    PushOperand("e#DIV/0!", 0);
                }
                break;

            case "DSTDEVP":
                if (count > 1) {
                    PushOperand(resulttypesum, Math.sqrt(sk / count));
                }
                else {
                    PushOperand("e#DIV/0!", 0);
                }
                break;

            case "DVAR":
                if (count > 1) {
                    PushOperand(resulttypesum, sk / (count - 1));
                }
                else {
                    PushOperand("e#DIV/0!", 0);
                }
                break;

            case "DVARP":
                if (count > 1) {
                    PushOperand(resulttypesum, sk / count);
                }
                else {
                    PushOperand("e#DIV/0!", 0);
                }
                break;

            case "DGET":
                if (count == 1) {
                    PushOperand(resulttypesum, sum);
                }
                else if (count == 0) {
                    PushOperand("e#VALUE!", 0);
                }
                else {
                    PushOperand("e#NUM!", 0);
                }
                break;

        }

        return;

    }

    FunctionClassList['stat']["DAVERAGE"] = [SocialCalcFormula.DSeriesFunctions, 3, "dfunc", "", "stat"];
    FunctionClassList['stat']["DCOUNT"] = [SocialCalcFormula.DSeriesFunctions, 3, "dfunc", "", "stat"];
    FunctionClassList['stat']["DCOUNTA"] = [SocialCalcFormula.DSeriesFunctions, 3, "dfunc", "", "stat"];
    FunctionClassList['stat']["DGET"] = [SocialCalcFormula.DSeriesFunctions, 3, "dfunc", "", "stat"];
    FunctionClassList['stat']["DMAX"] = [SocialCalcFormula.DSeriesFunctions, 3, "dfunc", "", "stat"];
    FunctionClassList['stat']["DMIN"] = [SocialCalcFormula.DSeriesFunctions, 3, "dfunc", "", "stat"];
    FunctionClassList['stat']["DPRODUCT"] = [SocialCalcFormula.DSeriesFunctions, 3, "dfunc", "", "stat"];
    FunctionClassList['stat']["DSTDEV"] = [SocialCalcFormula.DSeriesFunctions, 3, "dfunc", "", "stat"];
    FunctionClassList['stat']["DSTDEVP"] = [SocialCalcFormula.DSeriesFunctions, 3, "dfunc", "", "stat"];
    FunctionClassList['stat']["DSUM"] = [SocialCalcFormula.DSeriesFunctions, 3, "dfunc", "", "stat"];
    FunctionClassList['stat']["DVAR"] = [SocialCalcFormula.DSeriesFunctions, 3, "dfunc", "", "stat"];
    FunctionClassList['stat']["DVARP"] = [SocialCalcFormula.DSeriesFunctions, 3, "dfunc", "", "stat"];

    /*
    #
    # colnum = SocialCalcFormula.FieldToColnum(sheet, col1num, ncols, row1num, fieldname, fieldtype)
    #
    # If fieldname is a number, uses it, otherwise looks up string in cells in row to find field number
    #
    # If not found, returns 0.
    #
    */

    SocialCalcFormula.FieldToColnum = function (col1num, ncols, row1num, fieldname, fieldtype) {

        var colnum, cell, value;

        if (fieldtype.charAt(0) == "n") { // number - return it if legal
            colnum = fieldname - 0; // make sure a number
            if (colnum <= 0 || colnum > ncols) {
                return 0;
            }
            return Math.floor(colnum);
        }

        if (fieldtype.charAt(0) != "t") { // must be text otherwise
            return 0;
        }

        fieldname = fieldname ? fieldname.toLowerCase() : "";

        for (colnum = 0; colnum < ncols; colnum++) { // look through column headers for a match
            cell = sheet.GetAssuredCell(SocialCalc.crToCoord(col1num + colnum, row1num));
            value = cell.datavalue;
            value = (value + "").toLowerCase(); // ignore case
            if (value == fieldname) { // match
                return colnum + 1;
            }
        }
        return 0; // looked at all and no match

    }


    /*
    #
    # HLOOKUP(value, range, row, [rangelookup])
    # VLOOKUP(value, range, col, [rangelookup])
    # MATCH(value, range, [rangelookup])
    #
    */

    SocialCalcFormula.LookupFunctions = function (fname, operand, foperand) {

        var lookupvalue, range, offset, rangelookup, offsetvalue, rangeinfo;
        var c, r, cincr, rincr, previousOK, csave, rsave, cell, value, valuetype, cr, lookupvalue;

        var scf = SocialCalcFormula;
        var operand_value_and_type = scf.OperandValueAndType;
        var lookup_result_type = scf.LookupResultType;
        var typelookupplus = scf.TypeLookupTable.plus;

        var PushOperand = function (t, v) {
            operand.push({type: t, value: v});
        };

        lookupvalue = operand_value_and_type( foperand);
        if (typeof lookupvalue.value == "string") {
            lookupvalue.value = lookupvalue.value.toLowerCase();
        }

        range = scf.TopOfStackValueAndType(foperand);

        rangelookup = 1; // default to true or 1
        if (fname == "MATCH") {
            if (foperand.length) {
                rangelookup = scf.OperandAsNumber(foperand);
                if (rangelookup.type.charAt(0) != "n") {
                    PushOperand("e#VALUE!", 0);
                    return;
                }
                if (foperand.length) {
                    scf.FunctionArgsError(fname, operand);
                    return 0;
                }
                rangelookup = rangelookup.value - 0;
            }
        }
        else {
            offsetvalue = scf.OperandAsNumber( foperand);
            if (offsetvalue.type.charAt(0) != "n") {
                PushOperand("e#VALUE!", 0);
                return;
            }
            offsetvalue = Math.floor(offsetvalue.value);
            if (foperand.length) {
                rangelookup = scf.OperandAsNumber( foperand);
                if (rangelookup.type.charAt(0) != "n") {
                    PushOperand("e#VALUE!", 0);
                    return;
                }
                if (foperand.length) {
                    scf.FunctionArgsError(fname, operand);
                    return 0;
                }
                rangelookup = rangelookup.value ? 1 : 0; // convert to 1 or 0
            }
        }
        lookupvalue.type = lookupvalue.type.charAt(0); // only deal with general type
        if (lookupvalue.type == "n") { // if number, make sure a number
            lookupvalue.value = lookupvalue.value - 0;
        }

        if (range.type != "range") {
            scf.FunctionArgsError(fname, operand);
            return 0;
        }

        rangeinfo = scf.DecodeRangeParts(range.value, range.type);
        if (!rangeinfo) {
            PushOperand("e#REF!", 0);
            return;
        }

        c = 0;
        r = 0;
        cincr = 0;
        rincr = 0;
        if (fname == "HLOOKUP") {
            cincr = 1;
            if (offsetvalue > rangeinfo.nrows) {
                PushOperand("e#REF!", 0);
                return;
            }
        }
        else if (fname == "VLOOKUP") {
            rincr = 1;
            if (offsetvalue > rangeinfo.ncols) {
                PushOperand("e#REF!", 0);
                return;
            }
        }
        else if (fname == "MATCH") {
            if (rangeinfo.ncols > 1) {
                if (rangeinfo.nrows > 1) {
                    PushOperand("e#N/A", 0);
                    return;
                }
                cincr = 1;
            }
            else {
                rincr = 1;
            }
        }
        else {
            scf.FunctionArgsError(fname, operand);
            return 0;
        }
        if (offsetvalue < 1 && fname != "MATCH") {
            PushOperand("e#VALUE!", 0);
            return 0;
        }

        previousOK; // if 1, previous test was <. If 2, also this one wasn't

        while (1) {
            cr = SocialCalc.crToCoord(rangeinfo.col1num + c, rangeinfo.row1num + r);
            cell = rangeinfo.sheetdata.GetAssuredCell(cr);
            value = cell.datavalue;
            valuetype = cell.valuetype ? cell.valuetype.charAt(0) : "b"; // only deal with general types
            if (valuetype == "n") {
                value = value - 0; // make sure number
            }
            if (rangelookup) { // rangelookup type 1 or -1: look for within brackets for matches
                if (lookupvalue.type == "n" && valuetype == "n") {
                    if (lookupvalue.value == value) { // match
                        break;
                    }
                    if ((rangelookup > 0 && lookupvalue.value > value)
                        || (rangelookup < 0 && lookupvalue.value < value)) { // possible match: wait and see
                        previousOK = 1;
                        csave = c; // remember col and row of last OK
                        rsave = r;
                    }
                    else if (previousOK) { // last one was OK, this one isn't
                        previousOK = 2;
                        break;
                    }
                }

                else if (lookupvalue.type == "t" && valuetype == "t") {
                    value = typeof value == "string" ? value.toLowerCase() : "";
                    if (lookupvalue.value == value) { // match
                        break;
                    }
                    if ((rangelookup > 0 && lookupvalue.value > value)
                        || (rangelookup < 0 && lookupvalue.value < value)) { // possible match: wait and see
                        previousOK = 1;
                        csave = c;
                        rsave = r;
                    }
                    else if (previousOK) { // last one was OK, this one isn't
                        previousOK = 2;
                        break;
                    }
                }
            }
            else { // exact value matches
                if (lookupvalue.type == "n" && valuetype == "n") {
                    if (lookupvalue.value == value) { // match
                        break;
                    }
                }
                else if (lookupvalue.type == "t" && valuetype == "t") {
                    value = typeof value == "string" ? value.toLowerCase() : "";
                    if (lookupvalue.value == value) { // match
                        break;
                    }
                }
            }

            r += rincr;
            c += cincr;
            if (r >= rangeinfo.nrows || c >= rangeinfo.ncols) { // end of range to check, no exact match
                if (previousOK) { // at least one could have been OK
                    previousOK = 2;
                    break;
                }
                PushOperand("e#N/A", 0);
                return;
            }
        }

        if (previousOK == 2) { // back to last OK
            r = rsave;
            c = csave;
        }

        if (fname == "MATCH") {
            value = c + r + 1; // only one may be <> 0
            valuetype = "n";
        }
        else {
            cr = SocialCalc.crToCoord(rangeinfo.col1num + c + (fname == "VLOOKUP" ? offsetvalue - 1 : 0), rangeinfo.row1num + r + (fname == "HLOOKUP" ? offsetvalue - 1 : 0));
            cell = rangeinfo.sheetdata.GetAssuredCell(cr);
            value = cell.datavalue;
            valuetype = cell.valuetype;
        }
        PushOperand(valuetype, value);

        return;

    }

    FunctionClassList['lookup']["HLOOKUP"] = [SocialCalcFormula.LookupFunctions, -3, "hlookup", "", "lookup"];
    FunctionClassList['lookup']["MATCH"] = [SocialCalcFormula.LookupFunctions, -2, "match", "", "lookup"];
    FunctionClassList['lookup']["VLOOKUP"] = [SocialCalcFormula.LookupFunctions, -3, "vlookup", "", "lookup"];

    /*
    #
    # INDEX(range, rownum, colnum)
    #
    */

    SocialCalcFormula.IndexFunction = function (fname, operand, foperand) {

        var range, sheetname, indexinfo, rowindex, colindex, result, resulttype;

        var scf = SocialCalcFormula;

        var PushOperand = function (t, v) {
            operand.push({type: t, value: v});
        };

        range = scf.TopOfStackValueAndType(sheet, foperand); // get range
        if (range.type != "range") {
            scf.FunctionArgsError(fname, operand);
            return 0;
        }
        indexinfo = scf.DecodeRangeParts(sheet, range.value, range.type);
        if (indexinfo.sheetname) {
            sheetname = "!" + indexinfo.sheetname;
        }
        else {
            sheetname = "";
        }

        rowindex = {value: 0};
        colindex = {value: 0};

        if (foperand.length) { // look for row number
            rowindex = scf.OperandAsNumber(sheet, foperand);
            if (rowindex.type.charAt(0) != "n" || rowindex.value < 0) {
                PushOperand("e#VALUE!", 0);
                return;
            }
            if (foperand.length) { // look for col number
                colindex = scf.OperandAsNumber(sheet, foperand);
                if (colindex.type.charAt(0) != "n" || colindex.value < 0) {
                    PushOperand("e#VALUE!", 0);
                    return;
                }
                if (foperand.length) {
                    scf.FunctionArgsError(fname, operand);
                    return 0;
                }
            }
            else { // col number missing
                if (indexinfo.nrows == 1) { // if only one row, then rowindex is really colindex
                    colindex.value = rowindex.value;
                    rowindex.value = 0;
                }
            }
        }

        if (rowindex.value > indexinfo.nrows || colindex.value > indexinfo.ncols) {
            PushOperand("e#REF!", 0);
            return;
        }

        if (rowindex.value == 0) {
            if (colindex.value == 0) {
                if (indexinfo.nrows == 1 && indexinfo.ncols == 1) {
                    result = SocialCalc.crToCoord(indexinfo.col1num, indexinfo.row1num) + sheetname;
                    resulttype = "coord";
                }
                else {
                    result = SocialCalc.crToCoord(indexinfo.col1num, indexinfo.row1num) + sheetname + "|" +
                        SocialCalc.crToCoord(indexinfo.col1num + indexinfo.ncols - 1, indexinfo.row1num + indexinfo.nrows - 1) +
                        "|";
                    resulttype = "range";
                }
            }
            else {
                if (indexinfo.nrows == 1) {
                    result = SocialCalc.crToCoord(indexinfo.col1num + colindex.value - 1, indexinfo.row1num) + sheetname;
                    resulttype = "coord";
                }
                else {
                    result = SocialCalc.crToCoord(indexinfo.col1num + colindex.value - 1, indexinfo.row1num) + sheetname + "|" +
                        SocialCalc.crToCoord(indexinfo.col1num + colindex.value - 1, indexinfo.row1num + indexinfo.nrows - 1) +
                        "|";
                    resulttype = "range";
                }
            }
        }
        else {
            if (colindex.value == 0) {
                if (indexinfo.ncols == 1) {
                    result = SocialCalc.crToCoord(indexinfo.col1num, indexinfo.row1num + rowindex.value - 1) + sheetname;
                    resulttype = "coord";
                }
                else {
                    result = SocialCalc.crToCoord(indexinfo.col1num, indexinfo.row1num + rowindex.value - 1) + sheetname + "|" +
                        SocialCalc.crToCoord(indexinfo.col1num + indexinfo.ncols - 1, indexinfo.row1num + rowindex.value - 1) +
                        "|";
                    resulttype = "range";
                }
            }
            else {
                result = SocialCalc.crToCoord(indexinfo.col1num + colindex.value - 1, indexinfo.row1num + rowindex.value - 1) + sheetname;
                resulttype = "coord";
            }
        }

        PushOperand(resulttype, result);

        return;

    }

    FunctionClassList['lookup']["INDEX"] = [SocialCalcFormula.IndexFunction, -1, "index", "", "lookup"];

    /*
    #
    # COUNTIF(c1:c2,"criteria")
    # SUMIF(c1:c2,"criteria",[range2])
    #
    */

    SocialCalcFormula.CountifSumifFunctions = function (fname, operand, foperand) {

        var range, criteria, sumrange, f2operand, result, resulttype, value1, value2;
        var sum = 0;
        var resulttypesum = "";
        var count = 0;

        var scf = SocialCalcFormula;
        var operand_value_and_type = scf.OperandValueAndType;
        var lookup_result_type = scf.LookupResultType;
        var typelookupplus = scf.TypeLookupTable.plus;

        var PushOperand = function (t, v) {
            operand.push({type: t, value: v});
        };

        range = scf.TopOfStackValueAndType(sheet, foperand); // get range or coord
        criteria = scf.OperandAsText(sheet, foperand); // get criteria
        if (fname == "SUMIF") {
            if (foperand.length == 1) { // three arg form of SUMIF
                sumrange = scf.TopOfStackValueAndType(sheet, foperand);
            }
            else if (foperand.length == 0) { // two arg form
                sumrange = {value: range.value, type: range.type};
            }
            else {
                scf.FunctionArgsError(fname, operand);
                return 0;
            }
        }
        else {
            sumrange = {value: range.value, type: range.type};
        }

        if (criteria.type.charAt(0) == "n") {
            criteria.value = criteria.value + ""; // make text
        }
        else if (criteria.type.charAt(0) == "e") { // error
            criteria.value = null;
        }
        else if (criteria.type.charAt(0) == "b") { // blank here is undefined
            criteria.value = null;
        }

        if (range.type != "coord" && range.type != "range") {
            scf.FunctionArgsError(fname, operand);
            return 0;
        }

        if (fname == "SUMIF" && sumrange.type != "coord" && sumrange.type != "range") {
            scf.FunctionArgsError(fname, operand);
            return 0;
        }

        foperand.push(range);
        f2operand = []; // to allow for 3 arg form
        f2operand.push(sumrange);

        while (foperand.length) {
            value1 = operand_value_and_type(sheet, foperand);
            value2 = operand_value_and_type(sheet, f2operand);

            if (!scf.TestCriteria(value1.value, value1.type, criteria.value)) {
                continue;
            }

            count += 1;

            if (value2.type.charAt(0) == "n") {
                sum += value2.value - 0;
                resulttypesum = lookup_result_type(value2.type, resulttypesum || value2.type, typelookupplus);
            }
            else if (value2.type.charAt(0) == "e" && resulttypesum.charAt(0) != "e") {
                resulttypesum = value2.type;
            }
        }

        resulttypesum = resulttypesum || "n";

        if (fname == "SUMIF") {
            PushOperand(resulttypesum, sum);
        }
        else if (fname == "COUNTIF") {
            PushOperand("n", count);
        }

        return;

    }

    FunctionClassList['stat']["COUNTIF"] = [SocialCalcFormula.CountifSumifFunctions, 2, "rangec", "", "stat"];
    FunctionClassList['stat']["SUMIF"] = [SocialCalcFormula.CountifSumifFunctions, -2, "sumif", "", "stat"];

    /*
    #
    # IF(cond,truevalue,falsevalue)
    #
    */

    SocialCalcFormula.IfFunction = function (fname, operand, foperand) {

        var cond, t;

        cond = SocialCalcFormula.OperandValueAndType(foperand);
        t = cond.type.charAt(0);
        if (t != "n" && t != "b") {
            operand.push({type: "e#VALUE!", value: 0});
            return;
        }

        if (!cond.value) foperand.pop();
        operand.push(foperand.pop());
        if (cond.value) foperand.pop();

        return null;

    }

// Add to function list
    FunctionClassList['test']["IF"] = [SocialCalcFormula.IfFunction, 3, "iffunc", "", "test"];

    /*
    #
    # DATE(year,month,day)
    #
    */

    // SocialCalcFormula.DateFunction = function (fname, operand, foperand) {
    //
    //     var scf = SocialCalcFormula;
    //     var result = 0;
    //     var year = scf.OperandAsNumber(foperand);
    //     var month = scf.OperandAsNumber(foperand);
    //     var day = scf.OperandAsNumber(foperand);
    //     var resulttype = scf.LookupResultType(year.type, month.type, scf.TypeLookupTable.twoargnumeric);
    //     resulttype = scf.LookupResultType(resulttype, day.type, scf.TypeLookupTable.twoargnumeric);
    //     if (resulttype.charAt(0) == "n") {
    //         result = SocialCalc.FormatNumber.convert_date_gregorian_to_julian(
    //             Math.floor(year.value), Math.floor(month.value), Math.floor(day.value)
    //         ) - SocialCalc.FormatNumber.datevalues.julian_offset;
    //         resulttype = "nd";
    //     }
    //     scf.PushOperand(operand, resulttype, result);
    //     return;
    //
    // }

    FunctionClassList['datetime']["DATE"] = [SocialCalcFormula.DateFunction, 3, "date", "", "datetime"];

    /*
    #
    # TIME(hour,minute,second)
    #
    */

    SocialCalcFormula.TimeFunction = function (fname, operand, foperand) {

        var scf = SocialCalcFormula;
        var result = 0;
        var hours = scf.OperandAsNumber(foperand);
        var minutes = scf.OperandAsNumber(foperand);
        var seconds = scf.OperandAsNumber(foperand);
        var resulttype = scf.LookupResultType(hours.type, minutes.type, scf.TypeLookupTable.twoargnumeric);
        resulttype = scf.LookupResultType(resulttype, seconds.type, scf.TypeLookupTable.twoargnumeric);
        if (resulttype.charAt(0) == "n") {
            result = ((hours.value * 60 * 60) + (minutes.value * 60) + seconds.value) / (24 * 60 * 60);
            resulttype = "nt";
        }
        scf.PushOperand(operand, resulttype, result);
        return;

    }

    FunctionClassList['datetime']["TIME"] = [SocialCalcFormula.TimeFunction, 3, "hms", "", "datetime"];

    /*
    #
    # DAY(date)
    # MONTH(date)
    # YEAR(date)
    # WEEKDAY(date, [type])
    #
    */

    // SocialCalcFormula.DMYFunctions = function (fname, operand, foperand) {
    //
    //     var ymd, dtype, doffset;
    //     var scf = SocialCalcFormula;
    //     var result = 0;
    //
    //     var datevalue = scf.OperandAsNumber(foperand);
    //     var resulttype = scf.LookupResultType(datevalue.type, datevalue.type, scf.TypeLookupTable.oneargnumeric);
    //
    //     if (resulttype.charAt(0) == "n") {
    //         ymd = SocialCalc.FormatNumber.convert_date_julian_to_gregorian(
    //             Math.floor(datevalue.value + SocialCalc.FormatNumber.datevalues.julian_offset));
    //         switch (fname) {
    //             case "DAY":
    //                 result = ymd.day;
    //                 break;
    //
    //             case "MONTH":
    //                 result = ymd.month;
    //                 break;
    //
    //             case "YEAR":
    //                 result = ymd.year;
    //                 break;
    //
    //             case "WEEKDAY":
    //                 dtype = {value: 1};
    //                 if (foperand.length) { // get type if present
    //                     dtype = scf.OperandAsNumber(foperand);
    //                     if (dtype.type.charAt(0) != "n" || dtype.value < 1 || dtype.value > 3) {
    //                         scf.PushOperand(operand, "e#VALUE!", 0);
    //                         return;
    //                     }
    //                     if (foperand.length) { // extra args
    //                         scf.FunctionArgsError(fname, operand);
    //                         return;
    //                     }
    //                 }
    //                 doffset = 6;
    //                 if (dtype.value > 1) {
    //                     doffset -= 1;
    //                 }
    //                 result = Math.floor(datevalue.value + doffset) % 7 + (dtype.value < 3 ? 1 : 0);
    //                 break;
    //         }
    //     }
    //
    //     scf.PushOperand(operand, resulttype, result);
    //     return;
    //
    // }

    FunctionClassList['datetime']["DAY"] = [SocialCalcFormula.DMYFunctions, 1, "v", "", "datetime"];
    FunctionClassList['datetime']["MONTH"] = [SocialCalcFormula.DMYFunctions, 1, "v", "", "datetime"];
    FunctionClassList['datetime']["YEAR"] = [SocialCalcFormula.DMYFunctions, 1, "v", "", "datetime"];
    FunctionClassList['datetime']["WEEKDAY"] = [SocialCalcFormula.DMYFunctions, -1, "weekday", "", "datetime"];

    /*
    #
    # HOUR(datetime)
    # MINUTE(datetime)
    # SECOND(datetime)
    #
    */

    SocialCalcFormula.HMSFunctions = function (fname, operand, foperand) {

        var hours, minutes, seconds, fraction;
        var scf = SocialCalcFormula;
        var result = 0;

        var datetime = scf.OperandAsNumber(foperand);
        var resulttype = scf.LookupResultType(datetime.type, datetime.type, scf.TypeLookupTable.oneargnumeric);

        if (resulttype.charAt(0) == "n") {
            if (datetime.value < 0) {
                scf.PushOperand(operand, "e#NUM!", 0); // must be non-negative
                return;
            }
            fraction = datetime.value - Math.floor(datetime.value); // fraction of a day
            fraction *= 24;
            hours = Math.floor(fraction);
            fraction -= Math.floor(fraction);
            fraction *= 60;
            minutes = Math.floor(fraction);
            fraction -= Math.floor(fraction);
            fraction *= 60;
            seconds = Math.floor(fraction + (datetime.value >= 0 ? 0.5 : -0.5));
            if (fname == "HOUR") {
                result = hours;
            }
            else if (fname == "MINUTE") {
                result = minutes;
            }
            else if (fname == "SECOND") {
                result = seconds;
            }
        }

        scf.PushOperand(operand, resulttype, result);
        return;

    }

    FunctionClassList['datetime']["HOUR"] = [SocialCalcFormula.HMSFunctions, 1, "v", "", "datetime"];
    FunctionClassList['datetime']["MINUTE"] = [SocialCalcFormula.HMSFunctions, 1, "v", "", "datetime"];
    FunctionClassList['datetime']["SECOND"] = [SocialCalcFormula.HMSFunctions, 1, "v", "", "datetime"];

    /*
    #
    # EXACT(v1,v2)
    #
    */

    SocialCalcFormula.ExactFunction = function (fname, operand, foperand) {

        var scf = SocialCalcFormula;
        var result = 0;
        var resulttype = "nl";

        var value1 = scf.OperandValueAndType(foperand);
        var v1type = value1.type.charAt(0);
        var value2 = scf.OperandValueAndType(foperand);
        var v2type = value2.type.charAt(0);

        if (v1type == "t") {
            if (v2type == "t") {
                result = value1.value == value2.value ? 1 : 0;
            }
            else if (v2type == "b") {
                result = value1.value.length ? 0 : 1;
            }
            else if (v2type == "n") {
                result = value1.value == value2.value + "" ? 1 : 0;
            }
            else if (v2type == "e") {
                result = value2.value;
                resulttype = value2.type;
            }
            else {
                result = 0;
            }
        }
        else if (v1type == "n") {
            if (v2type == "n") {
                result = value1.value - 0 == value2.value - 0 ? 1 : 0;
            }
            else if (v2type == "b") {
                result = 0;
            }
            else if (v2type == "t") {
                result = value1.value + "" == value2.value ? 1 : 0;
            }
            else if (v2type == "e") {
                result = value2.value;
                resulttype = value2.type;
            }
            else {
                result = 0;
            }
        }
        else if (v1type == "b") {
            if (v2type == "t") {
                result = value2.value.length ? 0 : 1;
            }
            else if (v2type == "b") {
                result = 1;
            }
            else if (v2type == "n") {
                result = 0;
            }
            else if (v2type == "e") {
                result = value2.value;
                resulttype = value2.type;
            }
            else {
                result = 0;
            }
        }
        else if (v1type == "e") {
            result = value1.value;
            resulttype = value1.type;
        }

        scf.PushOperand(operand, resulttype, result);
        return;

    }

    FunctionClassList['text']["EXACT"] = [SocialCalcFormula.ExactFunction, 2, "", "", "text"];

    /*
    #
    # FIND(key,string,[start])
    # LEFT(string,[length])
    # LEN(string)
    # LOWER(string)
    # MID(string,start,length)
    # PROPER(string)
    # REPLACE(string,start,length,new)
    # REPT(string,count)
    # RIGHT(string,[length])
    # SUBSTITUTE(string,old,new,[which])
    # TRIM(string)
    # UPPER(string)
    #
    */

// SocialCalcFormula.ArgList has an array for each function, one entry for each possible arg (up to max).
// Min args are specified in SocialCalcFormula.FunctionList.
// If array element is 1 then it's a text argument, if it's 0 then it's numeric, if -1 then just get whatever's there
// Text values are manipulated as UTF-8, converting from and back to byte strings

    SocialCalcFormula.ArgList = {
        FIND: [1, 1, 0],
        LEFT: [1, 0],
        LEN: [1],
        LOWER: [1],
        MID: [1, 0, 0],
        PROPER: [1],
        REPLACE: [1, 0, 0, 1],
        REPT: [1, 0],
        RIGHT: [1, 0],
        SUBSTITUTE: [1, 1, 1, 0],
        TRIM: [1],
        UPPER: [1]
    };

    SocialCalcFormula.StringFunctions = function (fname, operand, foperand) {

        var i, value, offset, len, start, count;
        var scf = SocialCalcFormula;
        var result = 0;
        var resulttype = "e#VALUE!";

        var numargs = foperand.length;
        var argdef = scf.ArgList[fname];
        var operand_value = [];
        var operand_type = [];

        for (i = 1; i <= numargs; i++) { // go through each arg, get value and type, and check for errors
            if (i > argdef.length) { // too many args
                scf.FunctionArgsError(fname, operand);
                return;
            }
            if (argdef[i - 1] == 0) {
                value = scf.OperandAsNumber(foperand);
            }
            else if (argdef[i - 1] == 1) {
                value = scf.OperandAsText(foperand);
            }
            else if (argdef[i - 1] == -1) {
                value = scf.OperandValueAndType(foperand);
            }
            operand_value[i] = value.value;
            operand_type[i] = value.type;
            if (value.type.charAt(0) == "e") {
                scf.PushOperand(operand, value.type, result);
                return;
            }
        }

        switch (fname) {
            case "FIND":
                offset = operand_type[3] ? operand_value[3] - 1 : 0;
                if (offset < 0) {
                    result = "Start is before string"; // !! not displayed, no need to translate
                }
                else {
                    result = operand_value[2].indexOf(operand_value[1], offset); // (null string matches first char)
                    if (result >= 0) {
                        result += 1;
                        resulttype = "n";
                    }
                    else {
                        result = "Not found"; // !! not displayed, error is e#VALUE!
                    }
                }
                break;

            case "LEFT":
                len = operand_type[2] ? operand_value[2] - 0 : 1;
                if (len < 0) {
                    result = "Negative length";
                }
                else {
                    result = operand_value[1].substring(0, len);
                    resulttype = "t";
                }
                break;

            case "LEN":
                result = operand_value[1].length;
                resulttype = "n";
                break;

            case "LOWER":
                result = operand_value[1].toLowerCase();
                resulttype = "t";
                break;

            case "MID":
                start = operand_value[2] - 0;
                len = operand_value[3] - 0;
                if (len < 1 || start < 1) {
                    result = "Bad arguments";
                }
                else {
                    result = operand_value[1].substring(start - 1, start + len - 1);
                    resulttype = "t";
                }
                break;

            case "PROPER":
                result = operand_value[1].replace(/\b\w+\b/g, function (word) {
                    return word.substring(0, 1).toUpperCase() +
                        word.substring(1);
                }); // uppercase first character of words (see JavaScript, Flanagan, 5th edition, page 704)
                resulttype = "t";
                break;

            case "REPLACE":
                start = operand_value[2] - 0;
                len = operand_value[3] - 0;
                if (len < 0 || start < 1) {
                    result = "Bad arguments";
                }
                else {
                    result = operand_value[1].substring(0, start - 1) + operand_value[4] +
                        operand_value[1].substring(start - 1 + len);
                    resulttype = "t";
                }
                break;

            case "REPT":
                count = operand_value[2] - 0;
                if (count < 0) {
                    result = "Negative count";
                }
                else {
                    result = "";
                    for (; count > 0; count--) {
                        result += operand_value[1];
                    }
                    resulttype = "t";
                }
                break;

            case "RIGHT":
                len = operand_type[2] ? operand_value[2] - 0 : 1;
                if (len < 0) {
                    result = "Negative length";
                }
                else {
                    result = operand_value[1].slice(-len);
                    resulttype = "t";
                }
                break;

            case "SUBSTITUTE":
                fulltext = operand_value[1];
                oldtext = operand_value[2];
                newtext = operand_value[3];
                if (operand_value[4] != null) {
                    which = operand_value[4] - 0;
                    if (which <= 0) {
                        result = "Non-positive instance number";
                        break;
                    }
                }
                else {
                    which = 0;
                }
                count = 0;
                oldpos = 0;
                result = "";
                while (true) {
                    pos = fulltext.indexOf(oldtext, oldpos);
                    if (pos >= 0) {
                        count++; //!!!!!! old test just in case: if (count>1000) {alert(pos); break;}
                        result += fulltext.substring(oldpos, pos);
                        if (which == 0) {
                            result += newtext; // substitute
                        }
                        else if (which == count) {
                            result += newtext + fulltext.substring(pos + oldtext.length);
                            break;
                        }
                        else {
                            result += oldtext; // leave as was
                        }
                        oldpos = pos + oldtext.length;
                    }
                    else { // no more
                        result += fulltext.substring(oldpos);
                        break;
                    }
                }
                resulttype = "t";
                break;

            case "TRIM":
                result = operand_value[1];
                result = result.replace(/^ */, "");
                result = result.replace(/ *$/, "");
                result = result.replace(/ +/g, " ");
                resulttype = "t";
                break;

            case "UPPER":
                result = operand_value[1].toUpperCase();
                resulttype = "t";
                break;

        }

        scf.PushOperand(operand, resulttype, result);
        return;

    }

    FunctionClassList['text']["FIND"] = [SocialCalcFormula.StringFunctions, -2, "find", "", "text"];
    FunctionClassList['text']["LEFT"] = [SocialCalcFormula.StringFunctions, -2, "tc", "", "text"];
    FunctionClassList['text']["LEN"] = [SocialCalcFormula.StringFunctions, 1, "txt", "", "text"];
    FunctionClassList['text']["LOWER"] = [SocialCalcFormula.StringFunctions, 1, "txt", "", "text"];
    FunctionClassList['text']["MID"] = [SocialCalcFormula.StringFunctions, 3, "mid", "", "text"];
    FunctionClassList['text']["PROPER"] = [SocialCalcFormula.StringFunctions, 1, "v", "", "text"];
    FunctionClassList['text']["REPLACE"] = [SocialCalcFormula.StringFunctions, 4, "replace", "", "text"];
    FunctionClassList['text']["REPT"] = [SocialCalcFormula.StringFunctions, 2, "tc", "", "text"];
    FunctionClassList['text']["RIGHT"] = [SocialCalcFormula.StringFunctions, -1, "tc", "", "text"];
    FunctionClassList['text']["SUBSTITUTE"] = [SocialCalcFormula.StringFunctions, -3, "subs", "", "text"];
    FunctionClassList['text']["TRIM"] = [SocialCalcFormula.StringFunctions, 1, "v", "", "text"];
    FunctionClassList['text']["UPPER"] = [SocialCalcFormula.StringFunctions, 1, "v", "", "text"];

    /*
    #
    # is_functions:
    #
    # ISBLANK(value)
    # ISERR(value)
    # ISERROR(value)
    # ISLOGICAL(value)
    # ISNA(value)
    # ISNONTEXT(value)
    # ISNUMBER(value)
    # ISTEXT(value)
    #
    */

    SocialCalcFormula.IsFunctions = function (fname, operand, foperand) {

        var scf = SocialCalcFormula;
        var result = 0;
        var resulttype = "nl";

        var value = scf.OperandValueAndType(foperand);
        var t = value.type.charAt(0);

        switch (fname) {

            case "ISBLANK":
                result = value.type == "b" ? 1 : 0;
                break;

            case "ISERR":
                result = t == "e" ? (value.type == "e#N/A" ? 0 : 1) : 0;
                break;

            case "ISERROR":
                result = t == "e" ? 1 : 0;
                break;

            case "ISLOGICAL":
                result = value.type == "nl" ? 1 : 0;
                break;

            case "ISNA":
                result = value.type == "e#N/A" ? 1 : 0;
                break;

            case "ISNONTEXT":
                result = t == "t" ? 0 : 1;
                break;

            case "ISNUMBER":
                result = t == "n" ? 1 : 0;
                break;

            case "ISTEXT":
                result = t == "t" ? 1 : 0;
                break;
        }

        scf.PushOperand(operand, resulttype, result);

        return;

    }

    FunctionClassList['test']["ISBLANK"] = [SocialCalcFormula.IsFunctions, 1, "v", "", "test"];
    FunctionClassList['test']["ISERR"] = [SocialCalcFormula.IsFunctions, 1, "v", "", "test"];
    FunctionClassList['test']["ISERROR"] = [SocialCalcFormula.IsFunctions, 1, "v", "", "test"];
    FunctionClassList['test']["ISLOGICAL"] = [SocialCalcFormula.IsFunctions, 1, "v", "", "test"];
    FunctionClassList['test']["ISNA"] = [SocialCalcFormula.IsFunctions, 1, "v", "", "test"];
    FunctionClassList['test']["ISNONTEXT"] = [SocialCalcFormula.IsFunctions, 1, "v", "", "test"];
    FunctionClassList['test']["ISNUMBER"] = [SocialCalcFormula.IsFunctions, 1, "v", "", "test"];
    FunctionClassList['test']["ISTEXT"] = [SocialCalcFormula.IsFunctions, 1, "v", "", "test"];

    /*
    #
    # ntv_functions:
    #
    # N(value)
    # T(value)
    # VALUE(value)
    #
    */

    SocialCalcFormula.NTVFunctions = function (fname, operand, foperand) {

        var scf = SocialCalcFormula;
        var result = 0;
        var resulttype = "e#VALUE!";

        var value = scf.OperandValueAndType(foperand);
        var t = value.type.charAt(0);

        switch (fname) {

            case "N":
                result = t == "n" ? value.value - 0 : 0;
                resulttype = "n";
                break;

            case "T":
                result = t == "t" ? value.value + "" : "";
                resulttype = "t";
                break;

            case "VALUE":
                if (t == "n" || t == "b") {
                    result = value.value || 0;
                    resulttype = "n";
                }
                else if (t == "t") {
                    value = SocialCalc.DetermineValueType(value.value);
                    if (value.type.charAt(0) != "n") {
                        result = 0;
                        resulttype = "e#VALUE!";
                    }
                    else {
                        result = value.value - 0;
                        resulttype = "n";
                    }
                }
                break;
        }

        if (t == "e") { // error trumps
            resulttype = value.type;
        }

        scf.PushOperand(operand, resulttype, result);

        return;

    }

    FunctionClassList['math']["N"] = [SocialCalcFormula.NTVFunctions, 1, "v", "", "math"];
    FunctionClassList['text']["T"] = [SocialCalcFormula.NTVFunctions, 1, "v", "", "text"];
    FunctionClassList['text']["VALUE"] = [SocialCalcFormula.NTVFunctions, 1, "v", "", "text"];

    /*
    #
    # ABS(value)
    # ACOS(value)
    # ASIN(value)
    # ATAN(value)
    # COS(value)
    # DEGREES(value)
    # EVEN(value)
    # EXP(value)
    # FACT(value)
    # INT(value)
    # LN(value)
    # LOG10(value)
    # ODD(value)
    # RADIANS(value)
    # SIN(value)
    # SQRT(value)
    # TAN(value)
    #
    */

    SocialCalcFormula.Math1Functions = function (fname, operand, foperand) {

        var v1, value, f;
        var result = {};

        var scf = SocialCalcFormula;

        v1 = scf.OperandAsNumber(foperand);
        value = v1.value;
        result.type = scf.LookupResultType(v1.type, v1.type, scf.TypeLookupTable.oneargnumeric);

        if (result.type == "n") {
            switch (fname) {
                case "ABS":
                    value = Math.abs(value);
                    break;

                case "ACOS":
                    if (value >= -1 && value <= 1) {
                        value = Math.acos(value);
                    }
                    else {
                        result.type = "e#NUM!";
                    }
                    break;

                case "ASIN":
                    if (value >= -1 && value <= 1) {
                        value = Math.asin(value);
                    }
                    else {
                        result.type = "e#NUM!";
                    }
                    break;

                case "ATAN":
                    value = Math.atan(value);
                    break;

                case "COS":
                    value = Math.cos(value);
                    break;

                case "DEGREES":
                    value = value * 180 / Math.PI;
                    break;

                case "EVEN":
                    value = value < 0 ? -value : value;
                    if (value != Math.floor(value)) {
                        value = Math.floor(value + 1) + (Math.floor(value + 1) % 2);
                    }
                    else { // integer
                        value = value + (value % 2);
                    }
                    if (v1.value < 0) value = -value;
                    break;

                case "EXP":
                    value = Math.exp(value);
                    break;

                case "FACT":
                    f = 1;
                    value = Math.floor(value);
                    for (; value > 0; value--) {
                        f *= value;
                    }
                    value = f;
                    break;

                case "INT":
                    value = Math.floor(value); // spreadsheet INT is floor(), not int()
                    break;

                case "LN":
                    if (value <= 0) {
                        result.type = "e#NUM!";
                        result.error = SocialCalc.Constants.s_sheetfunclnarg;
                    }
                    value = Math.log(value);
                    break;

                case "LOG10":
                    if (value <= 0) {
                        result.type = "e#NUM!";
                        result.error = SocialCalc.Constants.s_sheetfunclog10arg;
                    }
                    value = Math.log(value) / Math.log(10);
                    break;

                case "ODD":
                    value = value < 0 ? -value : value;
                    if (value != Math.floor(value)) {
                        value = Math.floor(value + 1) + (1 - (Math.floor(value + 1) % 2));
                    }
                    else { // integer
                        value = value + (1 - (value % 2));
                    }
                    if (v1.value < 0) value = -value;
                    break;

                case "RADIANS":
                    value = value * Math.PI / 180;
                    break;

                case "SIN":
                    value = Math.sin(value);
                    break;

                case "SQRT":
                    if (value >= 0) {
                        value = Math.sqrt(value);
                    }
                    else {
                        result.type = "e#NUM!";
                    }
                    break;

                case "TAN":
                    if (Math.cos(value) != 0) {
                        value = Math.tan(value);
                    }
                    else {
                        result.type = "e#NUM!";
                    }
                    break;
            }
        }

        result.value = value;
        operand.push(result);

        return null;

    }

// Add to function list
    FunctionClassList['math']["ABS"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["ACOS"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["ASIN"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["ATAN"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["COS"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["DEGREES"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["EVEN"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["EXP"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["FACT"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["INT"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["LN"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["LOG10"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["ODD"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["RADIANS"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["SIN"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["SQRT"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];
    FunctionClassList['math']["TAN"] = [SocialCalcFormula.Math1Functions, 1, "v", "", "math"];


    /*
    #
    # ATAN2(x, y)
    # MOD(a, b)
    # POWER(a, b)
    # TRUNC(value, precision)
    #
    */

    SocialCalcFormula.Math2Functions = function (fname, operand, foperand) {

        var xval, yval, value, quotient, decimalscale, i;
        var result = {};

        var scf = SocialCalcFormula;

        xval = scf.OperandAsNumber( foperand);
        yval = scf.OperandAsNumber( foperand);
        value = 0;
        result.type = scf.LookupResultType(xval.type, yval.type, scf.TypeLookupTable.twoargnumeric);

        if (result.type == "n") {
            switch (fname) {
                case "ATAN2":
                    if (xval.value == 0 && yval.value == 0) {
                        result.type = "e#DIV/0!";
                    }
                    else {
                        result.value = Math.atan2(yval.value, xval.value);
                    }
                    break;

                case "POWER":
                    result.value = Math.pow(xval.value, yval.value);
                    if (isNaN(result.value)) {
                        result.value = 0;
                        result.type = "e#NUM!";
                    }
                    break;

                case "MOD": // en.wikipedia.org/wiki/Modulo_operation, etc.
                    if (yval.value == 0) {
                        result.type = "e#DIV/0!";
                    }
                    else {
                        quotient = xval.value / yval.value;
                        quotient = Math.floor(quotient);
                        result.value = xval.value - (quotient * yval.value);
                    }
                    break;

                case "TRUNC":
                    decimalscale = 1; // cut down to required number of decimal digits
                    if (yval.value >= 0) {
                        yval.value = Math.floor(yval.value);
                        for (i = 0; i < yval.value; i++) {
                            decimalscale *= 10;
                        }
                        result.value = Math.floor(Math.abs(xval.value) * decimalscale) / decimalscale;
                    }
                    else if (yval.value < 0) {
                        yval.value = Math.floor(-yval.value);
                        for (i = 0; i < yval.value; i++) {
                            decimalscale *= 10;
                        }
                        result.value = Math.floor(Math.abs(xval.value) / decimalscale) * decimalscale;
                    }
                    if (xval.value < 0) {
                        result.value = -result.value;
                    }
            }
        }

        operand.push(result);

        return null;

    }

// Add to function list
    FunctionClassList['math']["ATAN2"] = [SocialCalcFormula.Math2Functions, 2, "xy", "", "math"];
    FunctionClassList['math']["MOD"] = [SocialCalcFormula.Math2Functions, 2, "", "", "math"];
    FunctionClassList['math']["POWER"] = [SocialCalcFormula.Math2Functions, 2, "", "", "math"];
    FunctionClassList['math']["TRUNC"] = [SocialCalcFormula.Math2Functions, 2, "valpre", "", "math"];

    /*
    #
    # LOG(value,[base])
    #
    */

    SocialCalcFormula.LogFunction = function (fname, operand, foperand) {

        var value, value2;
        var result = {};

        var scf = SocialCalcFormula;

        result.value = 0;

        value = scf.OperandAsNumber(foperand);
        result.type = scf.LookupResultType(value.type, value.type, scf.TypeLookupTable.oneargnumeric);
        if (foperand.length == 1) {
            value2 = scf.OperandAsNumber(foperand);
            if (value2.type.charAt(0) != "n" || value2.value <= 0) {
                scf.FunctionSpecificError(fname, operand, "e#NUM!", SocialCalc.Constants.s_sheetfunclogsecondarg);
                return 0;
            }
        }
        else if (foperand.length != 0) {
            scf.FunctionArgsError(fname, operand);
            return 0;
        }
        else {
            value2 = {value: Math.E, type: "n"};
        }

        if (result.type == "n") {
            if (value.value <= 0) {
                scf.FunctionSpecificError(fname, operand, "e#NUM!", SocialCalc.Constants.s_sheetfunclogfirstarg);
                return 0;
            }
            result.value = Math.log(value.value) / Math.log(value2.value);
        }

        operand.push(result);

        return;

    }

    FunctionClassList['math']["LOG"] = [SocialCalcFormula.LogFunction, -1, "log", "", "math"];


    /*
    #
    # ROUND(value,[precision])
    #
    */

    SocialCalcFormula.RoundFunction = function (fname, operand, foperand) {

        var value2, decimalscale, scaledvalue, i;

        var scf = SocialCalcFormula;
        var result = 0;
        var resulttype = "e#VALUE!";

        var value = scf.OperandValueAndType(foperand);
        var resulttype = scf.LookupResultType(value.type, value.type, scf.TypeLookupTable.oneargnumeric);

        if (foperand.length == 1) {
            value2 = scf.OperandValueAndType(foperand);
            if (value2.type.charAt(0) != "n") {
                scf.FunctionSpecificError(fname, operand, "e#NUM!", SocialCalc.Constants.s_sheetfuncroundsecondarg);
                return 0;
            }
        }
        else if (foperand.length != 0) {
            scf.FunctionArgsError(fname, operand);
            return 0;
        }
        else {
            value2 = {value: 0, type: "n"}; // if no second arg, assume 0 for simple round
        }

        if (resulttype == "n") {
            value2.value = value2.value - 0;
            if (value2.value == 0) {
                result = Math.round(value.value);
            }
            else if (value2.value > 0) {
                decimalscale = 1; // cut down to required number of decimal digits
                value2.value = Math.floor(value2.value);
                for (i = 0; i < value2.value; i++) {
                    decimalscale *= 10;
                }
                scaledvalue = Math.round(value.value * decimalscale);
                result = scaledvalue / decimalscale;
            }
            else if (value2.value < 0) {
                decimalscale = 1; // cut down to required number of decimal digits
                value2.value = Math.floor(-value2.value);
                for (i = 0; i < value2.value; i++) {
                    decimalscale *= 10;
                }
                scaledvalue = Math.round(value.value / decimalscale);
                result = scaledvalue * decimalscale;
            }
        }

        scf.PushOperand(operand, resulttype, result);

        return;

    }

    FunctionClassList['math']["ROUND"] = [SocialCalcFormula.RoundFunction, -1, "vp", "", "math"];

    /*
    #
    # AND(v1,c1:c2,...)
    # OR(v1,c1:c2,...)
    #
    */

    SocialCalcFormula.AndOrFunctions = function (fname, operand, foperand) {

        var value1, result;

        var scf = SocialCalcFormula;
        var resulttype = "";

        if (fname == "AND") {
            result = 1;
        }
        else if (fname == "OR") {
            result = 0;
        }

        while (foperand.length) {
            value1 = scf.OperandValueAndType(foperand);
            if (value1.type.charAt(0) == "n") {
                value1.value = value1.value - 0;
                if (fname == "AND") {
                    result = value1.value != 0 ? result : 0;
                }
                else if (fname == "OR") {
                    result = value1.value != 0 ? 1 : result;
                }
                resulttype = scf.LookupResultType(value1.type, resulttype || "nl", scf.TypeLookupTable.propagateerror);
            }
            else if (value1.type.charAt(0) == "e" && resulttype.charAt(0) != "e") {
                resulttype = value1.type;
            }
        }
        if (resulttype.length < 1) {
            resulttype = "e#VALUE!";
            result = 0;
        }
        console.log(scf)
        scf.PushOperand(operand, resulttype, result);

        return;

    }

    FunctionClassList['test']["AND"] = [SocialCalcFormula.AndOrFunctions, -1, "vn", "", "test"];
    FunctionClassList['test']["OR"] = [SocialCalcFormula.AndOrFunctions, -1, "vn", "", "test"];

    /*
    #
    # NOT(value)
    #
    */

    SocialCalcFormula.NotFunction = function (fname, operand, foperand) {

        var result = 0;
        var scf = SocialCalcFormula;
        var value = scf.OperandValueAndType(foperand);
        var resulttype = scf.LookupResultType(value.type, value.type, scf.TypeLookupTable.propagateerror);

        if (value.type.charAt(0) == "n" || value.type == "b") {
            result = value.value - 0 != 0 ? 0 : 1; // do the "not" operation
            resulttype = "nl";
        }
        else if (value.type.charAt(0) == "t") {
            resulttype = "e#VALUE!";
        }

        scf.PushOperand(operand, resulttype, result);

        return;

    }

    FunctionClassList['test']["NOT"] = [SocialCalcFormula.NotFunction, 1, "v", "", "test"];

    /*
    #
    # CHOOSE(index,value1,value2,...)
    #
    */

    SocialCalcFormula.ChooseFunction = function (fname, operand, foperand) {

        var resulttype, count, value1;
        var result = 0;
        var scf = SocialCalcFormula;

        var cindex = scf.OperandAsNumber(foperand);

        if (cindex.type.charAt(0) != "n") {
            cindex.value = 0;
        }
        cindex.value = Math.floor(cindex.value);

        count = 0;
        while (foperand.length) {
            value1 = scf.TopOfStackValueAndType(foperand);
            count += 1;
            if (cindex.value == count) {
                result = value1.value;
                resulttype = value1.type;
                break;
            }
        }
        if (resulttype) { // found something
            scf.PushOperand(operand, resulttype, result);
        }
        else {
            scf.PushOperand(operand, "e#VALUE!", 0);
        }

        return;

    }

    FunctionClassList['lookup']["CHOOSE"] = [SocialCalcFormula.ChooseFunction, -2, "choose", "", "lookup"];

    /*
    #
    # COLUMNS(c1:c2)
    # ROWS(c1:c2)
    #
    */

    SocialCalcFormula.ColumnsRowsFunctions = function (fname, operand, foperand) {

        var resulttype, rangeinfo;
        var result = 0;
        var scf = SocialCalcFormula;

        var value1 = scf.TopOfStackValueAndType(foperand);

        if (value1.type == "coord") {
            result = 1;
            resulttype = "n";
        }

        else if (value1.type == "range") {
            rangeinfo = scf.DecodeRangeParts(value1.value);
            if (fname == "COLUMNS") {
                result = rangeinfo.ncols;
            }
            else if (fname == "ROWS") {
                result = rangeinfo.nrows;
            }
            resulttype = "n";
        }
        else {
            result = 0;
            resulttype = "e#VALUE!";
        }

        scf.PushOperand(operand, resulttype, result);

        return;

    }

    FunctionClassList['lookup']["COLUMNS"] = [SocialCalcFormula.ColumnsRowsFunctions, 1, "range", "", "lookup"];
    FunctionClassList['lookup']["ROWS"] = [SocialCalcFormula.ColumnsRowsFunctions, 1, "range", "", "lookup"];


    /*
    #
    # FALSE()
    # NA()
    # NOW()
    # PI()
    # TODAY()
    # TRUE()
    #
    */

    SocialCalcFormula.ZeroArgFunctions = function (fname, operand, foperand) {

        var startval, tzoffset, start_1_1_1970, seconds_in_a_day, nowdays;
        var result = {value: 0};

        switch (fname) {
            case "FALSE":
                result.type = "nl";
                result.value = 0;
                break;

            case "NA":
                result.type = "e#N/A";
                break;

            case "NOW":
                startval = new Date();
                tzoffset = startval.getTimezoneOffset();
                startval = startval.getTime() / 1000; // convert to seconds
                start_1_1_1970 = 25569; // Day number of 1/1/1970 starting with 1/1/1900 as 1
                seconds_in_a_day = 24 * 60 * 60;
                nowdays = start_1_1_1970 + startval / seconds_in_a_day - tzoffset / (24 * 60);
                result.value = nowdays;
                result.type = "ndt";
                //SocialCalcFormula.FreshnessInfo.volatile.NOW = true; // remember
                break;

            case "PI":
                result.type = "n";
                result.value = Math.PI;
                break;

            case "TODAY":
                startval = new Date();
                tzoffset = startval.getTimezoneOffset();
                startval = startval.getTime() / 1000; // convert to seconds
                start_1_1_1970 = 25569; // Day number of 1/1/1970 starting with 1/1/1900 as 1
                seconds_in_a_day = 24 * 60 * 60;
                nowdays = start_1_1_1970 + startval / seconds_in_a_day - tzoffset / (24 * 60);
                result.value = Math.floor(nowdays);
                result.type = "nd";
                //SocialCalcFormula.FreshnessInfo.volatile.TODAY = true; // remember
                break;

            case "TRUE":
                result.type = "nl";
                result.value = 1;
                break;

        }

        operand.push(result);

        return null;

    }

// Add to function list
    FunctionClassList['test']["FALSE"] = [SocialCalcFormula.ZeroArgFunctions, 0, "", "", "test"];
    FunctionClassList['test']["NA"] = [SocialCalcFormula.ZeroArgFunctions, 0, "", "", "test"];
    FunctionClassList['datetime']["NOW"] = [SocialCalcFormula.ZeroArgFunctions, 0, "", "", "datetime"];
    FunctionClassList['math']["PI"] = [SocialCalcFormula.ZeroArgFunctions, 0, "", "", "math"];
    FunctionClassList['datetime']["TODAY"] = [SocialCalcFormula.ZeroArgFunctions, 0, "", "", "datetime"];
    FunctionClassList['test']["TRUE"] = [SocialCalcFormula.ZeroArgFunctions, 0, "", "", "test"];

//
// * * * * * FINANCIAL FUNCTIONS * * * * *
//

    /*
    #
    # DDB(cost,salvage,lifetime,period,[method])
    #
    # Depreciation, method defaults to 2 for double-declining balance
    # See: http://en.wikipedia.org/wiki/Depreciation
    #
    */

    SocialCalcFormula.DDBFunction = function (fname, operand, foperand) {

        var method, depreciation, accumulateddepreciation, i;
        var scf = SocialCalcFormula;

        var cost = scf.OperandAsNumber(foperand);
        var salvage = scf.OperandAsNumber(foperand);
        var lifetime = scf.OperandAsNumber(foperand);
        var period = scf.OperandAsNumber(foperand);

        if (scf.CheckForErrorValue(operand, cost)) return;
        if (scf.CheckForErrorValue(operand, salvage)) return;
        if (scf.CheckForErrorValue(operand, lifetime)) return;
        if (scf.CheckForErrorValue(operand, period)) return;

        if (lifetime.value < 1) {
            scf.FunctionSpecificError(fname, operand, "e#NUM!", SocialCalc.Constants.s_sheetfuncddblife);
            return 0;
        }

        method = {value: 2, type: "n"};
        if (foperand.length > 0) {
            method = scf.OperandAsNumber(foperand);
        }
        if (foperand.length != 0) {
            scf.FunctionArgsError(fname, operand);
            return 0;
        }
        if (scf.CheckForErrorValue(operand, method)) return;

        depreciation = 0; // calculated for each period
        accumulateddepreciation = 0; // accumulated by adding each period's

        for (i = 1; i <= period.value - 0 && i <= lifetime.value; i++) { // calculate for each period based on net from previous
            depreciation = (cost.value - accumulateddepreciation) * (method.value / lifetime.value);
            if (cost.value - accumulateddepreciation - depreciation < salvage.value) { // don't go lower than salvage value
                depreciation = cost.value - accumulateddepreciation - salvage.value;
            }
            accumulateddepreciation += depreciation;
        }

        scf.PushOperand(operand, 'n$', depreciation);

        return;

    }

    FunctionClassList['financial']["DDB"] = [SocialCalcFormula.DDBFunction, -4, "ddb", "", "financial"];

    /*
    #
    # SLN(cost,salvage,lifetime)
    #
    # Depreciation for each period by straight-line method
    # See: http://en.wikipedia.org/wiki/Depreciation
    #
    */

    SocialCalcFormula.SLNFunction = function (fname, operand, foperand) {

        var depreciation;
        var scf = SocialCalcFormula;

        var cost = scf.OperandAsNumber(foperand);
        var salvage = scf.OperandAsNumber(foperand);
        var lifetime = scf.OperandAsNumber(foperand);

        if (scf.CheckForErrorValue(operand, cost)) return;
        if (scf.CheckForErrorValue(operand, salvage)) return;
        if (scf.CheckForErrorValue(operand, lifetime)) return;

        if (lifetime.value < 1) {
            scf.FunctionSpecificError(fname, operand, "e#NUM!", SocialCalc.Constants.s_sheetfuncslnlife);
            return 0;
        }

        depreciation = (cost.value - salvage.value) / lifetime.value;

        scf.PushOperand(operand, 'n$', depreciation);

        return;

    }

    FunctionClassList['financial']["SLN"] = [SocialCalcFormula.SLNFunction, 3, "csl", "", "financial"];

    /*
    #
    # SYD(cost,salvage,lifetime,period)
    #
    # Depreciation by Sum of Year's Digits method
    #
    */

    SocialCalcFormula.SYDFunction = function (fname, operand, foperand) {

        var depreciation, sumperiods;
        var scf = SocialCalcFormula;

        var cost = scf.OperandAsNumber(foperand);
        var salvage = scf.OperandAsNumber(foperand);
        var lifetime = scf.OperandAsNumber(foperand);
        var period = scf.OperandAsNumber(foperand);

        if (scf.CheckForErrorValue(operand, cost)) return;
        if (scf.CheckForErrorValue(operand, salvage)) return;
        if (scf.CheckForErrorValue(operand, lifetime)) return;
        if (scf.CheckForErrorValue(operand, period)) return;

        if (lifetime.value < 1 || period.value <= 0) {
            scf.PushOperand(operand, "e#NUM!", 0);
            return 0;
        }

        sumperiods = ((lifetime.value + 1) * lifetime.value) / 2; // add up 1 through lifetime
        depreciation = (cost.value - salvage.value) * (lifetime.value - period.value + 1) / sumperiods; // calc depreciation

        scf.PushOperand(operand, 'n$', depreciation);

        return;

    }

    FunctionClassList['financial']["SYD"] = [SocialCalcFormula.SYDFunction, 4, "cslp", "", "financial"];

    /*
    #
    # FV(rate, n, payment, [pv, [paytype]])
    # NPER(rate, payment, pv, [fv, [paytype]])
    # PMT(rate, n, pv, [fv, [paytype]])
    # PV(rate, n, payment, [fv, [paytype]])
    # RATE(n, payment, pv, [fv, [paytype, [guess]]])
    #
    # Following the Open Document Format formula specification:
    #
    #    PV = - Fv - (Payment * Nper) [if rate equals 0]
    #    Pv*(1+Rate)^Nper + Payment * (1 + Rate*PaymentType) * ( (1+Rate)^nper -1)/Rate + Fv = 0
    #
    # For each function, the formulas are solved for the appropriate value (transformed using
    # basic algebra).
    #
    */

    SocialCalcFormula.InterestFunctions = function (fname, operand, foperand) {

        var resulttype, result, dval, eval, fval;
        var pv, fv, rate, n, payment, paytype, guess, part1, part2, part3, part4, part5;
        var olddelta, maxloop, tries, deltaepsilon, rate, oldrate, m;

        var scf = SocialCalcFormula;

        var aval = scf.OperandAsNumber(foperand);
        var bval = scf.OperandAsNumber(foperand);
        var cval = scf.OperandAsNumber(foperand);

        resulttype = scf.LookupResultType(aval.type, bval.type, scf.TypeLookupTable.twoargnumeric);
        resulttype = scf.LookupResultType(resulttype, cval.type, scf.TypeLookupTable.twoargnumeric);
        if (foperand.length) { // optional arguments
            dval = scf.OperandAsNumber(foperand);
            resulttype = scf.LookupResultType(resulttype, dval.type, scf.TypeLookupTable.twoargnumeric);
            if (foperand.length) { // optional arguments
                eval = scf.OperandAsNumber(foperand);
                resulttype = scf.LookupResultType(resulttype, eval.type, scf.TypeLookupTable.twoargnumeric);
                if (foperand.length) { // optional arguments
                    if (fname != "RATE") { // only rate has 6 possible args
                        scf.FunctionArgsError(fname, operand);
                        return 0;
                    }
                    fval = scf.OperandAsNumber(shfoperand);
                    resulttype = scf.LookupResultType(resulttype, fval.type, scf.TypeLookupTable.twoargnumeric);
                }
            }
        }

        if (resulttype == "n") {
            switch (fname) {
                case "FV": // FV(rate, n, payment, [pv, [paytype]])
                    rate = aval.value;
                    n = bval.value;
                    payment = cval.value;
                    pv = dval != null ? dval.value : 0; // get value if present, or use default
                    paytype = eval != null ? (eval.value ? 1 : 0) : 0;
                    if (rate == 0) { // simple calculation if no interest
                        fv = -pv - (payment * n);
                    }
                    else {
                        fv = -(pv * Math.pow(1 + rate, n) + payment * (1 + rate * paytype) * ( Math.pow(1 + rate, n) - 1) / rate);
                    }
                    result = fv;
                    resulttype = 'n$';
                    break;

                case "NPER": // NPER(rate, payment, pv, [fv, [paytype]])
                    rate = aval.value;
                    payment = bval.value;
                    pv = cval.value;
                    fv = dval != null ? dval.value : 0;
                    paytype = eval != null ? (eval.value ? 1 : 0) : 0;
                    if (rate == 0) { // simple calculation if no interest
                        if (payment == 0) {
                            scf.PushOperand(operand, "e#NUM!", 0);
                            return;
                        }
                        n = (pv + fv) / (-payment);
                    }
                    else {
                        part1 = payment * (1 + rate * paytype) / rate;
                        part2 = pv + part1;
                        if (part2 == 0 || rate <= -1) {
                            scf.PushOperand(operand, "e#NUM!", 0);
                            return;
                        }
                        part3 = (part1 - fv) / part2;
                        if (part3 <= 0) {
                            scf.PushOperand(operand, "e#NUM!", 0);
                            return;
                        }
                        part4 = Math.log(part3);
                        part5 = Math.log(1 + rate); // rate > -1
                        n = part4 / part5;
                    }
                    result = n;
                    resulttype = 'n';
                    break;

                case "PMT": // PMT(rate, n, pv, [fv, [paytype]])
                    rate = aval.value;
                    n = bval.value;
                    pv = cval.value;
                    fv = dval != null ? dval.value : 0;
                    paytype = eval != null ? (eval.value ? 1 : 0) : 0;
                    if (n == 0) {
                        scf.PushOperand(operand, "e#NUM!", 0);
                        return;
                    }
                    else if (rate == 0) { // simple calculation if no interest
                        payment = (fv - pv) / n;
                    }
                    else {
                        payment = (0 - fv - pv * Math.pow(1 + rate, n)) / ((1 + rate * paytype) * ( Math.pow(1 + rate, n) - 1) / rate);
                    }
                    result = payment;
                    resulttype = 'n$';
                    break;

                case "PV": // PV(rate, n, payment, [fv, [paytype]])
                    rate = aval.value;
                    n = bval.value;
                    payment = cval.value;
                    fv = dval != null ? dval.value : 0;
                    paytype = eval != null ? (eval.value ? 1 : 0) : 0;
                    if (rate == -1) {
                        scf.PushOperand(operand, "e#DIV/0!", 0);
                        return;
                    }
                    else if (rate == 0) { // simple calculation if no interest
                        pv = -fv - (payment * n);
                    }
                    else {
                        pv = (-fv - payment * (1 + rate * paytype) * ( Math.pow(1 + rate, n) - 1) / rate) / (Math.pow(1 + rate, n));
                    }
                    result = pv;
                    resulttype = 'n$';
                    break;

                case "RATE": // RATE(n, payment, pv, [fv, [paytype, [guess]]])
                    n = aval.value;
                    payment = bval.value;
                    pv = cval.value;
                    fv = dval != null ? dval.value : 0;
                    paytype = eval != null ? (eval.value ? 1 : 0) : 0;
                    guess = fval != null ? fval.value : 0.1;

                    // rate is calculated by repeated approximations
                    // The deltas are used to calculate new guesses

                    maxloop = 100;
                    tries = 0;
                    delta = 1;
                    epsilon = 0.0000001; // this is close enough
                    rate = guess || 0.00000001; // zero is not allowed
                    while ((delta >= 0 ? delta : -delta) > epsilon && (rate != oldrate)) {
                        delta = fv + pv * Math.pow(1 + rate, n) + payment * (1 + rate * paytype) * ( Math.pow(1 + rate, n) - 1) / rate;
                        if (olddelta != null) {
                            m = (delta - olddelta) / (rate - oldrate) || .001; // get slope (not zero)
                            oldrate = rate;
                            rate = rate - delta / m; // look for zero crossing
                            olddelta = delta;
                        }
                        else { // first time - no old values
                            oldrate = rate;
                            rate = 1.1 * rate;
                            olddelta = delta;
                        }
                        tries++;
                        if (tries >= maxloop) { // didn't converge yet
                            scf.PushOperand(operand, "e#NUM!", 0);
                            return;
                        }
                    }
                    result = rate;
                    resulttype = 'n%';
                    break;
            }
        }

        scf.PushOperand(operand, resulttype, result);

        return;

    }

    FunctionClassList['financial']["FV"] = [SocialCalcFormula.InterestFunctions, -3, "fv", "", "financial"];
    FunctionClassList['financial']["NPER"] = [SocialCalcFormula.InterestFunctions, -3, "nper", "", "financial"];
    FunctionClassList['financial']["PMT"] = [SocialCalcFormula.InterestFunctions, -3, "pmt", "", "financial"];
    FunctionClassList['financial']["PV"] = [SocialCalcFormula.InterestFunctions, -3, "pv", "", "financial"];
    FunctionClassList['financial']["RATE"] = [SocialCalcFormula.InterestFunctions, -3, "rate", "", "financial"];

    /*
    #
    # NPV(rate,v1,v2,c1:c2,...)
    #
    */

    SocialCalcFormula.NPVFunction = function (fname, operand, foperand) {

        var resulttypenpv, rate, sum, factor, value1;

        var scf = SocialCalcFormula;

        var rate = scf.OperandAsNumber(foperand);
        if (scf.CheckForErrorValue(operand, rate)) return;

        sum = 0;
        resulttypenpv = "n";
        factor = 1;

        while (foperand.length) {
            value1 = scf.OperandValueAndType(foperand);
            if (value1.type.charAt(0) == "n") {
                factor *= (1 + rate.value);
                if (factor == 0) {
                    scf.PushOperand(operand, "e#DIV/0!", 0);
                    return;
                }
                sum += value1.value / factor;
                resulttypenpv = scf.LookupResultType(value1.type, resulttypenpv || value1.type, scf.TypeLookupTable.plus);
            }
            else if (value1.type.charAt(0) == "e" && resulttypenpv.charAt(0) != "e") {
                resulttypenpv = value1.type;
                break;
            }
        }

        if (resulttypenpv.charAt(0) == "n") {
            resulttypenpv = 'n$';
        }

        scf.PushOperand(operand, resulttypenpv, sum);

        return;

    }

    FunctionClassList['financial']["NPV"] = [SocialCalcFormula.NPVFunction, -2, "npv", "", "financial"];

    /*
    #
    # IRR(c1:c2,[guess])
    #
    */

    SocialCalcFormula.IRRFunction = function (fname, operand, foperand) {

        var value1, guess, oldsum, maxloop, tries, epsilon, rate, oldrate, m, sum, factor, i;
        var rangeoperand = [];
        var cashflows = [];

        var scf = SocialCalcFormula;

        rangeoperand.push(foperand.pop()); // first operand is a range

        while (rangeoperand.length) { // get values from range so we can do iterative approximations
            value1 = scf.OperandValueAndType(rangeoperand);
            if (value1.type.charAt(0) == "n") {
                cashflows.push(value1.value);
            }
            else if (value1.type.charAt(0) == "e") {
                scf.PushOperand(operand, "e#VALUE!", 0);
                return;
            }
        }

        if (!cashflows.length) {
            scf.PushOperand(operand, "e#NUM!", 0);
            return;
        }

        guess = {value: 0};

        if (foperand.length) { // guess is provided
            guess = scf.OperandAsNumber(foperand);
            if (guess.type.charAt(0) != "n" && guess.type.charAt(0) != "b") {
                scf.PushOperand(operand, "e#VALUE!", 0);
                return;
            }
            if (foperand.length) { // should be no more args
                scf.FunctionArgsError(fname, operand);
                return;
            }
        }

        guess.value = guess.value || 0.1;

        // rate is calculated by repeated approximations
        // The deltas are used to calculate new guesses

        maxloop = 20;
        tries = 0;
        epsilon = 0.0000001; // this is close enough
        rate = guess.value;
        sum = 1;

        while ((sum >= 0 ? sum : -sum) > epsilon && (rate != oldrate)) {
            sum = 0;
            factor = 1;
            for (i = 0; i < cashflows.length; i++) {
                factor *= (1 + rate);
                if (factor == 0) {
                    scf.PushOperand(operand, "e#DIV/0!", 0);
                    return;
                }
                sum += cashflows[i] / factor;
            }

            if (oldsum != null) {
                m = (sum - oldsum) / (rate - oldrate); // get slope
                oldrate = rate;
                rate = rate - sum / m; // look for zero crossing
                oldsum = sum;
            }
            else { // first time - no old values
                oldrate = rate;
                rate = 1.1 * rate;
                oldsum = sum;
            }
            tries++;
            if (tries >= maxloop) { // didn't converge yet
                scf.PushOperand(operand, "e#NUM!", 0);
                return;
            }
        }

        scf.PushOperand(operand, 'n%', rate);

        return;

    }

    FunctionClassList['financial']["IRR"] = [SocialCalcFormula.IRRFunction, -1, "irr", "", "financial"];
    this.FunctionList = {}
    for (fclass in FunctionClassList){
        for (fname in FunctionClassList[fclass]){
            this.FunctionList[fname]=FunctionClassList[fclass][fname]
        }
    }

//
// SHEET CACHE
//

    // SocialCalcFormula.SheetCache = {
    //
    //     // Sheet data: Attributes are each sheet in the cache with values of an object with:
    //     //
    //     //    sheet: sheet-obj (or null, meaning not found)
    //     //    recalcstate: constants.asloaded = as loaded
    //     //                 constants.recalcing = being recalced now
    //     //                 constants.recalcdone = recalc done
    //     //    name: name of sheet (in case just have object and don't know name)
    //     //
    //
    //     sheets: {},
    //
    //     // Waiting for loading:
    //     // If sheet is not in cache, this is set to the sheetname being loaded
    //     // so it can be tested in the recalc loop to start load and then wait until restarted.
    //     // Reset to null before restarting.
    //
    //     waitingForLoading: null,
    //
    //     // Constants to use for setting sheets[*].recalcstate:
    //
    //     constants: {asloaded: 0, recalcing: 1, recalcdone: 2},
    //
    //     loadsheet: null // (deprecated - use SocialCalc.RecalcInfo.LoadSheet)
    //
    // };

//
// othersheet = SocialCalcFormula.FindInSheetCache(sheetname)
//
// Returns a SocialCalc.Sheet object corresponding to string sheetname
// or null if the sheet is not available or in error.
//
// Each sheet is loaded only once and then stored in a cache.
// Loading is handled elsewhere, e.g., in the recalc loop.
//

    // SocialCalcFormula.FindInSheetCache = function (sheetname) {
    //
    //     var str;
    //     var sfsc = SocialCalcFormula.SheetCache;
    //
    //     var nsheetname = SocialCalcFormula.NormalizeSheetName(sheetname); // normalize different versions
    //
    //     if (sfsc.sheets[nsheetname]) { // a sheet by that name is in the cache already
    //         return sfsc.sheets[nsheetname].sheet; // return it
    //     }
    //
    //     if (sfsc.waitingForLoading) { // waiting already - only queue up one
    //         return null; // return not found
    //     }
    //
    //     if (sfsc.loadsheet) { // Deprecated old format synchronous callback
    //         alert("Using SocialCalcFormula.SheetCache.loadsheet - deprecated");
    //         return SocialCalcFormula.AddSheetToCache(nsheetname, sfsc.loadsheet(nsheetname));
    //     }
    //
    //     sfsc.waitingForLoading = nsheetname; // let recalc loop know that we have a sheet to load
    //
    //     return null; // return not found
    //
    // }

//
// newsheet = SocialCalcFormula.AddSheetToCache(sheetname, str)
//
// Adds a new sheet to the sheet cache.
// Returns the sheet object filled out with the str (a saved sheet).
//

    // SocialCalcFormula.AddSheetToCache = function (sheetname, str) {
    //
    //     var newsheet = null;
    //     var sfsc = SocialCalcFormula.SheetCache;
    //     var sfscc = sfsc.constants;
    //     var newsheetname = SocialCalcFormula.NormalizeSheetName(sheetname);
    //
    //     if (str) {
    //         newsheet = new SocialCalc.Sheet();
    //         newsheet.ParseSheetSave(str);
    //     }
    //
    //     sfsc.sheets[newsheetname] = {sheet: newsheet, recalcstate: sfscc.asloaded, name: newsheetname};
    //
    //     SocialCalcFormula.FreshnessInfo.sheets[newsheetname] = true;
    //
    //     return newsheet;
    //
    // }

//
// nsheet = SocialCalcFormula.NormalizeSheetName(sheetname)
//

    // SocialCalcFormula.NormalizeSheetName = function (sheetname) {
    //
    //     if (SocialCalc.Callbacks.NormalizeSheetName) {
    //         return SocialCalc.Callbacks.NormalizeSheetName(sheetname);
    //     }
    //     else {
    //         return sheetname.toLowerCase();
    //     }
    // }

//
// REMOTE FUNCTION INFO
//

    // SocialCalcFormula.RemoteFunctionInfo = {
    //
    //     // Waiting for server:
    //     // If waiting for an XHR response from the server, this is set to some non-blank status text
    //     // so it can be tested in the recalc loop to start load and then wait until restarted.
    //     // Reset to null before restarting.
    //
    //     waitingForServer: null
    //
    // };

//
// FRESHNESS INFO
//
// This information is generated during recalc.
// It may be used to help determine when the recalc data in a spreadsheet
// may be out of date.
// For example, it may be used to display a message like:
// "Dependent on sheet 'FOO' which was updated more recently than this printout"

    // SocialCalcFormula.FreshnessInfo = {
    //
    //     // For each external sheet referenced successfully an attribute of that name with value true.
    //
    //     sheets: {},
    //
    //     // For each volatile function that is called an attribute of that name with value true.
    //
    //     volatile: {},
    //
    //     // Set to false when started and true when recalc completes
    //
    //     recalc_completed: false
    //
    // };
    //
    // SocialCalcFormula.FreshnessInfoReset = function () {
    //
    //     var scffi = SocialCalcFormula.FreshnessInfo;
    //
    //     scffi.sheets = {};
    //     scffi.volatile = {};
    //     scffi.recalc_completed = false;
    //
    // }

//
// MISC ROUTINES
//

//
// result = SocialCalcFormula.PlainCoord(coord)
//
// Returns: coord without any $'s
//

    // SocialCalcFormula.PlainCoord = function (coord) {
    //
    //     if (coord.indexOf("$") == -1) return coord;
    //
    //     return coord.replace(/\$/g, ""); // remove any $'s
    //
    // }

//
// result = SocialCalcFormula.OrderRangeParts(coord1, coord2)
//
// Returns: {c1: col, r1: row, c2: col, r2 = row} with c1/r1 upper left
//

    // SocialCalcFormula.OrderRangeParts = function (coord1, coord2) {
    //
    //     var cr1, cr2;
    //     var result = {};
    //
    //     cr1 = SocialCalc.coordToCr(coord1);
    //     cr2 = SocialCalc.coordToCr(coord2);
    //     if (cr1.col > cr2.col) {
    //         result.c1 = cr2.col;
    //         result.c2 = cr1.col;
    //     }
    //     else {
    //         result.c1 = cr1.col;
    //         result.c2 = cr2.col;
    //     }
    //     if (cr1.row > cr2.row) {
    //         result.r1 = cr2.row;
    //         result.r2 = cr1.row;
    //     }
    //     else {
    //         result.r1 = cr1.row;
    //         result.r2 = cr2.row;
    //     }
    //
    //     return result;
    //
    // }

//
// cond = SocialCalcFormula.TestCriteria(value, type, criteria)
//
// Determines whether a value/type meets the criteria.
// A criteria can be a numeric value, text beginning with <, <=, =, >=, >, <>, text by itself is start of text to match.
// Used by a variety of functions, including the "D" functions (DSUM, etc.).
//
// Returns true or false
//

    // SocialCalcFormula.TestCriteria = function (value, type, criteria) {
    //
    //     var comparitor, basestring, basevalue, cond, testvalue;
    //
    //     if (criteria == null) { // undefined (e.g., error value) is always false
    //         return false;
    //     }
    //
    //     criteria = criteria + "";
    //     comparitor = criteria.charAt(0); // look for comparitor
    //     if (comparitor == "=" || comparitor == "<" || comparitor == ">") {
    //         basestring = criteria.substring(1);
    //     }
    //     else {
    //         comparitor = criteria.substring(0, 2);
    //         if (comparitor == "<=" || comparitor == "<>" || comparitor == ">=") {
    //             basestring = criteria.substring(2);
    //         }
    //         else {
    //             comparitor = "none";
    //             basestring = criteria;
    //         }
    //     }
    //
    //     basevalue = SocialCalc.DetermineValueType(basestring); // get type of value being compared
    //     if (!basevalue.type) { // no criteria base value given
    //         if (comparitor == "none") { // blank criteria matches nothing
    //             return false;
    //         }
    //         if (type.charAt(0) == "b") { // comparing to empty cell
    //             if (comparitor == "=") { // empty equals empty
    //                 return true;
    //             }
    //         }
    //         else {
    //             if (comparitor == "<>") { // "something" does not equal empty
    //                 return true;
    //             }
    //         }
    //         return false; // otherwise false
    //     }
    //
    //     cond = false;
    //
    //     if (basevalue.type.charAt(0) == "n" && type.charAt(0) == "t") { // criteria is number, but value is text
    //         testvalue = SocialCalc.DetermineValueType(value);
    //         if (testvalue.type.charAt(0) == "n") { // could be number - make it one
    //             value = testvalue.value;
    //             type = testvalue.type;
    //         }
    //     }
    //
    //     if (type.charAt(0) == "n" && basevalue.type.charAt(0) == "n") { // compare two numbers
    //         value = value - 0; // make sure numbers
    //         basevalue.value = basevalue.value - 0;
    //         switch (comparitor) {
    //             case "<":
    //                 cond = value < basevalue.value;
    //                 break;
    //
    //             case "<=":
    //                 cond = value <= basevalue.value;
    //                 break;
    //
    //             case "=":
    //             case "none":
    //                 cond = value == basevalue.value;
    //                 break;
    //
    //             case ">=":
    //                 cond = value >= basevalue.value;
    //                 break;
    //
    //             case ">":
    //                 cond = value > basevalue.value;
    //                 break;
    //
    //             case "<>":
    //                 cond = value != basevalue.value;
    //                 break;
    //         }
    //     }
    //
    //     else if (type.charAt(0) == "e") { // error on left
    //         cond = false;
    //     }
    //
    //     else if (basevalue.type.charAt(0) == "e") { // error on right
    //         cond = false;
    //     }
    //
    //     else { // text, maybe mixed with number or blank
    //         if (type.charAt(0) == "n") {
    //             value = SocialCalc.format_number_for_display(value, "n", "");
    //         }
    //         if (basevalue.type.charAt(0) == "n") {
    //             return false; // if number and didn't match already, isn't a match
    //         }
    //
    //         value = value ? value.toLowerCase() : "";
    //         basevalue.value = basevalue.value ? basevalue.value.toLowerCase() : "";
    //
    //         switch (comparitor) {
    //             case "<":
    //                 cond = value < basevalue.value;
    //                 break;
    //
    //             case "<=":
    //                 cond = value <= basevalue.value;
    //                 break;
    //
    //             case "=":
    //                 cond = value == basevalue.value;
    //                 break;
    //
    //             case "none":
    //                 cond = value.substring(0, basevalue.value.length) == basevalue.value;
    //                 break;
    //
    //             case ">=":
    //                 cond = value >= basevalue.value;
    //                 break;
    //
    //             case ">":
    //                 cond = value > basevalue.value;
    //                 break;
    //
    //             case "<>":
    //                 cond = value != basevalue.value;
    //                 break;
    //         }
    //     }
    //
    //     return cond;
    //
    // }
}

Formula.prototype.ParseFormulaIntoTokens = function (line) {
    console.log(line)
    var i, ch, chclass, haddecimal, last_token, last_token_type, last_token_text, t;

    //var scf = SocialCalcFormula;
    var scc = this.scc
    var parsestate = this.ParseState;
    var tokentype = this.TokenType;
    var charclass = this.CharClass;
    var charclasstable = this.CharClassTable;
    var uppercasetable = this.UpperCaseTable; // much faster than toUpperCase function
    var pushtoken = this.ParsePushToken;
    var coordregex = /^\$?[A-Z]{1,2}\$?[1-9]\d*$/i;

    var parseinfo = [];
    var str = "";
    var state = 0;
    var haddecimal = false;

    for (i = 0; i <= line.length; i++) {
        if (i < line.length) {
            ch = line.charAt(i);
            cclass = charclasstable[ch];
        }
        else {
            ch = "";
            cclass = charclass.eof;
        }

        if (state == parsestate.num) {
            if (cclass == charclass.num) {
                str += ch;
            }
            else if (cclass == charclass.numstart && !haddecimal) {
                haddecimal = true;
                str += ch;
            }
            else if (ch == "E" || ch == "e") {
                str += ch;
                haddecimal = false;
                state = parsestate.numexp1;
            }
            else { // end of number - save it
                pushtoken(parseinfo, str, tokentype.num, 0);
                haddecimal = false;
                state = 0;
            }
        }

        if (state == parsestate.numexp1) {
            if (cclass == parsestate.num) {
                state = parsestate.numexp2;
            }
            else if ((ch == '+' || ch == '-') && (uppercasetable[str.charAt(str.length - 1)] == 'E')) {
                str += ch;
            }
            else if (ch == 'E' || ch == 'e') {
                ;
            }
            else {
                pushtoken(parseinfo, this.s_parseerrexponent, tokentype.error, 0);
                state = 0;
            }
        }

        if (state == parsestate.numexp2) {
            if (cclass == charclass.num) {
                str += ch;
            }
            else { // end of number - save it
                pushtoken(parseinfo, str, tokentype.num, 0);
                state = 0;
            }
        }

        if (state == parsestate.alpha) {
            if (cclass == charclass.num) {
                state = parsestate.coord;
            }
            else if (cclass == charclass.alpha || ch == ".") { // alpha may be letters, numbers, "_", or "."
                str += ch;
            }
            else if (cclass == charclass.incoord) {
                state = parsestate.coord;
            }
            else if (cclass == charclass.op || cclass == charclass.numstart
                || cclass == charclass.space || cclass == charclass.eof) {
                pushtoken(parseinfo, str.toUpperCase(), tokentype.name, 0);
                state = 0;
            }
            else {
                pushtoken(parseinfo, scc.s_parseerrchar, tokentype.error, 0);
                state = 0;
            }
        }

        if (state == parsestate.coord) {
            if (cclass == charclass.num) {
                str += ch;
            }
            else if (cclass == charclass.incoord) {
                str += ch;
            }
            else if (cclass == charclass.alpha) {
                state = parsestate.alphanumeric;
            }
            else if (cclass == charclass.op || cclass == charclass.numstart ||
                cclass == charclass.eof || cclass == charclass.space) {
                if (coordregex.test(str)) {
                    t = tokentype.coord;
                }
                else {
                    t = tokentype.name;
                }
                pushtoken(parseinfo, str.toUpperCase(), t, 0);
                state = 0;
            }
            else {
                pushtoken(parseinfo, scc.s_parseerrchar, tokentype.error, 0);
                state = 0;
            }
        }


        if (state == parsestate.alphanumeric) {
            if (cclass == charclass.num || cclass == charclass.alpha) {
                str += ch;
            }
            else if (cclass == charclass.op || cclass == charclass.numstart
                || cclass == charclass.space || cclass == charclass.eof) {
                pushtoken(parseinfo, str.toUpperCase(), tokentype.name, 0);
                state = 0;
            }
            else {
                pushtoken(parseinfo, scc.s_parseerrchar, tokentype.error, 0);
                state = 0;
            }
        }

        if (state == parsestate.string) {
            if (cclass == charclass.quote) {
                state = parsestate.stringquote; // got quote in string: is it doubled (quote in string) or by itself (end of string)?
            }
            else if (cclass == charclass.eof) {
                pushtoken(parseinfo, scc.s_parseerrstring, tokentype.error, 0);
                state = 0;
            }
            else {
                str += ch;
            }
        }
        else if (state == parsestate.stringquote) { // note else if here
            if (cclass == charclass.quote) {
                str += '"';
                state = parsestate.string; // double quote: add one then continue getting string
            }
            else { // something else -- end of string
                pushtoken(parseinfo, str, tokentype.string, 0);
                state = 0; // drop through to process
            }
        }

        else if (state == parsestate.specialvalue) { // special values like #REF!
            if (str.charAt(str.length - 1) == "!") { // done - save value as a name
                pushtoken(parseinfo, str, tokentype.name, 0);
                state = 0; // drop through to process
            }
            else if (cclass == charclass.eof) {
                pushtoken(parseinfo, scc.s_parseerrspecialvalue, tokentype.error, 0);
                state = 0;
            }
            else {
                str += ch;
            }
        }

        if (state == 0) {
            if (cclass == charclass.num) {
                str = ch;
                state = parsestate.num;
            }
            else if (cclass == charclass.numstart) {
                str = ch;
                haddecimal = true;
                state = parsestate.num;
            }
            else if (cclass == charclass.alpha || cclass == charclass.incoord) {
                str = ch;
                state = parsestate.alpha;
            }
            else if (cclass == charclass.specialstart) {
                str = ch;
                state = parsestate.specialvalue;
            }
            else if (cclass == charclass.op) {
                str = ch;
                if (parseinfo.length > 0) {
                    last_token = parseinfo[parseinfo.length - 1];
                    last_token_type = last_token.type;
                    last_token_text = last_token.text;
                    if (last_token_type == charclass.op) {
                        if (last_token_text == '<' || last_token_text == ">") {
                            str = last_token_text + str;
                            parseinfo.pop();
                            if (parseinfo.length > 0) {
                                last_token = parseinfo[parseinfo.length - 1];
                                last_token_type = last_token.type;
                                last_token_text = last_token.text;
                            }
                            else {
                                last_token_type = charclass.eof;
                                last_token_text = "EOF";
                            }
                        }
                    }
                }
                else {
                    last_token_type = charclass.eof;
                    last_token_text = "EOF";
                }
                t = tokentype.op;
                if ((parseinfo.length == 0)
                    || (last_token_type == charclass.op && last_token_text != ')' && last_token_text != '%')) { // Unary operator
                    if (str == '-') { // M is unary minus
                        str = "M";
                        ch = "M";
                    }
                    else if (str == '+') { // P is unary plus
                        str = "P";
                        ch = "P";
                    }
                    else if (str == ')' && last_token_text == '(') { // null arg list OK
                        ;
                    }
                    else if (str != '(') { // binary-op open-paren OK, others no
                        t = tokentype.error;
                        str = this.s_parseerrtwoops;
                    }
                }
                else if (str.length > 1) {
                    if (str == '>=') { // G is >=
                        str = "G";
                        ch = "G";
                    }
                    else if (str == '<=') { // L is <=
                        str = "L";
                        ch = "L";
                    }
                    else if (str == '<>') { // N is <>
                        str = "N";
                        ch = "N";
                    }
                    else {
                        t = tokentype.error;
                        str = scc.s_parseerrtwoops;
                    }
                }
                pushtoken(parseinfo, str, t, ch);
                state = 0;
            }
            else if (cclass == charclass.quote) { // starting a string
                str = "";
                state = parsestate.string;
            }
            else if (cclass == charclass.space) { // store so can reconstruct spacing
                pushtoken(parseinfo, " ", tokentype.space, 0);
            }
            else if (cclass == charclass.eof) { // ignore -- needed to have extra loop to close out other things
            }
            else { // unknown class - such as unknown char
                pushtoken(parseinfo, scc.s_parseerrchar, tokentype.error, 0);
            }
        }
    }
    return parseinfo;

}
Formula.prototype.ParsePushToken = function (parseinfo, ttext, ttype, topcode) {

    parseinfo.push({text: ttext, type: ttype, opcode: topcode});

}

Formula.prototype.evaluate_parsed_formula = function(parseinfo,allowrangereturn) {

    var result;

    var scf = this
    var tokentype = scf.TokenType;

    var revpolish;
    var parsestack = [];

    var errortext = "";
    console.log(parseinfo)
    revpolish = this.ConvertInfixToPolish(parseinfo); // result is either an array or a string with error text
    console.log(revpolish)
    result = this.EvaluatePolish(parseinfo, revpolish, allowrangereturn);
    console.log(result)
    return result;

}
Formula.prototype.ConvertInfixToPolish = function(parseinfo) {

    var scf = this
    var scc = this.Constants;
    var tokentype = scf.TokenType;
    var token_precedence = scf.TokenPrecedence;

    var revpolish = [];
    var parsestack = [];

    var errortext = "";

    var function_start = -1;

    var i, pii, ttype, ttext, tprecedence, tstackprecedence;

    for (i=0; i<parseinfo.length; i++) {
        pii = parseinfo[i];
        ttype = pii.type;
        ttext = pii.text;
        if (ttype == tokentype.num || ttype == tokentype.coord || ttype == tokentype.string) {
            revpolish.push(i);
        }
        else if (ttype == tokentype.name) {
            parsestack.push(i);
            revpolish.push(function_start);
        }
        else if (ttype == tokentype.space) { // ignore
            continue;
        }
        else if (ttext == ',') {
            while (parsestack.length && parseinfo[parsestack[parsestack.length-1]].text != "(") {
                revpolish.push(parsestack.pop());
            }
            if (parsestack.length == 0) { // no ( -- error
                errortext = scc.s_parseerrmissingopenparen;
                break;
            }
        }
        else if (ttext == '(') {
            parsestack.push(i);
        }
        else if (ttext == ')') {
            while (parsestack.length && parseinfo[parsestack[parsestack.length-1]].text != "(") {
                revpolish.push(parsestack.pop());
            }
            if (parsestack.length == 0) { // no ( -- error
                errortext = scc.s_parseerrcloseparennoopen;
                break;
            }
            parsestack.pop();
            if (parsestack.length && parseinfo[parsestack[parsestack.length-1]].type == tokentype.name) {
                revpolish.push(parsestack.pop());
            }
        }
        else if (ttype == tokentype.op) {
            if (parsestack.length && parseinfo[parsestack[parsestack.length-1]].type == tokentype.name) {
                revpolish.push(parsestack.pop());
            }
            while (parsestack.length && parseinfo[parsestack[parsestack.length-1]].type == tokentype.op
            && parseinfo[parsestack[parsestack.length-1]].text != '(') {
                tprecedence = token_precedence[pii.opcode];
                tstackprecedence = token_precedence[parseinfo[parsestack[parsestack.length-1]].opcode];
                if (tprecedence >= 0 && tprecedence < tstackprecedence) {
                    break;
                }
                else if (tprecedence < 0) {
                    tprecedence = -tprecedence;
                    if (tstackprecedence < 0) tstackprecedence = -tstackprecedence;
                    if (tprecedence <= tstackprecedence) {
                        break;
                    }
                }
                revpolish.push(parsestack.pop());
            }
            parsestack.push(i);
        }
        else if (ttype == tokentype.error) {
            errortext = ttext;
            break;
        }
        else {
            errortext = "Internal error while processing parsed formula. ";
            break;
        }
    }
    while (parsestack.length>0) {
        if (parseinfo[parsestack[parsestack.length-1]].text == '(') {
            errortext = scc.s_parseerrmissingcloseparen;
            break;
        }
        revpolish.push(parsestack.pop());
    }

    if (errortext) {
        return errortext;
    }

    return revpolish;

}
Formula.prototype.EvaluatePolish = function(parseinfo, revpolish, allowrangereturn) {

    var scf = this
    var scc = this.scc
    var tokentype = this.TokenType
    var lookup_result_type = this.LookupResultType
    var typelookup = this.TypeLookupTable

    var operand_as_text = this.OperandAsText
    var operand_value_and_type = this.OperandValueAndType
    var operands_as_coord_on_sheet = this.OperandsAsCoordOnSheet
    //var format_number_for_display = SocialCalc.format_number_for_display || function(v, t, f) {return v+"";};

    var errortext = "";
    var function_start = -1;
    var missingOperandError = {value: "", type: "e#VALUE!"}//error: scc.s_parseerrmissingoperand};

    var operand = [];
    var PushOperand = function(t, v) {operand.push({type: t, value: v});};

    var i, rii, prii, ttype, ttext, value1, value2, tostype, tostype2, resulttype, valuetype, cond, vmatch, smatch;

    if (!parseinfo.length || (! (revpolish instanceof Array))) {
        return ({value: "", type: "e#VALUE!", error: (typeof revpolish == "string" ? revpolish : "")});
    }

    for (i=0; i<revpolish.length; i++) {
        rii = revpolish[i];
        if (rii == function_start) { // Remember the start of a function argument list
            PushOperand("start", 0);
            continue;
        }

        prii = parseinfo[rii];
        ttype = prii.type;
        ttext = prii.text;

        if (ttype == tokentype.num) {
            PushOperand("n", ttext-0);
        }

        // else if (ttype == tokentype.coord) {
        //     PushOperand("coord", ttext);
        // }

        // else if (ttype == tokentype.string) {
        //     PushOperand("t", ttext);
        // }

        else if (ttype == tokentype.op) {
            if (operand.length <= 0) { // Nothing on the stack...
                return missingOperandError;
                break; // done
            }

            // Unary minus

            if (ttext == 'M') {
                value1 = this.OperandAsNumber( operand);
                resulttype = this.LookupResultType(value1.type, value1.type, typelookup.unaryminus);
                PushOperand(resulttype, -value1.value);
            }

            // Unary plus

            else if (ttext == 'P') {
                value1 = this.OperandAsNumber( operand);
                resulttype = this.LookupResultType(value1.type, value1.type, typelookup.unaryplus);
                PushOperand(resulttype, value1.value);
            }

            // Unary % - percent, left associative

            else if (ttext == '%') {
                value1 = this.OperandAsNumber(operand);
                resulttype = this.LookupResultType(value1.type, value1.type, typelookup.unarypercent);
                PushOperand(resulttype, 0.01*value1.value);
            }

            // & - string concatenate

            // else if (ttext == '&') {
            //     if (operand.length <= 1) { // Need at least two things on the stack...
            //         return missingOperandError;
            //     }
            //     value2 = operand_as_text( operand);
            //     value1 = operand_as_text(operand);
            //     resulttype = this.LookupResultType(value1.type, value1.type, typelookup.concat);
            //     PushOperand(resulttype, value1.value + value2.value);
            // }

            // : - Range constructor

            // else if (ttext == ':') {
            //     if (operand.length <= 1) { // Need at least two things on the stack...
            //         return missingOperandError;
            //     }
            //     value1 = scf.OperandsAsRangeOnSheet( operand); // get coords even if use name on other sheet
            //     if (value1.error) { // not available
            //         errortext = errortext || value1.error;
            //     }
            //     PushOperand(value1.type, value1.value); // push sheetname with range on that sheet
            // }

            // ! - sheetname!coord

            // else if (ttext == '!') {
            //     if (operand.length <= 1) { // Need at least two things on the stack...
            //         return missingOperandError;
            //     }
            //     value1 = operands_as_coord_on_sheet( operand); // get coord even if name on other sheet
            //     if (value1.error) { // not available
            //         errortext = errortext || value1.error;
            //     }
            //     PushOperand(value1.type, value1.value); // push sheetname with coord or range on that sheet
            // }

            // Comparison operators: < L = G > N (< <= = >= > <>)

            else if (ttext == "<" || ttext == "L" || ttext == "=" || ttext == "G" || ttext == ">" || ttext == "N") {
                if (operand.length <= 1) { // Need at least two things on the stack...
                    errortext = scc.s_parseerrmissingoperand; // remember error
                    break;
                }
                value2 = operand_value_and_type( operand);
                value1 = operand_value_and_type( operand);
                if (value1.type.charAt(0) == "n" && value2.type.charAt(0) == "n") { // compare two numbers
                    cond = 0;
                    if (ttext == "<") { cond = value1.value < value2.value ? 1 : 0; }
                    else if (ttext == "L") { cond = value1.value <= value2.value ? 1 : 0; }
                    else if (ttext == "=") { cond = value1.value == value2.value ? 1 : 0; }
                    else if (ttext == "G") { cond = value1.value >= value2.value ? 1 : 0; }
                    else if (ttext == ">") { cond = value1.value > value2.value ? 1 : 0; }
                    else if (ttext == "N") { cond = value1.value != value2.value ? 1 : 0; }
                    PushOperand("nl", cond);
                }
                else if (value1.type.charAt(0) == "e") { // error on left
                    PushOperand(value1.type, 0);
                }
                else if (value2.type.charAt(0) == "e") { // error on right
                    PushOperand(value2.type, 0);
                }
                else { // text maybe mixed with numbers or blank
                    tostype = value1.type.charAt(0);
                    tostype2 = value2.type.charAt(0);
                    if (tostype == "n") {
                        value1.value = format_number_for_display(value1.value, "n", "");
                    }
                    else if (tostype == "b") {
                        value1.value = "";
                    }
                    if (tostype2 == "n") {
                        value2.value = format_number_for_display(value2.value, "n", "");
                    }
                    else if (tostype2 == "b") {
                        value2.value = "";
                    }
                    cond = 0;
                    value1.value = value1.value.toLowerCase(); // ignore case
                    value2.value = value2.value.toLowerCase();
                    if (ttext == "<") { cond = value1.value < value2.value ? 1 : 0; }
                    else if (ttext == "L") { cond = value1.value <= value2.value ? 1 : 0; }
                    else if (ttext == "=") { cond = value1.value == value2.value ? 1 : 0; }
                    else if (ttext == "G") { cond = value1.value >= value2.value ? 1 : 0; }
                    else if (ttext == ">") { cond = value1.value > value2.value ? 1 : 0; }
                    else if (ttext == "N") { cond = value1.value != value2.value ? 1 : 0; }
                    PushOperand("nl", cond);
                }
            }

            // Normal infix arithmethic operators: +, -. *, /, ^

            else { // what's left are the normal infix arithmetic operators
                if (operand.length <= 1) { // Need at least two things on the stack...
                    errortext = scc.s_parseerrmissingoperand; // remember error
                    break;
                }
                value2 = this.OperandAsNumber( operand);
                value1 = this.OperandAsNumber( operand);
                if (ttext == '+') {
                    resulttype = lookup_result_type(value1.type, value2.type, typelookup.plus);
                    PushOperand(resulttype, value1.value + value2.value);
                }
                else if (ttext == '-') {
                    resulttype = lookup_result_type(value1.type, value2.type, typelookup.plus);
                    PushOperand(resulttype, value1.value - value2.value);
                }
                else if (ttext == '*') {
                    resulttype = lookup_result_type(value1.type, value2.type, typelookup.plus);
                    PushOperand(resulttype, value1.value * value2.value);
                }
                else if (ttext == '/') {
                    if (value2.value != 0) {
                        PushOperand("n", value1.value / value2.value); // gives plain numeric result type
                    }
                    else {
                        PushOperand("e#DIV/0!", 0);
                    }
                }
                else if (ttext == '^') {
                    value1.value = Math.pow(value1.value, value2.value);
                    value1.type = "n"; // gives plain numeric result type
                    if (isNaN(value1.value)) {
                        value1.value = 0;
                        value1.type = "e#NUM!";
                    }
                    PushOperand(value1.type, value1.value);
                }
            }
        }

        // function or name

        else if (ttype == tokentype.name) {
            errortext = scf.CalculateFunction(ttext, operand);
            if (errortext) break;
        }

        // else {
        //     errortext = scc.s_InternalError+"Unknown token "+ttype+" ("+ttext+"). ";
        //     break;
        // }
    }

    // look at final value and handle special cases

    value = operand[0] ? operand[0].value : "";
    tostype = operand[0] ? operand[0].type : "";

    // if (tostype == "name") { // name - expand it
    //     value1 = SocialCalcFormula.LookupName(sheet, value);
    //     value = value1.value;
    //     tostype = value1.type;
    //     errortext = errortext || value1.error;
    // }

    // if (tostype == "coord") { // the value is a coord reference, get its value and type
    //     value1 = operand_value_and_type(sheet, operand);
    //     value = value1.value;
    //     tostype = value1.type;
    //     if (tostype == "b") {
    //         tostype = "n";
    //         value = 0;
    //     }
    // }

    if (operand.length > 1 && !errortext) { // something left - error
        errortext += scc.s_parseerrerrorinformula;
    }

    // set return type

    valuetype = tostype;

    if (tostype.charAt(0) == "e") { // error value
        errortext = errortext || tostype.substring(1) || scc.s_calcerrerrorvalueinformula;
    }
    else if (tostype == "range") {
        vmatch = value.match(/^(.*)\|(.*)\|/);
        smatch = vmatch[1].indexOf("!");
        if (smatch>=0) { // swap sheetname
            vmatch[1] = vmatch[1].substring(smatch+1) + "!" + vmatch[1].substring(0, smatch).toUpperCase();
        }
        else {
            vmatch[1] = vmatch[1].toUpperCase();
        }
        value = vmatch[1] + ":" + vmatch[2].toUpperCase();
        if (!allowrangereturn) {
            errortext = scc.s_formularangeresult+" "+value;
        }
    }

    if (errortext && valuetype.charAt(0) != "e") {
        value = errortext;
        valuetype = "e";
    }

    // look for overflow

    if (valuetype.charAt(0) == "n" && (isNaN(value) || !isFinite(value))) {
        value = 0;
        valuetype = "e#NUM!";
        errortext = isNaN(value) ? scc.s_calcerrnumericnan: scc.s_calcerrnumericoverflow;
    }

    return ({value: value, type: valuetype, error: errortext});

}

Formula.prototype.LookupResultType = function(type1, type2, typelookup) {

    var pos1, pos2, result;

    var table1 = typelookup[type1];

    if (!table1) {
        table1 = typelookup[type1.charAt(0)+'*'];
        if (!table1) {
            return "e#VALUE! (internal error, missing LookupResultType "+type1.charAt(0)+"*)"; // missing from table -- please add it
        }
    }
    pos1 = table1.indexOf("|"+type2+":");
    if (pos1 >= 0) {
        pos2 = table1.indexOf("|", pos1+1);
        if (pos2<0) return "e#VALUE! (internal error, incorrect LookupResultType "+table1+")";
        result = table1.substring(pos1+type2.length+2, pos2);
        if (result == "1") return type1;
        if (result == "2") return type2;
        return result;
    }
    pos1 = table1.indexOf("|"+type2.charAt(0)+"*:");
    if (pos1 >= 0) {
        pos2 = table1.indexOf("|", pos1+1);
        if (pos2<0) return "e#VALUE! (internal error, incorrect LookupResultType "+table1+")";
        result = table1.substring(pos1+4, pos2);
        if (result == "1") return type1;
        if (result == "2") return type2;
        return result;
    }
    return "e#VALUE!";

}

Formula.prototype.OperandAsNumber = function(operand) {

    var t, valueinfo;
    var operandinfo = this.OperandValueAndType(operand)

    t = operandinfo.type.charAt(0);

    if (t == "n") {
        operandinfo.value = operandinfo.value-0;
    }
    else if (t == "b") { // blank cell
        operandinfo.type = "n";
        operandinfo.value = 0;
    }
    else if (t == "e") { // error
        operandinfo.value = 0;
    }
    else {
        valueinfo = SocialCalc.DetermineValueType ? SocialCalc.DetermineValueType(operandinfo.value) :
            {value: operandinfo.value-0, type: "n"}; // if without rest of SocialCalc
        if (valueinfo.type.charAt(0) == "n") {
            operandinfo.value = valueinfo.value-0;
            operandinfo.type = valueinfo.type;
        }
        else {
            operandinfo.value = 0;
            operandinfo.type = valueinfo.type;
        }
    }

    return operandinfo;

}
Formula.prototype.OperandValueAndType = function(operand) {
    console.log(operand)
    console.log('a')
    var cellvtype, cell, pos, coordsheet;
    var scf = this;

    var result = {type: "", value: ""};

    var stacklen = operand.length;

    if (!stacklen) { // make sure something is there
        //result.error = SocialCalc.Constants.s_InternalError+"no operand on stack";
        return result;
    }

    result.value = operand[stacklen-1].value; // get top of stack
    result.type = operand[stacklen-1].type;
    operand.pop(); // we have data - pop stack

    // if (result.type == "name") {
    //     result = scf.LookupName( result.value);
    // }
    //
    // if (result.type == "range") {
    //     result = scf.StepThroughRangeDown(operand, result.value);
    // }
    //
    // if (result.type == "coord") { // value is a coord reference
    //     coordsheet = sheet;
    //     pos = result.value.indexOf("!");
    //     if (pos != -1) { // sheet reference
    //         coordsheet = scf.FindInSheetCache(result.value.substring(pos+1)); // get other sheet
    //         if (coordsheet == null) { // unavailable
    //             result.type = "e#REF!";
    //             result.error = SocialCalc.Constants.s_sheetunavailable+" "+result.value.substring(pos+1);
    //             result.value = 0;
    //             return result;
    //         }
    //         result.value = result.value.substring(0, pos); // get coord part
    //     }
    //
    //     if (coordsheet) {
    //         cell = coordsheet.cells[SocialCalcFormula.PlainCoord(result.value)];
    //         if (cell) {
    //             cellvtype = cell.valuetype; // get type of value in the cell it points to
    //             result.value = cell.datavalue;
    //         }
    //         else {
    //             cellvtype = "b";
    //         }
    //     }
    //     else {
    //         cellvtype = "e#N/A";
    //         result.value = 0;
    //     }
    //     result.type = cellvtype || "b";
    //     if (result.type == "b") { // blank
    //         result.value = 0;
    //     }
    // }

    return result;

}
Formula.prototype.CalculateFunction = function(fname, operand) {
    console.log(fname)
    var fobj, foperand, ffunc, argnum, ttext;
    var scf = this
    var ok = 1;
    var errortext = "";

    fobj = scf.FunctionList[fname];

    if (fobj) {
        foperand = [];
        ffunc = fobj[0];
        argnum = fobj[1];
        scf.CopyFunctionArgs(operand, foperand);
        if (argnum != 100) {
            if (argnum < 0) {
                if (foperand.length < -argnum) {
                    errortext = scf.FunctionArgsError(fname, operand);
                    return errortext;
                }
            }
            else {
                if (foperand.length != argnum) {
                    errortext = scf.FunctionArgsError(fname, operand);
                    return errortext;
                }
            }
        }
        errortext = ffunc(fname, operand, foperand);
    }

    else {
        ttext = fname;

        if (operand.length && operand[operand.length-1].type == "start") { // no arguments - name or zero arg function
            operand.pop();
            scf.PushOperand(operand, "name", ttext);
        }

        else {
            errortext = SocialCalc.Constants.s_sheetfuncunknownfunction+" " + ttext +". ";
        }
    }

    return errortext;

}
Formula.prototype.CopyFunctionArgs = function(operand, foperand) {

    var fobj, foperand, ffunc, argnum;
    var scf = this
    var ok = 1;
    var errortext = null;

    while (operand.length>0 && operand[operand.length-1].type != "start") { // get each arg
        foperand.push(operand.pop()); // copy it
    }
    operand.pop(); // get rid of "start"

    return;

}
Formula.prototype.FunctionArgsError = function(fname, operand) {

    var errortext = this.scc.s_calcerrincorrectargstofunction+" " + fname + ". ";
    this.PushOperand(operand, "e#VALUE!", errortext);

    return errortext;

}
Formula.prototype.DecodeRangeParts = function(sheetdata, range) {

    var value1, value2, pos1, pos2, sheet1, coordsheetdata, rp;

    var scf = SocialCalc.Formula;

    pos1 = range.indexOf("|");
    pos2 = range.indexOf("|", pos1+1);
    value1 = range.substring(0, pos1);
    value2 = range.substring(pos1+1, pos2);

    pos1 = value1.indexOf("!");
    if (pos1 != -1) {
        sheet1 = value1.substring(pos1+1);
        value1 = value1.substring(0, pos1);
    }
    else {
        sheet1 = "";
    }
    pos1 = value2.indexOf("!");
    if (pos1 != -1) {
        value2 = value2.substring(0, pos1);
    }

    coordsheetdata = sheetdata;
    if (sheet1) { // sheet reference
        coordsheetdata = scf.FindInSheetCache(sheet1);
        if (coordsheetdata == null) { // unavailable
            return null;
        }
    }

    rp = scf.OrderRangeParts(value1, value2);

    return {sheetdata: coordsheetdata, sheetname: sheet1, col1num: rp.c1, ncols: rp.c2-rp.c1+1, row1num: rp.r1, nrows: rp.r2-rp.r1+1}

}

module.exports.Formula = Formula

/***/ }),
/* 7 */
/***/ (function(module, exports, __webpack_require__) {

/**
 工具栏绑定事件
 */

var ToolEventHandlerModule = __webpack_require__(16)
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
/* 8 */
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
	var ToolModule = __webpack_require__(15)
	var Tool = ToolModule.Tool
	var tool = new Tool()

	//实例化初始化UndoStack对象
	var UndoStackModule = __webpack_require__(18)
	var UndoStack = UndoStackModule.UndoStack
	var undoStack = new UndoStack()

	//实例化初始化Sheet对象
	var SheetModule = __webpack_require__(12)
	var Sheet = SheetModule.Sheet
	var sheet = new Sheet(undoStack)

	// //实例化初始化SliderBar对象
	// var SliderBarModule = require('component/SliderBarHandler')

	//实例化初始化Workspace对象
	var WorkspaceModule = __webpack_require__(20)
	var Workspace = WorkspaceModule.Workspace
	this.workspace = new Workspace(tool, sheet)

	//实例化初始化WSRender对象
	var WSRenderModule = __webpack_require__(19)
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
/* 9 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * Created by Ian on 17/6/5.
 */

var WSManagerModule = __webpack_require__(8)
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
/* 10 */
/***/ (function(module, exports, __webpack_require__) {

var config = __webpack_require__(0)

// var UtilModule=require('Util')
// var Util=UtilModule.Util
// var util=new Util()

var CellModule = __webpack_require__(5)
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
                    // console.log(newId)
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
/* 11 */
/***/ (function(module, exports, __webpack_require__) {

var Formula = __webpack_require__(6).Formula

var FuncListRender =  function () {
    this.formula = new Formula()
}

FuncListRender.prototype.init = function (ele) {
    var formula = this.formula
    var funcTable = document.createElement('table')
    funcTable.style.height = 200 + 'px'
    funcTable.style.width = 500 + 'px'
    funcTable.border = '1px bold #000'
    funcTable.style.backgroundColor = '#DDD'
    ele.appendChild(funcTable)

    var funcTr = document.createElement('tr')
    var funcTdLeft = document.createElement('td')
    var funcTdRight = document.createElement('td')
    var funcSpanLeft = document.createElement('span')
    var funcSpanRight = document.createElement('span')
    var funcTr2 = document.createElement('tr')
    var funcTd2 = document.createElement('td')
    var funcDiv = document.createElement('div')
    var funcDiv2 = document.createElement('div')
    var funcClassSelect = document.createElement('select')
    var funcNameSelect = document.createElement('select')
    funcNameSelect.id = 'funcNameSelect'

    var sumOfClass = 0
    for (name in formula.FunctionClassList) {
        var op = document.createElement('option')
        op.innerHTML = name
        op.id = name
        funcClassSelect.appendChild(op)
        if (sumOfClass === 0) op.selected = true
        sumOfClass++

    }

    funcClassSelect.size = sumOfClass
    funcNameSelect.size = sumOfClass
    var sumOfFname = 0
    for (fname in  formula.FunctionClassList[funcClassSelect.value]) {
        var op = document.createElement('option')
        op.innerHTML = fname
        op.id = fname
        funcNameSelect.appendChild(op)
        if (sumOfFname === 0) op.selected = true
        sumOfFname++

    }
    funcDiv.innerHTML = funcNameSelect.value + '('+ formula.funcInfo['s_farg_' + formula.funcParem[funcNameSelect.value]] + ')'
    funcDiv2.innerHTML = formula.funcInfo['s_fdef_' + funcNameSelect.value]

    funcSpanLeft.innerHTML = 'Category'
    funcTdRight.innerHTML = 'Functions'
    funcTd2.colSpan = 2
    funcClassSelect.style.width = 250 + 'px'
    funcNameSelect.style.width = 250 + 'px'
    funcTr.appendChild(funcTdLeft)
    funcTr.appendChild(funcTdRight)
    funcTdLeft.appendChild(funcSpanLeft)
    funcTdLeft.appendChild(document.createElement('tr'))
    funcTdLeft.appendChild(funcClassSelect)
    funcTdRight.appendChild(funcSpanRight)
    funcTdRight.appendChild(document.createElement('tr'))
    funcTdRight.appendChild(funcNameSelect)
    funcTr2.appendChild(funcTd2)
    funcTd2.appendChild(funcDiv)
    funcTd2.appendChild(document.createElement('tr'))
    funcTd2.appendChild(funcDiv2)

    funcTable.appendChild(funcTr)
    funcTable.appendChild(funcTr2)

    funcClassSelect.onchange = function () {
        funcNameSelect.innerHTML = ''
        funcNameSelect.size = 7
        var sumOfFname =0
        for (fname in  formula.FunctionClassList[this.value]) {
            var op = document.createElement('option')
            op.innerHTML = fname
            op.id = fname
            funcNameSelect.appendChild(op)
            if (sumOfFname === 0) {
                funcDiv.innerHTML = fname + '('+ formula.funcInfo['s_farg_' + formula.funcParem[fname]]+ ')'
                funcDiv2.innerHTML = formula.funcInfo['s_fdef_' + fname]
                op.selected = true
            }
        }
    }
    funcNameSelect.onchange = function () {
        funcDiv.innerHTML = funcNameSelect.value +'('+ formula.funcInfo['s_farg_' + formula.funcParem[funcNameSelect.value]] + ')'
        funcDiv2.innerHTML = formula.funcInfo['s_fdef_' + funcNameSelect.value]
    }
}

module.exports.FuncListRender = FuncListRender

/***/ }),
/* 12 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * 表格对象
 */


var configModule = __webpack_require__(0)
var CellModule = __webpack_require__(5)
var Cell = CellModule.Cell
var sheetConfig = configModule.SheetConfig
// var UtilModule = require('Util')
// var Util = UtilModule.Util
// var util = new Util()
var CellRenderModule = __webpack_require__(10)
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
    this.isCtrlDown = false

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

Sheet.prototype.setIdAttr = function (id,cmd,value) {
    this.cellRender.renderCell(id,cmd,value)
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
/* 13 */
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
        if(!sheet.isPreview)  sheetEventHandler.keyDown(event)

        if (!sheet.isMultiLineEditing && !sheet.isEditing) return false
    }
    document.onkeyup = function (event) {
        if(event.which === 17) sheet.isKeyDown = false
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
SheetEventBinder.prototype.initFormulaButton = function (button) {
    button.onclick = function () {
        sheetEventHandler.formulaButonClick()
    }
}
module.exports.SheetEventBinder = SheetEventBinder

/***/ }),
/* 14 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * Created by Ian on 17/7/22.
 */


var config = __webpack_require__(0)
var style = __webpack_require__(1)
var renderPane = __webpack_require__(2).renderPane
var renderBar = __webpack_require__(2).renderBar
var renderToggle = __webpack_require__(2).renderToggle

var SliderBarHandler = function (sheet) {
	this.sheet = sheet
}

SliderBarHandler.prototype.autoOpen = function () {
	var sheet = this.sheet
	var cell = sheet.cells[sheet.firstCellid]

	if ((sheet.firstCellid !== sheet.lastCellid) || !cell) {
		return
	}

	var sliderBarDiv = document.getElementById(config.SlideBarConfig.id)
	var sheetTableDiv = document.getElementById(config.SheetTableDivConfig.id)
	var toggleDiv = document.getElementById(config.SlideBarConfig.toggleId)

	toggleDiv.isOpen = true

	renderBar(sliderBarDiv, sheetTableDiv, toggleDiv)
	renderToggle(toggleDiv)
}

SliderBarHandler.prototype.addCellAttr = function () {
	var sheet = this.sheet
	var cell = sheet.cells[sheet.firstCellid]

	if ((sheet.firstCellid !== sheet.lastCellid) || !cell) {
		return
	}

	// var paneDiv = document.getElementById(config.SlideBarConfig.sliderPaneConfig.cellAttr.paneId)
	var titleDiv = document.getElementById(config.SlideBarConfig.sliderPaneConfig.cellAttr.titleId)
	var arrowDiv = document.getElementById(config.SlideBarConfig.sliderPaneConfig.cellAttr.arrowId)
	var contentDiv = document.getElementById(config.SlideBarConfig.sliderPaneConfig.cellAttr.contentId)

	titleDiv.isOpen = true

	renderPane(arrowDiv, contentDiv, titleDiv)

	removeAllChild(contentDiv)

	var cellAttrTable = document.createElement('table')
	cellAttrTable.style = style.SliderTableStyle
	contentDiv.appendChild(cellAttrTable)

	var i = 0

	for (var attr in cell) {

		if(cell[attr] == false){continue}

		var attrTr = document.createElement('tr')
		cellAttrTable.appendChild(attrTr)

		i++
		if(i%2 === 1){
			attrTr.style = style.SliderEvenTrStyle
		}else{
			attrTr.style = style.SliderOddTrStyle
		}

		var attrTd = document.createElement('td')
		attrTd.innerHTML = attr
		attrTd.style = style.SliderTdStyle
		var valueTd = document.createElement('td')
		valueTd.innerHTML = cell[attr]
		valueTd.style = style.SliderTdStyle
		attrTr.appendChild(attrTd)
		attrTr.appendChild(valueTd)
	}
}

function removeAllChild(node) {

	while (node.hasChildNodes()) {
		node.removeChild(node.firstChild)
	}
}
module.exports.SliderBarHandler = SliderBarHandler

/***/ }),
/* 15 */
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
/* 16 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * 工具栏事件
 */
var config = __webpack_require__(0)


var SheetRenderModule = __webpack_require__(4)
var SheetRender = SheetRenderModule.SheetRender
var sheetRender = null
// var CellRenderModule = require('CellRender')
// var CellRender = CellRenderModule.CellRender
// var cellRender = null

var SheetEventHandlerModule = __webpack_require__(3)
var SheetEventHandler = SheetEventHandlerModule.SheetEventHandler
var sheetEventHandler = null

var ToolEventHandler = function (sheet) {
    this.sheet = sheet
    sheetRender = new SheetRender(sheet)
    //cellRender = new CellRender(sheet)
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
                    if (isNaN(e[key]["content"]) && e[key]["content"] != '' && e[key]["content"].indexOf('=') == 0) {
                        var formula = {
                            'coord': key,
                            'value': e[key]["content"]
                        }
                        needEditCells.push(formula)
                        this.sheet.cells[key]["formula"] = e[key]["content"]
                    }
                    // if(e[key]["formula"]!==undefined && e[key]["formula"]!=''){
                    //     var formula = {
                    //         'coord': key,
                    //         'value': e[key]["formula"]
                    //     }
                    //     needEditCells.push(formula)
                    // }
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
                        ajax.open("POST", "http://123.56.22.114:8080/qianmo-service/changeContent", true)
                        // ajax.open("POST", "http://localhost:8088/qianmo-service/changeContent", true)
                        needEditCells = JSON.stringify(needEditCells)
                        ajax.send(needEditCells)
                        ajax.onreadystatechange = function () {
                            if (ajax.readyState == 4 && ajax.status == 200) {
                                var CellList = JSON.parse(ajax.responseText)
                                for (var key in CellList) {
                                    var value = CellList[key].value.substring(1)
                                    // if (value.indexOf("+") != -1) {
                                    //     var result = 0
                                    //     var values = value.split("+")
                                    //     for (var v in values) {
                                    //         result += parseInt(values[v])
                                    //     }
                                    //     value = result
                                    // }
                                    e[CellList[key].coord]["content"] = value
                                }
                                sheet.isPreview = true
                                document.getElementById('previewDiv').style.display = 'block'
                                document.getElementById('editDiv').style.display = 'none'
                                document.getElementById('menuDiv').isDisabled = true
                                sheetRender.renderSheet(sheet, document.getElementById('sheetTable'))

                                for (cell in sheet.cells) {
                                    for (attr in sheet.cells[cell]) {
                                        if (attr != 'coord') sheet.setIdAttr(sheet.cells[cell].coord, attr, sheet.cells[cell][attr])
                                        //console.log(attr)
                                    }
                                }
                            }
                        }
                    }
                } else {
                    sheet.isPreview = true
                    document.getElementById('previewDiv').style.display = 'block'
                    document.getElementById('editDiv').style.display = 'none'
                    document.getElementById('menuDiv').isDisabled = true
                    sheetRender.renderSheet(sheet, document.getElementById('sheetTable'))

                    for (cell in sheet.cells) {
                        for (attr in sheet.cells[cell]) {
                            if (attr != 'coord') sheet.setIdAttr(sheet.cells[cell].coord, attr, sheet.cells[cell][attr])
                        }
                    }
                }
            } else {
                alert("无内容，不允许预览！")
            }

            break
        case "Edit":
            sheet.isPreview = false
            document.getElementById('previewDiv').style.display = 'none'
            document.getElementById('editDiv').style.display = 'block'
            //document.getElementById('menuDiv').style.display = 'block'
            sheetRender.renderSheet(sheet, document.getElementById('sheetTable'))
            // console.log(sheet.cells)
            for (cell in sheet.cells) {
                for (attr in sheet.cells[cell]) {
                    if (attr != 'coord') {
                        // console.log(sheet.cells[cell]['formula'])
                        if (attr === 'content' && sheet.cells[cell]['formula'] !== undefined && sheet.cells[cell]['formula'] !== '') {
                            sheet.setIdAttr(sheet.cells[cell].coord, attr, sheet.cells[cell]['formula'])
                        } else {
                            sheet.setIdAttr(sheet.cells[cell].coord, attr, sheet.cells[cell][attr])
                        }
                    }
                }
            }
            break
        case "Down":
            var myMask = document.getElementById('mask')
            myMask.style.display = "block"
            var e = this.sheet.cells
            var a = {}
            var CellList = []
            for (var key in e) {
                if (e[key].show) {
                    var col = key.split('_')[0]
                    var row = key.split('_')[1]
                    col = String.fromCharCode(col.charCodeAt(0) + e[key].colSpan - 1)
                    row = parseInt(row) + e[key].rowSpan - 1
                    var area = key + '_' + col + '_' + row
                    var content = e[key]["content"] === undefined ? "" : e[key]["content"] === null ? "" : e[key]["content"] === "" ? "" : e[key]["content"] + ''
                    var addCell = {
                        "cellName": key,
                        "area": area,
                        "content": content.replace(/&nbsp;/g, ''),
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
                        "indentation": e[key]["indentation"] === undefined ? 0 : e[key]["indentation"],
                        "alignment": e[key]["alignment"] === undefined ? "" : e[key]["alignment"],
                        "bold": e[key]["bold"] === undefined ? "" : e[key]["bold"],
                        "italic": e[key]["italic"] === undefined ? "" : e[key]["italic"],
                        "vertical": e[key]["vertical"] === undefined ? "" : e[key]["vertical"],
                        "wrapText": e[key]["wrapText"] === undefined ? "" : e[key]["wrapText"],
                        "width": e[key]["width"] === undefined ? "" : e[key]["width"],
                        "height": e[key]["height"] === undefined ? "" : e[key]["height"]
                    }
                    // console.log(addCell.content)
                    CellList.push(addCell)
                }
            }
            var json = JSON.stringify(CellList)
            var ajax
            if (window.XMLHttpRequest) {
                ajax = new XMLHttpRequest()
            } else if (window.ActiveXObject) {
                ajax = new window.ActiveXObject()
            } else {
                alert("请升级至最新版本的浏览器")
            }
            if (ajax != null) {
                ajax.open("POST", "http://123.56.22.114:8080/qianmo-service/excelDownload", true)
                // ajax.open("POST","http://localhost:8088/qianmo-service/excelDownload",true)
                ajax.onload = function () {
                    if (ajax.status == 200) {
                        var filename = "";
                        var disposition = ajax.getResponseHeader('Content-Disposition');
                        if (disposition && disposition.indexOf('attachment') !== -1) {
                            var filenameRegex = /filename[^;=\n]*=((['"]).*?\2|[^;\n]*)/;
                            var matches = filenameRegex.exec(disposition);
                            if (matches != null && matches[1]) filename = matches[1].replace(/['"]/g, '');
                        }
                        var type = ajax.getResponseHeader('Content-Type');
                        var blob = new Blob([this.response], {type: type});
                        if (typeof window.navigator.msSaveBlob !== 'undefined') {
                            // IE workaround for "HTML7007: One or more blob URLs were revoked by closing the blob for which they were created. These URLs will no longer resolve as the data backing the URL has been freed."
                            window.navigator.msSaveBlob(blob, filename);
                        } else {
                            var URL = window.URL || window.webkitURL;
                            var downloadUrl = URL.createObjectURL(blob);
                            if (filename) {
                                // use HTML5 a[download] attribute to specify filename
                                var a = document.createElement("a");
                                // safari doesn't support this yet
                                if (typeof a.download === 'undefined') {
                                    window.location = downloadUrl;
                                } else {
                                    a.href = downloadUrl;
                                    a.download = filename;
                                    document.body.appendChild(a);
                                    a.click();
                                }
                            } else {
                                window.location = downloadUrl;
                            }

                            setTimeout(function () {
                                URL.revokeObjectURL(downloadUrl);
                            }, 100); // cleanup
                        }
                    }
                    myMask.style.display = "none"
                }
                ajax.responseType = 'blob'
                ajax.setRequestHeader('Content-type', 'application/json;charset=utf-8')
                ajax.send(json)
            }
            break
        case  "Init":
            var parentNode = document.getElementById("QianMoApp")
            var param = parentNode.getAttribute("url")
            var url = ''
            var method = "GET"
            if (!param && param != '') {
                url = 'json.json'
                param = null
            } else {
                method = "POST"
                url = 'http://123.56.22.114:8080/qianmo-service/getContentJson'
                // url='http://localhost:8088/qianmo-service/getContentJson'
                param = param.substring(param.lastIndexOf("\\") + 1)
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
                                // if(key==='content' && e['formula']!==undefined && e['formula']!==''){
                                //     value=e['formula']
                                // }
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
                // console.log(sheet.cells)
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
            // console.log(sheet.firstCellid)
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
            cells = sheet.getColAndRow(sheet.firstCellid.split('_')[0] + '_1', String.fromCharCode(sheet.maxCol) + '_' + sheet.maxRow)
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
                if (action === 'deleteCol') {
                    var firstCopiedCellCol = cells.firstCellCol + 1
                    var firstCopiedCellRow = cells.firstCellRow
                } else {
                    var firstCopiedCellCol = cells.firstCellCol - 1
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
            cells = sheet.getColAndRow('A_' + sheet.firstCellid.split('_')[1], String.fromCharCode(sheet.maxCol) + '_' + sheet.maxRow)
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
                if (action === 'deleteRow') {
                    var firstCopiedCellCol = cells.firstCellCol
                    var firstCopiedCellRow = cells.firstCellRow + 1
                } else {
                    var firstCopiedCellCol = cells.firstCellCol
                    var firstCopiedCellRow = cells.firstCellRow - 1
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
        case 'func':
            if (document.getElementById('funcListDiv').style.display === 'none') document.getElementById('funcListDiv').style.display = 'block'
            else document.getElementById('funcListDiv').style.display = 'none'
            break
        default:
            alert('该功能未实现，请期待')
            break
    }
    if (action !== ('multiLine' || 'copy' ) && !sheet.isPreview) {
        sheetEventHandler.mouseDown(sheet.firstCellid)
        sheetEventHandler.mouseUp(document.getElementById(sheet.lastCellid))
    }
}


module.exports.ToolEventHandler = ToolEventHandler

/***/ }),
/* 17 */
/***/ (function(module, exports, __webpack_require__) {

/**
 * 工具栏渲染
 */
var config = __webpack_require__(0)
var style = __webpack_require__(1)
var ToolRender = function (sheet) {
    this.sheet = sheet
}
ToolRender.prototype.init = function (toolDiv, width, height) {
    var ToolEventBinderModule = __webpack_require__(7)
    var ToolEventBinder = ToolEventBinderModule.ToolEventBinder
    // var toolEventBinder = new ToolEventBinder(this.sheet)

    toolDiv.style.width = width + 'px'
    toolDiv.style.height = height + 'px'

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
    menuDiv.id = 'menuDiv'
    menuDiv.style = style.MeunDivStyle

    changeSelected(boxes, 'styleBox')
    renderTool('styleBox', buttonBox, sheet)

    styleBox.onclick = function () {
        if (!sheet.isPreview) {
            changeSelected(boxes, 'styleBox')
            renderTool('styleBox', buttonBox, sheet)
        }
    }
    toolBox.onclick = function () {
        if (!sheet.isPreview) {
            changeSelected(boxes, 'toolBox')
            renderTool('toolBox', buttonBox, sheet)
        }
    }
    dataBox.onclick = function () {
        if (!sheet.isPreview) {
            changeSelected(boxes, 'dataBox')
            renderTool('dataBox', buttonBox, sheet)
        }
    }

    for (var box in boxes) {
        menuDiv.appendChild(boxes[box].node)
    }

}

function renderTool(menu, buttonBox, sheet) {

    removeAllChild(buttonBox)

    var ToolEventBinderModule = __webpack_require__(7)
    var ToolEventBinder = ToolEventBinderModule.ToolEventBinder
    var toolEventBinder = new ToolEventBinder(sheet)

    var cmdMap = config.ToolConfig.cmdCodeMap
    var toolHtml = config.ToolConfig.toolHtml
    var styleHtml = config.ToolConfig.styleHtml
    var dataHtml = config.ToolConfig.dataHtml
    var defaultHtml = config.ToolConfig.defaultHtml

    var previewHtml = config.ToolConfig.previewHtml
    var previewDiv = document.createElement('div')
    previewHtml.forEach(function (innerhtml) {

        var buttonDiv = document.createElement('div')
        buttonDiv.id = innerhtml
        buttonDiv.style = style.ButtonDivStyle
        buttonDiv.innerHTML = innerhtml

        buttonDiv.onclick = function () {
            toolEventBinder.buttonClick(innerhtml)
        }
        previewDiv.appendChild(buttonDiv)
    })
    previewDiv.style.display = 'none'
    previewDiv.id = 'previewDiv'
    buttonBox.appendChild(previewDiv)

    var editDiv = document.createElement('div')
    defaultHtml.forEach(function (innerhtml) {
        console.log(innerhtml)
        var buttonDiv = document.createElement('div')
        buttonDiv.id = innerhtml
        buttonDiv.style = style.ButtonDivStyle
        buttonDiv.innerHTML = cmdMap[innerhtml]
        // if(innerhtml==='Init'){
        //     buttonDiv.style.display='none'
        // }
        buttonDiv.onclick = function () {
            toolEventBinder.buttonClick(innerhtml)
        }
        editDiv.appendChild(buttonDiv)
    })

    if (menu === 'styleBox') {

        for (var i = 0; i < styleHtml.length; i++) {

            if (styleHtml[i] === 'font') {
                renderFont(toolEventBinder, editDiv)
                continue
            }

            if (styleHtml[i] === 'align') {
                renderAlign(toolEventBinder, editDiv)
                continue
            }

            if (styleHtml[i] === 'border') {
                renderBorder(toolEventBinder, editDiv)
                continue
            }

            var styleButtonDiv = document.createElement('div')
            styleButtonDiv.id = styleHtml[i]
            styleButtonDiv.style = style.ButtonDivStyle
            styleButtonDiv.innerHTML = cmdMap[styleHtml[i]]

            editDiv.appendChild(styleButtonDiv)

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
                renderColorSelect(toolEventBinder, editDiv)
                continue
            }
            if (styleHtml[i] === 'fillColor') {
                renderBackgroundColor(toolEventBinder, styleButtonDiv)
                renderColorSelect(toolEventBinder, editDiv)
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

            editDiv.appendChild(toolButtonDiv)

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
            editDiv.appendChild(dataButtonDiv)
        }
    }
    editDiv.id = 'editDiv'
    buttonBox.appendChild(editDiv)

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
/* 18 */
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
/* 19 */
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
	WSDiv.style = __webpack_require__(1).WSStyle
	this.parNode.appendChild(WSDiv)

	//tool
	var ToolRenderModule = __webpack_require__(17)
	var ToolRender = ToolRenderModule.ToolRender
	var toolRender = new ToolRender(ws.Sheet)
	var tool = ws.Tool
	var toolDiv = document.createElement('div')
	toolRender.init(toolDiv, tool.width, tool.height)
	WSDiv.appendChild(toolDiv)

	//sheet
	var SheetRenderModule = __webpack_require__(4)
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
/* 20 */
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