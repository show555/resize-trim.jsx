/*===================================================================================================
	File Name: リサイズ＆トリム.jsx
	Title: リサイズ＆トリム
	Version: 1.5.0
	Author: show555
	Description: 選択したフォルダ内の画像を指定したサイズいっぱいにリサイズしトリミングする
	Includes: Underscore.js,
	          Underscore.string.js
===================================================================================================*/

#target photoshop

// Photoshopの設定単位を保存
var originalRulerUnits = app.preferences.rulerUnits;
// Photoshopの設定単位をピクセルに変更
app.preferences.rulerUnits = Units.PIXELS;
// Photoshopの不要なダイアログを表示させない
app.displayDialogs = DialogModes.NO;

// 初期設定
var settings = {
	folderPath: '',          // 対象フォルダのパスの初期値
	_fileTypes: {
		init: [ 'JPG' ],       // 対象ファイルタイプのデフォルトチェック JPG／GIF／PNG／EPS／TIFF／BMP
		regex: {
			JPG:  '\\.jpe?g',
			GIF:  '\\.gif',
			PNG:  '\\.png',
			EPS:  '\\.eps',
			TIFF: '\\.tiff?',
			BMP:  '\\.bmp',
			PDF:  '\\.pdf',
			PSD:  '\\.psd',
			AI:   '\\.ai'
		}
	},
	colorMode: 'RGB',        // カラーモードの初期値
	_quality: {
		jpgWeb: {
			init: 100,            // 保存画質（Web用JPG）の初期値
			min:  0,
			max:  100
		},
		jpgDtp: {
			init: 10,            // 保存画質（DTP用JPG）の初期値
			min:  0,
			max:  12
		},
	},
	trim: {
		init: '縦・横固定',
		type: {
			fix:    '縦・横固定',
			square: '正方形',
			flex:   '長辺・短辺'
		},
		width:  1000,           // 幅（長辺）のサイズの初期値
		height: 667,            // 高さ（短辺）のサイズの初期値
	},
	save: {
		init: 'JPG（WEB用）',  // 保存形式の初期値
		type: {
			jpgWeb: { label: 'JPG（WEB用）', extension: 'jpg' },
			jpgDtp: { label: 'JPG', extension: 'jpg' },
			eps:    { label: 'EPS', extension: 'eps' },
			png:    { label: 'PNG', extension: 'png' }
		},
		dir: 'thumb',          // 保存先のディレクトリ名
		overwrite: false,      // 同じ階層に保存する
		suffix: ''             // リサイズした画像の接尾辞
	},
	fileTypes: [],
	recursive: false,        // 指定フォルダを再帰的に処理するか
	trimType: '',
	saveType: '',
	quality: '',
	folderObj: {}
};

// ----------------------------------▼ Underscore.js ▼----------------------------------
#include "underscore.inc"
#include "underscore.string.inc"
_.mixin(_.string.exports());
// ----------------------------------▲ Underscore.js ▲----------------------------------

// 保存関数
var saveFunctions = {
	jpgWeb: function( theDoc, newFile, settings ) {
		var jpegOpt       = new ExportOptionsSaveForWeb();
		jpegOpt.format    = SaveDocumentType.JPEG;
		jpegOpt.optimized = true;
		jpegOpt.quality   = settings.quality;
		theDoc.exportDocument( newFile, ExportType.SAVEFORWEB, jpegOpt );
	},
	jpgDtp: function( theDoc, newFile, settings ) {
		var jpegOpt     = new JPEGSaveOptions();
		jpegOpt.quality = settings.quality;
		theDoc.saveAs( newFile, jpegOpt, true, Extension.LOWERCASE );
	},
	eps: function( theDoc, newFile, settings ) {
		var epsOpt               = new EPSSaveOptions();
		epsOpt.embedColorProfile = true;
		epsOpt.encoding          = SaveEncoding.JPEGMAXIMUM;
		epsOpt.halftoneScreen    = false;
		epsOpt.interpolation     = false;
		epsOpt.preview           = Preview.MACOSEIGHTBIT;
		epsOpt.psColorManagement = false;
		epsOpt.transferFunction  = false;
		epsOpt.transparentWhites = false;
		epsOpt.vectorData        = false;
		theDoc.saveAs( newFile, epsOpt, true, Extension.LOWERCASE );
	},
	png: function( theDoc, newFile, settings ) {
		var pngOpt        = new PNGSaveOptions();
		pngOpt.interlaced = false;
		theDoc.saveAs( newFile, pngOpt, true, Extension.LOWERCASE );
	}
};

// 実行フラグ
var do_flag = true;

// ---------------------------------- ダイアログ作成 ----------------------------------
// ダイアログオブジェクト
var uDlg = new Window( 'dialog', 'リサイズ＆トリム', { x:0, y:0, width:400, height:615 } );

// ダイアログを画面に対して中央揃えに
uDlg.center();

// パネル 対象フォルダ
uDlg.folderPnl           = uDlg.add( "panel",    { x:10,  y:10, width:380, height:80 }, "対象フォルダ" );
uDlg.folderPnl.path      = uDlg.add( "edittext", { x:25,  y:30, width:270, height:25 }, settings.folderPath );
uDlg.folderPnl.selectBtn = uDlg.add( "button",   { x:300, y:30, width:75,  height:25 }, "選択" );
uDlg.folderPnl.recursive = uDlg.add( "checkbox", { x:25,  y:60, width:350, height:20 }, "サブフォルダも含める" );
// “サブフォルダも含める”チェックボックスの初期値を設定
uDlg.folderPnl.recursive.value = settings.recursive;

// 対象フォルダ選択ボタンが押された時の処理
uDlg.folderPnl.selectBtn.onClick = function() {
	var oldPath = uDlg.folderPnl.path.text;
	uDlg.folderPnl.path.text = Folder.selectDialog( 'フォルダを選択してください' ) || oldPath;
}

// パネル 対象ファイルタイプ
uDlg.fileTypePnl     = uDlg.add( "panel", { x:10, y:100, width:380, height:80 }, "対象ファイルタイプ" );
uDlg.fileTypePnl.ext = [];
uDlg.fileTypePnl.ext.push( uDlg.add( "checkbox", { x:25,  y:120, width:50, height:25 }, "JPG" ) );
uDlg.fileTypePnl.ext.push( uDlg.add( "checkbox", { x:85,  y:120, width:50, height:25 }, "GIF" ) );
uDlg.fileTypePnl.ext.push( uDlg.add( "checkbox", { x:145, y:120, width:50, height:25 }, "PNG" ) );
uDlg.fileTypePnl.ext.push( uDlg.add( "checkbox", { x:205, y:120, width:50, height:25 }, "EPS" ) );
uDlg.fileTypePnl.ext.push( uDlg.add( "checkbox", { x:265, y:120, width:50, height:25 }, "TIFF" ) );
uDlg.fileTypePnl.ext.push( uDlg.add( "checkbox", { x:325, y:120, width:50, height:25 }, "BMP" ) );
uDlg.fileTypePnl.ext.push( uDlg.add( "checkbox", { x:25, y:145, width:50, height:25 }, "PDF" ) );
uDlg.fileTypePnl.ext.push( uDlg.add( "checkbox", { x:85, y:145, width:50, height:25 }, "PSD" ) );
uDlg.fileTypePnl.ext.push( uDlg.add( "checkbox", { x:145, y:145, width:50, height:25 }, "AI" ) );
_.each( uDlg.fileTypePnl.ext, function( item, key ) {
	if ( _.contains( settings._fileTypes.init, item.text ) ) {
		item.value = true;
	}
} );

// パネル トリム設定
var trimTypeList = _.values( settings.trim.type );
uDlg.trimPnl                 = uDlg.add( "panel",        { x:10,  y:190, width:380, height:120 }, "リサイズ＆トリム設定" );
uDlg.trimPnl.trimTypeText    = uDlg.add( "statictext",   { x:25,  y:215, width:85,  height:20  }, "トリムタイプ:" );
uDlg.trimPnl.trimType        = uDlg.add( "dropdownlist", { x:120, y:213, width:100, height:22  }, trimTypeList );
uDlg.trimPnl.widthLongText   = uDlg.add( "statictext",   { x:60,  y:245, width:50,  height:20  }, "幅:" );
uDlg.trimPnl.widthLong       = uDlg.add( "edittext",     { x:120, y:243, width:50,  height:22  }, settings.trim.width );
uDlg.trimPnl.widthLongUnit   = uDlg.add( "statictext",   { x:175, y:245, width:20,  height:20  }, "px" );
uDlg.trimPnl.heightShortText = uDlg.add( "statictext",   { x:60,  y:275, width:50,  height:20  }, "高さ:" );
uDlg.trimPnl.heightShort     = uDlg.add( "edittext",     { x:120, y:273, width:50,  height:22  }, settings.trim.height );
uDlg.trimPnl.heightShortUnit = uDlg.add( "statictext",   { x:175, y:275, width:20,  height:20  }, "px" );
uDlg.trimPnl.trimTypeText.justify = uDlg.trimPnl.widthLongText.justify = uDlg.trimPnl.heightShortText.justify = 'right';
// 幅／長辺の長さが変更された時の処理
uDlg.trimPnl.widthLong.onChange = function() {
	if ( uDlg.trimPnl.trimType.selection.text == '正方形' ) {
		uDlg.trimPnl.heightShort.text = uDlg.trimPnl.widthLong.text;
	}
}
// 保存形式の初期値を設定
uDlg.trimPnl.trimType.selection = _.indexOf( trimTypeList, settings.trim.init );
setTrimTypeSizes( settings.trim.init );
// 保存形式が変更された時の処理
uDlg.trimPnl.trimType.onChange = function() {
	setTrimTypeSizes( uDlg.trimPnl.trimType.selection.text );
}

// パネル 書き出し設定
var saveTypeList = _.pluck( settings.save.type, 'label' );
uDlg.resizePnl               = uDlg.add( "panel",        { x:10,  y:320, width:380, height:175 }, "書き出し設定" );
uDlg.resizePnl.colorModeText = uDlg.add( "statictext",   { x:25,  y:345, width:60,  height:20  }, "モード:" );
uDlg.resizePnl.saveTypeText  = uDlg.add( "statictext",   { x:25,  y:375, width:60,  height:20  }, "保存形式:" );
uDlg.resizePnl.saveType      = uDlg.add( "dropdownlist", { x:95,  y:373, width:110, height:22  }, saveTypeList );
uDlg.resizePnl.qualityText   = uDlg.add( "statictext",   { x:25,  y:405, width:60,  height:20  }, "画質:" );
uDlg.resizePnl.quality       = uDlg.add( "edittext",     { x:95,  y:403, width:50,  height:22  }, settings._quality.jpgDtp.init );
uDlg.resizePnl.qualityRange  = uDlg.add( "statictext",   { x:150, y:405, width:60,  height:20  }, "(0〜" + settings._quality.jpgDtp.max + ")" );
uDlg.resizePnl.qualitySlider = uDlg.add( "slider",       { x:205, y:400, width:170, height:20  }, settings._quality.jpgDtp.init, settings._quality.jpgDtp.min, settings._quality.jpgDtp.max );
uDlg.resizePnl.overwrite     = uDlg.add( "checkbox",     { x:25,  y:430, width:350, height:20  }, "リサイズ後同じ階層に保存する（同形式の場合上書き）" );
uDlg.resizePnl.saveDirText   = uDlg.add( "statictext",   { x:25,  y:457, width:60,  height:20  }, "保存先:" );
uDlg.resizePnl.saveDir       = uDlg.add( "edittext",     { x:95,  y:455, width:110, height:22  }, settings.save.dir );
uDlg.resizePnl.colorModeText.justify = uDlg.resizePnl.saveTypeText.justify = uDlg.resizePnl.qualityText.justify = uDlg.resizePnl.saveDirText.justify = 'right';
// カラーモード選択ラジオボタンの追加
uDlg.resizePnl.colorMode      = uDlg.add( "group", { x:95, y:345, width:245, height:20 } );
uDlg.resizePnl.colorMode.RGB  = uDlg.resizePnl.colorMode.add( "radiobutton",  { x:0,  y:0, width:50, height:20 }, "RGB" );
uDlg.resizePnl.colorMode.CMYK = uDlg.resizePnl.colorMode.add( "radiobutton",  { x:55, y:0, width:70, height:20 }, "CMYK" );
// カラーモードの初期値を設定
uDlg.resizePnl.colorMode[settings.colorMode].value = true;
// “対象ファイルに上書き保存”チェックボックスの初期値を設定
uDlg.resizePnl.overwrite.value = settings.save.overwrite;

// 保存形式の初期値を設定
uDlg.resizePnl.saveType.selection = _.indexOf( saveTypeList, settings.save.init );
setSaveTypeQuality( settings.save.init );
// 保存形式が変更された時の処理
uDlg.resizePnl.saveType.onChange = function() {
	setSaveTypeQuality( uDlg.resizePnl.saveType.selection.text );
}
// 画質のスライダーを動かしている時の処理
uDlg.resizePnl.qualitySlider.onChanging = function() {
	uDlg.resizePnl.quality.text = parseInt( uDlg.resizePnl.qualitySlider.value );
}
// 画質を入力した時の処理
uDlg.resizePnl.quality.onChange = function() {
	uDlg.resizePnl.qualitySlider.value = parseInt( uDlg.resizePnl.quality.text );
}
// “対象ファイルに上書き保存”をチェックした時の処理
uDlg.resizePnl.overwrite.onClick = function() {
	uDlg.resizePnl.saveDirText.enabled = !uDlg.resizePnl.overwrite.value;
	uDlg.resizePnl.saveDir.enabled     = !uDlg.resizePnl.overwrite.value;
}

// パネル 追加アクション
uDlg.actionPnl            = uDlg.add( "panel",        { x:10, y:505, width:380, height:60 }, "追加アクション" );
uDlg.actionPnl.actionList = uDlg.add( "dropdownlist", { x:25, y:527, width:350, height:22 }, {} );
actionList = getActionSets();
uDlg.actionPnl.actionList.add( 'item', 'なし' );
for (i = 0; i < actionList.length; i++) {
	uDlg.actionPnl.actionList.add( 'item', actionList[i] );
}
uDlg.actionPnl.actionList.selection = uDlg.actionPnl.actionList.items[0];

// キャンセルボタン
uDlg.cancelBtn = uDlg.add( "button", { x:95, y:575, width:100, height:25 }, "キャンセル", { name: "cancel" } );
// キャンセルボタンが押されたらキャンセル処理（ESCキー含む）
uDlg.cancelBtn.onClick = function() {
	// 実行フラグにfalseを代入
	do_flag = false;
	// ダイアログを閉じる
	uDlg.close();
}

// OKボタン
uDlg.okBtn = uDlg.add( "button", { x:205, y:575, width:100, height:25 }, "リサイズ実行", { name: "ok" } );
// OKボタンが押されたら各設定項目に不備がないかチェック
uDlg.okBtn.onClick = function() {
	// 各種項目の値を格納
	settings.folderPath     = uDlg.folderPnl.path.text;
	settings.folderObj      = new Folder( settings.folderPath );
	settings.recursive      = uDlg.folderPnl.recursive.value;
	settings.colorMode      = uDlg.resizePnl.colorMode.RGB.value ? 'RGB' : 'CMYK';
	settings.trimType       = getTrimTypeKey( uDlg.trimPnl.trimType.selection.text );
	settings.saveType       = getSaveTypeKey( uDlg.resizePnl.saveType.selection.text );
	settings.save.dir       = uDlg.resizePnl.saveDir.text;
	settings.save.overwrite = uDlg.resizePnl.overwrite.value;
	settings.trim.width     = parseInt( uDlg.trimPnl.widthLong.text );
	settings.trim.height    = parseInt( uDlg.trimPnl.heightShort.text );
	settings.fileTypes      = [];
	_.each( uDlg.fileTypePnl.ext, function( item ) {
		if ( item.value ) {
			settings.fileTypes.push( item.text );
		}
	} );

	// 対象フォルダが選択されているかチェック
	if ( !settings.folderPath ) {
		alert( '対象フォルダが選択されていません' );
		return false;
	}
	// 対象フォルダが存在するかチェック
	if ( !settings.folderObj.exists ) {
		alert( '対象フォルダが存在しません' );
		return false;
	}
	// 拡張子が最低1つは選択されているかチェック
	if ( settings.fileTypes.length < 1 ) {
		alert( '対象ファイルタイプが選択されていません' );
		return false;
	}
	// カラーモードがCMYKの時保存形式がPNGになっていないかチェック
	if ( settings.colorMode == 'CMYK' ) {
		if ( settings.saveType == 'jpgWeb' ) {
			alert( 'JPG（WEB用）形式で保存するためにはカラーモードはRGBでなければいけません' );
			return false;
		}
		if ( settings.saveType == 'png' ) {
			alert( 'PNG形式で保存するためにはカラーモードはRGBでなければいけません' );
			return false;
		}
	}
	// 幅／長辺の長さが入力されているかチェック
	if ( _.isNaN( settings.trim.width ) ) {
		alert( '幅／長辺の長さを整数で入力して下さい' );
		return false;
	}
	// 幅／長辺の長さが0より大きいかチェック
	if ( settings.trim.width < 1 ) {
		alert( '0より大きい幅／長辺の長さを入力しください' );
		return false;
	}
	if ( settings.trimType != 'square' ) {
		// 高さ／短辺の長さが入力されているかチェック
		if ( _.isNaN( settings.trim.height ) ) {
			alert( '高さ／短辺の長さを整数で入力して下さい' );
			return false;
		}
		// 高さ／短辺の長さが0より大きいかチェック
		if ( settings.trim.height < 1 ) {
			alert( '0より大きい高さ／短辺の長さを入力しください' );
			return false;
		}
		// 幅／長辺と高さ／短辺の長さが同じでないかチェック
		if ( settings.trim.height == settings.trim.width ) {
			alert( '同じ長さを設定する時はトリムモードを正方形にしてください' );
			return false;
		}
	}
	// トリムモードが長辺・短辺の時短辺の方が長くなっていないかチェック
	if ( settings.trimType == 'flex' ) {
		if ( settings.trim.height > settings.trim.width ) {
			alert( '長辺より短辺の方が長くなっています' );
			return false;
		}
	}
	// 保存形式がJPG／JPG（WEB用）の場合
	if ( _.contains( [ 'jpgWeb', 'jpgDtp' ], settings.saveType ) ) {
		// 画質が入力されているかチェック
		settings.quality = parseInt( uDlg.resizePnl.quality.text );
		if ( _.isNaN( settings.quality ) ) {
			alert( '画質を整数で入力して下さい' );
			return false;
		}
	}
	// 不備がなかった場合処理続行
	uDlg.close();
}

// ダイアログ表示
uDlg.show();

// ---------------------------------- メインリサイズ処理 ----------------------------------
if ( do_flag ) {
	// alert( 'フォルダ:' + uDlg.folderPnl.path.text + "\n" + '拡張子：' + settings.fileTypes.join( ', ' ) + "\n" + 'トリムタイプ：' + settings.trimType + "\n" + '幅／長辺：' + settings.trim.width + "\n" + '高さ／短辺：' + settings.trim.height + "\n" + 'モード：' + settings.colorMode + "\n" + '保存形式：' + settings.saveType + "\n" + '画質：' + settings.quality + "\n" + '保存先：' + settings.save.dir + "\n" + '上書きする' + settings.save.overwrite );

	// 複数の対象ファイルを取得するための正規表現オブジェクトを作成
	var extensions = [];
	_.each( settings.fileTypes, function( fileType ) {
		extensions.push( settings._fileTypes.regex[fileType] );
	} );
	var fileReg = new RegExp( '(' + extensions.join( '|' ) + ')$', 'i' );
	var files   = _getFileList( settings.folderObj );

	// 進捗バーを表示
	var ProgressPanel = CreateProgressPanel( files.length, 500, '処理中…', true );
	ProgressPanel.show();
	var i = 1;

	// 対象ファイルに対してリサイズ→保存のループ処理
	_.each( files, function( file ) {
		// キャンセルの場合処理中止
		if ( !do_flag ) return;
		// 進捗バーを更新
		ProgressPanel.val( i );

		// ファイルオープン
		if ( /\.pdf$|\.ai$/i.test( file ) ) {
			// PDFの場合
			var pdfObj = new File( file );
			var theDoc = app.open( pdfObj, pdfOpenOptions );
			var path   = pdfObj.path;
		} else {
			// PDF以外
			var theDoc = app.open( file );
			var path   = theDoc.path;
		}

		// カラーモードをRGBに変更
		theDoc.changeMode( ChangeMode[settings.colorMode] );
		//リサイズする
		var imageWidth   = theDoc.width.value,
		    imageHeight  = theDoc.height.value,
		    userWidth    = settings.trim.width;
		    userHeight   = settings.trim.height;
		    resizeWidth  = null,
		    resizeHeight = null,
		    cropWidth    = null,
		    cropHeight   = null;

		switch ( settings.trimType ) {
			case 'fix':
				var imageAspectRatio = imageWidth / imageHeight;
				var userAspectRatio  = userWidth / userHeight;
				if ( imageAspectRatio > userAspectRatio ) {
					resizeWidth  = imageWidth * userHeight / imageHeight;
					resizeHeight = userHeight;
				} else {
					resizeWidth  = userWidth;
					resizeHeight = imageHeight * userWidth / imageWidth;
				}
				cropWidth  = userWidth;
				cropHeight = userHeight;
				break;
			case 'square':
				if ( imageWidth > imageHeight ) {
					resizeWidth  = imageWidth * userHeight / imageHeight;
					resizeHeight = userWidth;
				} else {
					resizeWidth  = userWidth;
					resizeHeight = imageHeight * userWidth / imageWidth;
				}
				cropWidth  = userWidth;
				cropHeight = userHeight;
				break;
			case 'flex':
				var imageAspectRatio = imageWidth / imageHeight;
				var userAspectRatio  = ( ( imageWidth >= imageHeight ) ? userWidth : userHeight ) / ( ( imageHeight > imageWidth ) ? userWidth : userHeight );
				if ( imageWidth >= imageHeight ) {
					if ( imageAspectRatio > userAspectRatio ) {
						resizeWidth  = imageWidth * userHeight / imageHeight;
						resizeHeight = userHeight;
					} else {
						resizeWidth  = userWidth;
						resizeHeight = imageHeight * userWidth / imageWidth;
					}
				} else {
					if ( imageAspectRatio > userAspectRatio ) {
						resizeWidth  = imageWidth * userWidth / imageHeight;
						resizeHeight = userWidth;
					} else {
						resizeWidth  = userHeight;
						resizeHeight = imageHeight * userHeight / imageWidth;
					}
				}
				cropWidth  = ( imageWidth >= imageHeight ) ? userWidth : userHeight;
				cropHeight = ( imageHeight > imageWidth ) ? userWidth : userHeight;
				break;
		}
		theDoc.resizeImage( resizeWidth, resizeHeight, 72, ResampleMethod.BICUBICSMOOTHER );
		// 入力されたサイズでトリミング
		theDoc.resizeCanvas( cropWidth, cropHeight, AnchorPosition.MIDDLECENTER );
		// 追加アクション実行
		var isDoAction = (uDlg.actionPnl.actionList.selection.toString() !== 'なし') ? true : false;
		if ( isDoAction ) {
			selectAction = uDlg.actionPnl.actionList.selection.toString().split("::->>");
			app.doAction( selectAction[1], selectAction[0] );
		}
		// 保存先フォルダを作成
		var saveDir = !settings.save.overwrite ? new Folder( theDoc.path + '/' + settings.save.dir ) : new Folder( theDoc.path );
		if( !saveDir.exists ){
			saveDir.create();
		}
		// 保存用の新規オブジェクト作成
		var saveDirPath = !settings.save.overwrite ? settings.save.dir + '/' : '';
		var newFile     = new File( theDoc.path + '/' + saveDirPath + theDoc.name.replace( /\.\w+$/i, '' ) + settings.save.suffix + '.' + settings.save.type[settings.saveType].extension );
		// 保存形式ごとの関数を呼び出し
		saveFunctions[settings.saveType]( theDoc, newFile, settings );
		// ファイルクローズ
		theDoc.close( SaveOptions.DONOTSAVECHANGES );

		i++;
	} );
}
if ( ProgressPanel ) {
	ProgressPanel.close();
}

// Photoshopの設定単位を復元
app.preferences.rulerUnits = originalRulerUnits;

// ------------------------------------------ 関数 -----------------------------------------
function setSaveTypeQuality( saveType ) {
	var key = getSaveTypeKey( saveType );
	if ( _.contains( [ 'jpgWeb', 'jpgDtp' ], key ) ) {
		uDlg.resizePnl.quality.text           = settings._quality[key].init;
		uDlg.resizePnl.qualityRange.text      = "(" + settings._quality[key].min + "〜" + settings._quality[key].max + ")";
		uDlg.resizePnl.qualitySlider.minvalue = settings._quality[key].min;
		uDlg.resizePnl.qualitySlider.maxvalue = settings._quality[key].max;
		uDlg.resizePnl.qualitySlider.value    = settings._quality[key].init;
		uDlg.resizePnl.qualityText.enabled    = true;
		uDlg.resizePnl.qualityRange.enabled   = true;
		uDlg.resizePnl.quality.enabled        = true;
		uDlg.resizePnl.qualitySlider.enabled  = true;
	} else {
		uDlg.resizePnl.qualityText.enabled    = false;
		uDlg.resizePnl.qualityRange.enabled   = false;
		uDlg.resizePnl.quality.enabled        = false;
		uDlg.resizePnl.qualitySlider.enabled  = false;
	}
}

function getSaveTypeKey( saveType ) {
	var saveTypeKey;
	_.some( settings.save.type, function( value, key ) {
		if ( value.label == saveType ) {
			saveTypeKey = key;
			return true;
		}
		return false;
	} );
	return saveTypeKey;
}

function setTrimTypeSizes( trimType ) {
	var key = getTrimTypeKey( trimType );
	switch ( key ) {
		case 'fix':
			uDlg.trimPnl.widthLongText.text      = '幅:';
			uDlg.trimPnl.heightShortText.text    = '高さ:';
			uDlg.trimPnl.heightShort.enabled     = true;
			uDlg.trimPnl.heightShortText.enabled = true;
			uDlg.trimPnl.heightShortUnit.enabled = true;
			break;
		case 'square':
			uDlg.trimPnl.widthLongText.text      = '幅:';
			uDlg.trimPnl.heightShortText.text    = '高さ:';
			uDlg.trimPnl.heightShort.text        = uDlg.trimPnl.widthLong.text;
			uDlg.trimPnl.heightShort.enabled     = false;
			uDlg.trimPnl.heightShortText.enabled = false;
			uDlg.trimPnl.heightShortUnit.enabled = false;
			break;
		case 'flex':
			uDlg.trimPnl.widthLongText.text      = '長辺:';
			uDlg.trimPnl.heightShortText.text    = '短辺:';
			uDlg.trimPnl.heightShort.enabled     = true;
			uDlg.trimPnl.heightShortText.enabled = true;
			uDlg.trimPnl.heightShortUnit.enabled = true;
			break;
	}
}

function getTrimTypeKey( trimType ) {
	return _.invert( settings.trim.type )[trimType];
}

function CreateProgressPanel( myMaximumValue, myProgressBarWidth , progresTitle, useCancel ) {
	var progresTitle = typeof progresTitle == 'string' ? progresTitle : 'Processing...';
	myProgressPanel = new Window( 'palette', _.sprintf( "%s(%d/%d)", progresTitle, 1, myMaximumValue ) );
	myProgressPanel.myProgressBar = myProgressPanel.add( 'progressbar', [ 12, 12, myProgressBarWidth, 24 ], 0, myMaximumValue );
	if ( useCancel ) {
		myProgressPanel.cancel = myProgressPanel.add( 'button', undefined, 'キャンセル' );
		myProgressPanel.cancel.onClick = function() {
			try {
				do_flag = false;
				myProgressPanel.close();
			} catch(e) {
				alert(e);
			}
		}
	}
	var PP = {
		'ProgressPanel': myProgressPanel,
		'title': progresTitle,
		'show': function() { this.ProgressPanel.show() },
		'close': function() { this.ProgressPanel.close() },
		'max': myMaximumValue,
		'barwidth': myProgressBarWidth,
		'val': function( val ) {
			this.ProgressPanel.myProgressBar.value = val;
			if ( val < this.max ) {
				this.ProgressPanel.text = _.sprintf( "%s(%d/%d)", this.title, val+1, this.max );
			}
			this.ProgressPanel.update();
		}
	}
	return PP;
}

function _getFileList( path ) {
	var rv    = [],
	    files = path.getFiles();

	_.each( files, function( file ) {
		if ( file.alias ) { // alias
			return;
		}
		if ( file.constructor.name === 'File' ) { // file
			if ( fileReg.test( file.name ) ) {
				rv.push( file );
			}
		} else { // folder
			if ( settings.recursive && (file.name !== settings.save.dir) ) {
				rv.push.apply( rv, _getFileList( file ) );
			}
		}
	} );
	return rv;
}

function getActionSets() {
	cTID = function(s) {
		return app.charIDToTypeID(s);
	};
	sTID = function(s) {
		return app.stringIDToTypeID(s);
	};
	var i = 1;
	var sets = [];
	while (true) {
		var ref = new ActionReference();
		ref.putIndex(cTID("ASet"), i);
		var desc;
		var lvl = $.level;
		$.level = 0;
		try {
			desc = executeActionGet(ref);
		} catch (e) {
			break;
		} finally {
			$.level = lvl;
		}
		if (desc.hasKey(cTID("Nm  "))) {
			var set = {};
			set.index = i;
			set.name = desc.getString(cTID("Nm  "));
			set.toString = function() {
				return this.name;
			};
			set.count = desc.getInteger(cTID("NmbC"));
			set.actions = [];
			for (var j = 1; j <= set.count; j++) {
				var ref = new ActionReference();
				ref.putIndex(cTID('Actn'), j);
				ref.putIndex(cTID('ASet'), set.index);
				var adesc = executeActionGet(ref);
				var actName = adesc.getString(cTID('Nm  '));
				set.actions.push(actName);
			}
			sets.push(set);
		}
		i++;
	}

	function getActions(aset) {
		cTID = function(s) {
			return app.charIDToTypeID(s);
		};
		sTID = function(s) {
			return app.stringIDToTypeID(s);
		};
		var i = 1;
		var names = [];
		if (!aset) {
			throw "Action set must be specified";
		}
		while (true) {
			var ref = new ActionReference();
			ref.putIndex(cTID("ASet"), i);
			var desc;
			try {
				desc = executeActionGet(ref);
			} catch (e) {
				break;
			}
			if (desc.hasKey(cTID("Nm  "))) {
				var name = desc.getString(cTID("Nm  "));
				if (name == aset) {
					var count = desc.getInteger(cTID("NmbC"));
					var names = [];
					for (var j = 1; j <= count; j++) {
						var ref = new ActionReference();
						ref.putIndex(cTID('Actn'), j);
						ref.putIndex(cTID('ASet'), i);
						var adesc = executeActionGet(ref);
						var actName = adesc.getString(cTID('Nm  '));
						names.push(actName);
					}
					break;
				}
			}
			i++;
		}
		return names;
	};
	var ActionList = [];
	for (i = 0; i < sets.length; i++) ActionList.push([sets[i], getActions(sets[i])]);
	var AL = [];
	for (var i = 0; i < ActionList.length; i++) {
		for (var j = 0; j < ActionList[i].length; j++) {
			if (ActionList[i][j] instanceof Object) {
				var SetName = ActionList[i][j].name;
				try {
					for (var k = 0; k < ActionList[i][j].actions.length; k++) {
						AL.push(SetName + "::->>" + ActionList[i][j].actions[k]);
					}
				} catch (e) {}
			}
		}
	}
	return AL;
}