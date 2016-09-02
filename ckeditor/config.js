/*
Copyright (c) 2003-2011, CKSource - Frederico Knabben. All rights reserved.
For licensing, see LICENSE.html or http://ckeditor.com/license
*/

CKEDITOR.editorConfig = function( config )
{
	// Define changes to default configuration here. For example:
	// config.language = 'fr';
	// config.uiColor = '#AADC6E';
	config.language = 'zh-cn';
// 界面语言，默认为 'en'

    config.language = 'zh-cn';

// 设置宽高

    config.width = 700;

    config.height = 200;

// 编辑器样式，有三种：'kama'（默认）、'office2003'、'v2'

    config.skin = 'v2';

// 背景颜色

    config.uiColor = '#FFF';

// 工具栏（基础'Basic'、全能'Full'、自定义）plugins/toolbar/plugin.js

    //config.toolbar = 'Basic';
    config.toolbar = 'Full';
	config.toolbar_Full = [
	['Source','-','Save','NewPage','Preview','-','Templates'],
	['Cut','Copy','Paste','PasteText','PasteFromWord','-','Print', 'SpellChecker', 'Scayt'],
	['Undo','Redo','-','Find','Replace','-','SelectAll','RemoveFormat'],
	['Form', 'Checkbox', 'Radio', 'TextField', 'Textarea', 'Select', 'Button', 'ImageButton', 'HiddenField'],
	'/',
	['Bold','Italic','Underline','Strike','-','Subscript','Superscript'],
	['NumberedList','BulletedList','-','Outdent','Indent','Blockquote'],
	['JustifyLeft','JustifyCenter','JustifyRight','JustifyBlock'],
	['Link','Unlink','Anchor'],
	['Image','Flash','Table','HorizontalRule','Smiley','SpecialChar','PageBreak'],
	'/',
	['Styles','Format','Font','FontSize'],
	['TextColor','BGColor']
	];
	
	config.filebrowserBrowseUrl = '/ckeditor/ckfinder/ckfinder.html';
    config.filebrowserImageBrowseUrl = '/ckeditor/ckfinder/ckfinder.html?Type=Images';
    config.filebrowserFlashBrowseUrl = '/ckeditor/ckfinder/ckfinder.html?Type=Flash';
    config.filebrowserUploadUrl = '/ckeditor/ckfinder/core/connector/asp/connector.asp?command=QuickUpload&type=Images';
    config.filebrowserImageUploadUrl = '/Upfile_Photo.asp?t=fck&PhotoUrlID=1';
    config.filebrowserFlashUploadUrl = '/Upfile_Photo.asp?t=fck&PhotoUrlID=1';
};
