window.onload = function fncOnLoad() {
	try {

		// print_chapter.html?chapter=[id]
		var strChapter = document.location.search.substring(9);
		strChapter = decodeURI(strChapter);

		// toc.jsonをロード
		var t = eval(toc);

		var strHtml = "";
		var arrHtml = new Array();

		// 印刷依頼メッセージ
		var strPrintInstructions = "\t\t\t<div class=\"print_instructions\"><img src=\"../frame_images/bar_icon_print.gif\" />" + fncGetResourceByResourceId("print_instructions") + "</div>\n";

		var strTemp = "";
		var strDelimiterBegin	= "<!--CONTENT_START-->";
		var strDelimiterEnd		= "<!--CONTENT_END-->";

		// 印刷対象チャプターフラグ
		var bCurrentChapter;

		// チャプター分けされていない場合はすべて印刷対象
		var bPrintAll = false;
		if (strChapter == "ALL") {
			bPrintAll = true;
		}

		// 各モジュールで使用されているCSSのリスト
		var arrStyles = new Array();

		// toc.json要素分ループ
		var iLoopLength = t.length;
		for (var i = 0; i < iLoopLength; i++) {

			// 指定されたチャプターのidと一致
			if (	(t[i].id == strChapter)
				||	(bPrintAll)
			) {
				bCurrentChapter = true;

				// ウィンドウタイトルはチャプター名
				document.title = t[i].title;

				// スタイル番号を取得
				if (t[i].style && t[i].style != "") {
					if (arrStyles.indexOf(t[i].style) == -1) {
						arrStyles.push(t[i].style);
					}
				}

			// チャプターが一致するので印刷対象
			} else if (	(bCurrentChapter)
				&&	(1 < t[i].level)
				&&	(t[i].href)
			) {

				// 目次にないトピックは印刷しない
				if (t[i].show_toc == "n") {
					continue;
				}

				// スタイル番号を取得
				if (t[i].style && t[i].style != "") {
					if (arrStyles.indexOf(t[i].style) == -1) {
						arrStyles.push(t[i].style);
					}
				}

			// 印刷対象外
			} else {

				bCurrentChapter = false;

				// 以降処理不要
				if (bCurrentChapter) {
					break;
				}
				continue;
			}

			// 使用するAJAX収集手段を判定
			var bUseJqAjax = true;
			var ua = window.navigator.userAgent.toLowerCase();
			var isIE;
			if (ua.indexOf("msie") != -1 || ua.indexOf("trident") != -1) {
				isIE = true;
			}
			if (isIE) {
				var array;
				var version;
				array = /[msie|rv:]([\d\.]+)/.exec(ua);
				if (array) {
					version = array[1];
				} else {
					version = "";
				}

				// IE11の場合、jQueryを用いずAjax処理
				if (version >= 11) {
					//if (document.location.href.indexOf("file:///") != -1) { // ローカルコンテンツ
					bUseJqAjax = false;
					//}
				}
			}

			// AJAXによるコンテンツデータ収集
			if (bUseJqAjax) {
				strTemp = $.ajax({url:"../contents/" + t[i].href, async:false}).responseText;
			} else {
				strTemp = fncGetPageForPrintChapter("../contents/" + t[i].href);
			}

			// 不要部分の除去
			if (strTemp) {
				strTemp = strTemp.substring(
					strTemp.indexOf(strDelimiterBegin) + strDelimiterBegin.length + 27,
					strTemp.lastIndexOf(strDelimiterEnd) - 10
				);
				arrHtml.push(strTemp);
			}
		}

		// 収集したHTMLをページ内に配置
		strHtml = arrHtml.join("\t<div class=\"page_end\">&nbsp;</div>\n");
		document.body.innerHTML = strPrintInstructions + strHtml;

		// ウィンドウタイトルが末尾のトピックタイトルになることを防ぐ
		document.title = fncGetResourceByResourceId("title");

		// CSSを動的にロードする
		var iLoopLength = arrStyles.length;
		for (var i = 0; i < iLoopLength; i++) {
			if (arrStyles[i] == "") {
				continue;
			}
			var link = document.createElement("link");
			link.rel = "stylesheet";
			link.type = "text/css";
			link.href = "styles/style" + arrStyles[i] + "/style.css";
			document.getElementsByTagName("head")[0].appendChild(link);
		}

		// NOTE: 印刷命令はユーザーに委ねる
//		window.print();
//
//		// ウィンドウ終了
//		window.close();

	} catch (e) {
	}
}
function fncGetPageForPrintChapter(pageURL) {
	xmlhttp = createXMLHttpForPrintChapter();
	if (xmlhttp) {
		strReturnValue = null;
		xmlhttp.onreadystatechange = setPageDataForPrintChapter;
		xmlhttp.open('GET', pageURL, true);
		xmlhttp.send(null);
		if (strReturnValue) {
			return strReturnValue;
		}
	}
}
function setPageDataForPrintChapter() {
	if (xmlhttp.readyState == 4) {
		try {
			strReturnValue = xmlhttp.responseText;
		} catch (e) {
		}
	}
}
function createXMLHttpForPrintChapter() {
	try {
		return new ActiveXObject("Microsoft.XMLHTTP");
	} catch (e) {
		try {
			return new XMLHttpRequest();
		} catch (e) {
			return null;
		}
	}
	return null;
}