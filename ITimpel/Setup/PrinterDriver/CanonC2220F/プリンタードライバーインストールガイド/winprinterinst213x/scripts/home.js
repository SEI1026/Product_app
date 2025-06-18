window.onload = function fncOnLoad() {
	try {

		// resource.jsonのconstantでlang_codeが指定されている場合はドキュメントの言語属性を上書き
		// 例：
		// var constant = [
		// {
		// lang_code:		"zh-CN",
		// :
		var strLangCode = fncGetConstantByName("lang_code");
		if (strLangCode) {
			document.body.lang = strLangCode;
		}

		fncIncludeHeader();
		fncIncludeFooter();
		fncLoadResource();
		fncGenerateDynamicLink();
		fncResizeFrame();

		// 検索ボタンにクリックイベントをセット
		if (document.getElementById("id_search_button")) {
			document.getElementById("id_search_button").onclick = function() {

				var strSearchTextsOriginal = document.getElementById("id_search_texts").value;
				var strSearchTexts = strSearchTextsOriginal;

				// --------------------------------------------------------------------------------
				// キーワードとして有効な文字列が指定されているかを検査
				// --------------------------------------------------------------------------------

				// 未入力→許容（検索説明に誘導）
				if (	(strSearchTexts == "")
					||	(strSearchTexts == fncGetResourceByResourceId("enter_search_keyword"))
				) {
					strSearchTextsOriginal = "";
					document.getElementById("id_search_texts").value = "";
				}

				// 検索用クッキーをセット
				var strTabPosition = "2";
				fncSetCookie("TAB-POSITION", strTabPosition);
				fncSetCookie("SEARCH-KEYWORD", strSearchTextsOriginal);
				fncSetCookie("SEARCH-RESULT-SETTING", "");

				// 検索操作説明用トピックを呼び出し
				var strSearchHelpLinkName = fncGetConstantByName("search_help");
				var strSearchHelpTocId = fncGetTocIdByLinkName(strSearchHelpLinkName);

				// 見つからない場合は先頭のトピックを呼び出し
				if (!strSearchHelpTocId) {
					var t = eval(toc);
					strSearchHelpTocId = t[0].id;
				}
				fncOpenTopic(strSearchHelpTocId, 1);
			}
			document.getElementById("id_search_button").title = fncGetResourceByResourceId("search");
		}
		if (document.getElementById("id_search_texts")) {
			document.getElementById("id_search_texts").focus();

			// 検索キーワードの再現
			var strSavedSearchKeyword = fncGetCookie("SEARCH-KEYWORD");
			if (strSavedSearchKeyword) {
				document.getElementById("id_search_texts").value = strSavedSearchKeyword;
			}
		}

		fncSearchBox("hdr_srch_w_bg");

		// タブ位置の記憶（トップページから遷移した場合は、必ず目次タブに戻す）
		var strTabPosition = "0";
		fncSetCookie("TAB-POSITION", strTabPosition);

		// 「Contents」タブ表示時にスクロール位置を初期化する
		fncSetCookie("CONTENTS-SCROLL", "INIT");

	} catch (e) {
	}
}
window.onresize = fncResizeFrame;
window.onunload = function() {};
var strWindowType = "HOME";
var c = new Array();