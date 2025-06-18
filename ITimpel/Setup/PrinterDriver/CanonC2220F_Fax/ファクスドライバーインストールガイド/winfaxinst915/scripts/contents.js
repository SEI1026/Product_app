window.onload = function fncOnLoad() {
	try {

		// IEバージョンの判定
		var bIsIE6;
		if (navigator.userAgent.indexOf("MSIE 6") != -1) {
			bIsIE6 = true;
		}

		// ----------------------------------------------------------------------------------------
		// タイトルの指定
		// ----------------------------------------------------------------------------------------
		document.title = fncGetResourceByResourceId("sitemap") + " - " + fncGetResourceByResourceId("title");

		fncIncludeHeader();
		fncIncludeFooter();
		fncLoadResource();
		fncGenerateDynamicLink();

		// ----------------------------------------------------------------------------------------
		// カテゴリーのロード
		// ----------------------------------------------------------------------------------------
		var t = eval(toc);
		var nTLength = t.length;
		var strPrevGroup = "";
		var strDivMark = "::";
		var strGroupName = "";
		var arrCatHtml = new Array();
		var arrTocHtml = new Array();
		var arrCategories = new Array();
		var strResContentsLinkPrev = fncGetResourceByResourceId("contents_link_prev");
		var strResContentsLinkNext = fncGetResourceByResourceId("contents_link_next");
		var strResContentsLinkTop = fncGetResourceByResourceId("contents_link_top");
		var bH26 = false;
		arrCatHtml.push("<div class=\"title\">" + fncGetResourceByResourceId("sitemap") + "</div>");
		arrCatHtml.push("<table class=\"link\" cellspacing=\"0\" cellpadding=\"0\" id=\"id_top\">");

		for (var i = 0; i < nTLength; i++) {
			var toc_id = t[i].id;
			var toc_title = t[i].title;
			var toc_level = t[i].level;
			var toc_href = t[i].href;
			var toc_show_toc = t[i].show_toc;
			var toc_link_name = t[i].link_name;
			if (toc_show_toc == "n") {
				continue;
			}
			if (toc_level != 1) {
				continue;
			} else {
				arrCategories.push(t[i]);
			}
		}
		var j = 0;
		for (var i = 0; i < nTLength; i++) {
			var toc_id = t[i].id;
			var toc_title = t[i].title;
			var toc_level = t[i].level;
			var toc_href = t[i].href;
			var toc_show_toc = t[i].show_toc;
			var toc_link_name = t[i].link_name;

			if (toc_show_toc == "n") {
				continue;
			}

			// カテゴリーレベルのノード
			if (	(toc_level == 1)
				&&	(toc_link_name != "")
			) {
				var nDivPos = toc_link_name.indexOf(strDivMark);
				var toc_group = toc_link_name.substring(0, nDivPos);
				if (toc_link_name.indexOf(strDivMark) != -1) {
					toc_link_name = toc_link_name.substring(nDivPos + strDivMark.length);
				}

				// カテゴリーグループ列の表示
				if (strPrevGroup != toc_group) {
					strGroupName = fncGetResourceByResourceId(toc_group);

					if (strPrevGroup == "") {
						arrCatHtml.push("<tr>");
					} else {
						arrCatHtml.push("</td></tr><tr>");
					}
					arrCatHtml.push("<td class=\"group\">" + strGroupName + "</td><td class=\"category\">");
					if (i != 0) {
						if (bH26) {
							arrTocHtml.push("<div class=\"group_end\"></div>");
						}
					}
					arrTocHtml.push("<div class=\"group\"><span>" + strGroupName + "</span></div>");

					// 前回のカテゴリーグループに指定
					strPrevGroup = toc_group;

				}

				if (toc_group != "") {

					// カテゴリーの表示
					arrCatHtml.push("<a href=\"#id_" + toc_link_name + "\"><img src=\"../frame_images/cnt_link_next.gif\" />");
					arrCatHtml.push(toc_title);
					arrCatHtml.push("</a>");

					// NOTE: IE9は<wbr />のサポートをドロップ
					arrCatHtml.push("<span>");
					arrCatHtml.push("|");

					// NOTE: IE6の一部の環境で0x200Bが文字化け
					if (!bIsIE6) {
						arrCatHtml.push(String.fromCharCode(0x200B)); // <= for IE9
					}
					arrCatHtml.push("</span>");
					arrCatHtml.push("<wbr />"); // <= for SF
				}
			}

			if (toc_level == 1) {
				if (!bH26) {
					arrTocHtml.push("<div class=\"spacer\" style=\"border-top:0;\"></div>");
				} else {
					arrTocHtml.push("<div class=\"spacer\"></div>");
				}
			}

			arrTocHtml.push("<h" + toc_level + ">");

			// カテゴリータイトル <h1>
			if ((toc_level == 1) && (toc_link_name != "")) {
				bH26 = false;

				arrTocHtml.push("<table class=\"h1\" cellspacing=\"0\" cellpadding=\"0\">");
				arrTocHtml.push("<tr>");
				arrTocHtml.push("<th><span>");

				// タイトル
				if (!toc_href) {
					arrTocHtml.push(toc_title);
				} else {
					arrTocHtml.push("<a id=\"id_" + toc_link_name + "\" href=\"../contents/" + toc_href + "\">" + toc_title + "</a>");
				}
				arrTocHtml.push("</span></th>");
				arrTocHtml.push("<td style=\"width:100%;\"></td>");
				arrTocHtml.push("<td>");

				// 前へ
				if (j > 0) {
					var nDivPos = arrCategories[j - 1].link_name.indexOf(strDivMark);
					var prev_link_name = arrCategories[j - 1].link_name;
					if (nDivPos != -1) {
						prev_link_name = prev_link_name.substring(nDivPos + strDivMark.length);
					}
					arrTocHtml.push("<a href=\"#id_" + prev_link_name + "\"><img src=\"../frame_images/cnt_link_prev.gif\" />" + strResContentsLinkPrev + "</a> ");
				}

				// 次へ
				if (j < arrCategories.length - 1) {
					var nDivPos = arrCategories[j + 1].link_name.indexOf(strDivMark);
					var next_link_name = arrCategories[j + 1].link_name;
					if (nDivPos != -1) {
						next_link_name = next_link_name.substring(nDivPos + strDivMark.length);
					}
					arrTocHtml.push("<a href=\"#id_" + next_link_name + "\"><img src=\"../frame_images/cnt_link_next.gif\" />" + strResContentsLinkNext + "</a> ");
				}

				// このページのトップへ
				arrTocHtml.push("<a href=\"#id_top\"><img src=\"../frame_images/cnt_link_top.gif\" />" + strResContentsLinkTop + "</a> ");
				arrTocHtml.push("</td>");
				arrTocHtml.push("</tr>");
				arrTocHtml.push("</table>");
				j++;

			// <h2>～<h6>
			} else {
				bH26 = true;
				if (!toc_href) {
					arrTocHtml.push(toc_title);
				} else {
					arrTocHtml.push("<a id=\"id_" + toc_link_name + "\" href=\"../contents/" + toc_href + "\">" + toc_title + "</a>");
				}
			}
			arrTocHtml.push("</h" + toc_level + ">");
		}
		arrCatHtml.push("</td></tr></table>");
		if (!bH26) {
			arrTocHtml.push("<div class=\"bottom\" style=\"border:0;\"><a href=\"#id_top\"><img src=\"../frame_images/cnt_link_top.gif\" />" + strResContentsLinkTop + "</a></div>");
		} else {
			arrTocHtml.push("<div class=\"bottom\"><a href=\"#id_top\"><img src=\"../frame_images/cnt_link_top.gif\" />" + strResContentsLinkTop + "</a></div>");
		}
		document.getElementById("id_cat").innerHTML = arrCatHtml.join("");
		document.getElementById("id_toc").innerHTML = arrTocHtml.join("");
		fncResizeFrame();

		// タブ位置の記憶（Site mapから遷移した場合は、必ず目次タブに戻す）
		var strTabPosition = "0";
		fncSetCookie("TAB-POSITION", strTabPosition);

		// 「Contents」タブ表示時にスクロール位置を初期化する
		fncSetCookie("CONTENTS-SCROLL", "INIT");

	} catch (e) {
	}
}
window.onresize = fncResizeFrame;