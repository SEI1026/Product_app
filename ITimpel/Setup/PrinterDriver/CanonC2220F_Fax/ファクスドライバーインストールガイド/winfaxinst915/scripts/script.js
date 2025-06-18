var c = new Array();
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

		// CSS情報
		var objLink = document.getElementsByTagName("link");
		var nLinkLength = objLink.length;
		var strStyleNumber = "";

		// ----------------------------------------------------------------------------------------
		// 別ウィンドウにロードされた場合の処理
		// ----------------------------------------------------------------------------------------
		if (document.location.search.indexOf("?sub=yes") != -1) {

			// 別ウィンドウ用スタイル定義にCSSを切換える
			// NOTE: SF5.1ではalternate stylesheetが正しくロードされない
			for (var i = 0; i < nLinkLength; i++) {
				if (objLink[i].href.indexOf("frame_style.css") != -1) {
					objLink[i].href = "../styles/frame_sub.css";
					break;
				}
			}

			// 別ウィンドウにロードされた場合、閉じるボタンを末尾に追加
			// <div class="close"></div>
			var objCloseDiv = document.createElement("div");
			objCloseDiv.className = "close";
			objCloseDiv.id = "id_close";

			// <button class="close" onclick="" accesskey="c" title="[close]">[close]</button>
			var objCloseButton = document.createElement("button");
			objCloseButton.onclick = function() {
				window.close();
			}
			objCloseButton.onmouseover = function() {
				this.style.backgroundColor = "#FFFFFF";
			}
			objCloseButton.onmouseout = function() {
				this.style.backgroundColor = "#EFEFEF";
			}
			objCloseButton.className = "close";
			objCloseButton.innerHTML = fncGetResourceByResourceId("close");
			objCloseButton.title = fncGetResourceByResourceId("close");
			objCloseButton.accessKey = "c";

			// <div class="close">
			//	<button class="close" onclick="" accesskey="c">[close]</button>
			// </div>
			objCloseDiv.appendChild(objCloseButton);
			document.getElementById("id_content").appendChild(objCloseDiv);
			document.onkeypress = fncKeyPress;

			fncOnResize();

			// 以降の処理（リソースのロード、目次生成など）不要
			return;

		// ----------------------------------------------------------------------------------------
		// 第二階層として末端コンテンツを表示
		// ----------------------------------------------------------------------------------------
		} else {

			// index.htmlから起動していない場合、ウィンドウ名を復活させる
			// （お気に入りに登録された場合や、コンテンツHTMLを直接表示させた場合）
			if (window.name == "") {
				window.name = "canon_main_window";
			}

			// 通常のCSSをロード
			for (var i = 0; i < nLinkLength; i++) {
				if (objLink[i].href.indexOf("frame_style.css") != -1) {
					objLink[i].disabled = false;
				}
				if (objLink[i].href.indexOf("frame_sub.css") != -1) {
					objLink[i].disabled = true;
				}

				// スタイル番号を取得
				if (objLink[i].href.indexOf("styles/style") != -1) {

					// NOTE: hrefはIEでは相対パス、その他では絶対パスで取得される
					var arrStyleNumberPath = objLink[i].href.split("/");
					if (arrStyleNumberPath.length > 2) {
						strStyleNumber = arrStyleNumberPath[arrStyleNumberPath.length - 2];
						if (strStyleNumber != "") {
							strStyleNumber += "_";
						}
					}
				}
			}
		}

		// ヘッダー項目の生成
		fncIncludeHeader();

		// フッター項目の生成
		fncIncludeFooter();

		// リソースをロード
		fncLoadResource();

		// 動的リンク生成
		fncGenerateDynamicLink();

		// 検索初期表示エリア下部に表示する検索に関するヘルプリンク
		fncResetSearchDisplay();

		// 全角半角区別オプションの表示制御
		var search_option_multibyte = fncGetConstantByName("search_option_multibyte");
		if (search_option_multibyte != 1) {
			document.getElementById("id_search_options_multibyte").parentNode.style.visibility = "Hidden";
		}

		// 現在カテゴリーのID
		// TODO: ネームスペースを使用することになった場合、属性「toc_id」は「caesar:toc_id」となる
		var strCurrentChapterId = "";
		if (document.getElementById("id_level_1")) {
			strCurrentChapterId = document.getElementById("id_level_1").getAttribute("toc_id");
		}

		// パンくずのIDリスト
		var strBreadCrumbsTocIds = "";
		var strCurrentTocId = ""; // 現在ページのID

		// パンくずをたどり目次ツリーを展開
		var objBreadCrumb = document.getElementById("id_breadcrumbs");
		if (objBreadCrumb) {

			// パンくずの要素分繰り返し
			var iLoopLength = objBreadCrumb.getElementsByTagName("a").length;
			for (var i = 0; i < iLoopLength; i++) {

				// TODO: ネームスペースを使用することになった場合、属性「toc_id」は「caesar:toc_id」となる
				if (objBreadCrumb.getElementsByTagName("a")[i].getAttribute("toc_id")) {
					strBreadCrumbsTocIds += objBreadCrumb.getElementsByTagName("a")[i].getAttribute("toc_id") + ",";
					strCurrentTocId = objBreadCrumb.getElementsByTagName("a")[i].getAttribute("toc_id");
				}
			}

			// パンくずをh1タイトルの次に移動
			// NOTE: FireFoxではwhitespaceをchildNodesの勘定に含める
			var objTarget = document.getElementById("id_content").childNodes[0];

			// コンテンツ最初の段落がh1ではない場合、見つかるまで探索
			while (	(	(objTarget.className != strStyleNumber + "h1")
					&&	(objTarget.className != "h1")
				) &&	(objTarget.nextSibling)
			) {
				objTarget = objTarget.nextSibling;
			}
			if (	(objTarget.className == strStyleNumber + "h1")
				||	(objTarget.className == "h1")
			) {
				if (objTarget.className == "h1") {
					strStyleNumber = "";
				}

				// パンくずと文書番号の見た目をコントロール
				// 適用されているテンプレートによってスタイルを変更する
				var h1_style = objTarget.currentStyle || document.defaultView.getComputedStyle(objTarget, "");
				var h1_color = h1_style.backgroundColor;

				// PS Front系
				if (	(h1_color == "#4682b4") // IE
					||	(h1_color == "rgb(70, 130, 180)") // Other
				) {

					// パンくず
					if (objBreadCrumb) {

						// 配置
						objTarget.parentNode.insertBefore(objBreadCrumb, objTarget.nextSibling);
						objBreadCrumb.style.display = "Block";

						objBreadCrumb.style.position = "Static";
						objBreadCrumb.style.backgroundColor = "#EEEEFF";
						objBreadCrumb.style.padding = "1px 2px 2px 2px";
						objBreadCrumb.style.borderBottom = "Solid 5px #F5F5FF";
					}

					// 文書番号
					var objDocumentNumber = document.getElementById("id_document_number");
					if (objDocumentNumber) {

						// 配置
						objTarget.parentNode.insertBefore(objDocumentNumber, objTarget.nextSibling);

						// スタイル
						objDocumentNumber.style.backgroundColor = "#E3E3FF";
						objDocumentNumber.style.paddingRight = "2px";
					}

				// デフォルト
				} else {

					// パンくず
					if (objBreadCrumb) {

						// 配置
						objTarget.parentNode.insertBefore(objBreadCrumb, objTarget);
						objBreadCrumb.style.display = "Block";

						// スタイル
						objBreadCrumb.style.position = "Static";
						objBreadCrumb.style.backgroundColor = "#F3F8FA";
						objBreadCrumb.style.border = "None";
						objBreadCrumb.style.padding = "1px 10px 2px 2px";
						objBreadCrumb.style.marginBottom = "0px";
						objTarget.style.paddingTop = "15px";
					}

					// 文書番号
					var objDocumentNumber = document.getElementById("id_document_number");
					if (objDocumentNumber) {

						// 配置
						objTarget.parentNode.insertBefore(objDocumentNumber, objTarget.nextSibling);

						// スタイル
						objDocumentNumber.style.marginTop = "-12px";
					}

					// 先頭に戻るリンク
					var nLinkToTop = fncGetConstantByName("link_to_top");
					if (nLinkToTop == 1) {
						var strLinkToTop = "";
						var ltt = eval(link_to_top_title);
						if (!strLangCode) {
							strLangCode = document.getElementsByTagName('html')[0].attributes["xml:lang"].value;
						}
						var iLoopLength = 0;
						iLoopLength = ltt.length;
						for (var i = 0; i < iLoopLength; i++) {
							if (ltt[i].lang == strLangCode) {
								strLinkToTop = ltt[i].title;
								break;
							}
							if (ltt[i].lang == "default") {
								strLinkToTop = ltt[i].title;
							}
						}
						var link_to_top = document.createElement("div");
						var link_to_top_a = document.createElement("a");
						link_to_top_a.href = "#id_breadcrumbs";
						link_to_top_a.title = strLinkToTop;
						link_to_top_a.innerHTML = strLinkToTop;
						link_to_top.className = strStyleNumber + "link_to_top";
						link_to_top.appendChild(link_to_top_a);
						document.getElementById("id_content").appendChild(link_to_top);
					}
				}
			}
		}

		// 現在ページのIDをBODYにセット（検索時にヒット項目が現在ページかどうかを判定するため）
		if (strCurrentTocId != "") {
			document.body.toc_id = strCurrentTocId;
		}

		var bJoinChapters = false;
		document.getElementById("id_panel_toc").className = "apart";
		if (fncGetConstantByName("join_chapters") == 1) {
			bJoinChapters = true;

			// チャプターコントロールを削除
			if (document.getElementById("id_chapter_select")) {
				document.getElementById("id_chapter_select").parentNode.parentNode.removeChild(document.getElementById("id_chapter_select").parentNode);
			}

			// TOCツリーのインデント調整
			document.getElementById("id_panel_toc").className = "join";

		// チャプターコントロールがない
		} else if (!document.getElementById("id_chapter_select")) {
			bJoinChapters = true;
			document.getElementById("id_panel_toc").className = "join";

		// 組み込み目次のトップページ
		} else if (strWindowType == "HOME_TOC") {
			bJoinChapters = true;
			document.getElementById("id_panel_toc").className = "join";

			// チャプターコントロールを削除
			if (document.getElementById("id_chapter_select")) {
				document.getElementById("id_chapter_select").parentNode.parentNode.removeChild(document.getElementById("id_chapter_select").parentNode);
			}
		}

		var iHideLevel = 999;
		var bIsCurrentChapter = false;

		// ----------------------------------------------------------------------------------------
		// 目次描画処理
		// ----------------------------------------------------------------------------------------
		if (document.getElementById("id_toc")) {

			// チャプター冒頭のlevel_1末端コンテンツ（紹介ページ）の表示方法
			// resource.jsonのcoverに定義がない場合は、level_1末端コンテンツはプルダウンには表示する
			// が、その下のツリーには表示しない。定義がある場合は、ツリーにも表示。
			try {
				var v = eval(cover);
				var vLength = v.length;
			} catch (e) {
				var vLength = 0;
			}

			// 目次情報をJSONから取得
			var t = eval(toc);
			var tLength = t.length;
			var strScroll = fncGetCookie("CONTENTS-SCROLL");
			for (var i = 0; i < tLength; i++) {

				// チャプターをまとめて表示
				if (bJoinChapters) {

					// 目次に表示しない場合スキップ
					if (t[i].show_toc == "n") {
						continue;
					}

					// h1 - h6
					var objHn = document.createElement("h" + t[i].level);
					if (iHideLevel < t[i].level) {
						objHn.style.display = "None";
					} else {
						iHideLevel = 999;
					}

					// [+][-][ ]
					var objSign = document.createElement("img");

					// 子トピックあり
					if	(	(t[i + 1])
						&&	(t[i + 1].level > t[i].level)
						)
					{

						// [+][-]リンク要素の作成
						var objAnchorSign = document.createElement("a");

						// パンくず上に存在→展開
						if (strBreadCrumbsTocIds.indexOf(t[i].id + ",") != -1) {
							objSign.src = "../frame_images/toc_sign_1.gif";

						// 折り畳み表示
						} else {
							objSign.src = "../frame_images/toc_sign_2.gif";

							// 現在のレベル以降を非表示にする
							if (iHideLevel == 999) {
								iHideLevel = t[i].level;
							}
						}

						// [+][-]クリック時処理の定義
						objAnchorSign.onclick = fncSwitchTocWrapper;
						objAnchorSign.href = "#";
						objAnchorSign.className = "sign";

						// [+][-]リンク要素に[+][-]マークを挿入
						objAnchorSign.appendChild(objSign);

						// 目次項目に[+][-]リンク要素を挿入
						objHn.appendChild(objAnchorSign);

					// 子トピックなし
					} else {

						// 現在表示中のトピック
						if (strCurrentTocId == t[i].id) {
							objSign.src = "../frame_images/toc_sign_0_actv.gif";
						} else {
							objSign.src = "../frame_images/toc_sign_0.gif";
						}
						objHn.appendChild(objSign);
					}

					// トピックにHTMLが存在する場合
					if (t[i].href) {

						// HTMLへのリンク要素を作成
						var objAnchor = document.createElement("a");
						objAnchor.href = "../contents/" + t[i].href;
						objAnchor.innerHTML = t[i].title;

						// 現在表示中のトピック
						if (strCurrentTocId == t[i].id) {
							objAnchor.className = "current";
							objAnchor.id = "id_toc_current";
						}

						// 目次項目にHTMLへのリンク要素を挿入（[+][-]の後ろ）
						objHn.appendChild(objAnchor);

					// 階層だけのノードの場合
					} else {
						var objSpan = document.createElement("span");
						objSpan.innerHTML = t[i].title;
						objHn.appendChild(objSpan);
					}

					// マウスオーバー時のティップス情報
					objHn.title = t[i].title.replace(/&quot;/g, "\"").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");

					// 目次項目をドキュメントに挿入
					document.getElementById("id_toc").appendChild(objHn);

				// チャプターをプルダウンで切り替える
				} else {

					// h1階層はプルダウンにチャプターとして出力
					if (t[i].level == 1) {

						// 目次に出力しない場合スキップ
						if (t[i].show_toc == "n") {
							continue;
						}

						// <option>要素の作成
						var objOption = document.createElement("option");
						objOption.id = t[i].id;
						objOption.innerHTML = t[i].title;

						// プルダウンに要素を挿入
						document.getElementById("id_chapter_select").appendChild(objOption);

						// 当該チャプター以外の情報は出力しないようにするための制御
						if (strCurrentChapterId == t[i].id) {
							bIsCurrentChapter = true;

							// 当該チャプターを選択状態にする
							objOption.selected = true;

							// ページを離れたときに目次スクロール位置を記憶するために、現在のチャプターIDを保持
							document.getElementById("id_chapter_select").current_chapter_id = strCurrentChapterId;

						} else {
							bIsCurrentChapter = false;
						}

						// 現在のチャプター
						if (bIsCurrentChapter) {
							var strTocLinkName = t[i].link_name;

							// resource.jsonのcoverに現在のチャプター名が存在するかどうかを探索
							for (var j = 0; j < vLength; j++) {

								// toc.link_nameに「::」が含まれる場合、
								//「::」から前の文字列は上位グループ名
								//「::」から後の文字列はチャプター名
								var strDivMark = "::";
								var nDivPosition = strTocLinkName.indexOf(strDivMark);

								// toc.link_nameからチャプター名を取得
								if (nDivPosition != -1) {
									strTocLinkName = strTocLinkName.substring(nDivPosition + strDivMark.length);
								}

								// チャプター名合致
								if (v[j].cover_name == strTocLinkName) {

									// h2としてTOCに要素を追加
									var objHn = document.createElement("h2");
									var objSign = document.createElement("img");

									// 現在表示中のトピック
									if (strCurrentTocId == t[i].id) {
										objSign.src = "../frame_images/toc_sign_0_actv.gif";
									} else {
										objSign.src = "../frame_images/toc_sign_0.gif";
									}
									objHn.appendChild(objSign);

									// トピックにHTMLが存在する場合、リンク可能に
									if (t[i].href) {

										// HTMLへのリンク要素を作成
										var objAnchor = document.createElement("a");
										objAnchor.href = "../contents/" + t[i].href;
										objAnchor.innerHTML = v[j].cover_title;

										// 現在表示中のトピック
										if (strCurrentTocId == t[i].id) {
											objAnchor.className = "current";
											objAnchor.id = "id_toc_current";
										}

										// 目次項目にHTMLへのリンク要素を挿入（[+][-]の後ろ）
										objHn.appendChild(objAnchor);

										// マウスオーバー時のティップス情報
										objHn.title = v[j].cover_title;

										// 目次項目をドキュメントに挿入
										document.getElementById("id_toc").appendChild(objHn);
									}
									break;
								}
							}
						}

						// 「Top」や「Site map」から移動してきた場合は、スクロール位置を初期化する
						if (strScroll == "INIT") {
							// スクロールを初期化
							var nScrollTop = fncGetCookie("TOC-SCROLL-POSITION-TOP-" + t[i].id);
							if (nScrollTop != "" && nScrollTop != 0) {
								// Cookieが存在する場合は「0」に設定する
								fncSetCookie("TOC-SCROLL-POSITION-TOP-" + t[i].id, 0);
							}
							var nScrollLeft = fncGetCookie("TOC-SCROLL-POSITION-LEFT-" + t[i].id);
							if (nScrollLeft != "" && nScrollLeft != 0) {
								// Cookieが存在する場合は「0」に設定する
								fncSetCookie("TOC-SCROLL-POSITION-LEFT-" + t[i].id, 0);
							}
						}

					// チャプター下の目次階層
					} else {

						// 目次に表示しない場合スキップ
						if (t[i].show_toc == "n") {
							continue;
						}

						// 当該チャプターのみ情報を出力（他はスキップ）
						if (bIsCurrentChapter) {

							// h2 - h6
							var objHn = document.createElement("h" + t[i].level);
							if (iHideLevel < t[i].level) {
								objHn.style.display = "None";
							} else {
								iHideLevel = 999;
							}

							// [+][-][ ]
							var objSign = document.createElement("img");

							// 子トピックあり
							if	(	(t[i + 1])
								&&	(t[i + 1].level > t[i].level)
								)
							{

								// [+][-]リンク要素の作成
								var objAnchorSign = document.createElement("a");

								// パンくず上に存在→展開
								if (strBreadCrumbsTocIds.indexOf(t[i].id + ",") != -1) {
									objSign.src = "../frame_images/toc_sign_1.gif";

								// 折り畳み表示
								} else {
									objSign.src = "../frame_images/toc_sign_2.gif";

									// 現在のレベル以降を非表示にする
									if (iHideLevel == 999) {
										iHideLevel = t[i].level;
									}
								}

								// [+][-]クリック時処理の定義
								objAnchorSign.onclick = fncSwitchTocWrapper;
								objAnchorSign.href = "#";
								objAnchorSign.className = "sign";

								// [+][-]リンク要素に[+][-]マークを挿入
								objAnchorSign.appendChild(objSign);

								// 目次項目に[+][-]リンク要素を挿入
								objHn.appendChild(objAnchorSign);

							// 子トピックなし
							} else {

								// 現在表示中のトピック
								if (strCurrentTocId == t[i].id) {
									objSign.src = "../frame_images/toc_sign_0_actv.gif";
								} else {
									objSign.src = "../frame_images/toc_sign_0.gif";
								}
								objHn.appendChild(objSign);
							}

							// トピックにHTMLが存在する場合
							if (t[i].href) {

								// HTMLへのリンク要素を作成
								var objAnchor = document.createElement("a");
								objAnchor.href = "../contents/" + t[i].href;
								objAnchor.innerHTML = t[i].title;

								// 現在表示中のトピック
								if (strCurrentTocId == t[i].id) {
									objAnchor.className = "current";
									objAnchor.id = "id_toc_current";
								}

								// 目次項目にHTMLへのリンク要素を挿入（[+][-]の後ろ）
								objHn.appendChild(objAnchor);

							// 階層だけのノードの場合
							} else {
								var objSpan = document.createElement("span");
								objSpan.innerHTML = t[i].title;
								objHn.appendChild(objSpan);
							}

							// マウスオーバー時のティップス情報
							objHn.title = t[i].title.replace(/&quot;/g, "\"").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");

							// 目次項目をドキュメントに挿入
							document.getElementById("id_toc").appendChild(objHn);
						}
					}
				}
			}
			// 「Contents」タブ表示時のスクロール位置初期化フラグを元に戻す
			fncSetCookie("CONTENTS-SCROLL", "");
		}

		// ----------------------------------------------------------------------------------------
		// チャプター移動制御
		// ----------------------------------------------------------------------------------------
		if (	(!bJoinChapters)
			&&	(document.getElementById("id_chapter_select"))
		) {

			// チャプター選択プルダウン変更時のイベント定義
			document.getElementById("id_chapter_select").onchange = function() {
				document.getElementById("id_toc").scrollTop = 0;
				document.getElementById("id_toc").scrollLeft = 0;
				document.location.href = "../contents/" + this.options[this.selectedIndex].id + ".html";
			}
		}
		if (bJoinChapters) {
			strCurrentChapterId = "ALL";
		}

		// ----------------------------------------------------------------------------------------
		// 左パネルの表示・非表示
		// ----------------------------------------------------------------------------------------
		if (document.getElementById("id_res_bar_icon_toggle_panel")) {
			document.getElementById("id_res_bar_icon_toggle_panel").title = fncGetResourceByResourceId("toggle_panel");
			document.getElementById("id_res_bar_icon_toggle_panel").onclick = function() {

				// 非表示
				if (this.src.indexOf("hide") != -1) {

					// 非表示にされるボックスのスクロール位置を一時記憶
					if (document.getElementById("id_search_chapters")) {
						if (document.getElementById("id_panel_search").style.display.toLowerCase() != "none") {
							nTempSearchChaptersScrollTop = document.getElementById("id_search_chapters").scrollTop;
						}
					}
					if (document.getElementById("id_search_results")) {
						if (document.getElementById("id_panel_search").style.display.toLowerCase() != "none") {
							nTempSearchResultsScrollTop = document.getElementById("id_search_results").scrollTop;
						}
					}
					if (document.getElementById("id_toc")) {
						if (document.getElementById("id_panel_toc").style.display.toLowerCase() != "none") {
							nTempTocScrollTop = document.getElementById("id_toc").scrollTop;
						}
					}

					// 表示・非表示の切り替え
					document.getElementById("id_panel").style.display = "None";
					document.getElementById("id_tabs").style.display = "None";
					document.getElementById("id_content").style.width = fncGetWindowWidth() + "px";
					document.getElementById("id_content").style.left = "0px";
					this.src = "../frame_images/bar_icon_pnl_show.gif";
					this.style.left = "-10px";

				// 表示
				} else {
					document.getElementById("id_panel").style.display = "Block";
					if (document.all) {
						document.getElementById("id_tabs").style.display = "Block";
					} else {
						document.getElementById("id_tabs").style.display = "Table-Cell";
					}

					// NOTE: IE678でウィンドウサイズを300px以下まで狭くした時に幅指定が無効になる現象を避ける
					var nContentWidth = fncGetWindowWidth() - 300;
					if (nContentWidth < 0) {
						nContentWidth = 0;
					}
					document.getElementById("id_content").style.width = nContentWidth + "px";
					document.getElementById("id_content").style.left = "300px";
					this.src = "../frame_images/bar_icon_pnl_hide.gif";
					this.style.left = "283px";

					// 非表示にされていたボックスのスクロール位置を復元
					if (	(document.getElementById("id_search_chapters"))
						&&	(nTempSearchChaptersScrollTop != -1)
					) {
						document.getElementById("id_search_chapters").scrollTop = nTempSearchChaptersScrollTop;
					}
					if (	(document.getElementById("id_search_results"))
						&&	(nTempSearchResultsScrollTop != -1)
					) {
						document.getElementById("id_search_results").scrollTop = nTempSearchResultsScrollTop;
					}
					if (	(document.getElementById("id_toc"))
						&&	(nTempTocScrollTop != -1)
					) {
						document.getElementById("id_toc").scrollTop = nTempTocScrollTop;
					}
				}
			}

			if (fncGetCookie("LEFT_PANE_VISIBILITY") == "HIDE") {
				document.getElementById("id_panel").style.display = "None";
				document.getElementById("id_tabs").style.display = "None";
				document.getElementById("id_content").style.width = fncGetWindowWidth() + "px";
				document.getElementById("id_content").style.left = "0px";
				document.getElementById("id_res_bar_icon_toggle_panel").src = "../frame_images/bar_icon_pnl_show.gif";
				document.getElementById("id_res_bar_icon_toggle_panel").style.left = "-10px";
			}
		}

		if (strCurrentTocId != "") {
			var strSiblingTocIds = fncGetSiblingNodeId(strCurrentTocId);
			var strPreviousTocId = strSiblingTocIds.split(":")[0];
			var strNextTocId = strSiblingTocIds.split(":")[1];
		} else {
			var strPreviousTocId = "";
			var strNextTocId = "";
		}

		// ----------------------------------------------------------------------------------------
		// 前トピック
		// ----------------------------------------------------------------------------------------
		if (document.getElementById("id_res_bar_icon_previous")) {

			// 前トピックなし
			if (strPreviousTocId == "") {
				document.getElementById("id_res_bar_icon_previous").disabled = true;
				document.getElementById("id_res_bar_icon_previous").childNodes[0].src = "../frame_images/bar_icon_prev_dis.gif";
				document.getElementById("id_res_bar_icon_previous").style.cursor = "Default";
				document.getElementById("id_res_bar_icon_previous").childNodes[0].title = "";

			// 前トピックあり
			} else {
				document.getElementById("id_res_bar_icon_previous").onclick = function() {
					document.location.href = "../contents/" + strPreviousTocId + ".html";
				}
				document.getElementById("id_res_bar_icon_previous").onmouseover = function() {
					this.style.backgroundImage = "Url(\"../frame_images/bar_icon_bg_over.gif?" + Math.random() + "\")";
				}
				document.getElementById("id_res_bar_icon_previous").onmouseout = function() {
					this.style.backgroundImage = "None";
				}
			}

			// ラベル
			if (document.getElementById("id_res_bar_label_previous")) {
				document.getElementById("id_res_bar_label_previous").title = document.getElementById("id_res_bar_icon_previous").childNodes[0].title;

				// NOTE: Firefoxの場合<button><img title="～" /></button>のままでは「～」がポップアップで表示されない
				document.getElementById("id_res_bar_icon_previous").title = document.getElementById("id_res_bar_icon_previous").childNodes[0].title;
				if (!document.getElementById("id_res_bar_icon_previous").disabled) {
					document.getElementById("id_res_bar_label_previous").onmouseover = function() {
						document.getElementById("id_res_bar_icon_previous").onmouseover();
						document.getElementById("id_res_bar_label_previous").style.color = "#0000FF";
						document.getElementById("id_res_bar_label_previous").style.textDecoration = "Underline";
					}
					document.getElementById("id_res_bar_label_previous").onmouseout = function() {
						document.getElementById("id_res_bar_icon_previous").onmouseout();
						document.getElementById("id_res_bar_label_previous").style.color = "#0000FF";
						document.getElementById("id_res_bar_label_previous").style.textDecoration = "None";
					}
				} else {
					document.getElementById("id_res_bar_label_previous").style.color = "#808080";
					document.getElementById("id_res_bar_label_previous").style.cursor = "Default";
				}
			}
		}


		// ----------------------------------------------------------------------------------------
		// 次トピック
		// ----------------------------------------------------------------------------------------
		if (document.getElementById("id_res_bar_icon_next")) {

			// 次トピックなし
			if (strNextTocId == "") {
				document.getElementById("id_res_bar_icon_next").disabled = true;
				document.getElementById("id_res_bar_icon_next").childNodes[0].src = "../frame_images/bar_icon_next_dis.gif";
				document.getElementById("id_res_bar_icon_next").style.cursor = "Default";
				document.getElementById("id_res_bar_icon_next").childNodes[0].title = "";

			// 次トピックあり
			} else {
				document.getElementById("id_res_bar_icon_next").onclick = function() {
					document.location.href = "../contents/" + strNextTocId + ".html";
				}
				document.getElementById("id_res_bar_icon_next").onmouseover = function() {
					this.style.backgroundImage = "Url(\"../frame_images/bar_icon_bg_over.gif?" + Math.random() + "\")";
				}
				document.getElementById("id_res_bar_icon_next").onmouseout = function() {
					this.style.backgroundImage = "None";
				}
			}

			// ラベル
			if (document.getElementById("id_res_bar_label_next")) {
				document.getElementById("id_res_bar_label_next").title = document.getElementById("id_res_bar_icon_next").childNodes[0].title;

				// NOTE: Firefoxの場合<button><img title="～" /></button>のままでは「～」がポップアップで表示されない
				document.getElementById("id_res_bar_icon_next").title = document.getElementById("id_res_bar_icon_next").childNodes[0].title;
				if (!document.getElementById("id_res_bar_icon_next").disabled) {
					document.getElementById("id_res_bar_label_next").onmouseover = function() {
						document.getElementById("id_res_bar_icon_next").onmouseover();
						document.getElementById("id_res_bar_label_next").style.color = "#0000FF";
						document.getElementById("id_res_bar_label_next").style.textDecoration = "Underline";
					}
					document.getElementById("id_res_bar_label_next").onmouseout = function() {
						document.getElementById("id_res_bar_icon_next").onmouseout();
						document.getElementById("id_res_bar_label_next").style.color = "#0000FF";
						document.getElementById("id_res_bar_label_next").style.textDecoration = "None";
					}
				} else {
					document.getElementById("id_res_bar_label_next").style.color = "#808080";
					document.getElementById("id_res_bar_label_next").style.cursor = "Default";
				}
			}
		}

		// ----------------------------------------------------------------------------------------
		// 印刷ボタン
		// ----------------------------------------------------------------------------------------
		if (document.getElementById("id_res_bar_icon_print")) {
			document.getElementById("id_res_bar_icon_print").onclick = function() {
				window.print();
			}
			document.getElementById("id_res_bar_icon_print").onmouseover = function() {
				this.style.backgroundImage = "Url(\"../frame_images/bar_icon_bg_over.gif?" + Math.random() + "\")";
			}
			document.getElementById("id_res_bar_icon_print").onmouseout = function() {
				this.style.backgroundImage = "None";
			}

			// ラベル
			if (document.getElementById("id_res_bar_label_print")) {
				if (!document.getElementById("id_res_bar_icon_print").disabled) {
					document.getElementById("id_res_bar_label_print").title = document.getElementById("id_res_bar_icon_print").childNodes[0].title;

					// NOTE: Firefoxの場合<button><img title="～" /></button>のままでは「～」がポップアップで表示されない
					document.getElementById("id_res_bar_icon_print").title = document.getElementById("id_res_bar_icon_print").childNodes[0].title;
					document.getElementById("id_res_bar_label_print").onmouseover = function() {
						document.getElementById("id_res_bar_icon_print").onmouseover();
						document.getElementById("id_res_bar_label_print").style.color = "#0000FF";
						document.getElementById("id_res_bar_label_print").style.textDecoration = "Underline";
					}
					document.getElementById("id_res_bar_label_print").onmouseout = function() {
						document.getElementById("id_res_bar_icon_print").onmouseout();
						document.getElementById("id_res_bar_label_print").style.color = "#0000FF";
						document.getElementById("id_res_bar_label_print").style.textDecoration = "None";
					}
				} else {
					document.getElementById("id_res_bar_icon_print").childNodes[0].src = "../frame_images/bar_icon_print_dis.gif"
					document.getElementById("id_res_bar_icon_print").style.cursor = "Default";
					document.getElementById("id_res_bar_icon_print").title = "";
					document.getElementById("id_res_bar_icon_print").childNodes[0].title = "";
					document.getElementById("id_res_bar_label_print").style.color = "#808080";
					document.getElementById("id_res_bar_label_print").style.cursor = "Default";
				}
				document.getElementById("id_res_bar_label_print").style.borderRight = "0";
			}
		}

		// ----------------------------------------------------------------------------------------
		// タブの切り換え
		// ----------------------------------------------------------------------------------------
		if (	(document.getElementById("id_res_search"))
			&&	(document.getElementById("id_res_contents"))
		) {

			// 検索
			document.getElementById("id_res_search").onclick = function() {
				fncResSearchClick();
			}
			document.getElementById("id_res_search").onmouseover = function() {
				if (this.className == "tab_inactive") {
					this.className = "tab_inactive_over";

					// NOTE: FF4で目次/検索タブ上にマウスを乗せても2回目以降背景画像が正しく切り替わらない
					this.style.backgroundImage = "Url(\"../frame_images/tb_inactv_over.gif?" + Math.random() + "\")";
				}
			}
			document.getElementById("id_res_search").onmouseout = function() {
				if (this.className == "tab_inactive_over") {
					this.className = "tab_inactive";
					this.style.backgroundImage = "Url(\"../frame_images/tb_inactv.gif\")";
				}
			}

			// 目次
			document.getElementById("id_res_contents").onclick = function() {
				fncResContentsClick();
			}
			document.getElementById("id_res_contents").onmouseover = function() {
				if (this.className == "tab_inactive") {
					this.className = "tab_inactive_over";
					this.style.backgroundImage = "Url(\"../frame_images/tb_inactv_over.gif?" + Math.random() + "\")";
				}
			}
			document.getElementById("id_res_contents").onmouseout = function() {
				if (this.className == "tab_inactive_over") {
					this.className = "tab_inactive";
					this.style.backgroundImage = "Url(\"../frame_images/tb_inactv.gif\")";
				}
			}
		}

		// 印刷のイベント定義
		if (document.getElementById("id_res_print_all")) {
			document.getElementById("id_res_print_all").onclick = function() {

				// チャプター統合目次の場合はすべてのチャプターが印刷対象
				window.open("../contents/print_chapter.html?chapter=" + strCurrentChapterId);
			}
			var strPrintAllTitle = document.getElementById("id_res_print_all").innerText;
			if (!strPrintAllTitle) {
				strPrintAllTitle = document.getElementById("id_res_print_all").textContent;
			}
			if (strPrintAllTitle) {
				document.getElementById("id_res_print_all").title = strPrintAllTitle;
			}
		}

		// ----------------------------------------------------------------------------------------
		// 検索
		// ----------------------------------------------------------------------------------------
		if (strWindowType != "HOME") {

			// 検索ボタンにクリックイベントをセット
			document.getElementById("id_search_button").onclick = function() {
				fncSetCookie("SEARCH-RESULT-SETTING", "");
				fncDoSearch(1);
			}
			document.getElementById("id_search_button").title = fncGetResourceByResourceId("search");

			// 検索オプションを表示・非表示 (デフォルト:非表示)
			if (document.getElementById("id_search_options_label")) {
				document.getElementById("id_search_options_label").innerHTML = "<a href=\"#\" onclick=\"fncToggleSearchOptions();\"><img src=\"../frame_images/srch_opt_show.gif\" />" + fncGetResourceByResourceId("search_options_show") + "</a>";
			}

			// すべてのカテゴリーから検索・以下のカテゴリーから検索 (デフォルト:非表示)
			if (document.getElementById("id_search_options_search_scope_all")) {
				document.getElementById("id_search_options_search_scope_all").onclick = function() {
					fncSelectChaptersFromAll();
				}
			}

			// チャプター一覧 (デフォルト:非表示)
			if (document.getElementById("id_search_chapters")) {
				// チャプター一覧抽出のため目次情報をロード
				// toc.jsonのlevelが「1」のノードは「チャプター」とする
				var t = eval(toc);
				var iLoopLength = t.length;
				for (var i = 0; i < iLoopLength; i++) {

					// チャプター情報を抽出
					if (	(t[i].level == 1)
						&&	(t[i].show_toc != "n")
					) {
						c.push(t[i]);
					}
				}
				var iLoopLength = c.length;
				var arrChapterHtml = new Array();
				for (var i = 0; i < iLoopLength; i++) {
					arrChapterHtml.push(fncGenerateChapterCheckbox(c[i].id, c[i].title));
				}

				// チャプター一覧を配置 (デフォルトは非表示)
				document.getElementById("id_search_chapters").innerHTML = arrChapterHtml.join("");
			}

			// 第2検索ボタン
			if (document.getElementById("id_res_search_button")) {
				document.getElementById("id_res_search_button").onclick = function() {
					fncSetCookie("SEARCH-RESULT-SETTING", "");
					fncDoSearch(1);
				}
				document.getElementById("id_res_search_button").onmouseover = function() {
					this.style.backgroundColor = "#FFFFFF";
				}
				document.getElementById("id_res_search_button").onmouseout = function() {
					this.style.backgroundColor = "#EFEFEF";
				}
				var strSearchButtonTitle = document.getElementById("id_res_search_button").innerText;
				if (!strSearchButtonTitle) {
					strSearchButtonTitle = document.getElementById("id_res_search_button").textContent;
				}
				if (strSearchButtonTitle) {
					document.getElementById("id_res_search_button").title = strSearchButtonTitle;
				}
			}

			// 検索オプション開閉状態の再現
			if (fncGetCookie("SEARCH-OPTIONS") == "CLOSE") {
				document.getElementById("id_search_options").style.display = "None";
				document.getElementById("id_search_options_label").innerHTML = "<a href=\"#\" onclick=\"fncToggleSearchOptions();\" title=\"" + fncGetResourceByResourceId("search_options_show") + "\"><img src=\"../frame_images/srch_opt_show.gif\" title=\"[+]\" />" + fncGetResourceByResourceId("search_options_show") + "</a>";
			} else if (fncGetCookie("SEARCH-OPTIONS") == "OPEN") {
				document.getElementById("id_search_options").style.display = "Block";
				document.getElementById("id_search_options_label").innerHTML = "<a href=\"#\" onclick=\"fncToggleSearchOptions();\" title=\"" + fncGetResourceByResourceId("search_options_hide") + "\"><img src=\"../frame_images/srch_opt_hide.gif\" title=\"[-]\" />" + fncGetResourceByResourceId("search_options_hide") + "</a>";
			} else {
				document.getElementById("id_search_options").style.display = "None";
				document.getElementById("id_search_options_label").innerHTML = "<a href=\"#\" onclick=\"fncToggleSearchOptions();\" title=\"" + fncGetResourceByResourceId("search_options_show") + "\"><img src=\"../frame_images/srch_opt_show.gif\" title=\"[+]\" />" + fncGetResourceByResourceId("search_options_show") + "</a>";
			}

			// 検索オプションの再現
			var strSavedSearchOptionScope = fncGetCookie("SEARCH-SCOPE");
			if (strSavedSearchOptionScope == "CHAPTER") {

				document.getElementById("id_search_options_search_scope_chapter").checked = true;

				// 選択チャプターの再現
				var search_chapters = document.getElementById("id_search_chapters");
				if (search_chapters) {
					var strSavedSearchChapters = fncGetCookie("SEARCH-CHAPTERS");
					if (strSavedSearchChapters != "") {
						var objChapterCheckboxes = search_chapters.getElementsByTagName("input");
						for (var i = 0; i < iLoopLength; i++) {
							if (strSavedSearchChapters.indexOf(objChapterCheckboxes[i].id) != -1) {
								objChapterCheckboxes[i].checked = true;
							}
						}
					}
				}

			// デフォルトは「すべて検索」
			} else {
				document.getElementById("id_search_options_search_scope_all").checked = true;
			}

			// 大文字小文字の区別
			if (fncGetCookie("SEARCH-OPTIONS-CASE") == "TRUE") {
				document.getElementById("id_search_options_case").checked = true;
			}

			// ------------------------------------------------------------------------------------
			// 全角半角の区別
			// ------------------------------------------------------------------------------------

			// 「非表示」の場合は「区別する」
			if (document.getElementById("id_search_options_multibyte").parentNode.style.visibility.toLowerCase() == "hidden") {
				document.getElementById("id_search_options_multibyte").checked = true;

			// 明示的に「区別しない」に指定されている場合
			} else if (fncGetCookie("SEARCH-OPTIONS-MULTIBYTE") == "FALSE") {
				document.getElementById("id_search_options_multibyte").checked = true;
			} else {
				document.getElementById("id_search_options_multibyte").checked = false;
			}

			// 検索キーワードの再現
			var strSavedSearchKeyword = fncGetCookie("SEARCH-KEYWORD");
			if (strSavedSearchKeyword) {
				document.getElementById("id_search").value = strSavedSearchKeyword;
			}
			fncSearchBox("hdr_srch_w_bg");
		}

		// タブ選択状態の再現
		if (document.getElementById("id_res_search")) {
			if (fncGetCookie("TAB-POSITION") == "2") {
				if (document.getElementById("id_search").value != fncGetResourceByResourceId("enter_search_keyword")) {
					document.getElementById("id_search_results").innerHTML = "<div class=\"message\">" + fncGetResourceByResourceId("search_message_wait") + "<br /><img src=\"../frame_images/srch_wait.gif\" /></div>";

					// 検索結果に表示するページを設定
					var nPage = 1;
					var strSearchResultSetting = fncGetCookie("SEARCH-RESULT-SETTING");
					if ("" != strSearchResultSetting) {
						// 検索結果から選択された場合はCookieから再表示するページ数を取得
						var arrSearchResultSetting = strSearchResultSetting.split(":");
						if ("" != arrSearchResultSetting[0] && undefined != arrSearchResultSetting[0]) {
							nPage = arrSearchResultSetting[0];
						}
					}
					var ti = window.setTimeout("fncDoSearch(" + nPage + ")", 13);
				}
				fncResSearchClick();
			} else if (document.getElementById("id_res_contents")) {
				fncResContentsClick();
			}
		} else if (document.getElementById("id_res_contents")) {
			fncResContentsClick();
		}

		// すべてたたむリンクのイベント定義
		document.getElementById("id_res_close_toc_all").onclick = function() {
			fncOpenCloseAllToc(1);
		}
		var strCloseTocAllTitle = document.getElementById("id_res_close_toc_all").innerText;
		if (!strCloseTocAllTitle) {
			strCloseTocAllTitle = document.getElementById("id_res_close_toc_all").textContent;
		}
		if (strCloseTocAllTitle) {
			document.getElementById("id_res_close_toc_all").title = strCloseTocAllTitle;
		}

		// すべて開くリンクのイベント定義
		document.getElementById("id_res_open_toc_all").onclick = function() {
			fncOpenCloseAllToc(2);
		}
		var strOpenTocAllTitle = document.getElementById("id_res_open_toc_all").innerText;
		if (!strOpenTocAllTitle) {
			strOpenTocAllTitle = document.getElementById("id_res_open_toc_all").textContent;
		}
		if (strOpenTocAllTitle) {
			document.getElementById("id_res_open_toc_all").title = strOpenTocAllTitle;
		}

		// 詳細開閉処理
		var objAnchors = document.getElementsByTagName("a");
		var nAnchorLength = objAnchors.length;
		for (var i = 0; i < nAnchorLength; i++) {
			if (objAnchors[i].className == "open_close_next_sibling") {
				objAnchors[i].innerHTML = fncGetResourceByResourceId("open_next_sibling");
				objAnchors[i].className = "open_next_sibling";
				objAnchors[i].onclick = function() {
					fncOpenCloseNextSibling(this);
					return false;
				}
			} else if(objAnchors[i].className == "open_all") {
				objAnchors[i].innerHTML = fncGetResourceByResourceId("open_all");
				objAnchors[i].onclick = function() {
					fncOpenCloseAll("open");
					return false;
				}
			} else if(objAnchors[i].className == "close_all") {
				objAnchors[i].innerHTML = fncGetResourceByResourceId("close_all");
				objAnchors[i].onclick = function() {
					fncOpenCloseAll("close");
					return false;
				}
			}
		}

		// 各ペインをウィンドウサイズに合わせてサイズ調整
		fncOnResize();

		// リンク移動してきた場合、対象トピックまで目次をスクロール
		if (	(document.location.hash != "")
			&&	(document.getElementById("id_toc_current"))
		) {			
			document.getElementById("id_toc_current").scrollIntoView(true);
			document.getElementById("id_toc").scrollLeft = 0;

		// 用語集からリンク移動してきた場合、対象トピックまで目次をスクロール
		} else if (document.location.search.indexOf("&word=yes") != -1) {	
			document.getElementById("id_toc_current").scrollIntoView(true);
			document.getElementById("id_toc").scrollLeft = 0;

		// 目次をクリックして移動してきた場合、スクロール位置を復元する
		} else {

			// 目次のスクロール位置（チャプターごとに再現）
			if (document.getElementById("id_toc")) {
				if (strCurrentChapterId == "") {
					strCurrentChapterId = "ALL";
				}

				// NOTE: IE7以降でスクロール位置が正しく反映されない問題の対応としてタイマーを設定
				if (document.all) {
					var ti = window.setTimeout(
						function() {
							document.getElementById("id_toc").scrollTop = fncGetCookie("TOC-SCROLL-POSITION-TOP-" + strCurrentChapterId);
							document.getElementById("id_toc").scrollLeft = fncGetCookie("TOC-SCROLL-POSITION-LEFT-" + strCurrentChapterId);
						},
						1
					);
				} else {
					document.getElementById("id_toc").scrollTop = fncGetCookie("TOC-SCROLL-POSITION-TOP-" + strCurrentChapterId);
					document.getElementById("id_toc").scrollLeft = fncGetCookie("TOC-SCROLL-POSITION-LEFT-" + strCurrentChapterId);
				}

				// それでもなお隠れているときは頭出し
				// NOTE: この処理はhome_toc.htmlでは不要
				var ti = window.setTimeout(
					function() {
						if (	(document.getElementById("id_toc_current"))
							&&	(document.getElementById("id_toc_current").offsetTop < document.getElementById("id_toc").scrollTop)
						) {
							document.getElementById("id_toc").scrollTop = document.getElementById("id_toc_current").offsetTop;
						}
					},
					2
				);
			}
		}

		// ----------------------------------------------------------------------------------------
		// 検索結果からジャンプしてきた場合、ヒット文字列をハイライトさせる
		// ----------------------------------------------------------------------------------------
		if (document.location.search) {
			if (strWindowType != "HOME") {
				var ti = window.setTimeout("fncMarkupSearch()", 26);
			}
		} else {

			// IE6の場合リンクアンカーまでスクロールしないので、自前でアンカー位置にスクロール
			if (document.all) { // FF/SFを除外
				if (document.location.hash != "") {
					var strHash = document.location.hash.substring(1);
					if (document.all.item(strHash)) {
						document.all.item(strHash).scrollIntoView(true);
					}
				}
			}
		}

		// ----------------------------------------------------------------------------------------
		// 本文中の文字列の内、用語集のタイトルと合致するものをハイライト
		// ----------------------------------------------------------------------------------------
		if (	(fncGetConstantByName("markup_glossary") == 1)
			&&	(strWindowType != "HOME")
			&&	(strWindowType != "HOME_TOC")
		) {
			var ti = window.setTimeout("fncDoMarkupGlossary()", 1500);
		}

		// ----------------------------------------------------------------------------------------
		// IE7の場合にのみ、スクロールのバーの位置と高さが正しく表示されないのを防ぐ
		// ----------------------------------------------------------------------------------------
		if (	(document.all) // FF/SFを除外
			&&	(window.XMLHttpRequest)	// IE6を除外
		) {
			// <hr>タグをheightを指定した<div>タグ内に移動
			// NOTE: ここで<hr>を正規表現で<div ..><hr /></div>としてinnerHTMLで入れ替えると、スクリプトで設定したonclickイベントが消えてしまう
			var hrs = document.getElementById("id_content").getElementsByTagName("hr");
			var iLoopLength = hrs.length;
			for (var i = 0; i < iLoopLength; i++) {
				var objDiv = document.createElement("div");
				objDiv.style.height = "40px";
				objDiv.style.verticalAlign = "Middle";
				hrs[i].parentNode.insertBefore(objDiv, hrs[i]);
				objDiv.appendChild(hrs[i]);
			}
		}
	} catch (e) {
	}
}

// NOTE: IE678以外のブラウザーはdisplay属性で表示・非表示を切り替えた際、元のスクロール位置が保持されない
var nTempSearchChaptersScrollTop = -1;
var nTempSearchResultsScrollTop = -1;
var nTempTocScrollTop = -1;
function fncResContentsClick() {
	try {

		// 非表示にされるボックスのスクロール位置を一時記憶
		if (document.getElementById("id_search_chapters")) {
			nTempSearchChaptersScrollTop = document.getElementById("id_search_chapters").scrollTop;
		}
		if (document.getElementById("id_search_results")) {
			nTempSearchResultsScrollTop = document.getElementById("id_search_results").scrollTop;
		}

		// 表示・非表示の切り替え
		document.getElementById("id_res_contents").className = "tab_active";
		document.getElementById("id_res_contents").style.backgroundImage = "Url(\"../frame_images/tb_actv.gif\")";
		document.getElementById("id_res_search").className = "tab_inactive";
		document.getElementById("id_res_search").style.backgroundImage = "Url(\"../frame_images/tb_inactv.gif\")";
		document.getElementById("id_panel_toc").style.display = "Block";
		document.getElementById("id_panel_search").style.display = "None";

		// 非表示にされていたボックスのスクロール位置を復元
		if (	(document.getElementById("id_toc"))
			&&	(nTempTocScrollTop != -1)
		) {
			document.getElementById("id_toc").scrollTop = nTempTocScrollTop;
		}
	} catch (e) {
	}
}
function fncResSearchClick() {
	try {

		// 非表示にされるボックスのスクロール位置を一時記憶
		if (document.getElementById("id_toc")) {
			nTempTocScrollTop = document.getElementById("id_toc").scrollTop;
		}

		// 表示・非表示の切り替え
		document.getElementById("id_res_search").className = "tab_active";
		document.getElementById("id_res_search").style.backgroundImage = "Url(\"../frame_images/tb_actv.gif\")";
		document.getElementById("id_res_contents").className = "tab_inactive";
		document.getElementById("id_res_contents").style.backgroundImage = "Url(\"../frame_images/tb_inactv.gif\")";
		document.getElementById("id_panel_toc").style.display = "None";
		document.getElementById("id_panel_search").style.display = "Block";

		// 非表示にされていたボックスのスクロール位置を復元
		if (	(document.getElementById("id_search_chapters"))
			&&	(nTempSearchChaptersScrollTop != -1)
		) {
			document.getElementById("id_search_chapters").scrollTop = nTempSearchChaptersScrollTop;
		}
		if (	(document.getElementById("id_search_results"))
			&&	(nTempSearchResultsScrollTop != -1)
		) {
			document.getElementById("id_search_results").scrollTop = nTempSearchResultsScrollTop;
		}
	} catch (e) {
	}
}
function fncDoMarkupGlossary() {
	try {
		var arrGlossaryIndex = fncCreateGlossaryIndex();
		var strLangCode = document.getElementsByTagName('html')[0].attributes["xml:lang"].value;
		var nMode = 0;
		if (strLangCode.match(/ja|zh|ko/)) {
			nMode = 1;
		}
		fncMarkupGlossary(document.getElementById("id_content"), arrGlossaryIndex, nMode);
	} catch (e) {
	}
}

function fncMarkupSearch() {
	try {

		// 検索条件を引数から取得
		var strSearchText = document.location.search.split("?search=")[1];

		strSearchText = strSearchText.split("&word=yes")[0];

		// トップページから検索した時の検索説明ページはハイライトさせない
		if (strSearchText.indexOf("&marking=no") != -1) {
			return;
		} else {
			strSearchText = strSearchText.split("&marking=")[0];
		}
		if (strSearchText != "") {

			// 検索条件文字列をデコード
			strSearchText = decodeURIComponent(strSearchText);

			// 前後のスペースを除去しておく
			// NOTE: IEとそれ以外のブラウザーで、前後にスペースがあるかどうかでsplitの結果が変わる（IEの場合、空の要素は省かれる）
			strSearchText = strSearchText.trim();

			if (strSearchText == fncGetResourceByResourceId("enter_search_keyword")) {
				return;
			}

			// NOTE: ここでエスケープすると2重処理

			// 複数検索条件を分解（全半角スペース）
			var res = /[\s　]+/;
			var arrSearchText = strSearchText.split(res);

			// NOTE: Safari3でキーワード前後に全角スペースがあると、配列に空の要素が作成されバーストが発生
			var iLoopLength = arrSearchText.length;
			for (var i = 0; i < iLoopLength; i++) {
				if (arrSearchText[i] == "") {
					arrSearchText.splice(i, 1);
				}
			}

			// 完全一致検索のためにダブルクォーテーションで囲まれた文字列内のスペース代替文字を復元
			var iLoopLength = arrSearchText.length;
			for (var i = 0; i < iLoopLength; i++) {
				arrSearchText[i] = arrSearchText[i].replace(/___SPACE___/g, " ");
			}

			// マルチバイトの区別
			if (document.getElementById("id_search_options_multibyte")) {
				if (!document.getElementById("id_search_options_multibyte").checked) {
					for (var i = 0; i < iLoopLength; i++) {
						arrSearchText[i] = fncConvertSearchText(arrSearchText[i], false);
					}
				}
			}

			// 大文字小文字の区別
			var strSearchOptionCaseSensitive = "i"; // 正規表現のフラグ「i」→区別あり
			if (fncGetCookie("SEARCH-OPTIONS-CASE") == "TRUE") {
				strSearchOptionCaseSensitive = "";
			}

			// 用語の検索結果から表示する場合、単語単位でハイライトする
			var bWordMarking = false;
			if (document.location.search.indexOf("&word=yes") != -1) {
				bWordMarking = true;
			}

			// コンテンツ領域内のテキスト要素をハイライト
			fncMarkupText(
				document.getElementById("id_content"),
				arrSearchText,
				strSearchOptionCaseSensitive,
				bWordMarking
			);

			// 折りたたみをすべて展開
			fncOpenCloseAll("open");

			// 最初にヒットした文字列までスクロール
			if (document.getElementById("id_hit")) {
				document.getElementById("id_hit").scrollIntoView(true);
			}
		}
	} catch (e) {
	}
}

var glossary;
// ------------------------------------------------------------------------------------------------
// 用語集のタイトル部分だけを配列化
// ------------------------------------------------------------------------------------------------
function fncCreateGlossaryIndex() {
	try {

		// ----------------------------------------------------------------------------------------
		// 用語集 (glossary.json) の取得
		// ----------------------------------------------------------------------------------------
		if (!glossary) {
			glossary = $.ajax({url:"../jsons/glossary.json", async:false}).responseText;
		}

		// NOTE: var glossary = [{...}];
		//                      ^^^^^^^ (v.2では動的にjsonをロードするため[～]部分だけが必要)
		var g = eval(glossary.substring(15, glossary.length - 1));
		var iLoopLength = g.length;
		var arrGlossaryIndex = new Array();
		for (var i = 0; i < iLoopLength; i++) {
			var jLoopLength = g[i].words.length;
			for (var j = 0; j < jLoopLength; j++) {
				arrGlossaryIndex.push(g[i].words[j].word);
			}
		}
		return arrGlossaryIndex;
	} catch (e) {
	}
}

// ------------------------------------------------------------------------------------------------
// 本文中のハイライトされた用語集にマウスオーバーしたときにツールチップで用語説明を表示
// ------------------------------------------------------------------------------------------------
function fncGlossaryToolTip(target) {
	try {

		var strGlossaryWordName = target.innerHTML; // NOTE: innerTextはDOM非標準

		// NOTE: var glossary = [{...}];
		//                      ^^^^^^^ (v.2では動的にjsonをロードするため[～]部分だけが必要)
		var g = eval(glossary.substring(15, glossary.length - 1));
		var iLoopLength = g.length;
		var arrGlossaryDesc = new Array();
		for (var i = 0; i < iLoopLength; i++) {
			var jLoopLength = g[i].words.length;
			for (var j = 0; j < jLoopLength; j++) {
				if (strGlossaryWordName == g[i].words[j].word) {
					target.title = g[i].words[j].desc.replace(/<br\/>/g, "\r\n").replace(/&quot;/g, "\"").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
					return;
				}
			}
		}
	} catch (e) {
	}
}

// ------------------------------------------------------------------------------------------------
// 用語集マークアップ
// ------------------------------------------------------------------------------------------------
function fncMarkupGlossary(element, arrGlossaryIndex, nMode) {
	try {

		// 子ノード数分繰り返し
		var iLoopLength = element.childNodes.length;
		for (var i = 0; i < iLoopLength; i++) {
			var child = element.childNodes[i];

			// タイトル部分は対象外
			if ((child.className) && (child.className.match(/h[1-6]/) != null)) {
				continue;
			}

			// パンくず部分は対象外
			if (child.id == "id_breadcrumbs") {
				continue;
			}

			// #textまでたどり着いたらマークアップ処理
			if (child.nodeType == 3) { // #text

				var strNodeValue = child.nodeValue;

				// NOTE: Safariにおいて検索画面からジャンプした際に表レイアウトが崩れる現象を回避
				// トリムした結果文字列が残らない場合はマーキング処理を実行しない
				var strNodeValueTemp = strNodeValue.replace(/\t| |\n/g, "");
				if (strNodeValueTemp == "") {
					continue;
				}

				// マーキング対象の有無
				var bIsMarkedup = false;

				// 用語集タイトル分ループ
				var jLoopLength = arrGlossaryIndex.length;
				for (var j = 0; j < jLoopLength; j++) {

					var bSearchAsWord = true;

					// 用語集タイトルの取り出し
					var strGlossaryWord = arrGlossaryIndex[j];

					// 用語集タイトルをエスケープ
					var regexpEscapeJson = /([()])/g;
					if (regexpEscapeJson.exec(strGlossaryWord) != null) {
						strGlossaryWord = strGlossaryWord.replace(regexpEscapeJson, "\\$1");

						// 括弧が含まれている場合は単語検索しない
						bSearchAsWord = false;
					}

					// マルチバイト言語の場合は単語検索しない
					if (nMode == 1) {
						bSearchAsWord = false;
					}

					// 合致するかどうかを判定
					if (bSearchAsWord) {
						var re = new RegExp("\\b(" + strGlossaryWord + ")\\b", ""); // 単語レベルで合致しているかどうか
					} else {
						var re = new RegExp("(" + strGlossaryWord + ")", "");
					}
					if (re.exec(strNodeValue) != null) {

						// テキスト値に<～>で囲まれた文字列が含まれると、innerHTMLで戻すときにタグとして認識され、囲まれた文字列が表示されなくなってしまう現象を回避
						// またマーキング用のタグ文字列と部分合致するキーワードが検索された場合にマーキング用タグまで文字列置換されてしまうことを防ぐ
						// マーキング開始タグ:	⁅(U+2045(Left Square Bracket With Quill))
						// マーキング終了タグ:	⁆(U+2046(Right Square Bracket With Quill))
						//strNodeValue = strNodeValue.replace(re, "⁅$1⁆");
						strNodeValue = strNodeValue.replace(re, function($1, $2, $3) {
							return  String.fromCharCode(0x2045) + $1.replace(/ /g, "___CAESAR___") + String.fromCharCode(0x2046); // NOTE: 用語集マーキングのネストを避けるため、用語集文字列内のスペースを一旦置換（\\bによりヒットしなくなる）
						});

						bIsMarkedup = true;

						// ページ内ではじめに見つけた箇所だけにマークアップ
						arrGlossaryIndex[j] = "___DONE___";
					}
				}
				if (bIsMarkedup) {

					// テキスト値に<～>で囲まれた文字列が含まれると、innerHTMLで戻すときにタグとして認識され、囲まれた文字列が表示されなくなってしまう現象を回避
					strNodeValue = strNodeValue.replace(/</g, "&lt;");
					strNodeValue = strNodeValue.replace(/>/g, "&gt;");
					strNodeValue = strNodeValue.replace(/___CAESAR___/g, " ");

					// マーキングタグを復元
					if (child.parentNode.nodeName.toLowerCase() == "a") {
						strNodeValue = strNodeValue.replace(/\u2045/g, "<span class=\"glossary\" onmouseover=\"fncGlossaryToolTip(this);\">");
						strNodeValue = strNodeValue.replace(/\u2046/g, "</span>");
					} else {
						strNodeValue = strNodeValue.replace(/\u2045/g, "<a href=\"#\" onclick=\"fncOpenGlossary(this.innerHTML);\" class=\"glossary\" onmouseover=\"fncGlossaryToolTip(this);\">");
						strNodeValue = strNodeValue.replace(/\u2046/g, "</a>");
					}

					// マーキング済みの文字列に差し替え
					var newNode = document.createElement("span");

					// NOTE:innerHTMLとすることでトリムが発生してしまう。結果先頭がスペースの#textは、単語の間のスペースが詰まってしまう
					if (strNodeValue.substring(0, 1) == " ") {
						newNode.innerHTML = "&nbsp;" + strNodeValue.substring(1);
					} else {
						newNode.innerHTML = strNodeValue;
					}
					element.replaceChild(newNode, child);
				}
			// <div><span><a>はさらに子ノードを処理
			} else {
				fncMarkupGlossary(child, arrGlossaryIndex, nMode);
			}
		}
	} catch (e) {
	}
}
function fncOpenGlossary(word) {
	try {
		var iWidth = 640;
		var iHeight = 480;
		var iLeft = (screen.width / 2) - (iWidth / 2);
		var iTop = (screen.height / 2) - (iHeight / 2);
		var wnd = window.open(
			"../frame_htmls/glossary.html?word=" + encodeURIComponent(word),
			"canon_sub_window",
			"directories=no,location=no,menubar=no,status=no,toolbar=no,resizable=yes,width=" + iWidth + ",top=" + iTop + ",left=" + iLeft + ",height=" + iHeight
		);
		wnd.focus();
	} catch (e) {
	}
}

// ------------------------------------------------------------------------------------------------
// 本文中の検索キーワードと合致する文字列をハイライトする
// ------------------------------------------------------------------------------------------------
function fncMarkupText(element, arrSearchText, strSearchOptionCaseSensitive, bWordMarking) {
	try {

		// 子ノード数分繰り返し
		var nElementChildLength = element.childNodes.length;
		for (var i = 0; i < nElementChildLength; i++) {
			var child = element.childNodes[i];

			// パンくず部分は対象外
			if (child.id == "id_breadcrumbs") {
				continue;
			}

			// #textまでたどり着いたらマークアップ処理
			if (child.nodeType == 3) { // #text

				var strNodeValue = child.nodeValue;

				// NOTE: Safariにおいて検索画面からジャンプした際に表レイアウトが崩れる現象を回避
				// トリムした結果文字列が残らない場合はマーキング処理を実行しない
				var strNodeValueTemp = strNodeValue.replace(/\t| |\n/g, "");
				if (strNodeValueTemp == "") {
					continue;
				}

				// マーキング対象の有無
				var bIsMarkedup = false;

				// 各検索条件文字列にマーキング
				// 10種類のカラーバリエーションを循環
				var nMarkerColor = 0;
				var nSearchTextLength = arrSearchText.length;
				for (var j = 0; j < nSearchTextLength; j++) {

					// 1桁数字を検索した場合にカラーバリエーションクラス名まで文字列置換されてしまうことを防ぐ
					// 0-9の代わりにU+2080(Subscript Zero)-U+2089(Subscript Nine)を使用する
					switch (nMarkerColor) {
						case 0:
							strMarkerColor = String.fromCharCode(0x2080); // "₀";
							break;
						case 1:
							strMarkerColor = String.fromCharCode(0x2081); // "₁";
							break;
						case 2:
							strMarkerColor = String.fromCharCode(0x2082); // "₂";
							break;
						case 3:
							strMarkerColor = String.fromCharCode(0x2083); // "₃";
							break;
						case 4:
							strMarkerColor = String.fromCharCode(0x2084); // "₄";
							break;
						case 5:
							strMarkerColor = String.fromCharCode(0x2085); // "₅";
							break;
						case 6:
							strMarkerColor = String.fromCharCode(0x2086); // "₆";
							break;
						case 7:
							strMarkerColor = String.fromCharCode(0x2087); // "₇";
							break;
						case 8:
							strMarkerColor = String.fromCharCode(0x2088); // "₈";
							break;
						case 9:
							strMarkerColor = String.fromCharCode(0x2089); // "₉";
							break;
					}

					// キーワードの取り出し
					var strSearchText = arrSearchText[j];

					// キーワードに合致するかどうかを判定
					var strSearchTextParam = "(" + strSearchText + ")";
					if (bWordMarking) {
						// 単語単位で検索された文字列をマーキング
						strSearchTextParam = "\\b(" + strSearchText + ")\\b";
					}
					var re = new RegExp(strSearchTextParam, "g" + strSearchOptionCaseSensitive);

					if (re.exec(strNodeValue) != null) {

						// テキスト値に<～>で囲まれた文字列が含まれると、innerHTMLで戻すときにタグとして認識され、囲まれた文字列が表示されなくなってしまう現象を回避
						// またマーキング用のタグ文字列と部分合致するキーワードが検索された場合にマーキング用タグまで文字列置換されてしまうことを防ぐ
						// マーキング開始タグの開始:	⁅(U+2045(Left Square Bracket With Quill))
						// マーキング開始タグの終了:	⁆(U+2046(Right Square Bracket With Quill))
						// マーキング終了タグ:			₎(U+208E(Subscript Right Parenthesis))
						//var strNodeValue = strNodeValue.replace(re, "⁅" + strMarkerColor + "⁆$1₎");
						var strNodeValue = strNodeValue.replace(re, String.fromCharCode(0x2045) + strMarkerColor + String.fromCharCode(0x2046) + "$1" + String.fromCharCode(0x208E));
						bIsMarkedup = true;
					}
					nMarkerColor++;
					if (nMarkerColor >= 10) {
						nMarkerColor = 0;
					}
				}
				if (bIsMarkedup) {

					// テキスト値に<～>で囲まれた文字列が含まれると、innerHTMLで戻すときにタグとして認識され、囲まれた文字列が表示されなくなってしまう現象を回避
					strNodeValue = strNodeValue.replace(/</g, "&lt;");
					strNodeValue = strNodeValue.replace(/>/g, "&gt;");

					// マーキングタグを復元
					strNodeValue = strNodeValue.replace(/\u2045/g, "<span id=\"id_hit\" class=\"hit hit_"); // ⁅
					strNodeValue = strNodeValue.replace(/\u2046/g, "\">"); // ⁆
					strNodeValue = strNodeValue.replace(/\u2080/g, "0"); // ₀
					strNodeValue = strNodeValue.replace(/\u2081/g, "1"); // ₁
					strNodeValue = strNodeValue.replace(/\u2082/g, "2"); // ₂
					strNodeValue = strNodeValue.replace(/\u2083/g, "3"); // ₃
					strNodeValue = strNodeValue.replace(/\u2084/g, "4"); // ₄
					strNodeValue = strNodeValue.replace(/\u2085/g, "5"); // ₅
					strNodeValue = strNodeValue.replace(/\u2086/g, "6"); // ₆
					strNodeValue = strNodeValue.replace(/\u2087/g, "7"); // ₇
					strNodeValue = strNodeValue.replace(/\u2088/g, "8"); // ₈
					strNodeValue = strNodeValue.replace(/\u2089/g, "9"); // ₉
					strNodeValue = strNodeValue.replace(/\u208E/g, "</span>"); // ₎

					// マーキング済みの文字列に差し替え
					var newNode = document.createElement("span");
					newNode.innerHTML = strNodeValue;
					element.replaceChild(newNode, child);
				}

			// <div><span><a>はさらに子ノードを処理
			} else {
				fncMarkupText(child, arrSearchText, strSearchOptionCaseSensitive, bWordMarking);
			}
		}
	} catch (e) {
	}
}

// 各ペインをウィンドウサイズに合わせてサイズ調整
window.onresize = fncOnResize;
function fncOnResize() {
	try {

		// ウィンドウサイズの取得
		var w = fncGetWindowWidth();
		var h = fncGetWindowHeight();

		// メインウィンドウ
		if (document.location.search.indexOf("?sub=yes") == -1) {
			document.getElementById("id_panel").style.height = h - 93 + "px";
			document.getElementById("id_toc").style.height = h - 194 + "px";
			document.getElementById("id_content").style.height = h - 93 + "px";

			if (!document.getElementById("id_res_bar_icon_toggle_panel")) {
				document.getElementById("id_content").style.width = w - 300 + "px";
			} else if (document.getElementById("id_res_bar_icon_toggle_panel").src.indexOf("hide") != -1) {
				document.getElementById("id_content").style.width = w - 300 + "px";
			} else {
				document.getElementById("id_content").style.width = w - 0 + "px";
			}

			document.getElementById("id_footer").style.top = h - 26 + "px";
			document.getElementById("id_footer").style.width = w + "px";

			// 検索結果
			if (document.getElementById("id_search_options").style.display.toLowerCase() == "none") {
				document.getElementById("id_search_results").style.height = h - 191 + "px";
			} else {
				document.getElementById("id_search_results").style.height = h - 401 + "px";
			}

			// チャプターコントロールがない場合はTOC表示エリアを調整
			if (!document.getElementById("id_chapter_select")) {
				document.getElementById("id_toc").style.top = "22px";
				document.getElementById("id_toc").style.height = h - 147 + "px";
			}

			// 目次つきトップページにおける目次領域を除く幅指定
			if (strWindowType == "HOME_TOC") {
				$(".home").css("width", document.getElementById("id_content").style.width);
			}

		// 別ウィンドウ
		} else {
			document.getElementById("id_content").style.height = h - 28 + "px";
			document.getElementById("id_close").style.top = h - 28 + "px";
		}
	} catch (e) {
	}
}

// カテゴリー目次[+][-]開閉処理
function fncSwitchTocWrapper() {
	fncSwitchToc(this);
}
function fncSwitchToc(eSrc) {
	try {

		var iTargetLevel = 999; // 初期値
		var strDisplay = "";

		// [-] -> [+]
		if (eSrc.childNodes[0].src.lastIndexOf("toc_sign_1") != -1) {
			eSrc.childNodes[0].src = "../frame_images/toc_sign_2.gif";
			strDisplay = "None";

		// [+] -> [-]
		} else if(eSrc.childNodes[0].src.lastIndexOf("toc_sign_2") != -1) {
			eSrc.childNodes[0].src = "../frame_images/toc_sign_1.gif";
			strDisplay = "Block";
		}

		// 目次項目ループ
		var objHns = document.getElementById("id_toc").childNodes;
		var nHnLength = objHns.length;
		for (var i = 0; i < nHnLength; i++) {

			// 表示切替開始位置を探索
			if (document.getElementById("id_toc").childNodes[i] === eSrc.parentNode) {

				// 処理対象レベルを取得（クリックされた[+][-]がh2ならば、表示切替対象はh3）
				iTargetLevel = parseInt(eSrc.parentNode.nodeName.substring(1)) + 1;
				continue;
			}

			// [-] -> [+]がクリックされた場合、下階層すべてを非表示にする
			if (strDisplay == "None") {
				if (parseInt(objHns[i].nodeName.substring(1)) >= iTargetLevel) {

					// 表示を切り替え
					objHns[i].style.display = strDisplay;

					// [-][+] -> [+]
					if (objHns[i].childNodes[0].nodeName.toLowerCase() == "a") {
						objHns[i].childNodes[0].childNodes[0].src = "../frame_images/toc_sign_2.gif";
					}
				} else {
					if (iTargetLevel != 999) {
						break;
					}
					continue;
				}

			// [+] -> [-]がクリックされた場合、下階層のみを表示する
			} else {
				if (parseInt(objHns[i].nodeName.substring(1)) == iTargetLevel) {

					// 表示を切り替え
					objHns[i].style.display = strDisplay;
					continue;
				} else {
					if (	(iTargetLevel != 999)
						&&	(iTargetLevel > parseInt(objHns[i].nodeName.substring(1)))
					) {
						break;
					}
					continue;
				}
			}
		}
	} catch (e) {
	}
}

// カテゴリー目次一括開閉処理
function fncOpenCloseAllToc(nMethod) {
	try {
		var objSign = document.getElementById("id_toc").getElementsByTagName("img");
		var nSignLength = objSign.length;
		for (var i = 0; i < nSignLength; i++) {
			if (objSign[i].src.lastIndexOf("toc_sign_" + nMethod + ".gif") != -1) {
				if (document.all) {
					objSign[i].click();
				} else {
					fncSwitchToc(objSign[i].parentNode);
				}
			}
		}
	} catch (e) {
	}
}

// 詳細開閉処理
function fncOpenCloseNextSibling(eSrc) {
	try {
		var objNextSibling;
		objNextSibling = eSrc.parentNode.nextSibling;
		while (objNextSibling.nodeType != 1) {
			objNextSibling = objNextSibling.nextSibling;
		}
		if (objNextSibling.style && objNextSibling.style.display.toLowerCase() != "block") {
			objNextSibling.style.display = "block";
			eSrc.innerHTML = fncGetResourceByResourceId("close_next_sibling");
			eSrc.className = "close_next_sibling";
		} else {
			objNextSibling.style.display = "none";
			eSrc.innerHTML = fncGetResourceByResourceId("open_next_sibling");
			eSrc.className = "open_next_sibling";
		}
	} catch (e) {
	}
}

// 詳細一括開閉処理
function fncOpenCloseAll(strMethod) {
	try {
		var objDivs = document.getElementsByTagName("div");
		var nDivLength = objDivs.length;
		for (var i = 0; i < nDivLength; i++) {
			if (objDivs[i].className == "invisible") {
				var objPreviousSibling;
				objPreviousSibling = objDivs[i].previousSibling;
				while (objPreviousSibling.nodeType != 1) {
					objPreviousSibling = objPreviousSibling.previousSibling;
				}
				if (strMethod == "open") {
					objDivs[i].style.display = "block";
					if	(	(objPreviousSibling.childNodes[0].className == "open_next_sibling")
						||	(objPreviousSibling.childNodes[0].className == "close_next_sibling")
						)
					{
						objPreviousSibling.childNodes[0].innerHTML = fncGetResourceByResourceId("close_next_sibling");
						objPreviousSibling.childNodes[0].className = "close_next_sibling";
					}
				} else {
					objDivs[i].style.display = "none";
					if	(	(objPreviousSibling.childNodes[0].className == "open_next_sibling")
						||	(objPreviousSibling.childNodes[0].className == "close_next_sibling")
						)
					{
						objPreviousSibling.childNodes[0].innerHTML = fncGetResourceByResourceId("open_next_sibling");
						objPreviousSibling.childNodes[0].className = "open_next_sibling";
					}
				}
			}
		}
	} catch (e) {
	}
}

window.onbeforeunload = function fncOnBeforeUnLoad() {
	try {

		// ----------------------------------------------------------------------------------------
		// クッキーに情報を保存
		// ----------------------------------------------------------------------------------------

		// タブ位置の記憶
		var strTabPosition = "0";
		if (document.getElementById("id_res_contents").className == "tab_active") {
			strTabPosition = "1";
		} else if (document.getElementById("id_res_search").className == "tab_active") {
			strTabPosition = "2";
		}
		fncSetCookie("TAB-POSITION", strTabPosition);

		// 目次のスクロール位置（チャプターごとに記憶）
		var strCurrentChapterId = "ALL";
		if (document.getElementById("id_chapter_select")){
			strCurrentChapterId = document.getElementById("id_chapter_select").current_chapter_id;
		}
		fncSetCookie("TOC-SCROLL-POSITION-TOP-" + strCurrentChapterId, document.getElementById("id_toc").scrollTop);
		fncSetCookie("TOC-SCROLL-POSITION-LEFT-" + strCurrentChapterId, document.getElementById("id_toc").scrollLeft);

		// 検索キーワードの記憶
		var strSearchKeyword = document.getElementById("id_search").value;
		fncSetCookie("SEARCH-KEYWORD", strSearchKeyword);

		// 検索オプションの記憶
		if (document.getElementById("id_search_options_search_scope_all").checked) {
			fncSetCookie("SEARCH-SCOPE", "ALL");
			fncSetCookie("SEARCH-CHAPTERS", "");
		} else if (document.getElementById("id_search_options_search_scope_chapter").checked) {
			fncSetCookie("SEARCH-SCOPE", "CHAPTER");

			// 検索対象チャプターの記憶
			var objChapterCheckboxes = document.getElementById("id_search_chapters").getElementsByTagName("input");
			var arrSelectedChapters = new Array();
			var iLoopLength = objChapterCheckboxes.length;
			for (var i = 0; i < iLoopLength; i++) {
				if (objChapterCheckboxes[i].checked) {
					arrSelectedChapters.push(objChapterCheckboxes[i].id);
				}
			}
			fncSetCookie("SEARCH-CHAPTERS", arrSelectedChapters.join(","));
		}
		if (document.getElementById("id_search_options_case").checked) {
			fncSetCookie("SEARCH-OPTIONS-CASE", "TRUE");
		} else {
			fncSetCookie("SEARCH-OPTIONS-CASE", "FALSE");
		}

		if (!document.getElementById("id_search_options_multibyte").checked) {
			fncSetCookie("SEARCH-OPTIONS-MULTIBYTE", "TRUE");
		} else {
			fncSetCookie("SEARCH-OPTIONS-MULTIBYTE", "FALSE");
		}

		// 検索オプション開閉状態の記憶
		if (document.getElementById("id_search_options")) {
			if (document.getElementById("id_search_options").style.display.toLowerCase() == "none") {
				fncSetCookie("SEARCH-OPTIONS", "CLOSE");
			} else if (document.getElementById("id_search_options").style.display.toLowerCase() == "block") {
				fncSetCookie("SEARCH-OPTIONS", "OPEN");
			}
		}

		// 検索オプション開閉状態の記憶
		if (document.getElementById("id_search_options")) {
			if (document.getElementById("id_search_options").style.display.toLowerCase() == "none") {
				fncSetCookie("SEARCH-OPTIONS", "CLOSE");
			} else if (document.getElementById("id_search_options").style.display.toLowerCase() == "block") {
				fncSetCookie("SEARCH-OPTIONS", "OPEN");
			}
		}

		// 左パネルの表示・非表示状態の記憶
		if (document.getElementById("id_res_bar_icon_toggle_panel")) {

			// 非表示
			if (document.getElementById("id_res_bar_icon_toggle_panel").src.indexOf("hide") != -1) {
				fncSetCookie("LEFT_PANE_VISIBILITY", "SHOW");
			} else {
				fncSetCookie("LEFT_PANE_VISIBILITY", "HIDE");
			}
		}
	} catch (e) {
	}
}
var strWindowType = "MAIN";