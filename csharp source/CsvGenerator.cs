using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using static System.Net.WebRequestMethods;
using Excel = Microsoft.Office.Interop.Excel;

namespace csharp
{
    public class CsvGenerator
    {
        private string 店舗コード = "taiho-kagu";
        private string outputDirectory = $@"{System.Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)}\CSVTOOL\{UtilityClass.GetTime()}";
        private static string ecPath = System.IO.Path.GetDirectoryName(Application.ExecutablePath) + @"\ec_csv_tool\";
        List<string> skuErrorList = new List<string>();

        //バッチの開始
        public void CsvGenerate()
        {
            System.IO.Directory.CreateDirectory($@"{outputDirectory}\");

            //DataTableを作成する
            var dtRakutenItem = NewDatatableRakutenItem();
            var dtRakutenCategory = NewDatatableRakutenCategory();
            var dtYahooItem = NewDatatableYahooItem();
            var dtYahooOption = NewDatatableYahooOption();
            var dtYahooAuctionItem = NewDatatableYahooAuctionItem();
            var dtYahooAuctionOption = NewDatatableYahooAuctionOption();


            //楽天、ヤフーショッピング、ヤフオク用のCSVを作成する
            var dtMasterItemMain = UtilityClass.GetDataTableFromExcelOfPureTable(ecPath + @"item.xlsm", "Main");
            var dtMasterItemSKU = UtilityClass.GetDataTableFromExcelOfPureTable(ecPath + @"item.xlsm", "SKU");
            var dtMasterItemIcon = UtilityClass.GetDataTableFromExcelOfPureTable(ecPath + @"item.xlsm", "Icon");
            for (int i = 0; i < dtMasterItemMain.Rows.Count; i++)
            {
                var dRow = dtMasterItemMain.Rows[i];

                //コントロールカラムがpの場合、商品情報を作成せずスキップする
                if (dRow["コントロールカラム"].ToString() == "p" || dRow["mycode"].ToString() == "")
                {
                    continue;
                }
                Console.WriteLine($"{i}/{dtMasterItemMain.Rows.Count} 商品コード : {dRow["mycode"].ToString()}");

                //テーブル、かご横、SP関連商品の文字列を作成する
                var generateTable = GenerateTable(dRow);
                var generateKagoyoko = GenerateKagoyoko(dRow);
                var generateSpKanrenShohin = GenerateSpKanrenShohin(dRow);

                //DataRowに行を追加しデータを入力する
                dtRakutenItem = GenerateCsvRakutenItem(dtRakutenItem, dRow, generateTable, generateKagoyoko, generateSpKanrenShohin, dtMasterItemSKU);
                dtRakutenCategory = GenerateCsvRakutenCategory(dtRakutenCategory, dRow);
                dtYahooItem = GenerateCsvYahooItem(dtYahooItem, dRow, generateTable, generateKagoyoko, generateSpKanrenShohin, dtMasterItemSKU);
                dtYahooOption = GenerateCsvYahooOption(dtYahooOption, dRow, dtMasterItemSKU);

                //ヤフオ用のCSVを作成する
                dtYahooAuctionItem = GenerateCsvYahooAuctionItem(dtYahooAuctionItem, dRow, dtMasterItemSKU, dtMasterItemIcon, dtYahooItem);
                dtYahooAuctionOption = GenerateCsvYahooAuctionOption(dtYahooAuctionOption, dRow, dtMasterItemSKU, dtYahooOption);
            }

            //DataTableをCSVに変換して書き出す
            if (skuErrorList.Count == 0)
            {
                Console.WriteLine("rakuten_item : " + dtRakutenItem.Rows.Count + "行作成されました。");
                Console.WriteLine("rakuten_category : " + dtRakutenCategory.Rows.Count + "行作成されました。");
                Console.WriteLine("yahoo_item : " + dtYahooItem.Rows.Count + "行作成されました。");
                Console.WriteLine("yahoo_option : " + dtYahooOption.Rows.Count + "行作成されました。");

                Console.WriteLine("yahoo_auction_item : " + dtYahooAuctionItem.Rows.Count + "行作成されました。");
                Console.WriteLine("yahoo_auction_option : " + dtYahooAuctionOption.Rows.Count + "行作成されました。");

                UtilityClass.DataTableToCsv(dtRakutenItem, $@"{outputDirectory}\rakuten_normal-item.csv", true);
                UtilityClass.DataTableToCsv(dtRakutenCategory, $@"{outputDirectory}\rakuten_item-cat.csv", true);
                UtilityClass.DataTableToCsv(dtYahooItem, $@"{outputDirectory}\yahoo_item.csv", true);
                UtilityClass.DataTableToCsv(dtYahooOption, $@"{outputDirectory}\yahoo_option.csv", true);
                UtilityClass.DataTableToCsv(dtYahooAuctionItem, $@"{outputDirectory}\yahoo_auction_item.csv", true);
                UtilityClass.DataTableToCsv(dtYahooAuctionOption, $@"{outputDirectory}\yahoo_auction_option.csv", true);
            }
            else
            {
                //エラーを表示する
                foreach (var errorItem in skuErrorList)
                {
                    Console.WriteLine($"SKUコードが見つかりませんでした：{errorItem}");
                }
                MessageBox.Show($"SKUコードが見つからない商品が{skuErrorList.Count}件ありました。\r\nエラーが出たため、CSVデータは出力されません。\r\nコンソール画面のエラー商品一覧をコピーしてご活用ください。");
                foreach (var errorItem in skuErrorList)
                {
                    MessageBox.Show($"SKUコードが見つかりませんでした：{errorItem}");
                }
            }
        }

        //商品情報テーブル、かご横情報、スマホ関連商品テーブルの生成
        public string GenerateTable(DataRow dtRowItemMaster)
        {
            //HTMLテーブルの読み込み
            var generateTable = UtilityClass.ReadTextFile(ecPath + @"templates\table_temp.html");

            //特徴
            if (1 < dtRowItemMaster["特徴_1"].ToString().Length)
            {
                generateTable = generateTable.Replace("XXX_parent_特徴", UtilityClass.ReadTextFile(ecPath + @"templates\table_temp_parent_00_特徴.html"));
                generateTable = generateTable.Replace("XXX_特徴_1", dtRowItemMaster["特徴_1"].ToString());
            }

            //商品サイズ
            for (int i = 1; i < 9; i++)
            {
                if (1 < dtRowItemMaster["商品サイズ_1a"].ToString().Length && i == 1)
                {
                    generateTable = generateTable.Replace("XXX_parent_商品サイズ", UtilityClass.ReadTextFile(ecPath + @"templates\table_temp_parent_01_商品サイズ.html"));
                }
                if (1 < dtRowItemMaster["商品サイズ_" + i + "a"].ToString().Length)
                {
                    generateTable = generateTable.Replace("XXX_child_商品サイズ", UtilityClass.ReadTextFile(ecPath + @"templates\table_temp_child_01_商品サイズ.html") + "XXX_child_商品サイズ");
                    generateTable = generateTable.Replace("XXX_商品サイズa", dtRowItemMaster["商品サイズ_" + i + "a"].ToString());
                    generateTable = generateTable.Replace("XXX_商品サイズb", dtRowItemMaster["商品サイズ_" + i + "b"].ToString());
                }
            }

            //梱包サイズ
            if (1 < dtRowItemMaster["梱包サイズ_1"].ToString().Length)
            {
                generateTable = generateTable.Replace("XXX_parent_梱包サイズ", UtilityClass.ReadTextFile(ecPath + @"templates\table_temp_parent_02_梱包サイズ.html"));
                generateTable = generateTable.Replace("XXX_梱包サイズ_1", dtRowItemMaster["梱包サイズ_1"].ToString());
            }

            //材質
            for (int i = 1; i < 7; i++)
            {
                if (i == 1)
                {
                    if (1 < dtRowItemMaster["材質_1"].ToString().Length)
                    {
                        generateTable = generateTable.Replace("XXX_parent_材質", UtilityClass.ReadTextFile(ecPath + @"templates\table_temp_parent_03_材質.html"));
                        generateTable = generateTable.Replace("XXX_材質_1", dtRowItemMaster["材質_1"].ToString());
                    }
                }
                else
                {
                    if (1 < dtRowItemMaster["材質_" + i + "a"].ToString().Length)
                    {
                        generateTable = generateTable.Replace("XXX_child_材質", UtilityClass.ReadTextFile(ecPath + @"templates\table_temp_child_03_材質.html") + "XXX_child_材質");
                        generateTable = generateTable.Replace("XXX_材質a", dtRowItemMaster["材質_" + i + "a"].ToString());
                        generateTable = generateTable.Replace("XXX_材質b", dtRowItemMaster["材質_" + i + "b"].ToString());
                    }
                }
            }

            //色
            if (1 < dtRowItemMaster["色_1"].ToString().Length)
            {
                generateTable = generateTable.Replace("XXX_parent_色", UtilityClass.ReadTextFile(ecPath + @"templates\table_temp_parent_04_色.html"));
                generateTable = generateTable.Replace("XXX_色_1", dtRowItemMaster["色_1"].ToString());
            }

            //仕様
            for (int i = 1; i < 7; i++)
            {
                if (i == 1)
                {
                    if (1 < dtRowItemMaster["仕様_1"].ToString().Length)
                    {
                        generateTable = generateTable.Replace("XXX_parent_仕様", UtilityClass.ReadTextFile(ecPath + @"templates\table_temp_parent_05_仕様.html"));
                        generateTable = generateTable.Replace("XXX_仕様_1", dtRowItemMaster["仕様_1"].ToString());
                    }
                }
                else
                {
                    if (1 < dtRowItemMaster["仕様_" + i + "a"].ToString().Length)
                    {
                        generateTable = generateTable.Replace("XXX_child_仕様", UtilityClass.ReadTextFile(ecPath + @"templates\table_temp_child_05_仕様.html") + "XXX_child_仕様");
                        generateTable = generateTable.Replace("XXX_仕様a", dtRowItemMaster["仕様_" + i + "a"].ToString());
                        generateTable = generateTable.Replace("XXX_仕様b", dtRowItemMaster["仕様_" + i + "b"].ToString());
                    }
                }
            }

            //お届け状態
            if (1 < dtRowItemMaster["お届け状態_1"].ToString().Length)
            {
                generateTable = generateTable.Replace("XXX_parent_お届け状態", UtilityClass.ReadTextFile(ecPath + @"templates\table_temp_parent_06_お届け状態.html"));
                generateTable = generateTable.Replace("XXX_お届け状態_1", dtRowItemMaster["お届け状態_1"].ToString());
            }

            //注意事項
            if (dtRowItemMaster["注意事項"].ToString() != "-")
            {
                generateTable = generateTable.Replace("XXX_parent_注意事項", UtilityClass.ReadTextFile(ecPath + @"templates\table_temp_parent_07_注意事項.html"));
                generateTable = generateTable.Replace("XXX_注意事項", UtilityClass.ReadTextFile(ecPath + @"templates\【注意事項PC】" + dtRowItemMaster["注意事項"].ToString() + ".html"));
            }
            else
            {
                MessageBox.Show("エラー：注意事項が入力されていません。");
            }

            //シリーズ商品
            if (1 < dtRowItemMaster["シリーズURL"].ToString().Length)
            {
                generateTable = generateTable.Replace("XXX_parent_関連商品", UtilityClass.ReadTextFile(ecPath + @"templates\table_temp_parent_08_関連商品.html"));
                string シリーズURL = dtRowItemMaster[$"シリーズURL"].ToString().Replace("/", "");
                generateTable = generateTable.Replace("XXX_child_シリーズ商品", UtilityClass.ReadTextFile(ecPath + @"templates\table_temp_child_08_シリーズ商品.html") + "XXX_child_シリーズ商品");
                generateTable = generateTable.Replace("XXX_シリーズURL", シリーズURL);
            }
            //関連商品
            if (1 < dtRowItemMaster["関連商品_1a"].ToString().Length)
            {
                generateTable = generateTable.Replace("XXX_parent_関連商品", UtilityClass.ReadTextFile(ecPath + @"templates\table_temp_parent_08_関連商品.html"));
                for (int i = 1; i <= 15; i++)
                {
                    if (1 < dtRowItemMaster[$"関連商品_{i}a"].ToString().Length)
                    {
                        generateTable = generateTable.Replace("XXX_child_関連商品", UtilityClass.ReadTextFile(ecPath + @"templates\table_temp_child_08_関連商品.html") + "XXX_child_関連商品");
                        string 関連商品_名前 = dtRowItemMaster[$"関連商品_{i}a"].ToString();
                        string 関連商品_商品番号 = dtRowItemMaster[$"関連商品_{i}b"].ToString();
                        //関連商品_名前 = "<a href=\"./商品URL/" + 関連商品_商品番号 + ".itempage\" target=\"parent\">" + 関連商品_名前 + "</A>";
                        関連商品_商品番号 = "<a href=\"./商品URL/" + 関連商品_商品番号 + ".itempage\" target=\"parent\"><IMG src=\"./サムネイル/" + 関連商品_商品番号 + ".thum\" border=\"1\" width=\"120\"></A>";
                        generateTable = generateTable.Replace("XXX_関連商品a", 関連商品_商品番号);
                        generateTable = generateTable.Replace("XXX_関連商品b", 関連商品_名前);
                    }
                }
            }
            //シリーズ商品関連商品改行
            if (1 < dtRowItemMaster["シリーズURL"].ToString().Length && 1 < dtRowItemMaster["関連商品_1a"].ToString().Length)
            {
                generateTable = generateTable.Replace("XXX_改行", "\r\n<br>\r\n<br>\r\n");
            }
            else
            {
                generateTable = generateTable.Replace("XXX_改行", "");
            }


            //説明マーク
            if (1 < dtRowItemMaster["説明マーク_1"].ToString().Length)
            {
                generateTable = generateTable.Replace("XXX_parent_説明マーク", UtilityClass.ReadTextFile(ecPath + @"templates\table_temp_parent_09_説明マーク.html"));
                string 説明 = dtRowItemMaster["説明マーク_1"].ToString();
                string[] 説明_args = 説明.Split(' ');
                string 説明_結合 = "";
                for (int i = 0; i < 説明_args.Length; i++)
                {
                    if (説明_args[i] != " ")
                    {
                        説明_結合 = 説明_結合 + "<IMG src=\"./アイコン/" + 説明_args[i] + ".gif\" width=\"56\" height=\"56\" border=\"0\">";

                    }
                }
                generateTable = generateTable.Replace("XXX_説明マーク_1", 説明_結合);
            }

            //不要だった文字の削除
            generateTable = generateTable.Replace("XXX_parent_特徴", "");
            generateTable = generateTable.Replace("XXX_parent_商品サイズ", "");
            generateTable = generateTable.Replace("XXX_parent_梱包サイズ", "");
            generateTable = generateTable.Replace("XXX_parent_材質", "");
            generateTable = generateTable.Replace("XXX_parent_色", "");
            generateTable = generateTable.Replace("XXX_parent_仕様", "");
            generateTable = generateTable.Replace("XXX_parent_お届け状態", "");
            generateTable = generateTable.Replace("XXX_parent_注意事項", "");
            generateTable = generateTable.Replace("XXX_parent_関連商品", "");
            generateTable = generateTable.Replace("XXX_parent_シリーズ商品", "");
            generateTable = generateTable.Replace("XXX_parent_説明マーク", "");
            generateTable = generateTable.Replace("XXX_child_商品サイズ", "");
            generateTable = generateTable.Replace("XXX_child_材質", "");
            generateTable = generateTable.Replace("XXX_child_仕様", "");
            generateTable = generateTable.Replace("XXX_child_関連商品", "");
            generateTable = generateTable.Replace("XXX_child_シリーズ商品", "");

            return generateTable;
        }
        public string GenerateKagoyoko(DataRow dtRowItemMaster)
        {
            var generateKagoyoko = "";

            //特徴
            if (1 < dtRowItemMaster["特徴_1"].ToString().Length)
            {
                generateKagoyoko = generateKagoyoko + dtRowItemMaster["特徴_1"].ToString() + "\r\n";
            }

            //商品コード
            generateKagoyoko = generateKagoyoko + "\r\n【商品コード：" + dtRowItemMaster["mycode"].ToString() + "】" + "\r\n";

            //商品名
            if (false)
            {
                //商品名_正式表記の出力停止
                generateKagoyoko = generateKagoyoko + "\r\n" + "【商品名】\r\n " + dtRowItemMaster["商品名_正式表記"].ToString() + "\r\n";
            }


            //商品サイズ
            for (int i = 1; i < 9; i++)
            {
                if (1 < dtRowItemMaster["商品サイズ_1a"].ToString().Length && i == 1)
                {
                    //商品サイズ_1aの場合
                    generateKagoyoko = generateKagoyoko + "\r\n【商品サイズ】\r\n";
                }
                if (1 < dtRowItemMaster["商品サイズ_" + i + "a"].ToString().Length)
                {
                    //商品サイズ_2a～8a、商品サイズ_2b～8bの場合
                    generateKagoyoko = generateKagoyoko + dtRowItemMaster["商品サイズ_" + i + "a"].ToString() + "：" + dtRowItemMaster["商品サイズ_" + i + "b"].ToString() + "\r\n";
                }
            }

            //梱包サイズ
            if (1 < dtRowItemMaster["梱包サイズ_1"].ToString().Length)
            {
                generateKagoyoko = generateKagoyoko + "\r\n【梱包サイズ】\r\n" + dtRowItemMaster["梱包サイズ_1"].ToString() + "\r\n";
            }

            //材質
            if (1 < dtRowItemMaster["材質_1"].ToString().Length)
            {
                generateKagoyoko = generateKagoyoko + "\r\n【材質】\r\n" + dtRowItemMaster["材質_1"].ToString() + "\r\n";
            }

            //色
            if (1 < dtRowItemMaster["色_1"].ToString().Length)
            {
                generateKagoyoko = generateKagoyoko + "\r\n【色】\r\n" + dtRowItemMaster["色_1"].ToString() + "\r\n";
            }

            //仕様
            if (1 < dtRowItemMaster["仕様_1"].ToString().Length)
            {
                generateKagoyoko = generateKagoyoko + "\r\n【仕様】\r\n" + dtRowItemMaster["仕様_1"].ToString() + "\r\n";
            }

            //お届け状態
            if (1 < dtRowItemMaster["お届け状態_1"].ToString().Length)
            {
                generateKagoyoko = generateKagoyoko + "\r\n【お届け状況】\r\n" + dtRowItemMaster["お届け状態_1"].ToString() + "\r\n";
            }

            //末尾に送料無料の文言を付け加える
            generateKagoyoko = generateKagoyoko + "\r\n【送料】\r\n送料無料（北海道、東北、沖縄及び離島は、別途送料が掛かります）" + "\r\n";

            return generateKagoyoko;
        }
        public string GenerateSpKanrenShohin(DataRow dtRowItemMaster)
        {
            var generateSpKanrenShohin = "";
            if (1 < dtRowItemMaster["シリーズURL"].ToString().Length || 1 < dtRowItemMaster["関連商品_1a"].ToString().Length)
            {
                generateSpKanrenShohin = UtilityClass.ReadTextFile(ecPath + @"templates\sp_temp_kanrenshohin.html");
            }

            //シリーズ商品
            if (1 < dtRowItemMaster["シリーズURL"].ToString().Length)
            {
                string シリーズURL = dtRowItemMaster[$"シリーズURL"].ToString().Replace("/", "");
                generateSpKanrenShohin = generateSpKanrenShohin.Replace("XXX_シリーズURL", シリーズURL);
                generateSpKanrenShohin = generateSpKanrenShohin.Replace("<!--シリーズ商品-->", "");
                generateSpKanrenShohin = generateSpKanrenShohin.Replace("<!--/シリーズ商品-->", "");
            }
            else
            {
                generateSpKanrenShohin = System.Text.RegularExpressions.Regex.Replace(generateSpKanrenShohin, "<!--シリーズ商品-->[\\s\\S]*?<!--/シリーズ商品-->", "");
            }

            //関連商品
            if (1 < dtRowItemMaster["関連商品_1a"].ToString().Length)
            {
                for (int i = 1; i <= 15; i++)
                {
                    if (1 < dtRowItemMaster[$"関連商品_{i}a"].ToString().Length)
                    {
                        //テーブル内の共通ソース
                        string 関連商品_名前 = dtRowItemMaster[$"関連商品_{i}a"].ToString();
                        string 関連商品_商品番号 = dtRowItemMaster[$"関連商品_{i}b"].ToString();
                        string 関連商品_画像リンク = "<TD align=\"center\" valign=\"bottom\"><a href=\"./商品URL/" + 関連商品_商品番号 + ".itempage\" target=\"parent\"><img src=\"./サムネイル/" + 関連商品_商品番号 + ".thum\" width=\"100%\"></A></TD>";
                        string 関連商品_テキスト = "<TD align=\"center\" valign=\"top\"><font size=\"-1\">" + 関連商品_名前 + "</font></TD>";

                        //1商品目～3商品目の3商品は1行目に表示する
                        if (i == 1)
                        {
                            generateSpKanrenShohin = generateSpKanrenShohin.Replace($"TR1行", "TR");
                        }
                        if (1 <= i && i <= 3)
                        {
                            var 行数 = 1;
                            generateSpKanrenShohin = generateSpKanrenShohin.Replace($"XXX_child_関連商品SP用top{行数}行", 関連商品_画像リンク + $"\r\nXXX_child_関連商品SP用top{行数}行");
                            generateSpKanrenShohin = generateSpKanrenShohin.Replace($"XXX_child_関連商品SP用bottom{行数}行", 関連商品_テキスト + $"\r\nXXX_child_関連商品SP用bottom{行数}行");
                        }

                        //4商品目～6商品目の3商品は2行目に表示する
                        if (i == 4)
                        {
                            generateSpKanrenShohin = generateSpKanrenShohin.Replace("TR2行", "TR");
                        }
                        if (4 <= i && i <= 6)
                        {
                            var 行数 = 2;
                            generateSpKanrenShohin = generateSpKanrenShohin.Replace($"XXX_child_関連商品SP用top{行数}行", 関連商品_画像リンク + $"\r\nXXX_child_関連商品SP用top{行数}行");
                            generateSpKanrenShohin = generateSpKanrenShohin.Replace($"XXX_child_関連商品SP用bottom{行数}行", 関連商品_テキスト + $"\r\nXXX_child_関連商品SP用bottom{行数}行");
                        }

                        //7商品目～9商品目の3商品は3行目に表示する
                        if (i == 7)
                        {
                            generateSpKanrenShohin = generateSpKanrenShohin.Replace("TR3行", "TR");
                        }
                        if (7 <= i && i <= 9)
                        {
                            var 行数 = 3;
                            generateSpKanrenShohin = generateSpKanrenShohin.Replace($"XXX_child_関連商品SP用top{行数}行", 関連商品_画像リンク + $"\r\nXXX_child_関連商品SP用top{行数}行");
                            generateSpKanrenShohin = generateSpKanrenShohin.Replace($"XXX_child_関連商品SP用bottom{行数}行", 関連商品_テキスト + $"\r\nXXX_child_関連商品SP用bottom{行数}行");
                        }

                        //10商品目～12商品目の3商品は3行目に表示する
                        if (i == 10)
                        {
                            generateSpKanrenShohin = generateSpKanrenShohin.Replace("TR4行", "TR");
                        }
                        if (10 <= i && i <= 12)
                        {
                            var 行数 = 4;
                            generateSpKanrenShohin = generateSpKanrenShohin.Replace($"XXX_child_関連商品SP用top{行数}行", 関連商品_画像リンク + $"\r\nXXX_child_関連商品SP用top{行数}行");
                            generateSpKanrenShohin = generateSpKanrenShohin.Replace($"XXX_child_関連商品SP用bottom{行数}行", 関連商品_テキスト + $"\r\nXXX_child_関連商品SP用bottom{行数}行");
                        }

                        //13商品目～15商品目の3商品は3行目に表示する
                        if (i == 13)
                        {
                            generateSpKanrenShohin = generateSpKanrenShohin.Replace("TR5行", "TR");
                        }
                        if (13 <= i && i <= 15)
                        {
                            var 行数 = 5;
                            generateSpKanrenShohin = generateSpKanrenShohin.Replace($"XXX_child_関連商品SP用top{行数}行", 関連商品_画像リンク + $"\r\nXXX_child_関連商品SP用top{行数}行");
                            generateSpKanrenShohin = generateSpKanrenShohin.Replace($"XXX_child_関連商品SP用bottom{行数}行", 関連商品_テキスト + $"\r\nXXX_child_関連商品SP用bottom{行数}行");
                        }
                    }
                }

                generateSpKanrenShohin = generateSpKanrenShohin.Replace("<TR1行>", "");
                generateSpKanrenShohin = generateSpKanrenShohin.Replace("</TR1行>", "");
                generateSpKanrenShohin = generateSpKanrenShohin.Replace("XXX_child_関連商品SP用top1行", "");
                generateSpKanrenShohin = generateSpKanrenShohin.Replace("XXX_child_関連商品SP用bottom1行", "");
                generateSpKanrenShohin = generateSpKanrenShohin.Replace("<TR2行>", "");
                generateSpKanrenShohin = generateSpKanrenShohin.Replace("</TR2行>", "");
                generateSpKanrenShohin = generateSpKanrenShohin.Replace("XXX_child_関連商品SP用top2行", "");
                generateSpKanrenShohin = generateSpKanrenShohin.Replace("XXX_child_関連商品SP用bottom2行", "");
                generateSpKanrenShohin = generateSpKanrenShohin.Replace("<TR3行>", "");
                generateSpKanrenShohin = generateSpKanrenShohin.Replace("</TR3行>", "");
                generateSpKanrenShohin = generateSpKanrenShohin.Replace("XXX_child_関連商品SP用top3行", "");
                generateSpKanrenShohin = generateSpKanrenShohin.Replace("XXX_child_関連商品SP用bottom3行", "");
                generateSpKanrenShohin = generateSpKanrenShohin.Replace("<TR4行>", "");
                generateSpKanrenShohin = generateSpKanrenShohin.Replace("</TR4行>", "");
                generateSpKanrenShohin = generateSpKanrenShohin.Replace("XXX_child_関連商品SP用top4行", "");
                generateSpKanrenShohin = generateSpKanrenShohin.Replace("XXX_child_関連商品SP用bottom4行", "");
                generateSpKanrenShohin = generateSpKanrenShohin.Replace("<TR5行>", "");
                generateSpKanrenShohin = generateSpKanrenShohin.Replace("</TR5行>", "");
                generateSpKanrenShohin = generateSpKanrenShohin.Replace("XXX_child_関連商品SP用top5行", "");
                generateSpKanrenShohin = generateSpKanrenShohin.Replace("XXX_child_関連商品SP用bottom5行", "");

                //置換
                generateSpKanrenShohin = System.Text.RegularExpressions.Regex.Replace(generateSpKanrenShohin, "^$\r\n", "");
                generateSpKanrenShohin = System.Text.RegularExpressions.Regex.Replace(generateSpKanrenShohin, "^$\r", "");
                generateSpKanrenShohin = System.Text.RegularExpressions.Regex.Replace(generateSpKanrenShohin, "^$\n", "");
                generateSpKanrenShohin = generateSpKanrenShohin.Replace("<!--関連商品-->", "");
                generateSpKanrenShohin = generateSpKanrenShohin.Replace("<!--/関連商品-->", "");
            }
            else
            {
                generateSpKanrenShohin = System.Text.RegularExpressions.Regex.Replace(generateSpKanrenShohin, "<!--関連商品-->[\\s\\S]*?<!--/関連商品-->", "");
            }

            return generateSpKanrenShohin;
        }


        //空のDataTableに各モールのCSV用データを追加する
        public DataTable GenerateCsvRakutenItem(DataTable dtRakutenItem, DataRow dtRowItemMasterMain, string generateTable, string generateKagoyoko, string generateSpKanrenShohin, DataTable dtRowItemMasterSKU)
        {
            var 商品コード = dtRowItemMasterMain["mycode"].ToString();
            var rowsSKU = dtRowItemMasterSKU.Select($"商品コード = '{商品コード}'");

            var is別途送料地域 = true;
            if (dtRowItemMasterMain["R_別途送料地域選択肢"].ToString() == "" || dtRowItemMasterMain["R_別途送料地域選択肢"].ToString() == "-")
            {
                is別途送料地域 = false;
            }
            var is配達オプション = true;
            if (dtRowItemMasterMain["R_配達オプション選択肢"].ToString() == "" || dtRowItemMasterMain["R_配達オプション選択肢"].ToString() == "-")
            {
                is配達オプション = false;
            }

            {
                ////////////////////////////////////////////////////////////////////////////////////////////
                // 商品レベル
                ////////////////////////////////////////////////////////////////////////////////////////////
                var row商品レベル = dtRakutenItem.NewRow();
                row商品レベル["商品管理番号（商品URL）"] = 商品コード;
                row商品レベル["商品番号"] = 商品コード;
                row商品レベル["商品名"] = dtRowItemMasterMain["R_商品名"].ToString();
                row商品レベル["倉庫指定"] = "0";
                row商品レベル["サーチ表示"] = "1";
                row商品レベル["消費税"] = "1";
                row商品レベル["消費税率"] = "0.1";
                row商品レベル["販売期間指定（開始日時）"] = "";
                row商品レベル["販売期間指定（終了日時）"] = "";
                row商品レベル["ポイント変倍率"] = "";
                row商品レベル["ポイント変倍率適用期間（開始日時）"] = "";
                row商品レベル["ポイント変倍率適用期間（終了日時）"] = "";
                row商品レベル["注文ボタン"] = "1";
                row商品レベル["予約商品発売日"] = "";
                row商品レベル["商品問い合わせボタン"] = "1";
                row商品レベル["闇市パスワード"] = "";
                row商品レベル["在庫表示"] = "0";
                row商品レベル["代引料"] = "0";
                row商品レベル["ジャンルID"] = dtRowItemMasterMain["R_ジャンルID"].ToString();
                if (dtRowItemMasterMain["非製品属性タグID"].ToString() == "-")
                {
                    row商品レベル["非製品属性タグID"] = "";
                }
                else
                {
                    row商品レベル["非製品属性タグID"] = dtRowItemMasterMain["非製品属性タグID"].ToString();
                }

                row商品レベル["キャッチコピー"] = dtRowItemMasterMain["R_キャッチコピー"].ToString();

                //PC商品説明文の作成
                var PC用商品説明文 = "";
                if (generateTable.Length > 1)
                {
                    PC用商品説明文 = generateTable.ToString();
                    PC用商品説明文 = PC用商品説明文.Replace("./説明用/", "https://image.rakuten.co.jp/" + 店舗コード + "/cabinet/" + dtRowItemMasterMain["画像パス_楽天"].ToString() + "/");
                    PC用商品説明文 = PC用商品説明文.Replace("./サムネイル/", "https://image.rakuten.co.jp/" + 店舗コード + "/cabinet/" + dtRowItemMasterMain["画像パス_楽天"].ToString() + "/");
                    PC用商品説明文 = PC用商品説明文.Replace(".thum", ".jpg");
                    PC用商品説明文 = PC用商品説明文.Replace(".itempage", "/");
                    PC用商品説明文 = PC用商品説明文.Replace("./商品URL/", "https://item.rakuten.co.jp/" + 店舗コード + "/");
                    PC用商品説明文 = PC用商品説明文.Replace("./アイコン/", "https://www.rakuten.ne.jp/gold/" + 店舗コード + "/icon/");
                    PC用商品説明文 = PC用商品説明文.Replace("./特設ページ/", "https://www.rakuten.ne.jp/gold/" + 店舗コード + "/haitatu/");
                    PC用商品説明文 = PC用商品説明文.Replace("./注意事項/", "https://www.rakuten.ne.jp/gold/" + 店舗コード + "/haitatu/");
                    PC用商品説明文 = PC用商品説明文.Replace("./ページ/", "https://image.rakuten.co.jp/" + 店舗コード + "/cabinet/page/");
                    PC用商品説明文 = PC用商品説明文.Replace("./イメージス/", "https://www.rakuten.ne.jp/gold/" + 店舗コード + "/images/");
                    PC用商品説明文 = PC用商品説明文.Replace("./カテゴリ/", "https://item.rakuten.co.jp/" + 店舗コード + "/c/");
                    PC用商品説明文 = System.Text.RegularExpressions.Regex.Replace(PC用商品説明文, "<!--関連商品-->", "");
                    PC用商品説明文 = System.Text.RegularExpressions.Regex.Replace(PC用商品説明文, "<!--/関連商品-->", "");
                    PC用商品説明文 = System.Text.RegularExpressions.Regex.Replace(PC用商品説明文, "<!--Yahoo-->[\\s\\S]*?<!--/Yahoo-->", "");
                    PC用商品説明文 = System.Text.RegularExpressions.Regex.Replace(PC用商品説明文, "<!--楽天-->", "");
                    PC用商品説明文 = System.Text.RegularExpressions.Regex.Replace(PC用商品説明文, "<!--/楽天-->", "");
                }
                row商品レベル["PC用商品説明文"] = PC用商品説明文;
                UtilityClass.OutputHtml(outputDirectory + @"\" + "html_楽天_PC用商品説明文", "<br><hr><br>" + row商品レベル["PC用商品説明文"].ToString());

                //画像説明
                var 画像説明 = "<table border=\"0\" cellspacing=\"0\" cellpadding=\"2\"><tr><td>\r\n" + dtRowItemMasterMain["画像説明"].ToString() + "\r\n</td></tr></table>";

                //スマートフォン用商品説明文
                {
                    var load_spgentei_rakuten = UtilityClass.ReadTextFile(ecPath + @"templates\load_spgentei_rakuten.html");
                    var 注意事項SP = UtilityClass.ReadTextFile(ecPath + @"templates\【注意事項SP】" + dtRowItemMasterMain["注意事項"].ToString() + ".html");
                    var スマートフォン用商品説明文 = generateKagoyoko.Replace("\r\n", "<br>\r\n") + "<br>\r\n" + load_spgentei_rakuten + "<br>\r\n" + 画像説明 + "<br>\r\n" + 注意事項SP + "<br>\r\n<br>\r\n" + generateSpKanrenShohin;
                    スマートフォン用商品説明文 = スマートフォン用商品説明文.Replace("./説明用/", "https://image.rakuten.co.jp/" + 店舗コード + "/cabinet/" + dtRowItemMasterMain["画像パス_楽天"].ToString() + "/");
                    スマートフォン用商品説明文 = スマートフォン用商品説明文.Replace("./サムネイル/", "https://image.rakuten.co.jp/" + 店舗コード + "/cabinet/" + dtRowItemMasterMain["画像パス_楽天"].ToString() + "/");
                    スマートフォン用商品説明文 = スマートフォン用商品説明文.Replace(".thum", ".jpg");
                    スマートフォン用商品説明文 = スマートフォン用商品説明文.Replace(".itempage", "/");
                    スマートフォン用商品説明文 = スマートフォン用商品説明文.Replace("./商品URL/", "https://item.rakuten.co.jp/" + 店舗コード + "/");
                    スマートフォン用商品説明文 = スマートフォン用商品説明文.Replace("./アイコン/", "https://www.rakuten.ne.jp/gold/" + 店舗コード + "/icon/");
                    スマートフォン用商品説明文 = スマートフォン用商品説明文.Replace("./注意事項/", "https://www.rakuten.ne.jp/gold/" + 店舗コード + "/haitatu/");
                    スマートフォン用商品説明文 = スマートフォン用商品説明文.Replace("./ページ/", "https://image.rakuten.co.jp/" + 店舗コード + "/cabinet/page/");
                    スマートフォン用商品説明文 = スマートフォン用商品説明文.Replace("./イメージス/", "https://www.rakuten.ne.jp/gold/" + 店舗コード + "/images/");
                    スマートフォン用商品説明文 = スマートフォン用商品説明文.Replace("./特設ページ/", "https://www.rakuten.ne.jp/gold/" + 店舗コード + "/haitatu/");
                    スマートフォン用商品説明文 = スマートフォン用商品説明文.Replace("./カテゴリ/", "https://item.rakuten.co.jp/" + 店舗コード + "/c/");
                    スマートフォン用商品説明文 = スマートフォン用商品説明文.Replace("&nbsp;　", "");
                    row商品レベル["スマートフォン用商品説明文"] = スマートフォン用商品説明文;
                    UtilityClass.OutputHtml(outputDirectory + @"\" + "html_楽天_スマートフォン用商品説明文", "<br><hr><br>" + row商品レベル["スマートフォン用商品説明文"].ToString());
                }

                //PC販売説明文
                {
                    var PC販売説明文 = 画像説明;
                    PC販売説明文 = PC販売説明文.Replace("./説明用/", "https://image.rakuten.co.jp/" + 店舗コード + "/cabinet/" + dtRowItemMasterMain["画像パス_楽天"].ToString() + "/");
                    PC販売説明文 = PC販売説明文.Replace("./サムネイル/", "https://image.rakuten.co.jp/" + 店舗コード + "/cabinet/" + dtRowItemMasterMain["画像パス_楽天"].ToString() + "/");
                    PC販売説明文 = PC販売説明文.Replace(".thum", ".jpg");
                    PC販売説明文 = PC販売説明文.Replace(".itempage", "/");
                    PC販売説明文 = PC販売説明文.Replace("./商品URL/", "https://item.rakuten.co.jp/" + 店舗コード + "/");
                    PC販売説明文 = PC販売説明文.Replace("./アイコン/", "https://www.rakuten.ne.jp/gold/" + 店舗コード + "/icon/");
                    PC販売説明文 = PC販売説明文.Replace("./注意事項/", "https://www.rakuten.ne.jp/gold/" + 店舗コード + "/haitatu/");
                    PC販売説明文 = PC販売説明文.Replace("./ページ/", "https://image.rakuten.co.jp/" + 店舗コード + "/cabinet/page/");
                    PC販売説明文 = PC販売説明文.Replace("./イメージス/", "https://www.rakuten.ne.jp/gold/" + 店舗コード + "/images/");
                    PC販売説明文 = PC販売説明文.Replace("./カテゴリ/", "https://item.rakuten.co.jp/" + 店舗コード + "/c/");
                    row商品レベル["PC用販売説明文"] = PC販売説明文;
                    UtilityClass.OutputHtml(outputDirectory + @"\" + "html_楽天_PC用販売説明文", "<br><hr><br>" + row商品レベル["PC用販売説明文"].ToString());
                }

                //サムネイル
                var listサムネイル = UtilityClass.StaticGetFindcodeList一行(dtRowItemMasterMain["画像説明"].ToString(), "<IMG SRC=\"./説明用/(?<findcode>.*?)\" border=");
                var サムネイル枚数 = listサムネイル.Count;
                if (1 <= サムネイル枚数)
                {
                    var 画像パス = $"/{dtRowItemMasterMain["画像パス_楽天"].ToString()}/{listサムネイル[0]}";
                    row商品レベル["商品画像タイプ1"] = "cabinet";
                    row商品レベル["商品画像パス1"] = 画像パス;
                    row商品レベル["商品画像名（ALT）1"] = "";
                }
                if (2 <= サムネイル枚数)
                {
                    var 画像パス = $"/{dtRowItemMasterMain["画像パス_楽天"].ToString()}/{listサムネイル[1]}";
                    row商品レベル["商品画像タイプ2"] = "cabinet";
                    row商品レベル["商品画像パス2"] = 画像パス;
                    row商品レベル["商品画像名（ALT）2"] = "";
                }
                if (3 <= サムネイル枚数)
                {
                    var 画像パス = $"/{dtRowItemMasterMain["画像パス_楽天"].ToString()}/{listサムネイル[2]}";
                    row商品レベル["商品画像タイプ3"] = "cabinet";
                    row商品レベル["商品画像パス3"] = 画像パス;
                    row商品レベル["商品画像名（ALT）3"] = "";
                }
                if (4 <= サムネイル枚数)
                {
                    var 画像パス = $"/{dtRowItemMasterMain["画像パス_楽天"].ToString()}/{listサムネイル[3]}";
                    row商品レベル["商品画像タイプ4"] = "cabinet";
                    row商品レベル["商品画像パス4"] = 画像パス;
                    row商品レベル["商品画像名（ALT）4"] = "";
                }
                if (5 <= サムネイル枚数)
                {
                    var 画像パス = $"/{dtRowItemMasterMain["画像パス_楽天"].ToString()}/{listサムネイル[4]}";
                    row商品レベル["商品画像タイプ5"] = "cabinet";
                    row商品レベル["商品画像パス5"] = 画像パス;
                    row商品レベル["商品画像名（ALT）5"] = "";
                }
                if (6 <= サムネイル枚数)
                {
                    var 画像パス = $"/{dtRowItemMasterMain["画像パス_楽天"].ToString()}/{listサムネイル[5]}";
                    row商品レベル["商品画像タイプ6"] = "cabinet";
                    row商品レベル["商品画像パス6"] = 画像パス;
                    row商品レベル["商品画像名（ALT）6"] = "";
                }
                if (7 <= サムネイル枚数)
                {
                    var 画像パス = $"/{dtRowItemMasterMain["画像パス_楽天"].ToString()}/{listサムネイル[6]}";
                    row商品レベル["商品画像タイプ7"] = "cabinet";
                    row商品レベル["商品画像パス7"] = 画像パス;
                    row商品レベル["商品画像名（ALT）7"] = "";
                }
                if (8 <= サムネイル枚数)
                {
                    var 画像パス = $"/{dtRowItemMasterMain["画像パス_楽天"].ToString()}/{listサムネイル[7]}";
                    row商品レベル["商品画像タイプ8"] = "cabinet";
                    row商品レベル["商品画像パス8"] = 画像パス;
                    row商品レベル["商品画像名（ALT）8"] = "";
                }
                if (9 <= サムネイル枚数)
                {
                    var 画像パス = $"/{dtRowItemMasterMain["画像パス_楽天"].ToString()}/{listサムネイル[8]}";
                    row商品レベル["商品画像タイプ9"] = "cabinet";
                    row商品レベル["商品画像パス9"] = 画像パス;
                    row商品レベル["商品画像名（ALT）9"] = "";
                }
                if (10 <= サムネイル枚数)
                {
                    var 画像パス = $"/{dtRowItemMasterMain["画像パス_楽天"].ToString()}/{listサムネイル[9]}";
                    row商品レベル["商品画像タイプ10"] = "cabinet";
                    row商品レベル["商品画像パス10"] = 画像パス;
                    row商品レベル["商品画像名（ALT）10"] = "";
                }
                if (11 <= サムネイル枚数)
                {
                    var 画像パス = $"/{dtRowItemMasterMain["画像パス_楽天"].ToString()}/{listサムネイル[10]}";
                    row商品レベル["商品画像タイプ11"] = "cabinet";
                    row商品レベル["商品画像パス11"] = 画像パス;
                    row商品レベル["商品画像名（ALT）11"] = "";
                }
                if (12 <= サムネイル枚数)
                {
                    var 画像パス = $"/{dtRowItemMasterMain["画像パス_楽天"].ToString()}/{listサムネイル[11]}";
                    row商品レベル["商品画像タイプ12"] = "cabinet";
                    row商品レベル["商品画像パス12"] = 画像パス;
                    row商品レベル["商品画像名（ALT）12"] = "";
                }
                if (13 <= サムネイル枚数)
                {
                    var 画像パス = $"/{dtRowItemMasterMain["画像パス_楽天"].ToString()}/{listサムネイル[12]}";
                    row商品レベル["商品画像タイプ13"] = "cabinet";
                    row商品レベル["商品画像パス13"] = 画像パス;
                    row商品レベル["商品画像名（ALT）13"] = "";
                }
                if (14 <= サムネイル枚数)
                {
                    var 画像パス = $"/{dtRowItemMasterMain["画像パス_楽天"].ToString()}/{listサムネイル[13]}";
                    row商品レベル["商品画像タイプ14"] = "cabinet";
                    row商品レベル["商品画像パス14"] = 画像パス;
                    row商品レベル["商品画像名（ALT）14"] = "";
                }
                if (15 <= サムネイル枚数)
                {
                    var 画像パス = $"/{dtRowItemMasterMain["画像パス_楽天"].ToString()}/{listサムネイル[14]}";
                    row商品レベル["商品画像タイプ15"] = "cabinet";
                    row商品レベル["商品画像パス15"] = 画像パス;
                    row商品レベル["商品画像名（ALT）15"] = "";
                }
                if (16 <= サムネイル枚数)
                {
                    var 画像パス = $"/{dtRowItemMasterMain["画像パス_楽天"].ToString()}/{listサムネイル[15]}";
                    row商品レベル["商品画像タイプ16"] = "cabinet";
                    row商品レベル["商品画像パス16"] = 画像パス;
                    row商品レベル["商品画像名（ALT）16"] = "";
                }
                if (17 <= サムネイル枚数)
                {
                    var 画像パス = $"/{dtRowItemMasterMain["画像パス_楽天"].ToString()}/{listサムネイル[16]}";
                    row商品レベル["商品画像タイプ17"] = "cabinet";
                    row商品レベル["商品画像パス17"] = 画像パス;
                    row商品レベル["商品画像名（ALT）17"] = "";
                }
                if (18 <= サムネイル枚数)
                {
                    var 画像パス = $"/{dtRowItemMasterMain["画像パス_楽天"].ToString()}/{listサムネイル[17]}";
                    row商品レベル["商品画像タイプ18"] = "cabinet";
                    row商品レベル["商品画像パス18"] = 画像パス;
                    row商品レベル["商品画像名（ALT）18"] = "";
                }
                if (19 <= サムネイル枚数)
                {
                    var 画像パス = $"/{dtRowItemMasterMain["画像パス_楽天"].ToString()}/{listサムネイル[18]}";
                    row商品レベル["商品画像タイプ19"] = "cabinet";
                    row商品レベル["商品画像パス19"] = 画像パス;
                    row商品レベル["商品画像名（ALT）19"] = "";
                }
                if (20 <= サムネイル枚数)
                {
                    var 画像パス = $"/{dtRowItemMasterMain["画像パス_楽天"].ToString()}/{listサムネイル[19]}";
                    row商品レベル["商品画像タイプ20"] = "cabinet";
                    row商品レベル["商品画像パス20"] = 画像パス;
                    row商品レベル["商品画像名（ALT）20"] = "";
                }
                row商品レベル["動画"] = "";
                row商品レベル["白背景画像タイプ"] = "";
                row商品レベル["白背景画像パス"] = "";
                row商品レベル["商品情報レイアウト"] = "5";
                row商品レベル["ヘッダー・フッター・レフトナビ"] = "自動選択";
                row商品レベル["表示項目の並び順"] = "自動選択";
                row商品レベル["共通説明文（小）"] = "自動選択";
                row商品レベル["目玉商品"] = "自動選択";
                row商品レベル["共通説明文（大）"] = "自動選択";
                row商品レベル["レビュー本文表示"] = "2";
                row商品レベル["メーカー提供情報表示"] = "0";
                //row商品レベル["定期購入ボタン"] = "0";
                //row商品レベル["指定可能なお届け日・月ごとに日付を指定"] = "0";
                //row商品レベル["指定可能なお届け日・週ごとに曜日を指定"] = "0";

                //バリエーション選択肢定義_商品選択肢
                var list商品選択肢 = new List<string>();
                var SKU項目名 = dtRowItemMasterMain["R_SKU項目名"].ToString();
                var 別途送料地域項目名 = dtRowItemMasterMain["R_別途送料地域項目名"].ToString();
                var 配達オプション項目名 = dtRowItemMasterMain["R_配達オプション項目名"].ToString();

                Debug.WriteLine($"{商品コード}のSKUが{rowsSKU.Length}件見つかりました。");
                foreach (var rowSKU in rowsSKU)
                {
                    var 選択肢名 = rowSKU["選択肢名"].ToString();
                    list商品選択肢.Add(選択肢名);
                    Debug.WriteLine($"{商品コード}:{SKU項目名}:{選択肢名}");
                }
                var バリエーション選択肢定義_SKU項目名 = UtilityClass.ListToString(list商品選択肢, "|"); //×



                //バリエーション選択肢定義_別途送料地域 
                var list別途送料地域 = new List<string>();
                foreach (var 別途送料地域 in dtRowItemMasterMain["R_別途送料地域選択肢"].ToString().Split(' '))
                {
                    list別途送料地域.Add(別途送料地域.Split(':')[0]);
                }
                var バリエーション選択肢定義_別途送料地域 = UtilityClass.ListToString(list別途送料地域, "|");



                //バリエーション選択肢定義_配達オプション 
                var list配達オプション = new List<string>();
                foreach (var 配達オプション in dtRowItemMasterMain["R_配達オプション選択肢"].ToString().Split(' '))
                {
                    list配達オプション.Add(配達オプション.Split(':')[0]);
                }
                var バリエーション選択肢定義_配達オプション = UtilityClass.ListToString(list配達オプション, "|");



                if (is別途送料地域 && is配達オプション)
                {
                    if (rowsSKU.Length == 1)
                    {
                        //単色カラー
                        row商品レベル["バリエーション項目キー定義"] = "1|2";
                        row商品レベル["バリエーション項目名定義"] = $"{別途送料地域項目名}|{配達オプション項目名}";
                        row商品レベル["バリエーション1選択肢定義"] = バリエーション選択肢定義_別途送料地域;
                        row商品レベル["バリエーション2選択肢定義"] = バリエーション選択肢定義_配達オプション;
                        row商品レベル["バリエーション3選択肢定義"] = "";
                    }
                    else if (1 < rowsSKU.Length)
                    {
                        //複色カラー
                        row商品レベル["バリエーション項目キー定義"] = "1|2|3";
                        row商品レベル["バリエーション項目名定義"] = $"{SKU項目名}|{別途送料地域項目名}|{配達オプション項目名}";
                        row商品レベル["バリエーション1選択肢定義"] = バリエーション選択肢定義_SKU項目名;
                        row商品レベル["バリエーション2選択肢定義"] = バリエーション選択肢定義_別途送料地域;
                        row商品レベル["バリエーション3選択肢定義"] = バリエーション選択肢定義_配達オプション;
                    }
                    else
                    {
                        Console.WriteLine($"エラー：{商品コード}のSKUが定義されていません。");
                        //MessageBox.Show($"エラー：{商品コード}のSKUが定義されていません。");
                        skuErrorList.Add($"{商品コード}");
                    }
                }
                else if (is別途送料地域 && is配達オプション == false)
                {
                    if (rowsSKU.Length == 1)
                    {
                        //単色カラー
                        row商品レベル["バリエーション項目キー定義"] = "1";
                        row商品レベル["バリエーション項目名定義"] = $"{別途送料地域項目名}";
                        row商品レベル["バリエーション1選択肢定義"] = バリエーション選択肢定義_別途送料地域;
                        row商品レベル["バリエーション2選択肢定義"] = "";
                        row商品レベル["バリエーション3選択肢定義"] = "";
                    }
                    else if (1 < rowsSKU.Length)
                    {
                        //複色カラー
                        row商品レベル["バリエーション項目キー定義"] = "1|2";
                        row商品レベル["バリエーション項目名定義"] = $"{SKU項目名}|{別途送料地域項目名}";
                        row商品レベル["バリエーション1選択肢定義"] = バリエーション選択肢定義_SKU項目名;
                        row商品レベル["バリエーション2選択肢定義"] = バリエーション選択肢定義_別途送料地域;
                        row商品レベル["バリエーション3選択肢定義"] = "";
                    }
                    else
                    {
                        Console.WriteLine($"エラー：{商品コード}のSKUが定義されていません。");
                        //MessageBox.Show($"エラー：{商品コード}のSKUが定義されていません。");
                        skuErrorList.Add($"{商品コード}");
                    }
                }
                else if (is配達オプション && is別途送料地域 == false)
                {
                    if (rowsSKU.Length == 1)
                    {
                        //単色カラー
                        row商品レベル["バリエーション項目キー定義"] = "1";
                        row商品レベル["バリエーション項目名定義"] = $"{配達オプション項目名}";
                        row商品レベル["バリエーション1選択肢定義"] = バリエーション選択肢定義_配達オプション;
                        row商品レベル["バリエーション2選択肢定義"] = "";
                        row商品レベル["バリエーション3選択肢定義"] = "";
                    }
                    else if (1 < rowsSKU.Length)
                    {
                        //複色カラー
                        row商品レベル["バリエーション項目キー定義"] = "1|2";
                        row商品レベル["バリエーション項目名定義"] = $"{SKU項目名}|{配達オプション項目名}";
                        row商品レベル["バリエーション1選択肢定義"] = バリエーション選択肢定義_SKU項目名;
                        row商品レベル["バリエーション2選択肢定義"] = バリエーション選択肢定義_配達オプション;
                        row商品レベル["バリエーション3選択肢定義"] = "";
                    }
                    else
                    {
                        Console.WriteLine($"エラー：{商品コード}のSKUが定義されていません。");
                        //MessageBox.Show($"エラー：{商品コード}のSKUが定義されていません。");
                        skuErrorList.Add($"{商品コード}");
                    }
                }
                else
                {
                    if (rowsSKU.Length == 1)
                    {
                        //単色カラー
                        row商品レベル["バリエーション項目キー定義"] = "";
                        row商品レベル["バリエーション項目名定義"] = "";
                        row商品レベル["バリエーション1選択肢定義"] = "";
                        row商品レベル["バリエーション2選択肢定義"] = "";
                        row商品レベル["バリエーション3選択肢定義"] = "";
                    }
                    else if (1 < rowsSKU.Length)
                    {
                        //複色カラー
                        row商品レベル["バリエーション項目キー定義"] = "1";
                        row商品レベル["バリエーション項目名定義"] = $"{SKU項目名}";
                        row商品レベル["バリエーション1選択肢定義"] = バリエーション選択肢定義_SKU項目名;
                        row商品レベル["バリエーション2選択肢定義"] = "";
                        row商品レベル["バリエーション3選択肢定義"] = "";
                    }
                    else
                    {
                        Console.WriteLine($"エラー：{商品コード}のSKUが定義されていません。");
                        //MessageBox.Show($"エラー：{商品コード}のSKUが定義されていません。");
                        skuErrorList.Add($"{商品コード}");
                    }
                }

                dtRakutenItem.Rows.Add(row商品レベル);
            }

            ////////////////////////////////////////////////////////////////////////////////////////////
            // 商品オプションレベル
            ////////////////////////////////////////////////////////////////////////////////////////////
            {
                var 楽天_選択肢_プルダウン = "";

                var 選択肢 = UtilityClass.ReadTextFile(ecPath + @"templates\【選択肢】" + dtRowItemMasterMain["R_注意事項プルダウン"].ToString() + "_R.txt");
                if (dtRowItemMasterMain["R_商品プルダウン"].ToString().Length > 1)
                {
                    //商品固有のプルダウン選択肢がある場合、テンプレートのプルダウン選択肢と結合する
                    楽天_選択肢_プルダウン = dtRowItemMasterMain["R_商品プルダウン"].ToString();
                    楽天_選択肢_プルダウン = 楽天_選択肢_プルダウン + "\r\n" + "\r\n" + 選択肢;
                }
                else
                {
                    //商品固有のプルダウン選択肢がない場合、テンプレートのプルダウン選択肢のみとする
                    楽天_選択肢_プルダウン = 選択肢;
                }

                //楽天_選択肢_プルダウン = 楽天_選択肢_プルダウン.Replace("必ず選択してください ", "");

                var sr = new System.IO.StringReader(楽天_選択肢_プルダウン);
                while (sr.Peek() > 0)
                {
                    string 行 = sr.ReadLine().ToString();
                    string[] 楽天選択肢_プルダウン_配列 = 行.Split(' ');
                    if (行 == "")
                    {
                        continue;
                    }

                    var row商品オプションレベル = dtRakutenItem.NewRow();
                    row商品オプションレベル["商品管理番号（商品URL）"] = 商品コード;
                    row商品オプションレベル["商品オプション項目名"] = 楽天選択肢_プルダウン_配列[0];

                    if (2 <= 楽天選択肢_プルダウン_配列.Length)
                    {
                        row商品オプションレベル["商品オプション選択肢1"] = 楽天選択肢_プルダウン_配列[1];
                    }
                    if (3 <= 楽天選択肢_プルダウン_配列.Length)
                    {
                        row商品オプションレベル["商品オプション選択肢2"] = 楽天選択肢_プルダウン_配列[2];
                    }
                    if (4 <= 楽天選択肢_プルダウン_配列.Length)
                    {
                        row商品オプションレベル["商品オプション選択肢3"] = 楽天選択肢_プルダウン_配列[3];
                    }
                    if (5 <= 楽天選択肢_プルダウン_配列.Length)
                    {
                        row商品オプションレベル["商品オプション選択肢4"] = 楽天選択肢_プルダウン_配列[4];
                    }
                    if (6 <= 楽天選択肢_プルダウン_配列.Length)
                    {
                        row商品オプションレベル["商品オプション選択肢5"] = 楽天選択肢_プルダウン_配列[5];
                    }
                    if (7 <= 楽天選択肢_プルダウン_配列.Length)
                    {
                        row商品オプションレベル["商品オプション選択肢6"] = 楽天選択肢_プルダウン_配列[6];
                    }
                    if (8 <= 楽天選択肢_プルダウン_配列.Length)
                    {
                        row商品オプションレベル["商品オプション選択肢7"] = 楽天選択肢_プルダウン_配列[7];
                    }
                    if (9 <= 楽天選択肢_プルダウン_配列.Length)
                    {
                        row商品オプションレベル["商品オプション選択肢8"] = 楽天選択肢_プルダウン_配列[8];
                    }
                    row商品オプションレベル["選択肢タイプ"] = "s";
                    row商品オプションレベル["商品オプション選択必須"] = "1";
                    dtRakutenItem.Rows.Add(row商品オプションレベル);
                }
                sr.Close();
            }


            ////////////////////////////////////////////////////////////////////////////////////////////
            // SKUレベル
            ////////////////////////////////////////////////////////////////////////////////////////////
            {
                Debug.WriteLine($"{商品コード}のSKUが{rowsSKU.Length}件見つかりました。");
                var 連番 = 0;

                //商品選択肢×別途送料地域×配達オプション
                if (is別途送料地域 && is配達オプション)
                {
                    foreach (var rowSKU in rowsSKU)
                    {
                        //バリエーション選択肢定義_別途送料地域 
                        foreach (var 別途送料地域 in dtRowItemMasterMain["R_別途送料地域選択肢"].ToString().Split(' '))
                        {
                            var 別途送料地域名_地域名 = 別途送料地域.Split(':')[0];
                            var 別途送料地域名_料金 = 別途送料地域.Split(':')[1];

                            //バリエーション選択肢定義_配達オプション 
                            foreach (var 配達オプション in dtRowItemMasterMain["R_配達オプション選択肢"].ToString().Split(' '))
                            {
                                var rowSKUレベル = dtRakutenItem.NewRow();

                                var 配達オプション_オプション名 = 配達オプション.Split(':')[0];
                                var 配達オプション_料金 = 配達オプション.Split(':')[1];

                                連番 = 連番 + 1;
                                rowSKUレベル["商品管理番号（商品URL）"] = 商品コード;
                                rowSKUレベル["SKU管理番号"] = $"{rowSKU["SKUコード"].ToString()}_{連番}"; //10桁の商品管理番号+_+3桁の連番
                                rowSKUレベル["システム連携用SKU番号"] = rowSKU["SKUコード"].ToString();

                                //選択肢名
                                var SKUコード = rowSKU["SKUコード"].ToString();
                                if (商品コード == SKUコード)
                                {
                                    //単色
                                    rowSKUレベル["バリエーション項目キー1"] = "1";
                                    rowSKUレベル["バリエーション項目選択肢1"] = 別途送料地域名_地域名;
                                    rowSKUレベル["バリエーション項目キー2"] = "2";
                                    rowSKUレベル["バリエーション項目選択肢2"] = 配達オプション_オプション名;
                                    rowSKUレベル["バリエーション項目キー3"] = "";
                                    rowSKUレベル["バリエーション項目選択肢3"] = "";
                                }
                                else
                                {
                                    //複色
                                    rowSKUレベル["バリエーション項目キー1"] = "1";
                                    rowSKUレベル["バリエーション項目選択肢1"] = rowSKU["選択肢名"].ToString();
                                    rowSKUレベル["バリエーション項目キー2"] = "2";
                                    rowSKUレベル["バリエーション項目選択肢2"] = 別途送料地域名_地域名;
                                    rowSKUレベル["バリエーション項目キー3"] = "3";
                                    rowSKUレベル["バリエーション項目選択肢3"] = 配達オプション_オプション名;
                                }


                                var 販売価格 = int.Parse(dtRowItemMasterMain["当店通常価格_税込み"].ToString()) + int.Parse(別途送料地域名_料金) + int.Parse(配達オプション_料金);
                                var 販売価格説明 = $"{dtRowItemMasterMain["当店通常価格_税込み"].ToString()}+{別途送料地域名_料金}+{配達オプション_料金}";
                                rowSKUレベル["販売価格"] = 販売価格;
                                rowSKUレベル["表示価格"] = "";
                                rowSKUレベル["二重価格文言管理番号"] = "";
                                rowSKUレベル["注文受付数"] = "50";
                                rowSKUレベル["再入荷お知らせボタン"] = "0";
                                rowSKUレベル["のし対応"] = "0";
                                rowSKUレベル["在庫数"] = "999";
                                rowSKUレベル["在庫戻しフラグ"] = "0";
                                rowSKUレベル["在庫切れ時の注文受付"] = "0";
                                rowSKUレベル["在庫あり時納期管理番号"] = "1";
                                rowSKUレベル["在庫切れ時納期管理番号"] = "2";
                                rowSKUレベル["SKU倉庫指定"] = "0";
                                rowSKUレベル["配送方法セット管理番号"] = "";
                                //送料
                                if (dtRowItemMasterMain["送料形態"].ToString() == "送料無料")
                                {
                                    rowSKUレベル["送料"] = "1";
                                }
                                else if (dtRowItemMasterMain["送料形態"].ToString() == "条件付き送料無料")
                                {
                                    rowSKUレベル["送料"] = "0";
                                }
                                else
                                {
                                    rowSKUレベル["送料"] = "0";
                                }
                                rowSKUレベル["送料区分1"] = "";
                                rowSKUレベル["送料区分2"] = "";
                                rowSKUレベル["個別送料"] = "";
                                rowSKUレベル["地域別個別送料管理番号"] = "";
                                rowSKUレベル["単品配送設定使用"] = "";
                                rowSKUレベル["カタログID"] = "";
                                rowSKUレベル["カタログIDなしの理由"] = "5";
                                rowSKUレベル["セット商品用カタログID"] = "";
                                rowSKUレベル["SKU画像タイプ"] = "";

                                //画像パス
                                if (商品コード == SKUコード)
                                {
                                    //単色
                                    var SKU画像URL = $"/{dtRowItemMasterMain["画像パス_楽天"].ToString()}/{商品コード}.jpg";
                                    //rowSKUレベル["SKU画像パス"] = SKU画像URL;
                                    rowSKUレベル["SKU画像パス"] = "";
                                }
                                else
                                {
                                    //複色
                                    if (SKUコード.Length != 10)
                                    {
                                        MessageBox.Show($"エラー：{SKUコード}の文字数が10桁ではありません");
                                    }
                                    var 下3 = int.Parse(UtilityClass.Mid(SKUコード, 8, 1));
                                    var 下2 = int.Parse(UtilityClass.Mid(SKUコード, 9, 1));
                                    var 下1 = int.Parse(UtilityClass.Mid(SKUコード, 10, 1));

                                    var 番号 = int.Parse($"{下3}{下2}");
                                    if (下1 != 0)
                                    {
                                        MessageBox.Show($"SKUの数は99が最大です。{商品コード} 商品管理番号下3桁{下3}{下2}{下1}");
                                    }
                                    var SKU画像URL = $"/{dtRowItemMasterMain["画像パス_楽天"].ToString()}/{商品コード}_{番号}.jpg";

                                    //rowSKUレベル["SKU画像パス"] = SKU画像URL;
                                    rowSKUレベル["SKU画像パス"] = "";
                                }

                                rowSKUレベル["SKU画像名（ALT）"] = "";
                                rowSKUレベル["自由入力行（項目）1"] = "";
                                rowSKUレベル["自由入力行（値）1"] = "";
                                rowSKUレベル["自由入力行（項目）2"] = "";
                                rowSKUレベル["自由入力行（値）2"] = "";
                                //rowSKUレベル["定期商品販売価格"] = "";
                                //rowSKUレベル["初回価格"] = "";
                                for (var i = 1; i < 41; i++)
                                {
                                    //-はスルー
                                    if (rowSKU[$"商品属性（項目）{i}"].ToString() != "-")
                                    {
                                        rowSKUレベル[$"商品属性（項目）{i}"] = rowSKU[$"商品属性（項目）{i}"].ToString();
                                        rowSKUレベル[$"商品属性（値）{i}"] = rowSKU[$"商品属性（値）{i}"].ToString();
                                    }
                                    if (rowSKU[$"商品属性（単位）{i}"].ToString() != "-")
                                    {
                                        rowSKUレベル[$"商品属性（単位）{i}"] = rowSKU[$"商品属性（単位）{i}"].ToString();
                                    }
                                }
                                dtRakutenItem.Rows.Add(rowSKUレベル);
                            }
                        }
                    }
                }
                else if (is別途送料地域)
                {
                    foreach (var rowSKU in rowsSKU)
                    {
                        //バリエーション選択肢定義_別途送料地域 
                        foreach (var 別途送料地域 in dtRowItemMasterMain["R_別途送料地域選択肢"].ToString().Split(' '))
                        {
                            var 別途送料地域名_地域名 = 別途送料地域.Split(':')[0];
                            var 別途送料地域名_料金 = 別途送料地域.Split(':')[1];

                            //バリエーション選択肢定義_配達オプション 
                            var rowSKUレベル = dtRakutenItem.NewRow();

                            連番 = 連番 + 1;
                            rowSKUレベル["商品管理番号（商品URL）"] = 商品コード;
                            rowSKUレベル["SKU管理番号"] = $"{rowSKU["SKUコード"].ToString()}_{連番}"; //10桁の商品管理番号+_+3桁の連番
                            rowSKUレベル["システム連携用SKU番号"] = rowSKU["SKUコード"].ToString();

                            //選択肢名
                            var SKUコード = rowSKU["SKUコード"].ToString();
                            if (商品コード == SKUコード)
                            {
                                //単色
                                rowSKUレベル["バリエーション項目キー1"] = "1";
                                rowSKUレベル["バリエーション項目選択肢1"] = 別途送料地域名_地域名;
                                rowSKUレベル["バリエーション項目キー2"] = "";
                                rowSKUレベル["バリエーション項目選択肢2"] = "";
                                rowSKUレベル["バリエーション項目キー3"] = "";
                                rowSKUレベル["バリエーション項目選択肢3"] = "";
                            }
                            else
                            {
                                //複色
                                rowSKUレベル["バリエーション項目キー1"] = "1";
                                rowSKUレベル["バリエーション項目選択肢1"] = rowSKU["選択肢名"].ToString();
                                rowSKUレベル["バリエーション項目キー2"] = "2";
                                rowSKUレベル["バリエーション項目選択肢2"] = 別途送料地域名_地域名;
                                rowSKUレベル["バリエーション項目キー3"] = "";
                                rowSKUレベル["バリエーション項目選択肢3"] = "";
                            }

                            var 販売価格 = int.Parse(dtRowItemMasterMain["当店通常価格_税込み"].ToString()) + int.Parse(別途送料地域名_料金);
                            var 販売価格説明 = $"{dtRowItemMasterMain["当店通常価格_税込み"].ToString()}+{別途送料地域名_料金}";
                            rowSKUレベル["販売価格"] = 販売価格;
                            rowSKUレベル["表示価格"] = "";
                            rowSKUレベル["二重価格文言管理番号"] = "";
                            rowSKUレベル["注文受付数"] = "50";
                            rowSKUレベル["再入荷お知らせボタン"] = "0";
                            rowSKUレベル["のし対応"] = "0";
                            rowSKUレベル["在庫数"] = "999";
                            rowSKUレベル["在庫戻しフラグ"] = "0";
                            rowSKUレベル["在庫切れ時の注文受付"] = "0";
                            rowSKUレベル["在庫あり時納期管理番号"] = "1";
                            rowSKUレベル["在庫切れ時納期管理番号"] = "2";
                            rowSKUレベル["SKU倉庫指定"] = "0";
                            rowSKUレベル["配送方法セット管理番号"] = "";
                            //送料
                            if (dtRowItemMasterMain["送料形態"].ToString() == "送料無料")
                            {
                                rowSKUレベル["送料"] = "1";
                            }
                            else if (dtRowItemMasterMain["送料形態"].ToString() == "条件付き送料無料")
                            {
                                rowSKUレベル["送料"] = "0";
                            }
                            else
                            {
                                rowSKUレベル["送料"] = "0";
                            }
                            rowSKUレベル["送料区分1"] = "";
                            rowSKUレベル["送料区分2"] = "";
                            rowSKUレベル["個別送料"] = "";
                            rowSKUレベル["地域別個別送料管理番号"] = "";
                            rowSKUレベル["単品配送設定使用"] = "";
                            rowSKUレベル["カタログID"] = "";
                            rowSKUレベル["カタログIDなしの理由"] = "5";
                            rowSKUレベル["セット商品用カタログID"] = "";
                            rowSKUレベル["SKU画像タイプ"] = "";

                            //画像パス
                            if (商品コード == SKUコード)
                            {
                                //単色
                                var SKU画像URL = $"/{dtRowItemMasterMain["画像パス_楽天"].ToString()}/{商品コード}.jpg";
                                //rowSKUレベル["SKU画像パス"] = SKU画像URL;
                                rowSKUレベル["SKU画像パス"] = "";
                            }
                            else
                            {
                                //複色
                                if (SKUコード.Length != 10)
                                {
                                    MessageBox.Show($"エラー：{SKUコード}の文字数が10桁ではありません");
                                }
                                var 下3 = int.Parse(UtilityClass.Mid(SKUコード, 8, 1));
                                var 下2 = int.Parse(UtilityClass.Mid(SKUコード, 9, 1));
                                var 下1 = int.Parse(UtilityClass.Mid(SKUコード, 10, 1));

                                var 番号 = int.Parse($"{下3}{下2}");
                                if (下1 != 0)
                                {
                                    MessageBox.Show($"SKUの数は99が最大です。{商品コード} 商品管理番号下3桁{下3}{下2}{下1}");
                                }
                                var SKU画像URL = $"/{dtRowItemMasterMain["画像パス_楽天"].ToString()}/{商品コード}_{番号}.jpg";

                                //rowSKUレベル["SKU画像パス"] = SKU画像URL;
                                rowSKUレベル["SKU画像パス"] = "";
                            }

                            rowSKUレベル["SKU画像名（ALT）"] = "";
                            rowSKUレベル["自由入力行（項目）1"] = "";
                            rowSKUレベル["自由入力行（値）1"] = "";
                            rowSKUレベル["自由入力行（項目）2"] = "";
                            rowSKUレベル["自由入力行（値）2"] = "";
                            //rowSKUレベル["定期商品販売価格"] = "";
                            //rowSKUレベル["初回価格"] = "";
                            for (var i = 1; i < 41; i++)
                            {
                                //-はスルー
                                if (rowSKU[$"商品属性（項目）{i}"].ToString() != "-")
                                {
                                    rowSKUレベル[$"商品属性（項目）{i}"] = rowSKU[$"商品属性（項目）{i}"].ToString();
                                    rowSKUレベル[$"商品属性（値）{i}"] = rowSKU[$"商品属性（値）{i}"].ToString();
                                }
                                if (rowSKU[$"商品属性（単位）{i}"].ToString() != "-")
                                {
                                    rowSKUレベル[$"商品属性（単位）{i}"] = rowSKU[$"商品属性（単位）{i}"].ToString();
                                }
                            }
                            dtRakutenItem.Rows.Add(rowSKUレベル);
                        }
                    }
                }
                else if (is配達オプション)
                {
                    foreach (var rowSKU in rowsSKU)
                    {
                        //バリエーション選択肢定義_配達オプション 
                        foreach (var 配達オプション in dtRowItemMasterMain["R_配達オプション選択肢"].ToString().Split(' '))
                        {
                            var rowSKUレベル = dtRakutenItem.NewRow();

                            var 配達オプション_オプション名 = 配達オプション.Split(':')[0];
                            var 配達オプション_料金 = 配達オプション.Split(':')[1];


                            連番 = 連番 + 1;
                            rowSKUレベル["商品管理番号（商品URL）"] = 商品コード;
                            rowSKUレベル["SKU管理番号"] = $"{rowSKU["SKUコード"].ToString()}_{連番}"; //10桁の商品管理番号+_+3桁の連番
                            rowSKUレベル["システム連携用SKU番号"] = rowSKU["SKUコード"].ToString();

                            //選択肢名
                            var SKUコード = rowSKU["SKUコード"].ToString();
                            if (商品コード == SKUコード)
                            {
                                //単色
                                rowSKUレベル["バリエーション項目キー1"] = "1";
                                rowSKUレベル["バリエーション項目選択肢1"] = 配達オプション_オプション名;
                                rowSKUレベル["バリエーション項目キー2"] = "";
                                rowSKUレベル["バリエーション項目選択肢2"] = "";
                                rowSKUレベル["バリエーション項目キー3"] = "";
                                rowSKUレベル["バリエーション項目選択肢3"] = "";
                            }
                            else
                            {
                                //複色
                                rowSKUレベル["バリエーション項目キー1"] = "1";
                                rowSKUレベル["バリエーション項目選択肢1"] = rowSKU["選択肢名"].ToString();
                                rowSKUレベル["バリエーション項目キー2"] = "2";
                                rowSKUレベル["バリエーション項目選択肢2"] = 配達オプション_オプション名;
                                rowSKUレベル["バリエーション項目キー3"] = "";
                                rowSKUレベル["バリエーション項目選択肢3"] = "";
                            }


                            var 販売価格 = int.Parse(dtRowItemMasterMain["当店通常価格_税込み"].ToString()) + int.Parse(配達オプション_料金);
                            var 販売価格説明 = $"{dtRowItemMasterMain["当店通常価格_税込み"].ToString()}+{配達オプション_料金}";
                            rowSKUレベル["販売価格"] = 販売価格;
                            rowSKUレベル["表示価格"] = "";
                            rowSKUレベル["二重価格文言管理番号"] = "";
                            rowSKUレベル["注文受付数"] = "50";
                            rowSKUレベル["再入荷お知らせボタン"] = "0";
                            rowSKUレベル["のし対応"] = "0";
                            rowSKUレベル["在庫数"] = "999";
                            rowSKUレベル["在庫戻しフラグ"] = "0";
                            rowSKUレベル["在庫切れ時の注文受付"] = "0";
                            rowSKUレベル["在庫あり時納期管理番号"] = "1";
                            rowSKUレベル["在庫切れ時納期管理番号"] = "2";
                            rowSKUレベル["SKU倉庫指定"] = "0";
                            rowSKUレベル["配送方法セット管理番号"] = "";
                            //送料
                            if (dtRowItemMasterMain["送料形態"].ToString() == "送料無料")
                            {
                                rowSKUレベル["送料"] = "1";
                            }
                            else if (dtRowItemMasterMain["送料形態"].ToString() == "条件付き送料無料")
                            {
                                rowSKUレベル["送料"] = "0";
                            }
                            else
                            {
                                rowSKUレベル["送料"] = "0";
                            }
                            rowSKUレベル["送料区分1"] = "";
                            rowSKUレベル["送料区分2"] = "";
                            rowSKUレベル["個別送料"] = "";
                            rowSKUレベル["地域別個別送料管理番号"] = "";
                            rowSKUレベル["単品配送設定使用"] = "";
                            rowSKUレベル["カタログID"] = "";
                            rowSKUレベル["カタログIDなしの理由"] = "5";
                            rowSKUレベル["セット商品用カタログID"] = "";
                            rowSKUレベル["SKU画像タイプ"] = "";

                            //画像パス
                            if (商品コード == SKUコード)
                            {
                                //単色
                                var SKU画像URL = $"/{dtRowItemMasterMain["画像パス_楽天"].ToString()}/{商品コード}.jpg";
                                //rowSKUレベル["SKU画像パス"] = SKU画像URL;
                                rowSKUレベル["SKU画像パス"] = "";
                            }
                            else
                            {
                                //複色
                                if (SKUコード.Length != 10)
                                {
                                    MessageBox.Show($"エラー：{SKUコード}の文字数が10桁ではありません");
                                }
                                var 下3 = int.Parse(UtilityClass.Mid(SKUコード, 8, 1));
                                var 下2 = int.Parse(UtilityClass.Mid(SKUコード, 9, 1));
                                var 下1 = int.Parse(UtilityClass.Mid(SKUコード, 10, 1));

                                var 番号 = int.Parse($"{下3}{下2}");
                                if (下1 != 0)
                                {
                                    MessageBox.Show($"SKUの数は99が最大です。{商品コード} 商品管理番号下3桁{下3}{下2}{下1}");
                                }
                                var SKU画像URL = $"/{dtRowItemMasterMain["画像パス_楽天"].ToString()}/{商品コード}_{番号}.jpg";

                                //rowSKUレベル["SKU画像パス"] = SKU画像URL;
                                rowSKUレベル["SKU画像パス"] = "";
                            }

                            rowSKUレベル["SKU画像名（ALT）"] = "";
                            rowSKUレベル["自由入力行（項目）1"] = "";
                            rowSKUレベル["自由入力行（値）1"] = "";
                            rowSKUレベル["自由入力行（項目）2"] = "";
                            rowSKUレベル["自由入力行（値）2"] = "";
                            //rowSKUレベル["定期商品販売価格"] = "";
                            //rowSKUレベル["初回価格"] = "";
                            for (var i = 1; i < 41; i++)
                            {
                                //-はスルー
                                if (rowSKU[$"商品属性（項目）{i}"].ToString() != "-")
                                {
                                    rowSKUレベル[$"商品属性（項目）{i}"] = rowSKU[$"商品属性（項目）{i}"].ToString();
                                    rowSKUレベル[$"商品属性（値）{i}"] = rowSKU[$"商品属性（値）{i}"].ToString();
                                }
                                if (rowSKU[$"商品属性（単位）{i}"].ToString() != "-")
                                {
                                    rowSKUレベル[$"商品属性（単位）{i}"] = rowSKU[$"商品属性（単位）{i}"].ToString();
                                }
                            }
                            dtRakutenItem.Rows.Add(rowSKUレベル);
                        }
                    }
                }
                else
                {
                    foreach (var rowSKU in rowsSKU)
                    {
                        //バリエーション選択肢定義_別途送料地域 
                        var rowSKUレベル = dtRakutenItem.NewRow();

                        rowSKUレベル["商品管理番号（商品URL）"] = 商品コード;
                        rowSKUレベル["システム連携用SKU番号"] = rowSKU["SKUコード"].ToString();

                        //選択肢名
                        var SKUコード = rowSKU["SKUコード"].ToString();
                        if (商品コード == SKUコード)
                        {
                            //単色
                            rowSKUレベル["SKU管理番号"] = $"{rowSKU["SKUコード"].ToString()}"; //10桁の商品管理番号のみ
                            rowSKUレベル["バリエーション項目キー1"] = "";
                            rowSKUレベル["バリエーション項目選択肢1"] = "";
                            rowSKUレベル["バリエーション項目キー2"] = "";
                            rowSKUレベル["バリエーション項目選択肢2"] = "";
                            rowSKUレベル["バリエーション項目キー3"] = "";
                            rowSKUレベル["バリエーション項目選択肢3"] = "";
                        }
                        else
                        {
                            //複色
                            連番 = 連番 + 1;
                            rowSKUレベル["SKU管理番号"] = $"{rowSKU["SKUコード"].ToString()}_{連番}"; //10桁の商品管理番号+_+3桁の連番 連番は不要
                            rowSKUレベル["バリエーション項目キー1"] = "1";
                            rowSKUレベル["バリエーション項目選択肢1"] = rowSKU["選択肢名"].ToString();
                            rowSKUレベル["バリエーション項目キー2"] = "";
                            rowSKUレベル["バリエーション項目選択肢2"] = "";
                            rowSKUレベル["バリエーション項目キー3"] = "";
                            rowSKUレベル["バリエーション項目選択肢3"] = "";
                        }

                        var 販売価格 = int.Parse(dtRowItemMasterMain["当店通常価格_税込み"].ToString());
                        rowSKUレベル["販売価格"] = 販売価格;
                        rowSKUレベル["表示価格"] = "";
                        rowSKUレベル["二重価格文言管理番号"] = "";
                        rowSKUレベル["注文受付数"] = "50";
                        rowSKUレベル["再入荷お知らせボタン"] = "0";
                        rowSKUレベル["のし対応"] = "0";
                        rowSKUレベル["在庫数"] = "999";
                        rowSKUレベル["在庫戻しフラグ"] = "0";
                        rowSKUレベル["在庫切れ時の注文受付"] = "0";
                        rowSKUレベル["在庫あり時納期管理番号"] = "1";
                        rowSKUレベル["在庫切れ時納期管理番号"] = "2";
                        rowSKUレベル["SKU倉庫指定"] = "0";
                        rowSKUレベル["配送方法セット管理番号"] = "";
                        //送料
                        if (dtRowItemMasterMain["送料形態"].ToString() == "送料無料")
                        {
                            rowSKUレベル["送料"] = "1";
                        }
                        else if (dtRowItemMasterMain["送料形態"].ToString() == "条件付き送料無料")
                        {
                            rowSKUレベル["送料"] = "0";
                        }
                        else
                        {
                            rowSKUレベル["送料"] = "0";
                        }
                        rowSKUレベル["送料区分1"] = "";
                        rowSKUレベル["送料区分2"] = "";
                        rowSKUレベル["個別送料"] = "";
                        rowSKUレベル["地域別個別送料管理番号"] = "";
                        rowSKUレベル["単品配送設定使用"] = "";
                        rowSKUレベル["カタログID"] = "";
                        rowSKUレベル["カタログIDなしの理由"] = "5";
                        rowSKUレベル["セット商品用カタログID"] = "";
                        rowSKUレベル["SKU画像タイプ"] = "";

                        //画像パス
                        if (商品コード == SKUコード)
                        {
                            //単色
                            var SKU画像URL = $"/{dtRowItemMasterMain["画像パス_楽天"].ToString()}/{商品コード}.jpg";
                            //rowSKUレベル["SKU画像パス"] = SKU画像URL;
                            rowSKUレベル["SKU画像パス"] = "";
                        }
                        else
                        {
                            //複色
                            if (SKUコード.Length != 10)
                            {
                                MessageBox.Show($"エラー：{SKUコード}の文字数が10桁ではありません");
                            }
                            var 下3 = int.Parse(UtilityClass.Mid(SKUコード, 8, 1));
                            var 下2 = int.Parse(UtilityClass.Mid(SKUコード, 9, 1));
                            var 下1 = int.Parse(UtilityClass.Mid(SKUコード, 10, 1));

                            var 番号 = int.Parse($"{下3}{下2}");
                            if (下1 != 0)
                            {
                                MessageBox.Show($"SKUの数は99が最大です。{商品コード} 商品管理番号下3桁{下3}{下2}{下1}");
                            }
                            var SKU画像URL = $"/{dtRowItemMasterMain["画像パス_楽天"].ToString()}/{商品コード}_{番号}.jpg";

                            //rowSKUレベル["SKU画像パス"] = SKU画像URL;
                            rowSKUレベル["SKU画像パス"] = "";
                        }

                        rowSKUレベル["SKU画像名（ALT）"] = "";
                        rowSKUレベル["自由入力行（項目）1"] = "";
                        rowSKUレベル["自由入力行（値）1"] = "";
                        rowSKUレベル["自由入力行（項目）2"] = "";
                        rowSKUレベル["自由入力行（値）2"] = "";
                        //rowSKUレベル["定期商品販売価格"] = "";
                        //rowSKUレベル["初回価格"] = "";
                        for (var i = 1; i < 41; i++)
                        {
                            //-はスルー
                            if (rowSKU[$"商品属性（項目）{i}"].ToString() != "-")
                            {
                                rowSKUレベル[$"商品属性（項目）{i}"] = rowSKU[$"商品属性（項目）{i}"].ToString();
                                rowSKUレベル[$"商品属性（値）{i}"] = rowSKU[$"商品属性（値）{i}"].ToString();
                            }
                            if (rowSKU[$"商品属性（単位）{i}"].ToString() != "-")
                            {
                                rowSKUレベル[$"商品属性（単位）{i}"] = rowSKU[$"商品属性（単位）{i}"].ToString();
                            }
                        }
                        dtRakutenItem.Rows.Add(rowSKUレベル);
                    }
                }

            }

            return dtRakutenItem;
        }
        public DataTable GenerateCsvRakutenCategory(DataTable dtRakutenCategory, DataRow dRow)
        {
            string[] categoryArg =
            {
                    dRow["商品カテゴリ1"].ToString(),
                    dRow["商品カテゴリ2"].ToString(),
                    dRow["商品カテゴリ3"].ToString(),
                    dRow["商品カテゴリ4"].ToString(),
                    dRow["商品カテゴリ5"].ToString()
            };

            for (int i = 0; i < categoryArg.Length; i++)
            {
                if (1 < categoryArg[i].Length)
                {
                    var dtRowRakutenCategory = dtRakutenCategory.NewRow();
                    dtRowRakutenCategory["商品管理番号（商品URL）"] = dRow["mycode"].ToString();
                    dtRowRakutenCategory["表示先カテゴリ"] = categoryArg[i].Replace(":", "\\");
                    if (dRow["R_商品名"].ToString().Length > 1)
                    {
                        dtRowRakutenCategory["商品名"] = dRow["R_商品名"].ToString();
                    }
                    dtRowRakutenCategory["コントロールカラム"] = dRow["コントロールカラム"].ToString();
                    dtRowRakutenCategory["URL"] = "";
                    dtRowRakutenCategory["1ページ複数形式"] = "";
                    if (dRow["ソート"].ToString() == "-")
                    {
                        dtRowRakutenCategory["優先度"] = "";
                    }
                    else
                    {
                        dtRowRakutenCategory["優先度"] = dRow["ソート"].ToString();
                    }
                    dtRakutenCategory.Rows.Add(dtRowRakutenCategory);
                }
            }
            return dtRakutenCategory;
        }
        public DataTable GenerateCsvYahooItem(DataTable dtYahooItem, DataRow dtRowItemMasterMain, string generateTable, string generateKagoyoko, string generateSpKanrenShohin, DataTable dtRowItemMasterSKU)
        {
            var dtRowYahoo = dtYahooItem.NewRow();

            var 商品コード = dtRowItemMasterMain["mycode"].ToString();

            //code
            dtRowYahoo["code"] = 商品コード;

            //path
            string[] categoryArg =
            {
                dtRowItemMasterMain["商品カテゴリ1"].ToString(),
                dtRowItemMasterMain["商品カテゴリ2"].ToString(),
                dtRowItemMasterMain["商品カテゴリ3"].ToString(),
                dtRowItemMasterMain["商品カテゴリ4"].ToString(),
                dtRowItemMasterMain["商品カテゴリ5"].ToString()
            };
            string path = "";
            for (int i = 0; i < categoryArg.Length; i++)
            {
                if (categoryArg[i].Length > 1)
                {
                    path = path + categoryArg[i] + "\r\n";
                }
            }
            dtRowYahoo["path"] = path;

            //name
            if (dtRowItemMasterMain["Y_商品名"].ToString().Length > 1)
            {
                dtRowYahoo["name"] = dtRowItemMasterMain["Y_商品名"].ToString();
            }
            else
            {
                dtRowYahoo["name"] = "";
            }

            //headline
            if (dtRowItemMasterMain["Y_キャッチコピー"].ToString().Length > 1)
            {
                dtRowYahoo["headline"] = dtRowItemMasterMain["Y_キャッチコピー"].ToString();
            }
            else
            {
                dtRowYahoo["name"] = "";
            }

            //price
            dtRowYahoo["price"] = dtRowItemMasterMain["当店通常価格_税込み"].ToString();

            //original-price
            if (dtRowItemMasterMain["メーカー売価_画像"].ToString().ToString() == "-")
            {
                dtRowYahoo["original-price"] = "";
            }
            else
            {
                dtRowYahoo["original-price"] = dtRowItemMasterMain["メーカー売価_税込み"].ToString();
            }

            //original-price-evidence
            if (dtRowItemMasterMain["メーカー売価_画像"].ToString().ToString() == "-")
            {
                dtRowYahoo["original-price-evidence"] = "";
            }
            else
            {
                dtRowYahoo["original-price-evidence"] = dtRowItemMasterMain["メーカー売価_画像"].ToString();
            }

            //lead-time-instock
            dtRowYahoo["lead-time-instock"] = "1";



            var rowsSKU = dtRowItemMasterSKU.Select($"商品コード = '{商品コード}'");
            if (1 == rowsSKU.Length)
            {
                Debug.WriteLine($"{商品コード}は単色商品");
            }
            else
            {
                Debug.WriteLine($"{商品コード}は複色商品");
                var listSubcode = new List<string>();
                var listOptions = new List<string>();
                var SKU項目名 = dtRowItemMasterMain["Y_SKU項目名"].ToString();
                foreach (var rowSKU in rowsSKU)
                {
                    var SKUコード = rowSKU["SKUコード"].ToString();
                    var 選択肢名 = rowSKU["選択肢名"].ToString();
                    Debug.WriteLine($"★{SKU項目名}:{選択肢名}:{SKUコード}");

                    //カラーをお選びください:ホワイト=3100000010&カラーをお選びください:ブラウン=3100000020
                    //カラーをお選びください ホワイト ブラウン
                    listSubcode.Add($"{SKU項目名}:{選択肢名}={SKUコード}");
                    listOptions.Add($"{選択肢名}");
                }
                var strSubcode = UtilityClass.ListToString(listSubcode, "&");
                dtRowYahoo["sub-code"] = strSubcode;
                var strOptions = UtilityClass.ListToString(listOptions, " ");
                dtRowYahoo["options"] = $"{SKU項目名} {strOptions}";
            }

            //display
            dtRowYahoo["display"] = "1";

            //delivery , ship-weight
            if (dtRowItemMasterMain["送料形態"].ToString() == "送料無料")
            {
                dtRowYahoo["delivery"] = "1";
                dtRowYahoo["ship-weight"] = "";
            }
            else if (dtRowItemMasterMain["送料形態"].ToString() == "条件付き送料無料")
            {
                dtRowYahoo["delivery"] = "3";
                dtRowYahoo["ship-weight"] = "";
            }
            else
            {
                dtRowYahoo["delivery"] = "0";
                dtRowYahoo["ship-weight"] = "";
            }

            //product-category
            dtRowYahoo["product-category"] = dtRowItemMasterMain["Y_カテゴリID"].ToString();

            //spec1～spec6
            if (dtRowItemMasterMain["Y_spec1"].ToString().ToString() == "-")
            {
                dtRowYahoo["spec1"] = "";
            }
            else
            {
                dtRowYahoo["spec1"] = dtRowItemMasterMain["Y_spec1"].ToString();
            }

            if (dtRowItemMasterMain["Y_spec2"].ToString().ToString() == "-")
            {
                dtRowYahoo["spec2"] = "";
            }
            else
            {
                dtRowYahoo["spec2"] = dtRowItemMasterMain["Y_spec2"].ToString();
            }

            if (dtRowItemMasterMain["Y_spec3"].ToString().ToString() == "-")
            {
                dtRowYahoo["spec3"] = "";
            }
            else
            {
                dtRowYahoo["spec3"] = dtRowItemMasterMain["Y_spec3"].ToString();
            }

            if (dtRowItemMasterMain["Y_spec4"].ToString().ToString() == "-")
            {
                dtRowYahoo["spec4"] = "";
            }
            else
            {
                dtRowYahoo["spec4"] = dtRowItemMasterMain["Y_spec4"].ToString();
            }

            if (dtRowItemMasterMain["Y_spec5"].ToString().ToString() == "-")
            {
                dtRowYahoo["spec5"] = "";
            }
            else
            {
                dtRowYahoo["spec5"] = dtRowItemMasterMain["Y_spec5"].ToString();
            }

            if (dtRowItemMasterMain["Y_spec6"].ToString().ToString() == "-")
            {
                dtRowYahoo["spec6"] = "";
            }
            else
            {
                dtRowYahoo["spec6"] = dtRowItemMasterMain["Y_spec6"].ToString();
            }

            if (dtRowItemMasterMain["Y_spec7"].ToString().ToString() == "-")
            {
                dtRowYahoo["spec7"] = "";
            }
            else
            {
                dtRowYahoo["spec7"] = dtRowItemMasterMain["Y_spec7"].ToString();
            }

            if (dtRowItemMasterMain["Y_spec8"].ToString().ToString() == "-")
            {
                dtRowYahoo["spec8"] = "";
            }
            else
            {
                dtRowYahoo["spec8"] = dtRowItemMasterMain["Y_spec8"].ToString();
            }

            if (dtRowItemMasterMain["Y_spec9"].ToString().ToString() == "-")
            {
                dtRowYahoo["spec9"] = "";
            }
            else
            {
                dtRowYahoo["spec9"] = dtRowItemMasterMain["Y_spec9"].ToString();
            }

            if (dtRowItemMasterMain["Y_spec10"].ToString().ToString() == "-")
            {
                dtRowYahoo["spec10"] = "";
            }
            else
            {
                dtRowYahoo["spec10"] = dtRowItemMasterMain["Y_spec10"].ToString();
            }


            //sort
            if (dtRowItemMasterMain["ソート"].ToString().ToString() == "-")
            {
                dtRowYahoo["sort"] = "";
            }
            else
            {
                dtRowYahoo["sort"] = dtRowItemMasterMain["ソート"].ToString();
            }

            //meta-key
            if (dtRowItemMasterMain["Y_metakey"].ToString() == "-")
            {
                dtRowYahoo["meta-key"] = "";
            }
            else
            {
                dtRowYahoo["meta-key"] = dtRowItemMasterMain["Y_metakey"].ToString();
            }

            //meta-desc
            if (dtRowItemMasterMain["Y_metadesc"].ToString() == "-")
            {
                dtRowYahoo["meta-desc"] = "";
            }
            else
            {
                string metadesc = dtRowItemMasterMain["Y_metadesc"].ToString();
                metadesc = metadesc.Replace("<br>", "");
                metadesc = metadesc.Replace("<BR>", "");
                dtRowYahoo["meta-desc"] = metadesc;
            }

            //abstract
            if (dtRowItemMasterMain["Y_abstract"].ToString() == "-")
            {
                dtRowYahoo["abstract"] = "";
            }
            else
            {
                dtRowYahoo["abstract"] = dtRowItemMasterMain["Y_abstract"].ToString();
            }

            //relevant-links
            if (dtRowItemMasterMain["relevant_links"].ToString() == "-")
            {
                dtRowYahoo["relevant-links"] = "";
            }
            else
            {
                dtRowYahoo["relevant-links"] = dtRowItemMasterMain["relevant_links"].ToString();
            }

            //taxable
            dtRowYahoo["taxable"] = "1";

            //template
            dtRowYahoo["template"] = "IT02";

            //astk-code
            dtRowYahoo["astk-code"] = "0";

            //condition
            dtRowYahoo["condition"] = "0";

            //taojapan
            dtRowYahoo["taojapan"] = "1";

            //カゴ横説明
            if (generateKagoyoko.Length < 4)
            {
                MessageBox.Show("かご横説明がありません");
                generateKagoyoko = "";
            }

            //explanation
            var explanation = generateKagoyoko.Replace("&nbsp;　", "");
            explanation = explanation.Replace("<br>", "");
            explanation = Regex.Replace(explanation, "<.*?>", "");
            dtRowYahoo["explanation"] = explanation;
            UtilityClass.OutputHtml(outputDirectory + @"\" + "html_yahoo_explanation", "<br><hr><br>" + dtRowYahoo["explanation"].ToString());

            //画像説明
            string 画像説明 = "<table border=\"0\" cellspacing=\"0\" cellpadding=\"2\"><tr><td>\r\n" + dtRowItemMasterMain["画像説明"].ToString() + "\r\n</td></tr></table>";

            string yahoo_アディショナル = 画像説明.Replace("./説明用/", "https://shopping.c.yimg.jp/lib/" + 店舗コード + "/");
            yahoo_アディショナル = yahoo_アディショナル.Replace("./サムネイル/", "https://shopping.c.yimg.jp/lib/" + 店舗コード + "/");
            yahoo_アディショナル = yahoo_アディショナル.Replace(".thum", ".jpg");
            yahoo_アディショナル = yahoo_アディショナル.Replace("./商品URL/", "https://store.shopping.yahoo.co.jp/" + 店舗コード + "/");
            yahoo_アディショナル = System.Text.RegularExpressions.Regex.Replace(yahoo_アディショナル, "<!--楽天-->[\\s\\S]*?<!--/楽天-->", "");
            yahoo_アディショナル = System.Text.RegularExpressions.Regex.Replace(yahoo_アディショナル, "<!--Yahoo-->", "");
            yahoo_アディショナル = System.Text.RegularExpressions.Regex.Replace(yahoo_アディショナル, "<!--/Yahoo-->", "");

            //テーブル追加する caption=テーブル+画像説明
            string load_spgentei_yahoo = load_spgentei_yahoo = UtilityClass.ReadTextFile(ecPath + @"templates\load_spgentei_yahoo.html");

            string 注意事項 = UtilityClass.ReadTextFile(ecPath + @"templates\【注意事項SP】" + dtRowItemMasterMain["注意事項"].ToString() + ".html");

            //sp-additional
            var スマホ限定_yahoo = load_spgentei_yahoo + "<br>\r\n" + yahoo_アディショナル + "<br>\r\n" + 注意事項 + "<br>\r\n<br>\r\n" + generateSpKanrenShohin;
            スマホ限定_yahoo = スマホ限定_yahoo.Replace("./説明用/", "https://shopping.c.yimg.jp/lib/" + 店舗コード + "/");
            スマホ限定_yahoo = スマホ限定_yahoo.Replace("./サムネイル/", "https://shopping.c.yimg.jp/lib/" + 店舗コード + "/");
            スマホ限定_yahoo = スマホ限定_yahoo.Replace(".thum", ".jpg");
            スマホ限定_yahoo = スマホ限定_yahoo.Replace(".itempage", ".html");
            スマホ限定_yahoo = スマホ限定_yahoo.Replace("./商品URL/", "https://store.shopping.yahoo.co.jp/" + 店舗コード + "/");
            スマホ限定_yahoo = スマホ限定_yahoo.Replace("./アイコン/", "https://shopping.c.yimg.jp/lib/" + 店舗コード + "/setu");
            スマホ限定_yahoo = スマホ限定_yahoo.Replace("./特設ページ/", "https://store.shopping.yahoo.co.jp/" + 店舗コード + "/");
            スマホ限定_yahoo = スマホ限定_yahoo.Replace("./注意事項/", "https://shopping.c.yimg.jp/lib/" + 店舗コード + "/");
            スマホ限定_yahoo = スマホ限定_yahoo.Replace("./ページ/", "https://shopping.c.yimg.jp/lib/" + 店舗コード + "/");
            スマホ限定_yahoo = スマホ限定_yahoo.Replace("./カテゴリ/", "https://store.shopping.yahoo.co.jp/" + 店舗コード + "/");
            スマホ限定_yahoo = System.Text.RegularExpressions.Regex.Replace(スマホ限定_yahoo, "<!--楽天-->[\\s\\S]*?<!--/楽天-->", "");
            スマホ限定_yahoo = System.Text.RegularExpressions.Regex.Replace(スマホ限定_yahoo, "<!--Yahoo-->", "");
            スマホ限定_yahoo = System.Text.RegularExpressions.Regex.Replace(スマホ限定_yahoo, "<!--/Yahoo-->", "");
            スマホ限定_yahoo = System.Text.RegularExpressions.Regex.Replace(スマホ限定_yahoo, "<font.*?>", "");
            スマホ限定_yahoo = System.Text.RegularExpressions.Regex.Replace(スマホ限定_yahoo, "</font>", "");
            dtRowYahoo["sp-additional"] = スマホ限定_yahoo;
            UtilityClass.OutputHtml(outputDirectory + @"\" + "html_yahoo_sp-additional", "<br><hr><br>" + dtRowYahoo["sp-additional"].ToString());

            //caption , additional1
            string Yahoo_テーブル = "";
            if (generateTable.ToString().Length > 10)
            {
                Yahoo_テーブル = generateTable.ToString();
                //Yahoo_テーブル = System.Text.RegularExpressions.Regex.Replace(Yahoo_テーブル, "<!--関連商品-->[\\s\\S]*?<!--/関連商品-->", "");
                Yahoo_テーブル = System.Text.RegularExpressions.Regex.Replace(Yahoo_テーブル, "<!--関連商品-->", "");
                Yahoo_テーブル = System.Text.RegularExpressions.Regex.Replace(Yahoo_テーブル, "<!--/関連商品-->", "");
                //関連商品テーブル 改修
                Yahoo_テーブル = Yahoo_テーブル.Replace("./説明用/", "https://shopping.c.yimg.jp/lib/" + 店舗コード + "/");
                Yahoo_テーブル = Yahoo_テーブル.Replace("./サムネイル/", "https://shopping.c.yimg.jp/lib/" + 店舗コード + "/");
                Yahoo_テーブル = Yahoo_テーブル.Replace(".thum", ".jpg");
                Yahoo_テーブル = Yahoo_テーブル.Replace(".itempage", ".html");
                Yahoo_テーブル = Yahoo_テーブル.Replace("./商品URL/", "https://store.shopping.yahoo.co.jp/" + 店舗コード + "/");
                Yahoo_テーブル = Yahoo_テーブル.Replace("./アイコン/", "https://shopping.c.yimg.jp/lib/" + 店舗コード + "/setu");
                Yahoo_テーブル = Yahoo_テーブル.Replace("./特設ページ/", "https://store.shopping.yahoo.co.jp/" + 店舗コード + "/");
                Yahoo_テーブル = Yahoo_テーブル.Replace("./注意事項/", "https://shopping.c.yimg.jp/lib/" + 店舗コード + "/");
                Yahoo_テーブル = Yahoo_テーブル.Replace("./ページ/", "https://shopping.c.yimg.jp/lib/" + 店舗コード + "/");
                Yahoo_テーブル = Yahoo_テーブル.Replace("./イメージス/", "https://shopping.c.yimg.jp/lib/" + 店舗コード + "/");
                Yahoo_テーブル = Yahoo_テーブル.Replace("./カテゴリ/", "https://store.shopping.yahoo.co.jp/" + 店舗コード + "/");
                Yahoo_テーブル = System.Text.RegularExpressions.Regex.Replace(Yahoo_テーブル, "<!--楽天-->[\\s\\S]*?<!--/楽天-->", "");
                Yahoo_テーブル = System.Text.RegularExpressions.Regex.Replace(Yahoo_テーブル, "<!--Yahoo-->", "");
                Yahoo_テーブル = System.Text.RegularExpressions.Regex.Replace(Yahoo_テーブル, "<!--/Yahoo-->", "");
                dtRowYahoo["additional1"] = yahoo_アディショナル;
                UtilityClass.OutputHtml(outputDirectory + @"\" + "html_yahoo_アディショナル", "<br><hr><br>" + dtRowYahoo["additional1"].ToString());
                string caption = Yahoo_テーブル.Replace("<CENTER>", "");
                caption = caption.Replace("</CENTER>", "");
                dtRowYahoo["caption"] = caption;//pc_onry
                UtilityClass.OutputHtml(outputDirectory + @"\" + "html_yahoo_キャプション", "<br><hr><br>" + dtRowYahoo["caption"].ToString());
            }
            else
            {
                dtRowYahoo["caption"] = "";
                dtRowYahoo["additional1"] = "";
            }


            //item-image-urls
            var tmpItemImageUrls = "";
            for (var i = 0; i < 20; i++)
            {
                //サムネイル
                var listサムネイル = UtilityClass.StaticGetFindcodeList一行(dtRowItemMasterMain["画像説明"].ToString(), "<IMG SRC=\"./説明用/(?<findcode>.*?)\" border=");
                var サムネイル枚数 = listサムネイル.Count;
                if (i < サムネイル枚数)
                {
                    var url = $"https://shopping.c.yimg.jp/lib/taiho-kagu/{listサムネイル[i]};";
                    tmpItemImageUrls = $"{tmpItemImageUrls}{url}";
                }
                else
                {
                    var url = $";";
                    tmpItemImageUrls = $"{tmpItemImageUrls}{url}";

                }
            }
            dtRowYahoo["item-image-urls"] = tmpItemImageUrls;

            dtYahooItem.Rows.Add(dtRowYahoo);
            return dtYahooItem;
        }
        public DataTable GenerateCsvYahooOption(DataTable dtYahooOption, DataRow dtRowItemMasterMain, DataTable dtRowItemMasterSKU)
        {
            var 商品コード = dtRowItemMasterMain["mycode"].ToString();
            var rowsSKU = dtRowItemMasterSKU.Select($"商品コード = '{商品コード}'");
            Debug.WriteLine($"{商品コード}のSKUが{rowsSKU.Length}件見つかりました。");


            //SKU
            if (1 == rowsSKU.Length)
            {
                Debug.WriteLine($"{商品コード}は単色商品");
            }
            else
            {
                Debug.WriteLine($"{商品コード}は複色商品");
                foreach (var rowSKU in rowsSKU)
                {
                    var rowSKUレベル = dtYahooOption.NewRow();
                    var SKUコード = rowSKU["SKUコード"].ToString();


                    rowSKUレベル["code"] = 商品コード;
                    rowSKUレベル["sub-code"] = SKUコード;
                    rowSKUレベル["option-name-1"] = dtRowItemMasterMain["Y_SKU項目名"].ToString();
                    rowSKUレベル["option-value-1"] = rowSKU["選択肢名"].ToString();
                    rowSKUレベル["unselectable-1"] = 0;


                    //2番目3番目
                    if (SKUコード.Length != 10)
                    {
                        MessageBox.Show($"エラー：{SKUコード}の文字数が10桁ではありません");
                    }
                    var 下3 = int.Parse(UtilityClass.Mid(SKUコード, 8, 1));
                    var 下2 = int.Parse(UtilityClass.Mid(SKUコード, 9, 1));
                    var 下1 = int.Parse(UtilityClass.Mid(SKUコード, 10, 1));

                    var 番号 = int.Parse($"{下3}{下2}");
                    if (下1 != 0)
                    {
                        MessageBox.Show($"SKUの数は99が最大です。{商品コード} 商品管理番号下3桁{下3}{下2}{下1}");
                    }
                    rowSKUレベル["sub-code-img1"] = $"https://shopping.c.yimg.jp/lib/taiho-kagu/{商品コード}_{番号}.jpg";
                    rowSKUレベル["main-flag"] = "0";
                    rowSKUレベル["exist-flag"] = "1";
                    rowSKUレベル["option-charge-1"] = "";
                    dtYahooOption.Rows.Add(rowSKUレベル);
                }
            }

            if (1 < dtRowItemMasterMain["Y_商品プルダウン"].ToString().Length)
            {
                var Yahoo_商品プルダウン = dtRowItemMasterMain["Y_商品プルダウン"].ToString();
                Yahoo_商品プルダウン = Regex.Replace(Yahoo_商品プルダウン, "\r\n", "★改行★");
                Yahoo_商品プルダウン = Regex.Replace(Yahoo_商品プルダウン, "\r", "★改行★");
                Yahoo_商品プルダウン = Regex.Replace(Yahoo_商品プルダウン, "\n", "★改行★");

                var Yahoo_商品プルダウンList = UtilityClass.StringToList(Yahoo_商品プルダウン, "★改行★"); //Excelの改行は\n
                var counter = 0;
                foreach (var プルダウンLINE in Yahoo_商品プルダウンList)
                {
                    if (!プルダウンLINE.Contains(" "))
                    {
                        continue;
                    }
                    Debug.WriteLine($"{counter}行目：{プルダウンLINE}");
                    counter = counter + 1;

                    //プルダウンLINE
                    var splitListプルダウンライン = UtilityClass.StringToList(プルダウンLINE, " ");
                    for (int i = 0; i < splitListプルダウンライン.Count; i++)
                    {
                        if (i == 0)
                        {
                            continue;
                        }

                        var rowSKUレベル = dtYahooOption.NewRow();
                        rowSKUレベル["code"] = 商品コード;
                        rowSKUレベル["sub-code"] = "";
                        rowSKUレベル["option-name-1"] = splitListプルダウンライン[0];
                        rowSKUレベル["option-value-1"] = splitListプルダウンライン[i];

                        //選択必須 必ず選択してください。を設けられない？かも？
                        if (splitListプルダウンライン[i].Contains("選択して"))
                        {
                            rowSKUレベル["unselectable-1"] = 1;
                        }
                        else
                        {
                            rowSKUレベル["unselectable-1"] = 0;
                        }
                        rowSKUレベル["sub-code-img1"] = "";
                        rowSKUレベル["main-flag"] = "";
                        rowSKUレベル["exist-flag"] = "";
                        rowSKUレベル["option-charge-1"] = "";
                        dtYahooOption.Rows.Add(rowSKUレベル);
                    }
                }
            }

            if (1 < dtRowItemMasterMain["Y_別途送料地域選択肢"].ToString().Length)
            {
                foreach (var 別途送料地域 in dtRowItemMasterMain["Y_別途送料地域選択肢"].ToString().Split(' '))
                {
                    var rowSKUレベル = dtYahooOption.NewRow();

                    var 別途送料地域_地域名 = 別途送料地域.Split(':')[0];
                    var 別途送料地域_料金 = 別途送料地域.Split(':')[1];


                    rowSKUレベル["code"] = 商品コード;
                    rowSKUレベル["sub-code"] = "";
                    rowSKUレベル["option-name-1"] = dtRowItemMasterMain["Y_別途送料地域項目名"].ToString();
                    rowSKUレベル["option-value-1"] = 別途送料地域_地域名;

                    if (別途送料地域_地域名.Contains("選択して"))
                    {
                        rowSKUレベル["unselectable-1"] = 1;
                    }
                    else
                    {
                        rowSKUレベル["unselectable-1"] = 0;
                    }
                    rowSKUレベル["sub-code-img1"] = "";
                    rowSKUレベル["main-flag"] = "";
                    rowSKUレベル["exist-flag"] = "";
                    if (別途送料地域_料金 == "0")
                    {
                        rowSKUレベル["option-charge-1"] = "";
                    }
                    else
                    {
                        rowSKUレベル["option-charge-1"] = 別途送料地域_料金;
                    }
                    dtYahooOption.Rows.Add(rowSKUレベル);
                }
            }

            if (1 < dtRowItemMasterMain["Y_配達オプション選択肢"].ToString().Length)
            {
                foreach (var 配達オプション in dtRowItemMasterMain["Y_配達オプション選択肢"].ToString().Split(' '))
                {
                    var rowSKUレベル = dtYahooOption.NewRow();

                    var 配達オプション_オプション名 = 配達オプション.Split(':')[0];
                    var 配達オプション_料金 = 配達オプション.Split(':')[1];


                    rowSKUレベル["code"] = 商品コード;
                    rowSKUレベル["sub-code"] = "";
                    rowSKUレベル["option-name-1"] = dtRowItemMasterMain["Y_配達オプション項目名"].ToString();
                    rowSKUレベル["option-value-1"] = 配達オプション_オプション名;

                    if (配達オプション_オプション名.Contains("選択して"))
                    {
                        rowSKUレベル["unselectable-1"] = 1;
                    }
                    else
                    {
                        rowSKUレベル["unselectable-1"] = 0;
                    }
                    rowSKUレベル["sub-code-img1"] = "";
                    rowSKUレベル["main-flag"] = "";
                    rowSKUレベル["exist-flag"] = "";
                    if (配達オプション_料金 == "0")
                    {
                        rowSKUレベル["option-charge-1"] = "";
                    }
                    else
                    {
                        rowSKUレベル["option-charge-1"] = 配達オプション_料金;
                    }
                    dtYahooOption.Rows.Add(rowSKUレベル);
                }
            }

            if (1 < dtRowItemMasterMain["Y_注意事項プルダウン"].ToString().Length)
            {
                var 注意事項プルダウン = UtilityClass.ReadTextFile(ecPath + @"templates\【選択肢】" + dtRowItemMasterMain["Y_注意事項プルダウン"].ToString() + "_Y.txt");
                注意事項プルダウン = Regex.Replace(注意事項プルダウン, "\r\n", "★改行★");
                注意事項プルダウン = Regex.Replace(注意事項プルダウン, "\r", "★改行★");
                注意事項プルダウン = Regex.Replace(注意事項プルダウン, "\n", "★改行★");

                var Yahoo_商品プルダウンList = UtilityClass.StringToList(注意事項プルダウン, "★改行★"); //Excelの改行は\n
                var counter = 0;
                foreach (var プルダウンLINE in Yahoo_商品プルダウンList)
                {
                    if (!プルダウンLINE.Contains(" "))
                    {
                        continue;
                    }
                    Debug.WriteLine($"{counter}行目：{プルダウンLINE}");
                    counter = counter + 1;

                    //プルダウンLINE
                    var splitListプルダウンライン = UtilityClass.StringToList(プルダウンLINE, " ");
                    for (int i = 0; i < splitListプルダウンライン.Count; i++)
                    {
                        if (i == 0)
                        {
                            continue;
                        }

                        var rowSKUレベル = dtYahooOption.NewRow();
                        rowSKUレベル["code"] = 商品コード;
                        rowSKUレベル["sub-code"] = "";
                        rowSKUレベル["option-name-1"] = splitListプルダウンライン[0];
                        rowSKUレベル["option-value-1"] = splitListプルダウンライン[i];

                        //選択必須 必ず選択してください。を設けられない？かも？
                        if (splitListプルダウンライン[i].Contains("選択して"))
                        {
                            rowSKUレベル["unselectable-1"] = 1;
                        }
                        else
                        {
                            rowSKUレベル["unselectable-1"] = 0;
                        }
                        rowSKUレベル["sub-code-img1"] = "";
                        rowSKUレベル["main-flag"] = "";
                        rowSKUレベル["exist-flag"] = "";
                        rowSKUレベル["option-charge-1"] = "";
                        dtYahooOption.Rows.Add(rowSKUレベル);
                    }
                }
            }
            return dtYahooOption;
        }



        public DataTable GenerateCsvYahooAuctionItem(DataTable dtYahooAuctionItem, DataRow dtRowItemMasterMain, DataTable dtRowItemMasterSKU, DataTable dtRowItemMasterIcon, DataTable dtYahooItem)
        {
            var 商品コード = dtRowItemMasterMain["mycode"].ToString();
            var 色_1 = dtRowItemMasterMain["色_1"].ToString();
            var YA_商品名 = dtRowItemMasterMain["Y_商品名"].ToString(); //Yahooの商品名を取得
            var aucCategory = dtRowItemMasterMain["YA_カテゴリID"].ToString();
            var aucSuffix = dtRowItemMasterMain["YA_suffix"].ToString();
            var 注意事項ヤフオク = UtilityClass.ReadTextFile(ecPath + @"templates\【注意事項ヤフオク】.html");
            Debug.WriteLine($"{商品コード}");

            var rowsSKU = dtRowItemMasterSKU.Select($"商品コード = '{商品コード}'");
            Debug.WriteLine($"{商品コード}のSKUが{rowsSKU.Length}件見つかりました。");
            if (1 == rowsSKU.Length)
            {
                Debug.WriteLine($"ヤフオク : {商品コード}は単色商品");

                //行をコピー
                dtYahooAuctionItem.ImportRow(dtYahooItem.Select($"code = '{商品コード}'")[0]);
                var lastRow = dtYahooAuctionItem.Rows.Count - 1;

                //ヤフオク専用行追加
                if (true)
                {
                    dtYahooAuctionItem.Rows[lastRow]["auc-bcid"] = $"{商品コード}{aucSuffix}";
                    dtYahooAuctionItem.Rows[lastRow]["auc-category"] = $"{aucCategory}";
                    dtYahooAuctionItem.Rows[lastRow]["auc-store-keyword"] = $"{商品コード}";
                    dtYahooAuctionItem.Rows[lastRow]["auc-pref-code"] = $"40";
                    dtYahooAuctionItem.Rows[lastRow]["auc-city"] = $"";
                    dtYahooAuctionItem.Rows[lastRow]["condition"] = $"2";
                }

                //ヤフオク専用修正
                if (true)
                {
                    dtYahooAuctionItem.Rows[lastRow]["code"] = $"{商品コード}{aucSuffix}";
                    dtYahooAuctionItem.Rows[lastRow]["name"] = UtilityClass.文字数を制限(YA_商品名, " ", 120);
                    dtYahooAuctionItem.Rows[lastRow]["additional1"] = $"{注意事項ヤフオク.Replace("<<商品コード>>", $"{商品コード}{aucSuffix}")}【商品コード：{商品コード}{aucSuffix}】<br><br>{dtYahooAuctionItem.Rows[lastRow]["caption"]}";
                    dtYahooAuctionItem.Rows[lastRow]["explanation"] = $"{dtYahooAuctionItem.Rows[lastRow]["explanation"].ToString().Replace(商品コード, $"{商品コード}{aucSuffix}")}";
                    dtYahooAuctionItem.Rows[lastRow]["caption"] = $"";
                }

                //追加画像
                if (true)
                {
                    var サムネイル画像 = $"https://shopping.c.yimg.jp/lib/taiho-kagu/{商品コード}.jpg";
                    var 追加画像2枚目 = "https://shopping.c.yimg.jp/lib/taiho-kagu/auc_postage.jpg";
                    var 追加画像3枚目 = "https://shopping.c.yimg.jp/lib/taiho-kagu/auc_route.jpg";
                    var itemImage = $"{dtYahooAuctionItem.Rows[lastRow]["item-image-urls"]}";

                    dtYahooAuctionItem.Rows[lastRow]["item-image-urls"] = $"{dtYahooAuctionItem.Rows[lastRow]["item-image-urls"]}".Replace($"{サムネイル画像};", $"{サムネイル画像};{追加画像2枚目};{追加画像3枚目};");
                    dtYahooAuctionItem.Rows[lastRow]["item-image-urls"] = $"{dtYahooAuctionItem.Rows[lastRow]["item-image-urls"]}".Replace($".jpg;;;", $".jpg;");
                }

                //アイコンを置換する
                if (true)
                {
                    foreach (var rowIcon in dtRowItemMasterIcon.Select($"説明 <> ''"))
                    {
                        var キー = rowIcon["キー"].ToString();
                        var 説明 = rowIcon["説明"].ToString();
                        var アイコン画像 = $"<IMG src=\"https://shopping.c.yimg.jp/lib/taiho-kagu/setu{キー}.gif\" width=\"56\" height=\"56\" border=\"0\">";
                        dtYahooAuctionItem.Rows[lastRow]["additional1"] = $"{dtYahooAuctionItem.Rows[lastRow]["additional1"].ToString().Replace($"{アイコン画像}", $"●{説明} ")}";
                    }
                }

                //関連商品を削除する
                if(true)
                {
                    dtYahooAuctionItem.Rows[lastRow]["additional1"] = Regex.Replace($"{dtYahooAuctionItem.Rows[lastRow]["additional1"]}", "関連商品.*?説明マーク", "説明マーク", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                }
            }
            else
            {
                Debug.WriteLine($"ヤフオク : {商品コード}は複色商品");
                foreach (var rowSKU in rowsSKU)
                {
                    //管理番号が10桁のみ
                    var SKUコード = rowSKU["SKUコード"].ToString();
                    if (SKUコード.Length != 10)
                    {
                        MessageBox.Show($"エラー：{SKUコード}の文字数が10桁ではありません");
                    }
                    var 下3 = int.Parse(UtilityClass.Mid(SKUコード, 8, 1));
                    var 下2 = int.Parse(UtilityClass.Mid(SKUコード, 9, 1));
                    var 下1 = int.Parse(UtilityClass.Mid(SKUコード, 10, 1));
                    var 番号 = int.Parse($"{下3}{下2}");
                    if (下1 != 0)
                    {
                        MessageBox.Show($"SKUの数は99が最大です。{商品コード} 商品管理番号下3桁{下3}{下2}{下1}");
                    }
                    var 選択肢名 = rowSKU["選択肢名"].ToString();

                    //行をコピー
                    dtYahooAuctionItem.ImportRow(dtYahooItem.Select($"code = '{商品コード}'")[0]);
                    var lastRow = dtYahooAuctionItem.Rows.Count - 1;

                    //ヤフオク専用行追加
                    if (true)
                    {
                        dtYahooAuctionItem.Rows[lastRow]["auc-bcid"] = $"{SKUコード}{aucSuffix}";
                        dtYahooAuctionItem.Rows[lastRow]["auc-category"] = $"{aucCategory}";
                        dtYahooAuctionItem.Rows[lastRow]["auc-store-keyword"] = $"{SKUコード}";
                        dtYahooAuctionItem.Rows[lastRow]["auc-pref-code"] = $"40";
                        dtYahooAuctionItem.Rows[lastRow]["auc-city"] = $"";
                        dtYahooAuctionItem.Rows[lastRow]["condition"] = $"2";
                    }

                    //ヤフオク専用修正
                    if (true)
                    {
                        dtYahooAuctionItem.Rows[lastRow]["code"] = $"{SKUコード}{aucSuffix}";
                        var カラーバイト = 120 - Encoding.GetEncoding("Shift_JIS").GetByteCount($" {選択肢名}");
                        var 短い商品名 = UtilityClass.文字数を制限(YA_商品名, " ", カラーバイト);
                        dtYahooAuctionItem.Rows[lastRow]["name"] = $"{短い商品名} {選択肢名}"; //文字数を削除する
                        dtYahooAuctionItem.Rows[lastRow]["sub-code"] = $"";
                        dtYahooAuctionItem.Rows[lastRow]["options"] = $""; //optionCSVで作成 列ごと削除が良い
                        dtYahooAuctionItem.Rows[lastRow]["additional1"] = $"{注意事項ヤフオク.Replace("<<商品コード>>", $"{SKUコード}{aucSuffix}")}【商品コード：{SKUコード}{aucSuffix}】<br><br>{dtYahooAuctionItem.Rows[lastRow]["caption"]}<br>※当ページの販売カラーは「{選択肢名}」です。";
                        dtYahooAuctionItem.Rows[lastRow]["additional1"] = $"{dtYahooAuctionItem.Rows[lastRow]["additional1"].ToString().Replace(色_1, $"●{選択肢名}")}";
                        dtYahooAuctionItem.Rows[lastRow]["explanation"] = $"{dtYahooAuctionItem.Rows[lastRow]["explanation"].ToString().Replace(商品コード, $"{SKUコード}{aucSuffix}")}";
                        dtYahooAuctionItem.Rows[lastRow]["explanation"] = $"{dtYahooAuctionItem.Rows[lastRow]["explanation"].ToString().Replace(色_1, $"●{選択肢名}")}";
                        dtYahooAuctionItem.Rows[lastRow]["caption"] = $"";
                    }

                    //サムネイル処理
                    if (true)
                    {
                        //画像の順序を入れ替える
                        var サムネイル画像 = $"https://shopping.c.yimg.jp/lib/taiho-kagu/{商品コード}.jpg";
                        var SKU画像URL = $"https://shopping.c.yimg.jp/lib/taiho-kagu/{商品コード}_{番号}.jpg";
                        var itemImage = $"{dtYahooAuctionItem.Rows[lastRow]["item-image-urls"]}";
                        itemImage = itemImage.Replace($"{サムネイル画像};", "");
                        itemImage = itemImage.Replace($"{SKU画像URL};", "");

                        //追加画像
                        var 追加画像2枚目 = "https://shopping.c.yimg.jp/lib/taiho-kagu/auc_postage.jpg";
                        var 追加画像3枚目 = "https://shopping.c.yimg.jp/lib/taiho-kagu/auc_route.jpg";

                        dtYahooAuctionItem.Rows[lastRow]["item-image-urls"] = $"{SKU画像URL};{追加画像2枚目};{追加画像3枚目};{サムネイル画像};{itemImage}";
                        dtYahooAuctionItem.Rows[lastRow]["item-image-urls"] = $"{dtYahooAuctionItem.Rows[lastRow]["item-image-urls"]}".Replace($".jpg;;;", $".jpg;");
                    }

                    //不要画像を削除する
                    if (true)
                    {
                        foreach (var rowSKUXXX in dtRowItemMasterSKU.Select($"商品コード = '{商品コード}'"))
                        {
                            //管理番号が10桁のみ
                            var SKUコードXXX = rowSKUXXX["SKUコード"].ToString();
                            if (SKUコードXXX.Length != 10)
                            {
                                MessageBox.Show($"エラー：{SKUコードXXX}の文字数が10桁ではありません");
                            }
                            var 下3XXX = int.Parse(UtilityClass.Mid(SKUコードXXX, 8, 1));
                            var 下2XXX = int.Parse(UtilityClass.Mid(SKUコードXXX, 9, 1));
                            var 下1XXX = int.Parse(UtilityClass.Mid(SKUコードXXX, 10, 1));
                            var 番号XXX = int.Parse($"{下3XXX}{下2XXX}");
                            if (下1XXX != 0)
                            {
                                MessageBox.Show($"SKUの数は99が最大です。{商品コード} 商品管理番号XXX下3XXX桁{下3XXX}{下2XXX}{下1XXX}");
                            }
                            var 選択肢名XXX = rowSKUXXX["選択肢名"].ToString();
                            if (番号 != 番号XXX)
                            {
                                //不要画像
                                var 不要サムネイルURL = $"https://shopping.c.yimg.jp/lib/taiho-kagu/{商品コード}_{番号XXX}.jpg";
                                dtYahooAuctionItem.Rows[lastRow]["item-image-urls"] = $"{dtYahooAuctionItem.Rows[lastRow]["item-image-urls"].ToString().Replace($"{不要サムネイルURL};", "")};"; //末尾に;
                                var 不要ロングページURL = $"<IMG SRC=\"https://shopping.c.yimg.jp/lib/taiho-kagu/{商品コード}_{番号XXX}.jpg\" border=\"0\" width=\"100%\"><BR><BR>";
                                dtYahooAuctionItem.Rows[lastRow]["sp-additional"] = $"{dtYahooAuctionItem.Rows[lastRow]["sp-additional"].ToString().Replace(不要ロングページURL, "")}";
                            }
                        }
                    }

                    //アイコンを置換する
                    //
                    if (true)
                    {
                        foreach (var rowIcon in dtRowItemMasterIcon.Select($"説明 <> ''"))
                        {
                            var キー = rowIcon["キー"].ToString();
                            var 説明 = rowIcon["説明"].ToString();
                            var アイコン画像 = $"<IMG src=\"https://shopping.c.yimg.jp/lib/taiho-kagu/setu{キー}.gif\" width=\"56\" height=\"56\" border=\"0\">";
                            dtYahooAuctionItem.Rows[lastRow]["additional1"] = $"{dtYahooAuctionItem.Rows[lastRow]["additional1"].ToString().Replace($"{アイコン画像}", $"●{説明} ")}";
                        }
                    }

                    //関連商品を削除する
                    if (true)
                    {
                        dtYahooAuctionItem.Rows[lastRow]["additional1"] = Regex.Replace($"{dtYahooAuctionItem.Rows[lastRow]["additional1"]}", "関連商品.*?説明マーク", "説明マーク", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                    }
                }
            }
            return dtYahooAuctionItem;
        }


        public DataTable GenerateCsvYahooAuctionOption(DataTable dtYahooAuctionOption, DataRow dtRowItemMasterMain, DataTable dtRowItemMasterSKU, DataTable dtYahooOption)
        {
            var 商品コード = dtRowItemMasterMain["mycode"].ToString();
            var aucSuffix = dtRowItemMasterMain["YA_suffix"].ToString();
            Debug.WriteLine($"{商品コード}");

            var rowsSKU = dtRowItemMasterSKU.Select($"商品コード = '{商品コード}'");
            Debug.WriteLine($"{商品コード}のSKUが{rowsSKU.Length}件見つかりました。");
            if (1 == rowsSKU.Length)
            {
                Debug.WriteLine($"ヤフオク : {商品コード}は単色商品");
                foreach (var row in dtYahooOption.Select($"code = '{商品コード}'"))
                {
                    dtYahooAuctionOption.ImportRow(row);
                    var lastRow = dtYahooAuctionOption.Rows.Count - 1;
                    dtYahooAuctionOption.Rows[lastRow]["code"] = $"{商品コード}{aucSuffix}";
                }
            }
            else
            {
                Debug.WriteLine($"ヤフオク : {商品コード}は複色商品");
                foreach (var rowSKU in rowsSKU)
                {
                    //管理番号が10桁のみ
                    var SKUコード = rowSKU["SKUコード"].ToString();

                    foreach (var row in dtYahooOption.Select($"code = '{商品コード}'"))
                    {
                        if (row["sub-code"].ToString() == "")
                        {
                            dtYahooAuctionOption.ImportRow(row);
                            var lastRow = dtYahooAuctionOption.Rows.Count - 1;
                            dtYahooAuctionOption.Rows[lastRow]["code"] = $"{SKUコード}{aucSuffix}";
                        }
                    }
                }
            }
            return dtYahooAuctionOption;
        }





        public DataTable GenerateCsvYahooAuctionOptionGOMI(DataTable dtYahooAuctionOption, DataTable dtYahooOption, DataTable dtMasterItemMain)
        {














































            dtYahooAuctionOption = dtYahooOption.Copy();

            //カラーの出力を削除
            foreach (var row in dtYahooAuctionOption.Select($"[sub-code] <> ''"))
            {
                row.Delete();
            }

            //商品コードにsuffix
            foreach (var rowYahooAuctionOption in dtYahooAuctionOption.Select($"code <> ''"))
            {
                foreach (var rowMasterItemMain in dtMasterItemMain.Select($"mycode = '{rowYahooAuctionOption["code"]}'"))
                {
                    var suffix = $"{rowMasterItemMain["auc-suffix"]}";
                    rowYahooAuctionOption["code"] = $"{rowYahooAuctionOption["code"]}{suffix}";
                }
            }
            return dtYahooAuctionOption;
        }


        //空のヘッダ付きDataTableを作成する
        public DataTable NewDatatableRakutenItem()
        {
            var dtRakutenItem = new DataTable();

            //商品レベル
            dtRakutenItem.Columns.Add("商品管理番号（商品URL）", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品番号", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品名", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("倉庫指定", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("サーチ表示", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("消費税", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("消費税率", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("販売期間指定（開始日時）", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("販売期間指定（終了日時）", Type.GetType("System.String"));
            //dtRakutenItem.Columns.Add("運用型ポイント変倍", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("ポイント変倍率", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("ポイント変倍率適用期間（開始日時）", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("ポイント変倍率適用期間（終了日時）", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("注文ボタン", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("予約商品発売日", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品問い合わせボタン", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("闇市パスワード", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("在庫表示", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("代引料", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("ジャンルID", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("非製品属性タグID", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("キャッチコピー", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("PC用商品説明文", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("スマートフォン用商品説明文", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("PC用販売説明文", Type.GetType("System.String"));
            //dtRakutenItem.Columns.Add("医薬品説明文", Type.GetType("System.String"));
            //dtRakutenItem.Columns.Add("医薬品注意事項", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像タイプ1", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像パス1", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像名（ALT）1", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像タイプ2", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像パス2", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像名（ALT）2", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像タイプ3", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像パス3", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像名（ALT）3", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像タイプ4", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像パス4", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像名（ALT）4", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像タイプ5", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像パス5", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像名（ALT）5", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像タイプ6", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像パス6", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像名（ALT）6", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像タイプ7", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像パス7", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像名（ALT）7", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像タイプ8", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像パス8", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像名（ALT）8", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像タイプ9", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像パス9", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像名（ALT）9", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像タイプ10", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像パス10", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像名（ALT）10", Type.GetType("System.String"));


            dtRakutenItem.Columns.Add("商品画像タイプ11", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像パス11", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像名（ALT）11", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像タイプ12", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像パス12", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像名（ALT）12", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像タイプ13", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像パス13", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像名（ALT）13", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像タイプ14", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像パス14", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像名（ALT）14", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像タイプ15", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像パス15", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像名（ALT）15", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像タイプ16", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像パス16", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像名（ALT）16", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像タイプ17", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像パス17", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像名（ALT）17", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像タイプ18", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像パス18", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像名（ALT）18", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像タイプ19", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像パス19", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像名（ALT）19", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像タイプ20", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像パス20", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品画像名（ALT）20", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("動画", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("白背景画像タイプ", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("白背景画像パス", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品情報レイアウト", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("ヘッダー・フッター・レフトナビ", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("表示項目の並び順", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("共通説明文（小）", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("目玉商品", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("共通説明文（大）", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("レビュー本文表示", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("メーカー提供情報表示", Type.GetType("System.String"));
            //dtRakutenItem.Columns.Add("定期購入ボタン", Type.GetType("System.String"));
            //dtRakutenItem.Columns.Add("指定可能なお届け日・月ごとに日付を指定", Type.GetType("System.String"));
            //dtRakutenItem.Columns.Add("指定可能なお届け日・週ごとに曜日を指定", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("バリエーション項目キー定義", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("バリエーション項目名定義", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("バリエーション1選択肢定義", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("バリエーション2選択肢定義", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("バリエーション3選択肢定義", Type.GetType("System.String"));

            //商品オプションレベル
            dtRakutenItem.Columns.Add("選択肢タイプ", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品オプション項目名", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品オプション選択肢1", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品オプション選択肢2", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品オプション選択肢3", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品オプション選択肢4", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品オプション選択肢5", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品オプション選択肢6", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品オプション選択肢7", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品オプション選択肢8", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("商品オプション選択必須", Type.GetType("System.String"));

            //SKUレベル
            dtRakutenItem.Columns.Add("SKU管理番号", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("システム連携用SKU番号", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("バリエーション項目キー1", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("バリエーション項目選択肢1", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("バリエーション項目キー2", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("バリエーション項目選択肢2", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("バリエーション項目キー3", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("バリエーション項目選択肢3", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("販売価格", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("表示価格", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("二重価格文言管理番号", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("注文受付数", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("再入荷お知らせボタン", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("のし対応", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("在庫数", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("在庫戻しフラグ", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("在庫切れ時の注文受付", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("在庫あり時納期管理番号", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("在庫切れ時納期管理番号", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("SKU倉庫指定", Type.GetType("System.String"));
            //dtRakutenItem.Columns.Add("あす楽配送管理番号", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("配送方法セット管理番号", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("送料", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("送料区分1", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("送料区分2", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("個別送料", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("地域別個別送料管理番号", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("単品配送設定使用", Type.GetType("System.String"));
            //dtRakutenItem.Columns.Add("海外配送管理番号", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("カタログID", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("カタログIDなしの理由", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("セット商品用カタログID", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("SKU画像タイプ", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("SKU画像パス", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("SKU画像名（ALT）", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("自由入力行（項目）1", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("自由入力行（値）1", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("自由入力行（項目）2", Type.GetType("System.String"));
            dtRakutenItem.Columns.Add("自由入力行（値）2", Type.GetType("System.String"));
            //dtRakutenItem.Columns.Add("定期商品販売価格", Type.GetType("System.String"));
            //dtRakutenItem.Columns.Add("初回価格", Type.GetType("System.String"));
            for (var i = 1; i < 41; i++)
            {
                dtRakutenItem.Columns.Add($"商品属性（項目）{i}", Type.GetType("System.String"));
                dtRakutenItem.Columns.Add($"商品属性（値）{i}", Type.GetType("System.String"));
                dtRakutenItem.Columns.Add($"商品属性（単位）{i}", Type.GetType("System.String"));
            }
            return dtRakutenItem;
        }
        public DataTable NewDatatableRakutenCategory()
        {
            var dtRakutenCategory = new DataTable();
            dtRakutenCategory.Columns.Add("コントロールカラム", Type.GetType("System.String"));
            dtRakutenCategory.Columns.Add("商品管理番号（商品URL）", Type.GetType("System.String"));
            dtRakutenCategory.Columns.Add("商品名", Type.GetType("System.String"));
            dtRakutenCategory.Columns.Add("表示先カテゴリ", Type.GetType("System.String"));
            dtRakutenCategory.Columns.Add("優先度", Type.GetType("System.String"));
            dtRakutenCategory.Columns.Add("URL", Type.GetType("System.String"));
            dtRakutenCategory.Columns.Add("1ページ複数形式", Type.GetType("System.String"));
            return dtRakutenCategory;
        }
        public DataTable NewDatatableYahooItem()
        {
            var dtYahooItem = new DataTable();
            dtYahooItem.Columns.Add("path", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("name", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("code", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("sub-code", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("original-price", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("price", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("sale-price", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("options", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("headline", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("caption", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("abstract", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("explanation", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("additional1", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("additional2", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("additional3", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("relevant-links", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("ship-weight", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("taxable", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("release-date", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("temporary-point-term", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("point-code", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("meta-key", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("meta-desc", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("template", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("sale-period-start", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("sale-period-end", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("sale-limit", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("sp-code", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("pr-rate", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("brand-code", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("person-code", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("yahoo-product-code", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("product-code", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("jan", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("isbn", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("delivery", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("astk-code", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("condition", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("taojapan", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("product-category", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("spec1", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("spec2", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("spec3", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("spec4", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("spec5", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("spec6", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("spec7", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("spec8", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("spec9", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("spec10", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("display", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("sort", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("sp-additional", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("sort_priority", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("original-price-evidence", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("lead-time-instock", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("lead-time-outstock", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("item-image-urls", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("auc-bcid", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("auc-category", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("auc-store-keyword", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("auc-pref-code", Type.GetType("System.String"));
            dtYahooItem.Columns.Add("auc-city", Type.GetType("System.String"));
            return dtYahooItem;
        }

        public DataTable NewDatatableYahooAuctionItem()
        {
            return NewDatatableYahooItem();
        }

        public DataTable NewDatatableYahooOption()
        {
            var dtYahooOption = new DataTable();
            dtYahooOption.Columns.Add("code", Type.GetType("System.String"));
            dtYahooOption.Columns.Add("sub-code", Type.GetType("System.String"));
            dtYahooOption.Columns.Add("option-name-1", Type.GetType("System.String"));
            dtYahooOption.Columns.Add("option-value-1", Type.GetType("System.String"));
            dtYahooOption.Columns.Add("unselectable-1", Type.GetType("System.String"));
            dtYahooOption.Columns.Add("sub-code-img1", Type.GetType("System.String"));
            dtYahooOption.Columns.Add("main-flag", Type.GetType("System.String"));
            dtYahooOption.Columns.Add("exist-flag", Type.GetType("System.String"));
            dtYahooOption.Columns.Add("option-charge-1", Type.GetType("System.String"));
            return dtYahooOption;
        }
        public DataTable NewDatatableYahooAuctionOption()
        {
            return NewDatatableYahooOption();
        }

        //出力したYahoo1号店のCSVデータをコピーし、2号店用に文字を"taiho-kagu"から"taiho-kagu2"へ置換する。
        public void CsvCopyYahooItem2号店()
        {
            var yahooSr = new StreamReader(outputDirectory + @"\yahoo_item.csv", Encoding.GetEncoding("Shift_JIS"));
            var yahooStr = yahooSr.ReadToEnd();
            yahooSr.Close();
            yahooStr = yahooStr.Replace("taiho-kagu", "taiho-kagu2");
            yahooStr = yahooStr.Replace("https://shopping.geocities.jp/taiho-kagu2/", "https://shopping.geocities.jp/taiho-kagu/"); //トリプルのみtaiho-kagu1
            System.IO.StreamWriter yahoo_sw = new System.IO.StreamWriter(outputDirectory + @"\yahoo2_item.csv", false, System.Text.Encoding.GetEncoding("shift_jis"));
            yahoo_sw.Write(yahooStr);
            yahoo_sw.Close();
            Console.WriteLine("Yahoo2号店のItemCSVを作成しました。");
        }
        public void CsvCopyYahooOption2号店()
        {
            var yahooSr = new StreamReader(outputDirectory + @"\yahoo_option.csv", Encoding.GetEncoding("Shift_JIS"));
            var yahooStr = yahooSr.ReadToEnd();
            yahooSr.Close();
            yahooStr = yahooStr.Replace("taiho-kagu", "taiho-kagu2");
            System.IO.StreamWriter yahoo_sw = new System.IO.StreamWriter(outputDirectory + @"\yahoo2_option.csv", false, System.Text.Encoding.GetEncoding("shift_jis"));
            yahoo_sw.Write(yahooStr);
            yahoo_sw.Close();
            Console.WriteLine("Yahoo2号店のOptionCSVを作成しました。");
        }
        public void 完了()
        {
            Process.Start(outputDirectory);
        }



    }
}
