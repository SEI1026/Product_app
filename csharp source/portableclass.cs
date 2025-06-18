using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Reflection;
using System.IO;
using System.IO.Compression;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Net;
using System.Web;
using System.Drawing;
using Microsoft.VisualBasic;
using Ionic.Zip;
using Ionic.Zlib;
using MySql.Data.MySqlClient;
using System.Data;
using System.Data.SqlClient;
using System.Threading;
using System.Net.Http;
using System.Collections;
using System.Diagnostics;

namespace csharp
{
    class portableclass
    {
        public static void ecdatabase_set_generater()
        {
            //データベースへ接続
            MySqlConnection conn = new MySqlConnection();
            MySqlCommand command = new MySqlCommand();
            conn.ConnectionString = "server=localhost;user id=mareki;Password=MUqvHc59UKDrEdz8;persist security info=True;database=marekidatabase";
            string select_table = "ecitem";
            conn.Open();

            //テキストリーダー
            Console.WriteLine("FileName(csv) : ");
            string txt_path = @"F:\Users\kiyoka\Desktop\" + Console.ReadLine() + ".csv";
            System.IO.StreamReader cReader1 = (new System.IO.StreamReader(txt_path, System.Text.Encoding.Default));
            string[] csvArray = new string[100000];
            string set_name = "";
            string solo_name = "";
            string[] solo_name_array = { };
            int count = 0;
            int 通常価格_int = 0;
            while (cReader1.Peek() > 0)
            {
                csvArray[count] = cReader1.ReadLine();
                if (csvArray[count].Length > 0)
                {
                    set_name = csvArray[count];
                    solo_name = "";
                    //SELECT文を設定
                    command.CommandText = "SELECT * FROM " + select_table + " WHERE mycode = '" + set_name + "' ORDER BY RAND() LIMIT " + 1 + ";";
                    command.Connection = conn;
                    //SQLを実行し、データを格納し、接続の解除
                    MySqlDataReader readerx = command.ExecuteReader();
                    for (; readerx.Read(); )
                    {
                        solo_name = readerx.GetString("tmp_単品売り");
                    }
                    Console.WriteLine(set_name);
                    Console.WriteLine(solo_name);
                    readerx.Close();
                    solo_name_array = solo_name.Split(',');

                    string[] alp = new string[] {
                        "商品名_正式表記",
                        "p_卸価格_税抜き",
                        "メーカー売価_税抜",
                        "メーカー売価_税込",
                        "通常価格",
                        //"セール価格",
                        "tbl_サイズ_cm",
                        "tbl_主材_表記",
                        "tbl_カラー",
                        "tbl_塗装",
                        "tbl_生産国",
                        "tbl_重量",
                        "tbl_table_opt1",
                        "tbl_table_opt2",
                        "tbl_table_opt3",
                        "tbl_table_opt4",
                        "tbl_table_opt5",
                        "tbl_table_opt6"
                    };

                    string[,] alp_value = new string[,] {
                        {"","","","","","","","","","","","","","","","","","","","","",""}, //[0,*]
                        {"","","","","","","","","","","","","","","","","","","","","",""}, //[1,*]
                        {"","","","","","","","","","","","","","","","","","","","","",""},
                        {"","","","","","","","","","","","","","","","","","","","","",""},
                        {"","","","","","","","","","","","","","","","","","","","","",""},
                        {"","","","","","","","","","","","","","","","","","","","","",""},
                        {"","","","","","","","","","","","","","","","","","","","","",""},
                        {"","","","","","","","","","","","","","","","","","","","","",""} //[8,*]
                    };

                    for (int i = 0; i < solo_name_array.Length; i++)
                    {
                        //SELECT文を設定
                        command.CommandText = "SELECT * FROM " + select_table + " WHERE mycode = '" + solo_name_array[i] + "' ORDER BY RAND() LIMIT " + 1 + ";";
                        command.Connection = conn;

                        //SQLを実行し、データを格納し、接続の解除
                        MySqlDataReader reader = command.ExecuteReader();
                        for (; reader.Read(); )
                        {
                            Console.WriteLine(solo_name_array[i]);
                            for (int c = 0; c < alp.Length; c++)
                            {
                                Console.Write("[" + solo_name_array[i] + "]" + "[" + i + "," + c + "]");
                                Console.Write(alp[c] + " : ");
                                alp_value[i, c] = reader.GetString(alp[c]);
                                Console.WriteLine(alp_value[i, c]);
                                //Console.WriteLine(alp_value[i]);
                            }
                        }
                        reader.Close();
                    }
                    //マンションを合体する。
                    string[] alp_gattai = new string[] {
                        "","","","","","","","","","","","","","","","","","","","","",""
                    };
                    string[] flag = new string[] {
                        "","","","","","","","","","","","","","","","","","","","","",""
                    };
                    for (int c = 0; c < alp.Length; c++)
                    {
                        for (int i = 0; i < solo_name_array.Length; i++)
                        {
                            //MessageBox.Show(alp_value[0, c] + " = " + alp_value[i, c]);
                            if (alp_value[0, c] == alp_value[i, c])
                            {
                                flag[c] = "true";
                            }
                            else
                            {
                                flag[c] = "false";
                            }
                        }
                        for (int i = 0; i < solo_name_array.Length; i++)
                        {
                            if (alp[c] == "通常価格")
                            {
                                通常価格_int = 通常価格_int + int.Parse(alp_value[i, c]);
                                alp_gattai[c] = 通常価格_int.ToString();
                            }
                            else
                            {
                                if (flag[c] == "true")
                                {
                                    alp_gattai[c] = alp_value[i, c];
                                }
                                if (flag[c] == "false")
                                {
                                    alp_gattai[c] += alp_value[i, 0] + " : " + alp_value[i, c] + "\r\n";
                                }
                            }
                        }
                    }
                    for (int c = 0; c < alp.Length; c++)
                    {
                        Console.WriteLine("alp_gattai += [" + "\r\n" + alp_gattai[c] + "\r\n" + "]");
                        string sql = "UPDATE " + select_table + " SET " + alp[c] + "='" + alp_gattai[c] + "' WHERE mycode='" + set_name + "';";
                        functionclass.sqlpost(sql);
                        Console.WriteLine("");
                    }
                }
                System.Threading.Thread.Sleep(1000);
                count += 1;
                通常価格_int = 0;
            }
            conn.Close();
            cReader1.Close();
            MessageBox.Show("完了");
        }
        public static void ecitem_algorithm_generater()
        {
            //テキストリーダー
            string table_choise = "ecitem";
            string txt_path = @"F:\Users\kiyoka\Desktop\" + Console.ReadLine() + ".csv";

            System.IO.StreamReader cReader1 = (new System.IO.StreamReader(txt_path, System.Text.Encoding.Default));
            while (cReader1.Peek() > 0)
            {
                string mycode = cReader1.ReadLine();

                //データベースへ接続
                MySqlConnection conn = new MySqlConnection();
                MySqlCommand command = new MySqlCommand();
                conn.ConnectionString = "server=localhost;user id=mareki;Password=MUqvHc59UKDrEdz8;persist security info=True;database=marekidatabase";
                string select_table = table_choise;
                conn.Open();

                //SELECT文を設定
                int ncount = 1;
                command.CommandText = "SELECT * FROM " + select_table + " WHERE mycode = '" + mycode + "' ORDER BY RAND() LIMIT " + ncount + ";";
                command.Connection = conn;

                //SQLデータを格納
                MySqlDataReader reader = command.ExecuteReader();


                //商品テーブル & 関連商品テーブル作り
                string 商品名_正式表記 = "";
                string tbl_送料表記 = "";
                string tbl_テーブルヘッド = "";
                string tbl_テーブルフット = "";
                string tbl_サイズ_cm = "";
                string tbl_主材_表記 = "";
                string tbl_カラー = "";
                string tbl_塗装 = "";
                string tbl_生産国 = "";
                string tbl_重量 = "";
                string[] tbl_table_opt1 = { };
                string[] tbl_table_opt2 = { };
                string[] tbl_table_opt3 = { };
                string[] tbl_table_opt4 = { };
                string[] tbl_table_opt5 = { };
                string[] tbl_table_opt6 = { };
                string tbl_お届け状態 = "";
                string tbl_注意 = "";
                string tbl_備考1 = "";
                string tbl_備考2 = "";
                string tbl_関連ワード = "";
                string[] tmp_クロスセルリンク = { };
                string[] tmp_別シリーズ紹介 = { };
                string[] tmp_アップセルリンク = { };
                string[] tmp_セット売り = { };
                string[] tmp_単品売り = { };


                //関連カテゴリテーブル作り
                string cat_商品ジャンル1 = "";
                string cat_商品ジャンル2 = "";
                string cat_商品ジャンル3 = "";
                string cat_同ジャンル_セット = "";
                string cat_全ジャンル_材料別 = "";
                string cat_同ジャンル_サイズ別 = "";
                string cat_同ジャンル_価格別 = "";
                string cat_同シリーズ = "";
                string cat_同テイスト = "";
                string cat_全ジャンル_贈答向け = "";
                string cat_おすすめカテゴリ1 = "";
                string cat_おすすめカテゴリ2 = "";
                string cat_おすすめカテゴリ3 = "";

                //SEO
                string seo_B = "";
                string seo_E = "";
                string seo_A = "";
                string seo_自由文字 = "";
                string seo_ペルソナ = "";
                string seo_使用場所目的 = "";
                string tmp_商品名_可変 = "";
                string tmp_キャッチコピー_可変 = "";
                string tbl_メーカー名 = "";
                string sys_送料形態 = "";
                string tbl_主材_SEO = "";

                //サムネイル生成
                string sys_サムネイル_規則 = "";

                Boolean tbl_table_opt1_flag = false;
                Boolean tbl_table_opt2_flag = false;
                Boolean tbl_table_opt3_flag = false;
                Boolean tbl_table_opt4_flag = false;
                Boolean tbl_table_opt5_flag = false;
                Boolean tbl_table_opt6_flag = false;
                Boolean tmp_クロスセルリンク_flag = false;
                Boolean tmp_別シリーズ紹介_flag = false;
                Boolean tmp_アップセルリンク_flag = false;
                Boolean tmp_セット売り_flag = false;
                Boolean tmp_単品売り_flag = false;
                for (int i = 0; reader.Read(); i++)
                {
                    tbl_送料表記 = reader.GetString("tbl_送料表記");
                    商品名_正式表記 = reader.GetString("商品名_正式表記");
                    tbl_テーブルヘッド = reader.GetString("tbl_テーブルヘッド");
                    tbl_テーブルフット = reader.GetString("tbl_テーブルフット");
                    tbl_サイズ_cm = reader.GetString("tbl_サイズ_cm");
                    tbl_主材_表記 = reader.GetString("tbl_主材_表記");
                    tbl_カラー = reader.GetString("tbl_カラー");
                    tbl_塗装 = reader.GetString("tbl_塗装");
                    tbl_生産国 = reader.GetString("tbl_生産国");
                    tbl_重量 = reader.GetString("tbl_重量");

                    //string[] split_string_arg = { "tbl_table_opt1", "tbl_table_opt2", "tbl_table_opt3", "tbl_table_opt4", "tbl_table_opt5", "tbl_table_opt6", "tmp_クロスセルリンク", "tmp_別シリーズ紹介", "tmp_アップセルリンク", "tmp_セット売り", "tmp_単品売り" };
                    //string[] split_string_arg_flag = {"tbl_table_opt1_flag","tbl_table_opt2_flag","tbl_table_opt3_flag","tbl_table_opt4_flag","tbl_table_opt5_flag","tbl_table_opt6_flag","tmp_クロスセルリンク_flag","tmp_別シリーズ紹介_flag","tmp_アップセルリンク_flag","tmp_セット売り_flag","tmp_単品売り_flag"};
                    if (0 <= reader.GetString("tbl_table_opt1").IndexOf("|"))
                    {
                        tbl_table_opt1 = reader.GetString("tbl_table_opt1").Split('|');
                        tbl_table_opt1_flag = true;
                    }
                    if (0 <= reader.GetString("tbl_table_opt2").IndexOf("|"))
                    {
                        tbl_table_opt2 = reader.GetString("tbl_table_opt2").Split('|');
                        tbl_table_opt2_flag = true;
                    }
                    if (0 <= reader.GetString("tbl_table_opt3").IndexOf("|"))
                    {
                        tbl_table_opt3 = reader.GetString("tbl_table_opt3").Split('|');
                        tbl_table_opt3_flag = true;
                    }
                    if (0 <= reader.GetString("tbl_table_opt4").IndexOf("|"))
                    {
                        tbl_table_opt4 = reader.GetString("tbl_table_opt4").Split('|');
                        tbl_table_opt4_flag = true;
                    }
                    if (0 <= reader.GetString("tbl_table_opt5").IndexOf("|"))
                    {
                        tbl_table_opt5 = reader.GetString("tbl_table_opt5").Split('|');
                        tbl_table_opt5_flag = true;
                    }
                    if (0 <= reader.GetString("tbl_table_opt6").IndexOf("|"))
                    {
                        tbl_table_opt6 = reader.GetString("tbl_table_opt6").Split('|');
                        tbl_table_opt6_flag = true;
                    }
                    if (0 <= reader.GetString("tmp_クロスセルリンク").IndexOf(","))
                    {
                        tmp_クロスセルリンク = reader.GetString("tmp_クロスセルリンク").Split(',');
                        tmp_クロスセルリンク_flag = true;
                    }
                    if (0 <= reader.GetString("tmp_別シリーズ紹介").IndexOf(","))
                    {
                        tmp_別シリーズ紹介 = reader.GetString("tmp_別シリーズ紹介").Split(',');
                        tmp_別シリーズ紹介_flag = true;
                    }
                    if (0 <= reader.GetString("tmp_アップセルリンク").IndexOf(","))
                    {
                        tmp_アップセルリンク = reader.GetString("tmp_アップセルリンク").Split(',');
                        tmp_アップセルリンク_flag = true;
                    }
                    if (0 <= reader.GetString("tmp_セット売り").IndexOf(","))
                    {
                        tmp_セット売り = reader.GetString("tmp_セット売り").Split(',');
                        tmp_セット売り_flag = true;
                    }
                    if (0 <= reader.GetString("tmp_単品売り").IndexOf(","))
                    {
                        tmp_単品売り = reader.GetString("tmp_単品売り").Split(',');
                        tmp_単品売り_flag = true;
                    }

                    tbl_お届け状態 = reader.GetString("tbl_お届け状態");
                    tbl_注意 = reader.GetString("tbl_注意");
                    tbl_備考1 = reader.GetString("tbl_備考1");
                    tbl_備考2 = reader.GetString("tbl_備考2");
                    tbl_関連ワード = reader.GetString("tbl_関連ワード");

                    cat_商品ジャンル1 = reader.GetString("cat_商品ジャンル1");
                    cat_商品ジャンル2 = reader.GetString("cat_商品ジャンル2");
                    cat_商品ジャンル3 = reader.GetString("cat_商品ジャンル3");
                    cat_同ジャンル_セット = reader.GetString("cat_同ジャンル_セット");
                    cat_全ジャンル_材料別 = reader.GetString("cat_全ジャンル_材料別");
                    cat_同ジャンル_サイズ別 = reader.GetString("cat_同ジャンル_サイズ別");
                    cat_同ジャンル_価格別 = reader.GetString("cat_同ジャンル_価格別");
                    cat_同シリーズ = reader.GetString("cat_同シリーズ");
                    cat_同テイスト = reader.GetString("cat_同テイスト");
                    cat_全ジャンル_贈答向け = reader.GetString("cat_全ジャンル_贈答向け");
                    cat_おすすめカテゴリ1 = reader.GetString("cat_おすすめカテゴリ1");
                    cat_おすすめカテゴリ2 = reader.GetString("cat_おすすめカテゴリ2");
                    cat_おすすめカテゴリ3 = reader.GetString("cat_おすすめカテゴリ3");

                    //SEO
                    seo_B = reader.GetString("seo_B");
                    seo_E = reader.GetString("seo_E");
                    seo_A = reader.GetString("seo_A");
                    seo_自由文字 = reader.GetString("seo_自由文字");
                    seo_ペルソナ = reader.GetString("seo_ペルソナ");
                    seo_使用場所目的 = reader.GetString("seo_使用場所目的");

                    //TMP
                    tmp_商品名_可変 = reader.GetString("tmp_商品名_可変");
                    tmp_キャッチコピー_可変 = reader.GetString("tmp_キャッチコピー_可変");

                    //SYS
                    sys_送料形態 = reader.GetString("sys_送料形態");
                    sys_サムネイル_規則 = reader.GetString("sys_サムネイル_規則");

                    tbl_メーカー名 = reader.GetString("tbl_メーカー名");
                    tbl_主材_SEO = reader.GetString("tbl_主材_SEO");


                }
                Console.WriteLine(mycode);
                reader.Close();
                conn.Close();
                string kagoyoko_generate = "";
                string table_generate = "";
                string table_generate_kanren = "";
                string table_generate_category = "";
                string rakuten_itemname_generate = "";
                string rakuten_pccatchcopy_generate = "";
                string rakuten_mbcatchcopy_generate = "";
                string yahoo_itemname_generate = "";
                string yahoo_metadesc_generate = "";
                string yahoo_metakey_generate = "";
                string yahoo_catchcopy_generate = "";
                string thumbnail_generate = "";

                //gen_追加説明 @ 商品テーブル&カゴ横
                table_generate = "<table border=\"1\" width=\"100%\" style=\"border:1px solid #595959;border-collapse:collapse; font-size:13px;margin:0px 0px 10px 0px;\"><tbody>";
                //table_generate = "<table border=\"1\" width=\"750px\" style=\"border:1px solid #595959;border-collapse:collapse; font-size:13px;margin:0px 0px 10px 0px;\"><tbody>";
                table_generate = table_generate + "<tr><td colspan=\"2\" style=\"padding:10px; text-align:center; background-color:#000000; color:#ffffff; font-weight:bold;\">" + "商品詳細" + "</td></tr>";

                if (tbl_送料表記.Length > 1)
                {
                    kagoyoko_generate = kagoyoko_generate + tbl_送料表記 + "\r\n" + "\r\n";
                }
                if (tbl_テーブルヘッド.Length > 1)
                {
                    kagoyoko_generate = kagoyoko_generate + "\r\n" + tbl_テーブルヘッド + "\r\n";
                    table_generate = table_generate + "<tr><td colspan=\"2\" style=\"padding:10px;\">" + tbl_テーブルヘッド + "</td></tr>";
                }
                if (商品名_正式表記.Length > 1)
                {
                    kagoyoko_generate = kagoyoko_generate + "商品名" + " : " + 商品名_正式表記 + "\r\n";
                    table_generate = table_generate + "<tr><td width=\"100px\">" + "商品名" + "</td><td>" + 商品名_正式表記 + "</td></tr>";
                }
                if (tbl_サイズ_cm.Length > 1)
                {
                    kagoyoko_generate = kagoyoko_generate + "サイズ" + " : " + tbl_サイズ_cm + "\r\n";
                    table_generate = table_generate + "<tr><td>" + "サイズ" + "</td><td>" + tbl_サイズ_cm + "</td></tr>";
                }
                if (tbl_主材_表記.Length > 1)
                {
                    kagoyoko_generate = kagoyoko_generate + "主材" + " : " + tbl_主材_表記 + "\r\n";
                    table_generate = table_generate + "<tr><td>" + "主材" + "</td><td>" + tbl_主材_表記 + "</td></tr>";
                }
                if (tbl_カラー.Length > 1)
                {
                    kagoyoko_generate = kagoyoko_generate + "カラー" + " : " + tbl_カラー + "\r\n";
                    table_generate = table_generate + "<tr><td>" + "カラー" + "</td><td>" + tbl_カラー + "</td></tr>";
                }
                if (tbl_塗装.Length > 1)
                {
                    kagoyoko_generate = kagoyoko_generate + "塗装" + " : " + tbl_塗装 + "\r\n";
                    table_generate = table_generate + "<tr><td>" + "塗装" + "</td><td>" + tbl_塗装 + "</td></tr>";
                }
                if (tbl_生産国.Length > 1)
                {
                    kagoyoko_generate = kagoyoko_generate + "生産国" + " : " + tbl_生産国 + "\r\n";
                    table_generate = table_generate + "<tr><td>" + "生産国" + "</td><td>" + tbl_生産国 + "</td></tr>";
                }
                if (tbl_重量.Length > 1)
                {
                    kagoyoko_generate = kagoyoko_generate + "重量" + " : " + tbl_重量 + "\r\n";
                    table_generate = table_generate + "<tr><td>" + "重量" + "</td><td>" + tbl_重量 + "</td></tr>";
                }

                if (tbl_table_opt1_flag && tbl_table_opt1[0].Length > 1)
                {
                    kagoyoko_generate = kagoyoko_generate + tbl_table_opt1[0] + " : " + tbl_table_opt1[1] + "\r\n";
                    table_generate = table_generate + "<tr><td>" + tbl_table_opt1[0] + "</td><td>" + tbl_table_opt1[1] + "</td></tr>";
                }
                if (tbl_table_opt2_flag && tbl_table_opt2[0].Length > 1)
                {
                    kagoyoko_generate = kagoyoko_generate + tbl_table_opt2[0] + " : " + tbl_table_opt2[1] + "\r\n";
                    table_generate = table_generate + "<tr><td>" + tbl_table_opt2[0] + "</td><td>" + tbl_table_opt2[1] + "</td></tr>";
                }
                if (tbl_table_opt3_flag && tbl_table_opt3[0].Length > 1)
                {
                    kagoyoko_generate = kagoyoko_generate + tbl_table_opt3[0] + " : " + tbl_table_opt3[1] + "\r\n";
                    table_generate = table_generate + "<tr><td>" + tbl_table_opt3[0] + "</td><td>" + tbl_table_opt3[1] + "</td></tr>";
                }
                if (tbl_table_opt4_flag && tbl_table_opt4[0].Length > 1)
                {
                    kagoyoko_generate = kagoyoko_generate + tbl_table_opt4[0] + " : " + tbl_table_opt4[1] + "\r\n";
                    table_generate = table_generate + "<tr><td>" + tbl_table_opt4[0] + "</td><td>" + tbl_table_opt4[1] + "</td></tr>";
                }
                if (tbl_table_opt5_flag && tbl_table_opt5[0].Length > 1)
                {
                    kagoyoko_generate = kagoyoko_generate + tbl_table_opt5[0] + " : " + tbl_table_opt5[1] + "\r\n";
                    table_generate = table_generate + "<tr><td>" + tbl_table_opt5[0] + "</td><td>" + tbl_table_opt5[1] + "</td></tr>";
                }
                if (tbl_table_opt6_flag && tbl_table_opt6[0].Length > 1)
                {
                    kagoyoko_generate = kagoyoko_generate + tbl_table_opt6[0] + " : " + tbl_table_opt6[1] + "\r\n";
                    table_generate = table_generate + "<tr><td>" + tbl_table_opt6[0] + "</td><td>" + tbl_table_opt6[1] + "</td></tr>";
                }
                if (tbl_送料表記.Length > 1)
                {
                    table_generate = table_generate + "<tr><td>" + "送料" + "</td><td>" + tbl_送料表記 + "</td></tr>";
                }
                if (tbl_お届け状態.Length > 1)
                {
                    kagoyoko_generate = kagoyoko_generate + "\r\n" + "お届け状態 : " + tbl_お届け状態 + "\r\n" + "\r\n";
                    table_generate = table_generate + "<tr><td>" + "お届け状態" + "</td><td>" + tbl_お届け状態 + "</td></tr>";
                }
                if (tbl_注意.Length > 1)
                {
                    kagoyoko_generate = kagoyoko_generate + "\r\n" + "ご注意 : " + tbl_注意 + "\r\n" + "\r\n";
                    table_generate = table_generate + "<tr><td>" + "ご注意" + "</td><td>" + tbl_注意 + "</td></tr>";
                }
                if (tbl_備考1.Length > 1)
                {
                    if (tbl_備考2.Length > 1)
                    {
                        kagoyoko_generate = kagoyoko_generate + "\r\n" + "備考1 : " + tbl_備考1 + "\r\n" + "\r\n";
                        table_generate = table_generate + "<tr><td>" + "備考1" + "</td><td>" + tbl_備考1 + "</td></tr>";
                    }
                    else
                    {
                        kagoyoko_generate = kagoyoko_generate + "\r\n" + "備考 : " + tbl_備考1 + "\r\n" + "\r\n";
                        table_generate = table_generate + "<tr><td>" + "備考" + "</td><td>" + tbl_備考1 + "</td></tr>";
                    }
                }
                if (tbl_備考2.Length > 1)
                {
                    kagoyoko_generate = kagoyoko_generate + "\r\n" + "備考2 : " + tbl_備考2 + "\r\n" + "\r\n";
                    table_generate = table_generate + "<tr><td>" + "備考2" + "</td><td>" + tbl_備考2 + "</td></tr>";
                }
                if (tbl_テーブルフット.Length > 1)
                {
                    kagoyoko_generate = kagoyoko_generate + tbl_テーブルフット + "\r\n";
                    table_generate = table_generate + "<tr><td colspan=\"2\" style=\"padding:10px;\">" + tbl_テーブルフット + "</td></tr>";
                }
                if (tbl_関連ワード.Length > 1)
                {
                    kagoyoko_generate = kagoyoko_generate + "\r\n" + tbl_関連ワード;
                }







                table_generate = table_generate + "</tbody></table>";

                //gen_追加説明 @ 関連商品テーブル
                table_generate_kanren = "<table border=\"1\" width=\"100%\" style=\"border:1px solid #595959;border-collapse:collapse; font-size:13px;margin:0px 0px 10px 0px;\"><tbody>";
                table_generate_kanren = table_generate_kanren + "<tr><td colspan=\"2\" style=\"padding:10px; text-align:center; background-color:#000000; color:#ffffff; font-weight:bold;\">" + "関連商品" + "</td></tr>";
                string flag = "false";// スプリット

                //商品番号･パスの置換
                string 店舗コード = functionclass.正規表現("(?<findcode>.*?)--", mycode);
                string 商品コード = functionclass.正規表現(".*?--(?<findcode>.*?)$", mycode);
                店舗コード = 店舗コード.Replace("roomkoubou", "kagunoroomkoubou");



                if (tmp_クロスセルリンク_flag && tmp_クロスセルリンク[0].Length > 1)
                {
                    table_generate_kanren = table_generate_kanren + "<tr><td width=\"100px\">" + "関連商品" + "</td><td>";
                    for (int i = 0; i < tmp_クロスセルリンク.Length; i++)
                    {
                        table_generate_kanren = table_generate_kanren + "<a href=\"./商品URL/" + tmp_クロスセルリンク[i] + ".html\"><img src=\"./サムネイル/" + tmp_クロスセルリンク[i] + ".jpg\" width=\"146px\" height=\"146px\"></a> ";
                        flag = "true";
                    }
                    table_generate_kanren = table_generate_kanren + "</td></tr>";
                }
                if (tmp_別シリーズ紹介_flag && tmp_別シリーズ紹介[0].Length > 1)
                {
                    table_generate_kanren = table_generate_kanren + "<tr><td>" + "関連商品" + "</td><td>";
                    for (int i = 0; i < tmp_別シリーズ紹介.Length; i++)
                    {
                        table_generate_kanren = table_generate_kanren + "<a href=\"./商品URL/" + tmp_別シリーズ紹介[i] + ".html\"><img src=\"./サムネイル/" + tmp_別シリーズ紹介[i] + ".jpg\" width=\"146px\" height=\"146px\"></a> ";
                        flag = "true";
                    }
                    table_generate_kanren = table_generate_kanren + "</td></tr>";
                }
                if (tmp_アップセルリンク_flag && tmp_アップセルリンク[0].Length > 1)
                {
                    table_generate_kanren = table_generate_kanren + "<tr><td>" + "関連商品" + "</td><td>";
                    for (int i = 0; i < tmp_アップセルリンク.Length; i++)
                    {
                        table_generate_kanren = table_generate_kanren + "<a href=\"./商品URL/" + tmp_アップセルリンク[i] + ".html\"><img src=\"./サムネイル/" + tmp_アップセルリンク[i] + ".jpg\" width=\"146px\" height=\"146px\"></a> ";
                        flag = "true";
                    }
                    table_generate_kanren = table_generate_kanren + "</td></tr>";
                }
                if (tmp_セット売り_flag && tmp_セット売り[0].Length > 1)
                {
                    table_generate_kanren = table_generate_kanren + "<tr><td>" + "セットでの販売" + "</td><td>";
                    for (int i = 0; i < tmp_セット売り.Length; i++)
                    {
                        table_generate_kanren = table_generate_kanren + "<a href=\"./商品URL/" + tmp_セット売り[i] + ".html\"><img src=\"./サムネイル/" + tmp_セット売り[i] + ".jpg\" width=\"146px\" height=\"146px\"></a>";
                        flag = "true";
                    }
                    table_generate_kanren = table_generate_kanren + "</td></tr>";
                }
                if (tmp_単品売り_flag && tmp_単品売り[0].Length > 1)
                {
                    table_generate_kanren = table_generate_kanren + "<tr><td>" + "単品での販売" + "</td><td>";
                    for (int i = 0; i < tmp_単品売り.Length; i++)
                    {
                        table_generate_kanren = table_generate_kanren + "<a href=\"./商品URL/" + tmp_単品売り[i] + ".html\"><img src=\"./サムネイル/" + tmp_単品売り[i] + ".jpg\" width=\"146px\" height=\"146px\"></a>";
                        flag = "true";
                    }
                    table_generate_kanren = table_generate_kanren + "</td></tr>";
                }
                table_generate_kanren = table_generate_kanren + "</tbody></table>";

                //gen_追加説明 @ カテゴリテーブル
                //maruko--等はカテゴリを作らない
                if (0 <= mycode.IndexOf("maruko--") || 0 <= mycode.IndexOf("roomkoubou--") || 0 <= mycode.IndexOf("zen--"))
                {
                    Console.WriteLine("カテゴリテーブルは作りません");
                }
                else
                {
                    table_generate_category = "<table border=\"1\" width=\"100%\" style=\"border:1px solid #595959;border-collapse:collapse; font-size:13px;margin:0px 0px 10px 0px;\"><tbody>";
                    //table_generate_category = "<table border=\"1\" width=\"750px\" style=\"border:1px solid #595959;border-collapse:collapse; font-size:13px;margin:0px 0px 10px 0px;\"><tbody>";
                    table_generate_category = table_generate_category + "<tr><td colspan=\"2\" style=\"padding:10px; text-align:center; background-color:#000000; color:#ffffff; font-weight:bold;\">" + "関連カテゴリ" + "</td></tr>";
                    string[] カテゴリ列挙 = { cat_商品ジャンル1, cat_商品ジャンル2, cat_商品ジャンル3, cat_同ジャンル_セット, cat_全ジャンル_材料別, cat_同ジャンル_サイズ別, cat_同ジャンル_価格別, cat_同シリーズ, cat_同テイスト, cat_全ジャンル_贈答向け, cat_おすすめカテゴリ1, cat_おすすめカテゴリ2, cat_おすすめカテゴリ3 };
                    int カテゴリ2列 = 1;
                    for (int i = 0; i < カテゴリ列挙.Length; i++)
                    {
                        if (カテゴリ列挙[i].Length > 1)
                        {
                            string tmp_tikan = "";
                            string[] rep_yahoo_roomkoubou_cat_path = { "ＤＩＹ工房:テーブル脚", "ソファ:1Pソファ", "ベッド:安心の国産", "ダイニング:ダイニングセット・2人用", "TVボード:ロータイプ", "ソファ:2Pソファ", "ソファ:3Pソファ", "ダイニング:ダイニングセット・4人用", "ダイニング:ダイニング4人用", "TVボード:コーナータイプ", "ダイニング:ダイニングセット・6人用", "ＤＩＹ工房", "テーブル:サイドテーブル", "テーブル:こたつ", "その他", "テーブル:ちゃぶ台", "収納:ふとん収納", "デザイナーズ家具:イームズ", "その他:インテリア雑貨", "チェア:オフィスチェア", "その他:ラグ", "ソファ:カウチソファ", "デザイナーズ家具:カスティリオーニ", "その他:ガーデンファニチャー", "収納:キッチン収納", "ソファ:コーナーソファ", "ダイニング:コーナーダイニング", "サイズオーダー特集:北欧パイン無垢材ラック", "収納:リビングボード", "サイズオーダー特集:無垢材デスク　セシル", "ベッド:シングル", "デザイナーズ家具:ジョージ・ネルソン", "ソファ:スツール・オットマン", "チェア:ベンチ", "ベッド:セミダブル", "テーブル:センターテーブル", "ソファ", "ソファ:ソファーベッド", "ベッド:ソファーベッド", "ダイニング", "チェア:ダイニングチェア", "ベッド:ダブル", "ダイニング:ダイニングチェア", "収納:チェスト", "チェア", "ダイニング:ダイニングテーブル", "テーブル", "TVボード", "その他:ドレッサー、ミラー", "ベッド:ナイトテーブル", "TVボード:ハイタイプ", "収納:ハンガーラック", "ダイニング:バーカウンター、バーチェア", "ＤＩＹ工房:バターミルクペイント", "チェア:パーソナルチェア", "デスク", "デスク:パソコンデスク", "収納:すき間収納", "ベッド", "その他:ペット関連", "ベッド:マットレス", "デザイナーズ家具:ル・コルビジェ", "ソファ:フロアソファ", "ベッド:ロフト", "ベッド:ワイドダブル、クイーン", "材料から探す:アルダー材の商品", "シリーズで選ぶ:anthemシリーズ", "デスク:学習机", "収納:桐収納", "チェア:座椅子・フロアチェア", "テーブル:リビングテーブル・座卓", "キッズ家具", "ベッド:2段、ベビー用", "ＤＩＹ工房:プラネットカラー", "収納", "ＤＩＹ工房:各種集成材", "デスク:書斎机", "収納:本棚", "材料から探す:ブラックチェリー材の商品", "シリーズで選ぶ:BOSCOシリーズ", "カントリー家具:塗装済み", "ベッド:寝具、布団", "シリーズで選ぶ:CABANAシリーズ", "収納:壁面収納", "デスク:堀田木工", "収納:民芸調家具", "カントリー家具:無塗装", "シリーズで選ぶ:CHESTERシリーズ", "シリーズで選ぶ:DUNKシリーズ", "シリーズで選ぶ:emo.シリーズ", "ベッド:セミシングル", "収納:雛人形収納", "材料から探す:ひのき材の商品", "シリーズで選ぶ:homaシリーズ", "シリーズで選ぶ:木と風シリーズ", "ＤＩＹ工房:こたつ天板", "サイズオーダー特集:北欧パイン無垢材テーブル レオン", "ＤＩＹ工房:脚立", "シリーズで選ぶ:noraシリーズ", "材料から探す:ホワイトオーク材の商品", "シリーズで選ぶ:OCTAシリーズ", "ＤＩＹ工房:大阪ガスケミカル", "ＤＩＹ工房:大阪塗料", "シリーズで選ぶ:Organic Modern Life", "ＤＩＹ工房:オスモ/ウッドワックス", "材料から探す:パイン材の商品", "材料から探す:レッドオーク材の商品", "サイズオーダー特集", "サイズオーダー特集:ウォールナットダイニングテーブル マーク", "チェア:スツール", "材料から探す:杉材の商品", "その他:テーブルマット", "材料から探す:タモ材の商品", "収納:TEL台　FAX台", "ＤＩＹ工房:装飾タイル", "ＤＩＹ工房:ターナー色彩", "材料から探す:ウォールナット材の商品", "ＤＩＹ工房:和信化学工業", "材料から探す", "アジアンテイスト", "アンティーク調家具", "カントリー家具", "シャビー、フレンチカントリー", "シリーズで選ぶ", "デザイナーズ家具", "北欧スタイル家具", "和風家具", "ichiba" };
                            string[] rep_yahoo_roomkoubou_cat_url = { "1a2cc56a5ab.html", "1pa5bda5d5.html", "2ba0f44cb2b.html", "2bfcdcdd1a.html", "2d15c5abaae.html", "2pa5bda5d5.html", "3pa5bda5d5.html", "4bfcdcdd1a.html", "4ca5a205a5a.html", "55cba21aba3.html", "6bfcdcdd1a.html", "a3c4a3c9a3.html", "a45aaa2555b.html", "a4b3a4bfa4.html", "a4bda4cec2.html", "a4c1a4e3a4.html", "a4d5a4c8a4.html", "a5a4a1bca5.html", "a5a4a5f3a5.html", "a5aaa5d5a5.html", "a5aba1bca5.html", "a5aba5a6a5.html", "a5aba5b9a5.html", "a5aca1bca5.html", "a5ada5c3a5.html", "a5b3a1bca5.html", "a5b3a1bca52.html", "a5b5a5a4a5.html", "a5b5a5a4a52.html", "a5b5a5a4a54.html", "a5b7a5f3a5.html", "a5b8a5e7a1.html", "a5b9a5c4a1.html", "a5b9a5c6a5.html", "a5bba5dfa5.html", "a5bba5f3a5.html", "a5bda5d5a5.html", "a5bda5d5a52.html", "a5bda5d5a53.html", "a5c0a5a4a5.html", "a5c0a5a4a52.html", "a5c0a5d6a5.html", "a5c1a5a7a5.html", "a5c1a5a7a52.html", "a5c1a5a7a53.html", "a5c6a1bca5.html", "a5c6a1bca52.html", "a5c6a5eca5.html", "a5c9a5eca5.html", "a5caa5a4a52.html", "a5cfa5a4a5.html", "a5cfa5f3a5.html", "a5d0a1bca5.html", "a5d0a5bfa1.html", "a5d1a1bca5.html", "a5d1a5bda5.html", "a5d1a5bda52.html", "a5d5a5eaa1.html", "a5d9a5c3a5.html", "a5daa5c3a5.html", "a5dea5c3a5.html", "a5eba1a6a5.html", "a5eda1bca5.html", "a5eda5d5a5.html", "a5efa5a4a5.html", "alder.html", "anthem.html", "b3d8bdacb43.html", "b6cdbcfdc7.html", "bac2b0d8bb.html", "bac2c2ee.html", "bbd2b6a1.html", "bbd2b6a1cd.html", "bcabc1b3c5.html", "bcfdc7bc.html", "bdb8c0aeba.html", "bdf1bad8b4.html", "bdf1c3aa.html", "blackcherr.html", "bosco.html", "c5c9c1f5ba.html", "c9dbc3c4.html", "cabana28a5.html", "cac9ccccbc.html", "cbd9c5c4cc.html", "ccb1b7ddb2.html", "ccb5c5c9c1.html", "chestera5b.html", "dunk.html", "emo.html", "fbadb5a25a5.html", "hina.html", "hinoki.html", "homaa1caa5.html", "kitokaze.html", "kotatutt.html", "leon.html", "lucano.html", "mam.html", "oak.html", "octaa5b7a5.html", "oosakagus.html", "oosakatory.html", "organicmod.html", "osmo2fa5a6.html", "pine.html", "redoak.html", "sizeorder2.html", "steelleg.html", "stool.html", "sugi.html", "tablemat.html", "tamo.html", "telc2e6a1a.html", "tile.html", "turner.html", "walnut.html", "washin.html", "zairyoukar.html", "http:/kagunoroomkoubou/a5a2a5b8a5.html", "http:/kagunoroomkoubou/a5a2a5f3a5.html", "http:/kagunoroomkoubou/a5aba5f3a5.html", "http:/kagunoroomkoubou/a5b7a5e3a5.html", "http:/kagunoroomkoubou/a5b7a5eaa1.html", "http:/kagunoroomkoubou/a5c7a5b6a5.html", "http:/kagunoroomkoubou/cbccb2a4a5.html", "http:/kagunoroomkoubou/cfc2c9f7b2.html", "http:/kagunoroomkoubou/ichiba.html" };


                            //YahooRoomkoubou
                            for (int catloop = 0; catloop < rep_yahoo_roomkoubou_cat_path.Length; catloop++)
                            {
                                if (カテゴリ列挙[i] == rep_yahoo_roomkoubou_cat_path[catloop])
                                    tmp_tikan = rep_yahoo_roomkoubou_cat_url[catloop];
                            }
                            if (カテゴリ2列 == 1 && tmp_tikan.Length > 1)
                            {
                                table_generate_category = table_generate_category + "<tr><td><a href=\"http://store.shopping.yahoo.co.jp/kagunoroomkoubou/" + tmp_tikan + "\">" + カテゴリ列挙[i] + "</a></td>";
                                カテゴリ2列++;
                            }
                            else if (カテゴリ2列 == 2 && tmp_tikan.Length > 1)
                            {
                                table_generate_category = table_generate_category + "<td><a href=\"http://store.shopping.yahoo.co.jp/kagunoroomkoubou/" + tmp_tikan + "\">" + カテゴリ列挙[i] + "</a></td></tr>";
                                カテゴリ2列--;
                            }
                        }
                    }
                    if (カテゴリ2列 == 2)
                        table_generate_category = table_generate_category + "<td></td></tr>";
                    table_generate_category = table_generate_category + "</tbody></table>";
                }


                //SEOジェネレート
                //楽天@PC用キャッチコピー 174byte
                //楽天@モバイル用キャッチコピー 60byte
                //楽天@商品名 255byte
                string[] 商品名列挙 = { tmp_商品名_可変, 商品名_正式表記, tbl_主材_SEO, sys_送料形態, cat_商品ジャンル1, cat_商品ジャンル2, cat_商品ジャンル3, cat_同テイスト };
                //string[] 商品名列挙 = { tmp_商品名_可変, 商品名_正式表記, tbl_主材_SEO, sys_送料形態, cat_商品ジャンル1, cat_商品ジャンル2, cat_商品ジャンル3, cat_同テイスト, tbl_メーカー名, tbl_塗装 };
                for (int i = 0; i < 商品名列挙.Length; i++)
                {
                    if (商品名列挙[i].Length > 1)
                    {
                        //重複を削除するために : をスペースに置換して、配列格納後重複削除するか？
                        string[] スプリッタ = 商品名列挙[i].Split(':');
                        商品名列挙[i] = スプリッタ[スプリッタ.Length - 1];


                        if (Encoding.GetEncoding("Shift_JIS").GetByteCount(yahoo_itemname_generate) + Encoding.GetEncoding("Shift_JIS").GetByteCount(商品名列挙[i]) <= 60) //理想は60
                        {
                            yahoo_itemname_generate = yahoo_itemname_generate + 商品名列挙[i];
                        }
                        else
                        {
                            if (Encoding.GetEncoding("Shift_JIS").GetByteCount(yahoo_itemname_generate) < 10) //10は小さすぎ
                            {
                                if (Encoding.GetEncoding("Shift_JIS").GetByteCount(yahoo_itemname_generate) + Encoding.GetEncoding("Shift_JIS").GetByteCount(商品名列挙[i]) <= 160) //それなら長く
                                    yahoo_itemname_generate = yahoo_itemname_generate + 商品名列挙[i];
                                else
                                    yahoo_itemname_generate = "文字数制限エラー";
                            }
                            else
                            {
                                break;
                            }
                        }
                        yahoo_itemname_generate = yahoo_itemname_generate + " ";
                    }
                }
                for (int i = 0; i < 商品名列挙.Length; i++)
                {
                    if (商品名列挙[i].Length > 1)
                    {
                        //重複を削除するために : をスペースに置換して、配列格納後重複削除するか？
                        string[] スプリッタ = 商品名列挙[i].Split(':');
                        商品名列挙[i] = スプリッタ[スプリッタ.Length - 1];
                        if (Encoding.GetEncoding("Shift_JIS").GetByteCount(rakuten_itemname_generate) + Encoding.GetEncoding("Shift_JIS").GetByteCount(商品名列挙[i]) <= 255)
                            rakuten_itemname_generate = rakuten_itemname_generate + 商品名列挙[i];
                        else
                            break;
                        rakuten_itemname_generate = rakuten_itemname_generate + " ";
                    }
                }
                string[] キャッチコピー列挙 = { tmp_キャッチコピー_可変, seo_B, seo_E, seo_A, seo_自由文字, seo_ペルソナ, seo_使用場所目的 };
                for (int i = 0; i < キャッチコピー列挙.Length; i++)
                {
                    if (キャッチコピー列挙[i].Length > 1)
                    {
                        if (Encoding.GetEncoding("Shift_JIS").GetByteCount(rakuten_pccatchcopy_generate) + Encoding.GetEncoding("Shift_JIS").GetByteCount(キャッチコピー列挙[i]) <= 174)
                            rakuten_pccatchcopy_generate = rakuten_pccatchcopy_generate + キャッチコピー列挙[i];
                        else
                            break;
                        rakuten_pccatchcopy_generate = rakuten_pccatchcopy_generate + " ";
                    }
                }
                for (int i = 0; i < キャッチコピー列挙.Length; i++)
                {
                    if (キャッチコピー列挙[i].Length > 1)
                    {
                        if (Encoding.GetEncoding("Shift_JIS").GetByteCount(rakuten_mbcatchcopy_generate) + Encoding.GetEncoding("Shift_JIS").GetByteCount(キャッチコピー列挙[i]) <= 60) //60厳守
                            rakuten_mbcatchcopy_generate = rakuten_mbcatchcopy_generate + キャッチコピー列挙[i];
                        else
                            break;
                        rakuten_mbcatchcopy_generate = rakuten_mbcatchcopy_generate + " ";
                    }
                }
                for (int i = 0; i < キャッチコピー列挙.Length; i++)
                {
                    if (キャッチコピー列挙[i].Length > 1)
                    {
                        if (Encoding.GetEncoding("Shift_JIS").GetByteCount(yahoo_catchcopy_generate) + Encoding.GetEncoding("Shift_JIS").GetByteCount(キャッチコピー列挙[i]) <= 60)
                        {
                            yahoo_catchcopy_generate = yahoo_catchcopy_generate + キャッチコピー列挙[i];
                            Console.WriteLine(yahoo_catchcopy_generate.Length + キャッチコピー列挙[i].Length);
                        }
                        else
                        {
                            break;
                        }
                        yahoo_catchcopy_generate = yahoo_catchcopy_generate + " ";
                    }
                }

                string[] metadesc列挙 = { tbl_テーブルヘッド };
                for (int i = 0; i < metadesc列挙.Length; i++)
                {
                    if (metadesc列挙[i].Length > 1)
                    {
                        //重複を削除するために : をスペースに置換して、配列格納後重複削除するか？
                        string[] スプリッタ = metadesc列挙[i].Split(':');
                        metadesc列挙[i] = スプリッタ[スプリッタ.Length - 1];
                        if (Encoding.GetEncoding("Shift_JIS").GetByteCount(yahoo_metadesc_generate) + Encoding.GetEncoding("Shift_JIS").GetByteCount(metadesc列挙[i]) <= 160)
                            yahoo_metadesc_generate = yahoo_metadesc_generate + metadesc列挙[i];
                        else
                            break;
                        yahoo_metadesc_generate = yahoo_metadesc_generate + " ";
                    }
                    if (yahoo_metadesc_generate.Length > 1)
                        yahoo_metadesc_generate = yahoo_metadesc_generate.Remove(yahoo_metadesc_generate.Length - 1);
                }

                string[] metakey列挙 = { 商品名_正式表記, seo_B, seo_E, seo_A, cat_商品ジャンル1, tbl_主材_SEO };
                for (int i = 0; i < metakey列挙.Length; i++)
                {
                    if (metakey列挙[i].Length > 1)
                    {
                        //重複を削除するために : をスペースに置換して、配列格納後重複削除するか？
                        string[] スプリッタ = metakey列挙[i].Split(':');
                        metakey列挙[i] = スプリッタ[スプリッタ.Length - 1];
                        if (Encoding.GetEncoding("Shift_JIS").GetByteCount(yahoo_metakey_generate) + Encoding.GetEncoding("Shift_JIS").GetByteCount(metakey列挙[i]) <= 160)
                            yahoo_metakey_generate = yahoo_metakey_generate + metakey列挙[i];
                        else
                            break;
                        yahoo_metakey_generate = yahoo_metakey_generate + " ";
                    }
                }
                yahoo_metakey_generate = yahoo_metakey_generate.Replace(" ", "|");
                if (yahoo_metakey_generate.Length > 1)
                    yahoo_metakey_generate = yahoo_metakey_generate.Remove(yahoo_metakey_generate.Length - 1);


                //サムネイルジェネレート
                string ec_dir_path = "";
                if (mycode.Contains("roomkoubou--"))
                    ec_dir_path = @"D:\03.Ec_GodBot\00.設計_system\system_ShohinPageGenerator\roomkoubou\";
                if (mycode.Contains("mokuto--"))
                    ec_dir_path = @"D:\03.Ec_GodBot\00.設計_system\system_ShohinPageGenerator\mokuto\";
                if (mycode.Contains("marukomaruko--"))
                    ec_dir_path = @"D:\03.Ec_GodBot\00.設計_system\system_ShohinPageGenerator\marukomaruko\";
                if (mycode.Contains("zen--"))
                    ec_dir_path = @"D:\03.Ec_GodBot\00.設計_system\system_ShohinPageGenerator\zen\";
                string 一時mycode = mycode.Replace("roomkoubou--", "");
                一時mycode = 一時mycode.Replace("mokuto--", "");
                一時mycode = 一時mycode.Replace("marukomaruko--", "");
                一時mycode = 一時mycode.Replace("zen--", "zenshop");
                string[] files = System.IO.Directory.GetFiles(ec_dir_path + "サムネイル", 一時mycode + ".jpg", System.IO.SearchOption.AllDirectories);
                for (int i = 0; i < files.Length; i++)
                {
                    thumbnail_generate = thumbnail_generate + "<div class=\"thumbnail\"><img src=\"./サムネイル/" + Path.GetFileName(files[i]) + "\"></div> ";
                    Console.WriteLine(files[i]);
                }
                string[] files2 = System.IO.Directory.GetFiles(ec_dir_path + "サムネイル", 一時mycode + "_*.jpg", System.IO.SearchOption.AllDirectories);
                for (int i = 0; i < files2.Length; i++)
                {
                    thumbnail_generate = thumbnail_generate + "<div class=\"thumbnail\"><img src=\"./サムネイル/" + Path.GetFileName(files2[i]) + "\"></div> ";
                    Console.WriteLine(files2[i]);
                }

                //最終調整
                if (flag == "true")
                    table_generate = table_generate + table_generate_kanren;
                table_generate = table_generate + table_generate_category;
                kagoyoko_generate = kagoyoko_generate.Replace("\r\n", "temp_kaigyou");
                kagoyoko_generate = kagoyoko_generate.Replace("\n", "\r\n");
                kagoyoko_generate = kagoyoko_generate.Replace("\r", "\r\n");
                kagoyoko_generate = kagoyoko_generate.Replace("temp_kaigyou", "\r\n");
                table_generate = table_generate.Replace("\r\n", "<br>");
                table_generate = table_generate.Replace("\n", "<br>");
                table_generate = table_generate.Replace("\r", "<br>");

                kagoyoko_generate = kagoyoko_generate.Trim(' ', '　', '|', ',');
                table_generate = table_generate.Trim(' ', '　', '|', ',');
                rakuten_itemname_generate = rakuten_itemname_generate.Trim(' ', '　', '|', ',');
                rakuten_pccatchcopy_generate = rakuten_pccatchcopy_generate.Trim(' ', '　', '|', ',');
                rakuten_mbcatchcopy_generate = rakuten_mbcatchcopy_generate.Trim(' ', '　', '|', ',');
                yahoo_itemname_generate = yahoo_itemname_generate.Trim(' ', '　', '|', ',');
                yahoo_metadesc_generate = yahoo_metadesc_generate.Trim(' ', '　', '|', ',');
                yahoo_metakey_generate = yahoo_metakey_generate.Trim(' ', '　', '|', ',');
                yahoo_catchcopy_generate = yahoo_catchcopy_generate.Trim(' ', '　', '|', ',');
                thumbnail_generate = thumbnail_generate.Trim(' ', '　', '|', ',');

                string sql_kagoyoko_generate = "UPDATE " + select_table + " SET " + "gen_カゴ横説明" + "='" + kagoyoko_generate + "' WHERE mycode='" + mycode + "';";
                string sql_table_generate = "UPDATE " + select_table + " SET " + "gen_追加説明" + "='" + table_generate + "' WHERE mycode='" + mycode + "';";
                string sql_rakuten_itemname_generate = "UPDATE " + select_table + " SET " + "gen_楽天_商品名" + "='" + rakuten_itemname_generate + "' WHERE mycode='" + mycode + "';";
                string sql_rakuten_pccatchcopy_generate = "UPDATE " + select_table + " SET " + "gen_楽天_PCキャッチコピー" + "='" + rakuten_pccatchcopy_generate + "' WHERE mycode='" + mycode + "';";
                string sql_rakuten_mbcatchcopy_generate = "UPDATE " + select_table + " SET " + "gen_楽天_MBキャッチコピー" + "='" + rakuten_mbcatchcopy_generate + "' WHERE mycode='" + mycode + "';";
                string sql_yahoo_itemname_generate = "UPDATE " + select_table + " SET " + "gen_Yahoo_商品名" + "='" + yahoo_itemname_generate + "' WHERE mycode='" + mycode + "';";
                string sql_yahoo_metadesc_generate = "UPDATE " + select_table + " SET " + "gen_Yahoo_metadesc" + "='" + yahoo_metadesc_generate + "' WHERE mycode='" + mycode + "';";
                string sql_yahoo_metakey_generate = "UPDATE " + select_table + " SET " + "gen_Yahoo_metakey" + "='" + yahoo_metakey_generate + "' WHERE mycode='" + mycode + "';";
                string sql_yahoo_catchcopy_generate = "UPDATE " + select_table + " SET " + "gen_Yahoo_キャッチコピー" + "='" + yahoo_catchcopy_generate + "' WHERE mycode='" + mycode + "';";
                string sql_thumbnail_generate = "UPDATE " + select_table + " SET " + "gen_サムネイル" + "='" + thumbnail_generate + "' WHERE mycode='" + mycode + "';";
                functionclass.sqlpost(sql_kagoyoko_generate);
                functionclass.sqlpost(sql_table_generate);
                functionclass.sqlpost(sql_rakuten_itemname_generate);
                functionclass.sqlpost(sql_rakuten_pccatchcopy_generate);
                functionclass.sqlpost(sql_rakuten_mbcatchcopy_generate);
                functionclass.sqlpost(sql_yahoo_itemname_generate);
                functionclass.sqlpost(sql_yahoo_metadesc_generate);
                functionclass.sqlpost(sql_yahoo_metakey_generate);
                functionclass.sqlpost(sql_yahoo_catchcopy_generate);
                functionclass.sqlpost(sql_thumbnail_generate);
                Console.WriteLine("Database Uped");
                //catch (System.Exception)
            }
            cReader1.Close();
            Console.WriteLine("完了");
        }
        public static void ecitem_localbatch_generater()
        {
            //宣言
            string ec_dir_path = "";
            string x_code = "";

            //テキストリーダー
            string table_choise = "ecitem";
            Console.WriteLine("デスクトップにある、A列に商品番号のみを書いたCSVファイルを指定してください。");
            Console.Write("FileName(.csv) : ");
            string txt_path = @"F:\Users\kiyoka\Desktop\" + Console.ReadLine() + ".csv";
            System.IO.StreamReader cReader1 = (new System.IO.StreamReader(txt_path, System.Text.Encoding.Default));
            while (cReader1.Peek() > 0)
            {
                string mycode = cReader1.ReadLine();
                //データベースへ接続
                MySqlConnection conn = new MySqlConnection();
                MySqlCommand command = new MySqlCommand();
                conn.ConnectionString = "server=localhost;user id=mareki;Password=MUqvHc59UKDrEdz8;persist security info=True;database=marekidatabase";
                string select_table = table_choise;
                conn.Open();

                //SELECT文を設定
                int ncount = 1;
                command.CommandText = "SELECT * FROM " + select_table + " WHERE mycode = '" + mycode + "' ORDER BY RAND() LIMIT " + ncount + ";";
                command.Connection = conn;

                //SQLを実行し、データを格納し、接続の解除
                MySqlDataReader reader = command.ExecuteReader();
                string gen_サムネイル = "";
                string seo_画像説明_メイン = "";
                string gen_追加説明 = "";
                string gen_楽天_商品名 = "";
                string gen_楽天_PCキャッチコピー = "";
                string gen_楽天_MBキャッチコピー = "";
                string org_楽天_商品名 = "";
                string org_楽天_PCキャッチコピー = "";
                string org_楽天_MBキャッチコピー = "";
                string gen_Yahoo_商品名 = "";
                string gen_Yahoo_キャッチコピー = "";
                string org_Yahoo_商品名 = "";
                string org_Yahoo_キャッチコピー = "";
                string gen_Yahoo_metadesc = "";
                string gen_Yahoo_metakey = "";
                string gen_カゴ横説明 = "";
                string p_当店通常価格_税込み = "";

                for (int i = 0; reader.Read(); i++)
                {
                    gen_サムネイル = reader.GetString("gen_サムネイル");
                    seo_画像説明_メイン = reader.GetString("seo_画像説明_メイン");
                    gen_追加説明 = reader.GetString("gen_追加説明");
                    gen_楽天_商品名 = reader.GetString("gen_楽天_商品名");
                    gen_楽天_PCキャッチコピー = reader.GetString("gen_楽天_PCキャッチコピー");
                    gen_楽天_MBキャッチコピー = reader.GetString("gen_楽天_MBキャッチコピー");
                    org_楽天_商品名 = reader.GetString("org_楽天_商品名");
                    org_楽天_PCキャッチコピー = reader.GetString("org_楽天_PCキャッチコピー");
                    org_楽天_MBキャッチコピー = reader.GetString("org_楽天_MBキャッチコピー");
                    gen_Yahoo_商品名 = reader.GetString("gen_Yahoo_商品名");
                    gen_Yahoo_キャッチコピー = reader.GetString("gen_Yahoo_キャッチコピー");
                    org_Yahoo_商品名 = reader.GetString("org_Yahoo_商品名");
                    org_Yahoo_キャッチコピー = reader.GetString("org_Yahoo_キャッチコピー");
                    gen_Yahoo_metadesc = reader.GetString("gen_Yahoo_metadesc");
                    gen_Yahoo_metakey = reader.GetString("gen_Yahoo_metakey");
                    gen_カゴ横説明 = reader.GetString("gen_カゴ横説明");
                    p_当店通常価格_税込み = reader.GetString("p_当店通常価格_税込み");
                }
                Console.WriteLine(mycode);
                reader.Close();
                conn.Close();

                //改行置換
                gen_カゴ横説明 = gen_カゴ横説明.Replace("\r\n", "temp_kaigyou");
                gen_カゴ横説明 = gen_カゴ横説明.Replace("\n", "\r\n");
                gen_カゴ横説明 = gen_カゴ横説明.Replace("\r", "\r\n");
                gen_カゴ横説明 = gen_カゴ横説明.Replace("temp_kaigyou", "\r\n");
                gen_カゴ横説明 = gen_カゴ横説明.Replace("\r\n", "<br>");


                //テキスト読み込み
                if (mycode.Contains("roomkoubou--"))
                    ec_dir_path = @"D:\03.Ec_GodBot\00.設計_system\system_ShohinPageGenerator\roomkoubou\";
                if (mycode.Contains("mokuto--"))
                    ec_dir_path = @"D:\03.Ec_GodBot\00.設計_system\system_ShohinPageGenerator\mokuto\";
                if (mycode.Contains("marukomaruko--"))
                    ec_dir_path = @"D:\03.Ec_GodBot\00.設計_system\system_ShohinPageGenerator\marukomaruko\";
                if (mycode.Contains("zen--"))
                    ec_dir_path = @"D:\03.Ec_GodBot\00.設計_system\system_ShohinPageGenerator\zen\";
                System.IO.StreamReader sr = new System.IO.StreamReader(ec_dir_path + "ジェネレートテンプレート.html", System.Text.Encoding.GetEncoding("utf-8"));
                string s = sr.ReadToEnd();
                sr.Close();
                Console.WriteLine(s);
                string s_rep = s.Replace("xxx_商品コード", mycode);
                s_rep = s_rep.Replace("xxx_gen_サムネイル", gen_サムネイル);
                s_rep = s_rep.Replace("xxx_seo_画像説明_メイン", seo_画像説明_メイン);
                s_rep = s_rep.Replace("xxx_gen_追加説明", gen_追加説明);
                s_rep = s_rep.Replace("xxx_gen_楽天_商品名", gen_楽天_商品名);
                s_rep = s_rep.Replace("xxx_gen_楽天_PCキャッチコピー", gen_楽天_PCキャッチコピー);
                s_rep = s_rep.Replace("xxx_gen_楽天_MBキャッチコピー", gen_楽天_MBキャッチコピー);
                s_rep = s_rep.Replace("xxx_org_楽天_商品名", org_楽天_商品名);
                s_rep = s_rep.Replace("xxx_org_楽天_PCキャッチコピー", org_楽天_PCキャッチコピー);
                s_rep = s_rep.Replace("xxx_org_楽天_MBキャッチコピー", org_楽天_MBキャッチコピー);
                s_rep = s_rep.Replace("xxx_gen_Yahoo_商品名", gen_Yahoo_商品名);
                s_rep = s_rep.Replace("xxx_gen_Yahoo_キャッチコピー", gen_Yahoo_キャッチコピー);
                s_rep = s_rep.Replace("xxx_org_Yahoo_商品名", org_Yahoo_商品名);
                s_rep = s_rep.Replace("xxx_org_Yahoo_キャッチコピー", org_Yahoo_キャッチコピー);
                s_rep = s_rep.Replace("xxx_gen_Yahoo_metadesc", gen_Yahoo_metadesc);
                s_rep = s_rep.Replace("xxx_gen_Yahoo_metakey", gen_Yahoo_metakey);
                s_rep = s_rep.Replace("xxx_gen_カゴ横説明", gen_カゴ横説明);
                s_rep = s_rep.Replace("xxx_p_当店通常価格_税込み", p_当店通常価格_税込み);

                s_rep = s_rep.Replace("byte_seo_画像説明_メイン", Encoding.GetEncoding("Shift_JIS").GetByteCount(seo_画像説明_メイン).ToString());
                s_rep = s_rep.Replace("byte_gen_追加説明", Encoding.GetEncoding("Shift_JIS").GetByteCount(gen_追加説明).ToString());
                s_rep = s_rep.Replace("byte_gen_楽天_商品名", Encoding.GetEncoding("Shift_JIS").GetByteCount(gen_楽天_商品名).ToString());
                s_rep = s_rep.Replace("byte_gen_楽天_PCキャッチコピー", Encoding.GetEncoding("Shift_JIS").GetByteCount(gen_楽天_PCキャッチコピー).ToString());
                s_rep = s_rep.Replace("byte_gen_楽天_MBキャッチコピー", Encoding.GetEncoding("Shift_JIS").GetByteCount(gen_楽天_MBキャッチコピー).ToString());
                s_rep = s_rep.Replace("byte_org_楽天_商品名", Encoding.GetEncoding("Shift_JIS").GetByteCount(org_楽天_商品名).ToString());
                s_rep = s_rep.Replace("byte_org_楽天_PCキャッチコピー", Encoding.GetEncoding("Shift_JIS").GetByteCount(org_楽天_PCキャッチコピー).ToString());
                s_rep = s_rep.Replace("byte_org_楽天_MBキャッチコピー", Encoding.GetEncoding("Shift_JIS").GetByteCount(org_楽天_MBキャッチコピー).ToString());
                s_rep = s_rep.Replace("byte_gen_Yahoo_商品名", Encoding.GetEncoding("Shift_JIS").GetByteCount(gen_Yahoo_商品名).ToString());
                s_rep = s_rep.Replace("byte_gen_Yahoo_キャッチコピー", Encoding.GetEncoding("Shift_JIS").GetByteCount(gen_Yahoo_キャッチコピー).ToString());
                s_rep = s_rep.Replace("byte_org_Yahoo_商品名", Encoding.GetEncoding("Shift_JIS").GetByteCount(org_Yahoo_商品名).ToString());
                s_rep = s_rep.Replace("byte_org_Yahoo_キャッチコピー", Encoding.GetEncoding("Shift_JIS").GetByteCount(org_Yahoo_キャッチコピー).ToString());
                s_rep = s_rep.Replace("byte_gen_Yahoo_metadesc", Encoding.GetEncoding("Shift_JIS").GetByteCount(gen_Yahoo_metadesc).ToString());
                s_rep = s_rep.Replace("byte_gen_Yahoo_metakey", Encoding.GetEncoding("Shift_JIS").GetByteCount(gen_Yahoo_metakey).ToString());
                s_rep = s_rep.Replace("byte_gen_カゴ横説明", Encoding.GetEncoding("Shift_JIS").GetByteCount(gen_カゴ横説明).ToString());
                s_rep = s_rep.Replace("byte_p_当店通常価格_税込み", Encoding.GetEncoding("Shift_JIS").GetByteCount(p_当店通常価格_税込み).ToString());

                //MessageBox.Show(s_rep);

                if (4900 < Encoding.GetEncoding("Shift_JIS").GetByteCount(seo_画像説明_メイン) + Encoding.GetEncoding("Shift_JIS").GetByteCount(gen_追加説明))
                    s_rep = s_rep.Replace("chk_seo_画像説明_メイン+gen_追加説明", "エラー");
                else
                    s_rep = s_rep.Replace("chk_seo_画像説明_メイン+gen_追加説明", "合格");

                if (256 < Encoding.GetEncoding("Shift_JIS").GetByteCount(gen_楽天_商品名))
                    s_rep = s_rep.Replace("chk_gen_楽天_商品名", "エラー");
                else
                    s_rep = s_rep.Replace("chk_gen_楽天_商品名", "合格");

                if (256 < Encoding.GetEncoding("Shift_JIS").GetByteCount(gen_楽天_PCキャッチコピー))
                    s_rep = s_rep.Replace("chk_gen_楽天_PCキャッチコピー", "エラー");
                else
                    s_rep = s_rep.Replace("chk_gen_楽天_PCキャッチコピー", "合格");

                if (256 < Encoding.GetEncoding("Shift_JIS").GetByteCount(gen_楽天_MBキャッチコピー))
                    s_rep = s_rep.Replace("chk_gen_楽天_MBキャッチコピー", "エラー");
                else
                    s_rep = s_rep.Replace("chk_gen_楽天_MBキャッチコピー", "合格");

                if (256 < Encoding.GetEncoding("Shift_JIS").GetByteCount(org_楽天_商品名))
                    s_rep = s_rep.Replace("chk_org_楽天_商品名", "エラー");
                else
                    s_rep = s_rep.Replace("chk_org_楽天_商品名", "合格");

                if (256 < Encoding.GetEncoding("Shift_JIS").GetByteCount(org_楽天_PCキャッチコピー))
                    s_rep = s_rep.Replace("chk_org_楽天_PCキャッチコピー", "エラー");
                else
                    s_rep = s_rep.Replace("chk_org_楽天_PCキャッチコピー", "合格");

                if (256 < Encoding.GetEncoding("Shift_JIS").GetByteCount(org_楽天_MBキャッチコピー))
                    s_rep = s_rep.Replace("chk_org_楽天_MBキャッチコピー", "エラー");
                else
                    s_rep = s_rep.Replace("chk_org_楽天_MBキャッチコピー", "合格");

                if (160 < Encoding.GetEncoding("Shift_JIS").GetByteCount(gen_Yahoo_商品名))
                    s_rep = s_rep.Replace("chk_gen_Yahoo_商品名", "エラー");
                else if (60 < Encoding.GetEncoding("Shift_JIS").GetByteCount(gen_Yahoo_商品名))
                    s_rep = s_rep.Replace("chk_gen_Yahoo_商品名", "合格");
                else
                    s_rep = s_rep.Replace("chk_gen_Yahoo_商品名", "合格");

                if (160 < Encoding.GetEncoding("Shift_JIS").GetByteCount(org_Yahoo_商品名))
                    s_rep = s_rep.Replace("chk_org_Yahoo_商品名", "エラー");
                else if (60 < Encoding.GetEncoding("Shift_JIS").GetByteCount(org_Yahoo_商品名))
                    s_rep = s_rep.Replace("chk_org_Yahoo_商品名", "合格");
                else
                    s_rep = s_rep.Replace("chk_org_Yahoo_商品名", "合格");

                if (60 < Encoding.GetEncoding("Shift_JIS").GetByteCount(gen_Yahoo_キャッチコピー))
                    s_rep = s_rep.Replace("chk_gen_Yahoo_キャッチコピー", "エラー");
                else
                    s_rep = s_rep.Replace("chk_gen_Yahoo_キャッチコピー", "合格");

                if (60 < Encoding.GetEncoding("Shift_JIS").GetByteCount(org_Yahoo_キャッチコピー))
                    s_rep = s_rep.Replace("chk_org_Yahoo_キャッチコピー", "エラー");
                else
                    s_rep = s_rep.Replace("chk_org_Yahoo_キャッチコピー", "合格");

                if (160 < Encoding.GetEncoding("Shift_JIS").GetByteCount(gen_Yahoo_metadesc))
                    s_rep = s_rep.Replace("chk_gen_Yahoo_metadesc", "エラー");
                else
                    s_rep = s_rep.Replace("chk_gen_Yahoo_metadesc", "合格");

                if (160 < Encoding.GetEncoding("Shift_JIS").GetByteCount(gen_Yahoo_metakey))
                    s_rep = s_rep.Replace("chk_gen_Yahoo_metakey", "エラー");
                else
                    s_rep = s_rep.Replace("chk_gen_Yahoo_metakey", "合格");

                if (5000 < Encoding.GetEncoding("Shift_JIS").GetByteCount(gen_カゴ横説明))
                    s_rep = s_rep.Replace("chk_gen_カゴ横説明", "エラー");
                else
                    s_rep = s_rep.Replace("chk_gen_カゴ横説明", "合格");


                //書き込むファイルが既に存在している場合は、上書きする
                //System.IO.StreamWriter sw = new System.IO.StreamWriter(ec_dir_path + mycode + ".html",false,System.Text.Encoding.GetEncoding("utf-8"));
                //sw.Write(s_rep);
                //sw.Close();
                x_code = x_code + s_rep + "<br><br><hr><hr><hr><br><br>";
                Console.WriteLine("Database Uped");
            }
            cReader1.Close();
            //書き込むファイルが既に存在している場合は、上書きする
            System.IO.StreamWriter swx = new System.IO.StreamWriter(ec_dir_path + "統合.html", false, System.Text.Encoding.GetEncoding("utf-8"));
            swx.Write(x_code);
            swx.Close();

            Console.WriteLine("完了");
        }
        public static void ecitem_encoder()
        {
            //ここでカテゴリをコンバートする。 DIR_IDもコンバートする？ EXCELで表にして。




            //テキストリーダー
            string table_choise = "ecitem";
            Console.WriteLine("デスクトップにある、A列に商品番号のみを書いたCSVファイルを指定してください。");
            Console.Write("FileName(.csv) : ");
            string txt_path = @"F:\Users\kiyoka\Desktop\" + Console.ReadLine() + ".csv";
            System.IO.StreamReader cReader1 = (new System.IO.StreamReader(txt_path, System.Text.Encoding.Default));

            Console.Write("outpath : ");
            string outputfilename = Console.ReadLine();
            string outputfilename_yahoo = outputfilename + "_yahoo";
            string outputfilename_rakuten_item = outputfilename + "_rakuten_item";
            string outputfilename_rakuten_item_soukodasi = outputfilename + "_rakuten_item_soukodasi";
            string outputfilename_rakuten_select = outputfilename + "_rakuten_select";
            string outputfilename_rakuten_category = outputfilename + "_rakuten_category";
            string outputfilename_rakuten_auction = outputfilename + "_rakuten_auction";
            string outputfilename_rakuten_auction_soukodasi = outputfilename + "_rakuten_auction_soukodasi";


            outputfilename_yahoo = @"F:\Users\kiyoka\Desktop\" + outputfilename_yahoo + ".csv";
            outputfilename_rakuten_item = @"F:\Users\kiyoka\Desktop\" + outputfilename_rakuten_item + ".csv";
            outputfilename_rakuten_item_soukodasi = @"F:\Users\kiyoka\Desktop\" + outputfilename_rakuten_item_soukodasi + ".csv";
            outputfilename_rakuten_select = @"F:\Users\kiyoka\Desktop\" + outputfilename_rakuten_select + ".csv";
            outputfilename_rakuten_category = @"F:\Users\kiyoka\Desktop\" + outputfilename_rakuten_category + ".csv";
            outputfilename_rakuten_auction = @"F:\Users\kiyoka\Desktop\" + outputfilename_rakuten_auction + ".csv";
            outputfilename_rakuten_auction_soukodasi = @"F:\Users\kiyoka\Desktop\" + outputfilename_rakuten_auction_soukodasi + ".csv";

            //データテーブルの作成
            DataSet ds2_yahoo = new DataSet();
            DataTable dt2_yahoo = new DataTable();
            DataSet ds2_rakuten_item = new DataSet();
            DataTable dt2_rakuten_item = new DataTable();
            DataSet ds2_rakuten_item_soukodasi = new DataSet();
            DataTable dt2_rakuten_item_soukodasi = new DataTable();
            DataSet ds2_rakuten_select = new DataSet();
            DataTable dt2_rakuten_select = new DataTable();
            DataSet ds2_rakuten_category = new DataSet();
            DataTable dt2_rakuten_category = new DataTable();
            DataSet ds2_rakuten_auction = new DataSet();
            DataTable dt2_rakuten_auction = new DataTable();
            DataSet ds2_rakuten_auction_soukodasi = new DataSet();
            DataTable dt2_rakuten_auction_soukodasi = new DataTable();

            // 列定義@Yahoo
            dt2_yahoo.Columns.Add("path", Type.GetType("System.String")); //ok
            dt2_yahoo.Columns.Add("name", Type.GetType("System.String")); //ok
            dt2_yahoo.Columns.Add("code", Type.GetType("System.String")); //ok
            dt2_yahoo.Columns.Add("sub-code", Type.GetType("System.String"));
            dt2_yahoo.Columns.Add("price", Type.GetType("System.String")); //ok
            dt2_yahoo.Columns.Add("sale-price", Type.GetType("System.String"));
            dt2_yahoo.Columns.Add("options", Type.GetType("System.String"));
            dt2_yahoo.Columns.Add("headline", Type.GetType("System.String")); //ok
            dt2_yahoo.Columns.Add("explanation", Type.GetType("System.String"));
            dt2_yahoo.Columns.Add("caption", Type.GetType("System.String")); //ok
            dt2_yahoo.Columns.Add("additional1", Type.GetType("System.String")); //ok
            dt2_yahoo.Columns.Add("sp-additional", Type.GetType("System.String"));
            //dt2_yahoo.Columns.Add("point-code", Type.GetType("System.String"));
            dt2_yahoo.Columns.Add("meta-key", Type.GetType("System.String")); //ok
            dt2_yahoo.Columns.Add("meta-desc", Type.GetType("System.String")); //ok
            dt2_yahoo.Columns.Add("display", Type.GetType("System.String")); //ok
            dt2_yahoo.Columns.Add("delivery", Type.GetType("System.String")); //ok
            dt2_yahoo.Columns.Add("product-category", Type.GetType("System.String"));
            dt2_yahoo.Columns.Add("spec1", Type.GetType("System.String"));
            dt2_yahoo.Columns.Add("spec2", Type.GetType("System.String"));
            dt2_yahoo.Columns.Add("spec3", Type.GetType("System.String"));
            dt2_yahoo.Columns.Add("spec4", Type.GetType("System.String"));
            dt2_yahoo.Columns.Add("spec5", Type.GetType("System.String"));
            dt2_yahoo.Columns.Add("spec6", Type.GetType("System.String"));
            dt2_yahoo.Columns.Add("sort", Type.GetType("System.String"));
            DataRow row_yahoo = dt2_yahoo.NewRow();

            // 列定義@rakuten_item
            dt2_rakuten_item.Columns.Add("コントロールカラム", Type.GetType("System.String")); //ok
            dt2_rakuten_item.Columns.Add("商品管理番号（商品URL）", Type.GetType("System.String")); //ok
            dt2_rakuten_item.Columns.Add("商品番号", Type.GetType("System.String")); //ok
            dt2_rakuten_item.Columns.Add("全商品ディレクトリID", Type.GetType("System.String")); //ok
            dt2_rakuten_item.Columns.Add("タグID", Type.GetType("System.String")); //ok
            dt2_rakuten_item.Columns.Add("PC用キャッチコピー", Type.GetType("System.String")); //ok
            dt2_rakuten_item.Columns.Add("モバイル用キャッチコピー", Type.GetType("System.String")); //ok
            dt2_rakuten_item.Columns.Add("商品名", Type.GetType("System.String")); //ok
            dt2_rakuten_item.Columns.Add("販売価格", Type.GetType("System.String")); //ok
            dt2_rakuten_item.Columns.Add("表示価格", Type.GetType("System.String")); //空
            dt2_rakuten_item.Columns.Add("消費税", Type.GetType("System.String")); //FIX
            dt2_rakuten_item.Columns.Add("送料", Type.GetType("System.String")); //FIX
            dt2_rakuten_item.Columns.Add("個別送料", Type.GetType("System.String")); //FIX
            dt2_rakuten_item.Columns.Add("送料区分1", Type.GetType("System.String")); //FIX
            dt2_rakuten_item.Columns.Add("送料区分2", Type.GetType("System.String")); //FIX
            dt2_rakuten_item.Columns.Add("代引料", Type.GetType("System.String")); //FIX
            dt2_rakuten_item.Columns.Add("倉庫指定", Type.GetType("System.String")); //FIX
            dt2_rakuten_item.Columns.Add("商品情報レイアウト", Type.GetType("System.String")); //FIX
            dt2_rakuten_item.Columns.Add("注文ボタン", Type.GetType("System.String")); //FIX
            dt2_rakuten_item.Columns.Add("資料請求ボタン", Type.GetType("System.String")); //FIX
            dt2_rakuten_item.Columns.Add("商品問い合わせボタン", Type.GetType("System.String")); //FIX
            dt2_rakuten_item.Columns.Add("再入荷お知らせボタン", Type.GetType("System.String")); //FIX
            dt2_rakuten_item.Columns.Add("のし対応", Type.GetType("System.String")); //FIX
            dt2_rakuten_item.Columns.Add("PC用商品説明文", Type.GetType("System.String")); //ok
            dt2_rakuten_item.Columns.Add("モバイル用商品説明文", Type.GetType("System.String")); //ok
            dt2_rakuten_item.Columns.Add("スマートフォン用商品説明文", Type.GetType("System.String")); //ok
            dt2_rakuten_item.Columns.Add("PC用販売説明文", Type.GetType("System.String")); //ok
            dt2_rakuten_item.Columns.Add("商品画像URL", Type.GetType("System.String")); //ok
            dt2_rakuten_item.Columns.Add("商品画像名（ALT）", Type.GetType("System.String")); //FIX
            dt2_rakuten_item.Columns.Add("動画", Type.GetType("System.String")); //fix
            dt2_rakuten_item.Columns.Add("販売期間指定", Type.GetType("System.String")); //fix
            dt2_rakuten_item.Columns.Add("注文受付数", Type.GetType("System.String")); //fix
            dt2_rakuten_item.Columns.Add("在庫タイプ", Type.GetType("System.String")); //fix
            dt2_rakuten_item.Columns.Add("在庫数", Type.GetType("System.String")); //fix
            dt2_rakuten_item.Columns.Add("在庫数表示", Type.GetType("System.String")); //fix
            dt2_rakuten_item.Columns.Add("項目選択肢別在庫用横軸項目名", Type.GetType("System.String")); //fix
            dt2_rakuten_item.Columns.Add("項目選択肢別在庫用縦軸項目名", Type.GetType("System.String")); //fix
            dt2_rakuten_item.Columns.Add("項目選択肢別在庫用残り表示閾値", Type.GetType("System.String")); //fix
            dt2_rakuten_item.Columns.Add("RAC番号", Type.GetType("System.String")); //fix
            dt2_rakuten_item.Columns.Add("サーチ非表示", Type.GetType("System.String")); //fix
            dt2_rakuten_item.Columns.Add("闇市パスワード", Type.GetType("System.String")); //fix
            dt2_rakuten_item.Columns.Add("カタログID", Type.GetType("System.String")); //fix
            dt2_rakuten_item.Columns.Add("在庫戻しフラグ", Type.GetType("System.String")); //fix
            dt2_rakuten_item.Columns.Add("在庫切れ時の注文受付", Type.GetType("System.String")); //fix
            dt2_rakuten_item.Columns.Add("在庫あり時納期管理番号", Type.GetType("System.String")); //fix
            dt2_rakuten_item.Columns.Add("在庫切れ時納期管理番号", Type.GetType("System.String")); //fix
            dt2_rakuten_item.Columns.Add("予約商品発売日", Type.GetType("System.String")); //fix
            dt2_rakuten_item.Columns.Add("ポイント変倍率", Type.GetType("System.String")); //fix
            dt2_rakuten_item.Columns.Add("ポイント変倍率適用期間", Type.GetType("System.String")); //fix
            dt2_rakuten_item.Columns.Add("ヘッダー・フッター・レフトナビ", Type.GetType("System.String")); //fix
            dt2_rakuten_item.Columns.Add("表示項目の並び順", Type.GetType("System.String")); //fix
            dt2_rakuten_item.Columns.Add("共通説明文（小）", Type.GetType("System.String")); //fix
            dt2_rakuten_item.Columns.Add("目玉商品", Type.GetType("System.String")); //fix
            dt2_rakuten_item.Columns.Add("共通説明文（大）", Type.GetType("System.String")); //fix
            dt2_rakuten_item.Columns.Add("レビュー本文表示", Type.GetType("System.String")); //fix
            dt2_rakuten_item.Columns.Add("サイズ表リンク", Type.GetType("System.String")); //fix
            dt2_rakuten_item.Columns.Add("二重価格文言管理番号", Type.GetType("System.String")); //fix
            DataRow row_rakuten_item = dt2_rakuten_item.NewRow();

            // 列定義@rakuten_item_倉庫だし
            dt2_rakuten_item_soukodasi.Columns.Add("コントロールカラム", Type.GetType("System.String")); //ok
            dt2_rakuten_item_soukodasi.Columns.Add("商品管理番号（商品URL）", Type.GetType("System.String")); //ok
            dt2_rakuten_item_soukodasi.Columns.Add("倉庫指定", Type.GetType("System.String")); //FIX
            DataRow row_rakuten_item_soukodasi = dt2_rakuten_item_soukodasi.NewRow();

            // 列定義@rakuten_auction_倉庫だし
            dt2_rakuten_auction_soukodasi.Columns.Add("コントロールカラム", Type.GetType("System.String")); //ok
            dt2_rakuten_auction_soukodasi.Columns.Add("商品管理番号（商品URL）", Type.GetType("System.String")); //ok
            dt2_rakuten_auction_soukodasi.Columns.Add("倉庫指定", Type.GetType("System.String")); //FIX
            DataRow row_rakuten_auction_soukodasi = dt2_rakuten_auction_soukodasi.NewRow();

            // 列定義@rakuten_select
            dt2_rakuten_select.Columns.Add("項目選択肢用コントロールカラム", Type.GetType("System.String")); //ok
            dt2_rakuten_select.Columns.Add("商品管理番号（商品URL）", Type.GetType("System.String")); //ok
            dt2_rakuten_select.Columns.Add("選択肢タイプ", Type.GetType("System.String"));
            dt2_rakuten_select.Columns.Add("Select/Checkbox用項目名", Type.GetType("System.String"));
            dt2_rakuten_select.Columns.Add("Select/Checkbox用選択肢", Type.GetType("System.String"));
            dt2_rakuten_select.Columns.Add("項目選択肢別在庫用横軸選択肢", Type.GetType("System.String"));
            dt2_rakuten_select.Columns.Add("項目選択肢別在庫用横軸選択肢子番号", Type.GetType("System.String"));
            dt2_rakuten_select.Columns.Add("項目選択肢別在庫用縦軸選択肢", Type.GetType("System.String"));
            dt2_rakuten_select.Columns.Add("項目選択肢別在庫用縦軸選択肢子番号", Type.GetType("System.String"));
            dt2_rakuten_select.Columns.Add("項目選択肢別在庫用取り寄せ可能表示", Type.GetType("System.String"));
            dt2_rakuten_select.Columns.Add("項目選択肢別在庫用在庫数", Type.GetType("System.String"));
            dt2_rakuten_select.Columns.Add("在庫戻しフラグ", Type.GetType("System.String"));
            dt2_rakuten_select.Columns.Add("在庫切れ時の注文受付", Type.GetType("System.String"));
            dt2_rakuten_select.Columns.Add("在庫あり時納期管理番号", Type.GetType("System.String"));
            dt2_rakuten_select.Columns.Add("在庫切れ時納期管理番号", Type.GetType("System.String"));
            dt2_rakuten_select.Columns.Add("タグID", Type.GetType("System.String"));
            dt2_rakuten_select.Columns.Add("画像URL", Type.GetType("System.String"));
            DataRow row_rakuten_select = dt2_rakuten_select.NewRow();

            // 列定義@rakuten_category
            dt2_rakuten_category.Columns.Add("コントロールカラム", Type.GetType("System.String")); //ok
            dt2_rakuten_category.Columns.Add("商品管理番号（商品URL）", Type.GetType("System.String")); //ok
            dt2_rakuten_category.Columns.Add("商品名", Type.GetType("System.String")); //ok
            dt2_rakuten_category.Columns.Add("表示先カテゴリ", Type.GetType("System.String")); //ok←コンバートが未だ
            dt2_rakuten_category.Columns.Add("優先度", Type.GetType("System.String"));
            dt2_rakuten_category.Columns.Add("URL", Type.GetType("System.String"));
            dt2_rakuten_category.Columns.Add("1ページ複数形式", Type.GetType("System.String"));
            DataRow row_rakuten_category = dt2_rakuten_category.NewRow();

            // 列定義@rakuten_auction
            dt2_rakuten_auction.Columns.Add("コントロールカラム", Type.GetType("System.String")); //ok
            dt2_rakuten_auction.Columns.Add("商品管理番号（商品URL）", Type.GetType("System.String")); //ok
            dt2_rakuten_auction.Columns.Add("商品番号", Type.GetType("System.String")); //ok
            dt2_rakuten_auction.Columns.Add("全商品ディレクトリID", Type.GetType("System.String")); //ok
            dt2_rakuten_auction.Columns.Add("PC用キャッチコピー", Type.GetType("System.String")); //ok
            dt2_rakuten_auction.Columns.Add("モバイル用キャッチコピー", Type.GetType("System.String")); //ok
            dt2_rakuten_auction.Columns.Add("商品名", Type.GetType("System.String")); //ok
            dt2_rakuten_auction.Columns.Add("商品状態", Type.GetType("System.String"));
            dt2_rakuten_auction.Columns.Add("開始価格", Type.GetType("System.String"));
            dt2_rakuten_auction.Columns.Add("最高入札価格", Type.GetType("System.String"));
            dt2_rakuten_auction.Columns.Add("即落価格", Type.GetType("System.String"));
            dt2_rakuten_auction.Columns.Add("決済方法", Type.GetType("System.String"));
            dt2_rakuten_auction.Columns.Add("消費税", Type.GetType("System.String")); //ok
            dt2_rakuten_auction.Columns.Add("送料", Type.GetType("System.String")); //ok
            dt2_rakuten_auction.Columns.Add("個別送料", Type.GetType("System.String")); //ok
            dt2_rakuten_auction.Columns.Add("代引料", Type.GetType("System.String")); //ok
            dt2_rakuten_auction.Columns.Add("個別代引料", Type.GetType("System.String"));
            dt2_rakuten_auction.Columns.Add("倉庫指定", Type.GetType("System.String")); //ok
            dt2_rakuten_auction.Columns.Add("商品問い合わせボタン", Type.GetType("System.String")); //FIX
            dt2_rakuten_auction.Columns.Add("モバイル表示", Type.GetType("System.String"));
            dt2_rakuten_auction.Columns.Add("PC用商品説明文", Type.GetType("System.String")); //ok
            dt2_rakuten_auction.Columns.Add("モバイル用商品説明文", Type.GetType("System.String")); //ok
            dt2_rakuten_auction.Columns.Add("PC用販売説明文", Type.GetType("System.String")); //ok
            dt2_rakuten_auction.Columns.Add("商品画像URL", Type.GetType("System.String")); //ok
            dt2_rakuten_auction.Columns.Add("商品画像名（ALT）", Type.GetType("System.String")); //FIX
            dt2_rakuten_auction.Columns.Add("商品アイコン", Type.GetType("System.String"));
            dt2_rakuten_auction.Columns.Add("入札期間設定（開始日時）", Type.GetType("System.String"));
            dt2_rakuten_auction.Columns.Add("入札期間設定（終了日時）", Type.GetType("System.String"));
            dt2_rakuten_auction.Columns.Add("最大入札個数", Type.GetType("System.String"));
            dt2_rakuten_auction.Columns.Add("出品個数", Type.GetType("System.String"));
            dt2_rakuten_auction.Columns.Add("入札履歴表示", Type.GetType("System.String"));
            dt2_rakuten_auction.Columns.Add("自動再出品回数", Type.GetType("System.String"));
            dt2_rakuten_auction.Columns.Add("自動延長フラグ", Type.GetType("System.String"));
            dt2_rakuten_auction.Columns.Add("早期終了フラグ", Type.GetType("System.String"));
            dt2_rakuten_auction.Columns.Add("ヘッダー・フッター・レフトナビ", Type.GetType("System.String"));
            dt2_rakuten_auction.Columns.Add("表示項目の並び順", Type.GetType("System.String"));
            dt2_rakuten_auction.Columns.Add("共通説明文（小）", Type.GetType("System.String"));
            dt2_rakuten_auction.Columns.Add("共通説明文（大）", Type.GetType("System.String"));
            DataRow row_rakuten_auction = dt2_rakuten_auction.NewRow();



            while (cReader1.Peek() > 0)
            {
                //データベースのアウトプット
                using (MySqlDataAdapter adapter = new MySqlDataAdapter())
                {
                    string mycode = cReader1.ReadLine();
                    //データベースへ接続
                    MySqlConnection conn = new MySqlConnection();
                    MySqlCommand command = new MySqlCommand();
                    conn.ConnectionString = "server=localhost;user id=mareki;Password=MUqvHc59UKDrEdz8;persist security info=True;database=marekidatabase";
                    conn.Open();

                    //SELECT文を設定
                    int ncount = 1;
                    command.CommandText = "SELECT * FROM " + table_choise + " WHERE mycode = '" + mycode + "' ORDER BY RAND() LIMIT " + ncount + ";";
                    command.Connection = conn;



                    //SQLを実行し、データを格納し、接続の解除
                    MySqlDataReader reader = command.ExecuteReader();
                    for (int i = 0; reader.Read(); i++)
                    {
                        row_yahoo = dt2_yahoo.NewRow();
                        row_rakuten_item = dt2_rakuten_item.NewRow();
                        row_rakuten_item_soukodasi = dt2_rakuten_item_soukodasi.NewRow();
                        row_rakuten_auction = dt2_rakuten_auction.NewRow();
                        row_rakuten_auction_soukodasi = dt2_rakuten_auction_soukodasi.NewRow();

                        //商品番号･パスの置換
                        string 店舗コード = functionclass.正規表現("(?<findcode>.*?)--", reader.GetString("mycode"));
                        string 商品コード = functionclass.正規表現(".*?--(?<findcode>.*?)$", reader.GetString("mycode"));
                        店舗コード = 店舗コード.Replace("roomkoubou", "kagunoroomkoubou");
                        店舗コード = 店舗コード.Replace("maruko", "marukomaruko");
                        店舗コード = 店舗コード.Replace("zen", "zenshop");
                        Console.Write(店舗コード + " : " + 商品コード);

                        //商品番号
                        row_yahoo["code"] = 商品コード;
                        row_rakuten_item["商品管理番号（商品URL）"] = 商品コード;
                        row_rakuten_item_soukodasi["商品管理番号（商品URL）"] = 商品コード;
                        row_rakuten_auction["商品管理番号（商品URL）"] = 商品コード;
                        row_rakuten_auction_soukodasi["商品管理番号（商品URL）"] = 商品コード;
                        row_rakuten_item["商品番号"] = 商品コード;
                        row_rakuten_auction["商品番号"] = 商品コード;

                        //商品番号サブ
                        if (reader.GetString("sub_code").Length > 1)
                            row_yahoo["sub-code"] = reader.GetString("sub_code");


                        //カテゴリ
                        string[] loader_category_arg = { "cat_商品ジャンル1", "cat_商品ジャンル2", "cat_商品ジャンル3", "cat_全ジャンル_材料別", "cat_同シリーズ", "cat_同テイスト" };
                        string loader_category = "";
                        for (int x = 0; x < loader_category_arg.Length; x++)
                        {
                            if (reader.GetString(loader_category_arg[x]).Length > 1)
                            {
                                loader_category = loader_category + reader.GetString(loader_category_arg[x]);
                                //row_rakuten["zz_カテゴリ"] = reader.GetString("cat_商品ジャンル1");
                                loader_category = loader_category + "\r\n";
                            }
                        }
                        row_yahoo["path"] = loader_category;

                        //商品名
                        if (reader.GetString("org_Yahoo_商品名").Length > 1)
                        {
                            row_yahoo["name"] = reader.GetString("org_Yahoo_商品名");
                        }
                        else if (reader.GetString("gen_Yahoo_商品名").Length > 1)
                        {
                            row_yahoo["name"] = reader.GetString("gen_Yahoo_商品名");
                        }
                        if (reader.GetString("org_楽天_商品名").Length > 1)
                        {
                            row_rakuten_item["商品名"] = reader.GetString("org_楽天_商品名");
                            row_rakuten_auction["商品名"] = reader.GetString("org_楽天_商品名");
                        }
                        else if (reader.GetString("gen_楽天_商品名").Length > 1)
                        {
                            row_rakuten_item["商品名"] = reader.GetString("gen_楽天_商品名");
                            row_rakuten_auction["商品名"] = reader.GetString("gen_楽天_商品名");
                        }

                        //キャッチコピー
                        string fix_headline_yahoo = "";
                        if (reader.GetString("tmp_キャッチコピー_可変").Length > 1)
                            fix_headline_yahoo = reader.GetString("tmp_キャッチコピー_可変") + " ";
                        if (reader.GetString("org_Yahoo_キャッチコピー").Length > 1)
                        {
                            fix_headline_yahoo = fix_headline_yahoo + reader.GetString("org_Yahoo_キャッチコピー");
                        }
                        else if (reader.GetString("gen_Yahoo_キャッチコピー").Length > 1)
                        {
                            fix_headline_yahoo = fix_headline_yahoo + reader.GetString("gen_Yahoo_キャッチコピー");
                        }
                        row_yahoo["headline"] = fix_headline_yahoo;

                        string fix_catchcopy_pc_rakuten = "";
                        if (reader.GetString("tmp_キャッチコピー_可変").Length > 1)
                        {
                            fix_catchcopy_pc_rakuten = reader.GetString("tmp_キャッチコピー_可変") + " ";
                        }
                        if (reader.GetString("org_楽天_PCキャッチコピー").Length > 1)
                        {
                            fix_catchcopy_pc_rakuten = reader.GetString("org_楽天_PCキャッチコピー");
                        }
                        else if (reader.GetString("gen_楽天_PCキャッチコピー").Length > 1)
                        {
                            fix_catchcopy_pc_rakuten = reader.GetString("gen_楽天_PCキャッチコピー");
                        }
                        row_rakuten_item["PC用キャッチコピー"] = fix_catchcopy_pc_rakuten;
                        row_rakuten_auction["PC用キャッチコピー"] = fix_catchcopy_pc_rakuten;

                        string fix_catchcopy_mb_rakuten = "";
                        if (reader.GetString("tmp_キャッチコピー_可変").Length > 1)
                        {
                            fix_catchcopy_mb_rakuten = reader.GetString("tmp_キャッチコピー_可変") + " ";
                        }
                        if (reader.GetString("org_楽天_mbキャッチコピー").Length > 1)
                        {
                            fix_catchcopy_mb_rakuten = reader.GetString("org_楽天_mbキャッチコピー");
                        }
                        else if (reader.GetString("gen_楽天_mbキャッチコピー").Length > 1)
                        {
                            fix_catchcopy_mb_rakuten = reader.GetString("gen_楽天_mbキャッチコピー");
                        }
                        row_rakuten_item["モバイル用キャッチコピー"] = fix_catchcopy_mb_rakuten;
                        row_rakuten_auction["モバイル用キャッチコピー"] = fix_catchcopy_mb_rakuten;


                        //販売価格
                        row_yahoo["price"] = reader.GetString("p_当店通常価格_税込み");
                        row_rakuten_item["販売価格"] = reader.GetString("p_当店通常価格_税込み");


                        // 選択肢
                        if (reader.GetString("tbl_選択肢").Length > 1)
                            row_yahoo["options"] = reader.GetString("tbl_選択肢");


                        //カゴ横説明
                        string kagoyokosetumei = "";
                        if (reader.GetString("gen_カゴ横説明").Length > 4)
                            kagoyokosetumei = reader.GetString("gen_カゴ横説明");
                        else
                            kagoyokosetumei = "";
                        row_yahoo["explanation"] = kagoyokosetumei;
                        row_rakuten_item["PC用商品説明文"] = functionclass.改行BR化処理(kagoyokosetumei);
                        row_rakuten_auction["PC用商品説明文"] = functionclass.改行BR化処理(kagoyokosetumei);
                        row_rakuten_item["モバイル用商品説明文"] = "携帯電話からは、商品の詳細が表示できません。 大変申し訳ございませんが、ご購入前には必ずパソコンからご確認ください。";
                        row_rakuten_auction["モバイル用商品説明文"] = "携帯電話からは、商品の詳細が表示できません。 大変申し訳ございませんが、ご購入前には必ずパソコンからご確認ください。";
                        row_rakuten_item["スマートフォン用商品説明文"] = functionclass.改行BR化処理(kagoyokosetumei);


                        //画像説明 & キャプション
                        string fix_caption = "";
                        if (reader.GetString("tmp_画像説明_ヘッド").Length > 1)
                            fix_caption = reader.GetString("tmp_画像説明_ヘッド") + "<br>";
                        if (reader.GetString("seo_画像説明_メイン").Length > 1)
                            fix_caption = fix_caption + reader.GetString("seo_画像説明_メイン");
                        if (reader.GetString("tmp_画像説明_フット").Length > 1)
                            fix_caption = fix_caption + "<br>" + reader.GetString("tmp_画像説明_フット");
                        string yahoo_キャプション = fix_caption.Replace("./説明用/", "http://shopping.c.yimg.jp/lib/" + 店舗コード + "/");
                        yahoo_キャプション = yahoo_キャプション.Replace("./サムネイル/", "http://item.shopping.c.yimg.jp/i/g/" + 店舗コード + "_");
                        yahoo_キャプション = yahoo_キャプション.Replace("./商品URL/", "http://store.shopping.yahoo.co.jp/" + 店舗コード + "/");

                        row_yahoo["caption"] = yahoo_キャプション;
                        string rakuten_販売説明文 = fix_caption;

                        //テーブル説明 & 追加説明
                        if (reader.GetString("org_追加説明").Length > 10)
                        {
                            row_yahoo["additional1"] = reader.GetString("org_追加説明");
                            row_yahoo["sp-additional"] = reader.GetString("org_追加説明");
                            rakuten_販売説明文 = rakuten_販売説明文 + "<br>" + reader.GetString("org_追加説明");
                        }
                        else if (reader.GetString("gen_追加説明").Length > 10)
                        {
                            string yahoo_販売説明文 = reader.GetString("gen_追加説明");
                            yahoo_販売説明文 = yahoo_販売説明文.Replace("./サムネイル/", "http://item.shopping.c.yimg.jp/i/g/" + 店舗コード + "_");
                            yahoo_販売説明文 = yahoo_販売説明文.Replace(".jpg", "");
                            yahoo_販売説明文 = yahoo_販売説明文.Replace("./商品URL/", "http://store.shopping.yahoo.co.jp/" + 店舗コード + "/");

                            row_yahoo["additional1"] = yahoo_販売説明文;
                            row_yahoo["sp-additional"] = yahoo_販売説明文;

                            rakuten_販売説明文 = rakuten_販売説明文 + "<br>" + reader.GetString("gen_追加説明");
                            rakuten_販売説明文 = rakuten_販売説明文.Replace("./サムネイル/", "http://image.rakuten.co.jp/" + 店舗コード + "/cabinet/" + reader.GetString("sys_画像パス") + "/");
                            rakuten_販売説明文 = rakuten_販売説明文.Replace("./説明用/", "http://image.rakuten.co.jp/" + 店舗コード + "/cabinet/" + reader.GetString("sys_画像パス") + "/");
                            rakuten_販売説明文 = rakuten_販売説明文.Replace(".html", "/"); // .htmlを全てスラッシュにすると不具合が出るが、とりあえずok
                            rakuten_販売説明文 = rakuten_販売説明文.Replace("./商品URL/", "http://item.rakuten.co.jp/" + 店舗コード + "/");
                        }
                        else
                        {
                            row_yahoo["additional1"] = "";
                            row_yahoo["sp-additional"] = "";
                        }
                        row_rakuten_item["PC用販売説明文"] = rakuten_販売説明文;
                        row_rakuten_auction["PC用販売説明文"] = rakuten_販売説明文;

                        //サムネイル
                        string サムネイル_楽天_置換 = reader.GetString("gen_サムネイル");
                        サムネイル_楽天_置換 = サムネイル_楽天_置換.Replace("<div class=\"thumbnail\"><img src=\"./サムネイル/", "http://image.rakuten.co.jp/" + 店舗コード + "/cabinet/" + reader.GetString("sys_画像パス") + "/");
                        サムネイル_楽天_置換 = サムネイル_楽天_置換.Replace("\"></div>", "");
                        row_rakuten_item["商品画像URL"] = サムネイル_楽天_置換;
                        row_rakuten_auction["商品画像URL"] = サムネイル_楽天_置換;
                        row_rakuten_item["商品画像名（ALT）"] = "";
                        row_rakuten_auction["商品画像名（ALT）"] = "";

                        /*
                        string 商品目的 = "";
                        string 商品注目度 = "";
                        */
                        row_yahoo["meta-desc"] = reader.GetString("gen_Yahoo_metadesc");
                        row_yahoo["meta-key"] = reader.GetString("gen_Yahoo_metakey");



                        //表示･倉庫
                        //if (reader.GetString("sys_表示状態") == "表示")
                        //    row_yahoo["display"] = "1";
                        //else
                        //    row_yahoo["display"] = "0";
                        row_yahoo["display"] = "1"; //1=表示する。
                        row_rakuten_item["倉庫指定"] = "1"; //0=販売中/1=倉庫に入れる
                        row_rakuten_item_soukodasi["倉庫指定"] = "0"; //0=販売中/1=倉庫に入れる
                        row_rakuten_auction["倉庫指定"] = "1";
                        row_rakuten_auction_soukodasi["倉庫指定"] = "0";

                        if (reader.GetString("sys_送料形態") == "送料無料")
                            row_yahoo["delivery"] = "1";
                        else if (reader.GetString("sys_送料形態") == "条件付き送料無料")
                            row_yahoo["delivery"] = "3";
                        else
                            row_yahoo["delivery"] = "0";

                        //送料
                        //row_yahoo["delivery"] = "1"; //送料無料
                        row_rakuten_item["送料"] = "1"; //送料込み
                        row_rakuten_auction["送料"] = "1"; //送料込み

                        //消費税
                        row_rakuten_item["消費税"] = "1"; //1=税込,0=税別
                        row_rakuten_auction["消費税"] = "1";

                        row_yahoo["product-category"] = reader.GetString("dirid_yshop");
                        if (reader.GetString("spec1_yshop") == "x")
                            row_yahoo["spec1"] = "";
                        else
                            row_yahoo["spec1"] = reader.GetString("spec1_yshop");
                        if (reader.GetString("spec2_yshop") == "x")
                            row_yahoo["spec2"] = "";
                        else
                            row_yahoo["spec2"] = reader.GetString("spec2_yshop");
                        if (reader.GetString("spec3_yshop") == "x")
                            row_yahoo["spec3"] = "";
                        else
                            row_yahoo["spec3"] = reader.GetString("spec3_yshop");
                        if (reader.GetString("spec4_yshop") == "x")
                            row_yahoo["spec4"] = "";
                        else
                            row_yahoo["spec4"] = reader.GetString("spec4_yshop");
                        if (reader.GetString("spec5_yshop") == "x")
                            row_yahoo["spec5"] = "";
                        else
                            row_yahoo["spec5"] = reader.GetString("spec5_yshop");
                        if (reader.GetString("spec6_yshop") == "x")
                            row_yahoo["spec6"] = "";
                        else
                            row_yahoo["spec6"] = reader.GetString("spec6_yshop");


                        //ソート･優先順位
                        if (reader.GetString("sys_ソート") == "x")
                            row_yahoo["sort"] = "";
                        else
                            row_yahoo["sort"] = reader.GetString("sys_ソート");

                        row_rakuten_item["全商品ディレクトリID"] = reader.GetString("dirid_rakuten");
                        row_rakuten_auction["全商品ディレクトリID"] = reader.GetString("dirid_rakuten");
                        row_rakuten_item["タグID"] = reader.GetString("tagid_rakuten");

                        //item
                        row_rakuten_item["コントロールカラム"] = "n"; //空 n=新規登録,u=更新,d=商品の全削除
                        row_rakuten_item["個別送料"] = "";
                        row_rakuten_auction["個別送料"] = "";
                        row_rakuten_item["送料区分1"] = "";
                        row_rakuten_item["送料区分2"] = "";
                        row_rakuten_item["代引料"] = "0"; //0=代引料込み,1代引き料別
                        row_rakuten_auction["代引料"] = "0";
                        row_rakuten_item["商品情報レイアウト"] = "1"; //1=初期レイアウト
                        row_rakuten_item["注文ボタン"] = "1"; //1=通常ボタンをつける
                        row_rakuten_item["資料請求ボタン"] = "0"; //0=ボタンをつけない
                        row_rakuten_item["商品問い合わせボタン"] = "1"; //1=ボタンをつける
                        row_rakuten_auction["商品問い合わせボタン"] = "1"; //1=ボタンをつける
                        row_rakuten_item["再入荷お知らせボタン"] = "0"; //0=ボタンをつけない
                        row_rakuten_item["のし対応"] = "0"; //0=対応しない
                        row_rakuten_item["動画"] = ""; //空
                        row_rakuten_item["販売期間指定"] = ""; //空
                        row_rakuten_item["注文受付数"] = "-1"; //-1
                        row_rakuten_item["RAC番号"] = ""; //空
                        row_rakuten_item["サーチ非表示"] = "0"; //0=表示する,1=表示しない
                        row_rakuten_item["闇市パスワード"] = ""; //空
                        row_rakuten_item["カタログID"] = ""; //空
                        row_rakuten_item["在庫戻しフラグ"] = ""; //空
                        row_rakuten_item["在庫切れ時の注文受付"] = ""; //空
                        row_rakuten_item["在庫あり時納期管理番号"] = ""; //空
                        row_rakuten_item["在庫切れ時納期管理番号"] = ""; //空？？
                        row_rakuten_item["予約商品発売日"] = ""; //空
                        row_rakuten_item["ポイント変倍率"] = ""; //空
                        row_rakuten_item["ポイント変倍率適用期間"] = ""; //空
                        row_rakuten_item["ヘッダー・フッター・レフトナビ"] = "自動選択";
                        row_rakuten_item["表示項目の並び順"] = "自動選択";
                        row_rakuten_item["共通説明文（小）"] = "自動選択";
                        row_rakuten_item["目玉商品"] = "自動選択";
                        row_rakuten_item["共通説明文（大）"] = "自動選択";
                        row_rakuten_item["レビュー本文表示"] = "2"; //0=表示しない,1=表示する,2=デザイン設定での設定を使用
                        //row_rakuten_item["あす楽配送管理番号"] = ""; //空？？ 0=設定しない,1以上=あす楽管理番号
                        row_rakuten_item["サイズ表リンク"] = ""; //空
                        row_rakuten_item["二重価格文言管理番号"] = ""; //1=当店通常価格,2=メーカー希望小売価格,4=商品価格ナビのデータ参照,0=自動選択

                        //item_soukodasi
                        row_rakuten_item_soukodasi["コントロールカラム"] = "u"; //空 n=新規登録,u=更新,d=商品の全削除
                        row_rakuten_auction_soukodasi["コントロールカラム"] = "u"; //空 n=新規登録,u=更新,d=商品の全削除

                        //select
                        Boolean val_flag = true;
                        string[] 選択肢_バリエーション1軸 = reader.GetString("tbl_選択肢_商品バリエーション1軸").Split(' ');
                        if (選択肢_バリエーション1軸.Length > 1)
                        {
                            row_rakuten_item["在庫タイプ"] = "2"; //0=在庫設定しない,1=通常在庫設定,2=項目選択肢別在庫設定
                            row_rakuten_item["在庫数"] = ""; //在庫タイプが1の時に指定
                            row_rakuten_item["在庫数表示"] = ""; //空(0=表示しない)
                            row_rakuten_item["項目選択肢別在庫用横軸項目名"] = 選択肢_バリエーション1軸[0]; //在庫タイプが2の時に指定 必須
                            row_rakuten_item["項目選択肢別在庫用縦軸項目名"] = ""; //在庫タイプが2の時に指定 任意
                            row_rakuten_item["項目選択肢別在庫用残り表示閾値"] = ""; //在庫タイプが2の時に指定 -1=残り在庫数表示,0=表示しない,1～n=△で表示

                            for (int x = 1; x < 選択肢_バリエーション1軸.Length; x++)
                            {
                                row_rakuten_select = dt2_rakuten_select.NewRow();
                                row_rakuten_select["商品管理番号（商品URL）"] = 商品コード;
                                row_rakuten_select["項目選択肢用コントロールカラム"] = "n"; //空 n=新規登録,u=更新,d=削除
                                row_rakuten_select["Select/Checkbox用項目名"] = ""; //sかcの時入力必須
                                row_rakuten_select["Select/Checkbox用選択肢"] = ""; //sかcの時入力必須
                                row_rakuten_select["タグID"] = ""; //空
                                row_rakuten_select["画像URL"] = ""; //空
                                row_rakuten_select["選択肢タイプ"] = "i"; //s=セレクトボックス,c=チェックボックス,i=項目選択肢別在庫(チェックボックス)
                                row_rakuten_select["項目選択肢別在庫用横軸選択肢"] = 選択肢_バリエーション1軸[x]; //iの時入力必須
                                row_rakuten_select["項目選択肢別在庫用横軸選択肢子番号"] = ""; //空
                                row_rakuten_select["項目選択肢別在庫用縦軸選択肢"] = "";  //iの時入力必須
                                row_rakuten_select["項目選択肢別在庫用縦軸選択肢子番号"] = ""; //空
                                row_rakuten_select["項目選択肢別在庫用取り寄せ可能表示"] = "0"; //iの時 0を入力
                                row_rakuten_select["項目選択肢別在庫用在庫数"] = "100"; //iの時入力必須 iの時100を入力する
                                row_rakuten_select["在庫戻しフラグ"] = "0"; //0=利用しない,1=利用する 0を入力
                                row_rakuten_select["在庫切れ時の注文受付"] = "0"; //0=受け付けない,1=受け付ける 0を入力
                                row_rakuten_select["在庫あり時納期管理番号"] = "★★※注意 : 手動で設定※★★"; //空も有り
                                row_rakuten_select["在庫切れ時納期管理番号"] = "★★※注意 : 手動で設定※★★"; //空も有り
                                dt2_rakuten_select.Rows.Add(row_rakuten_select);
                            }
                            val_flag = false;
                        }
                        string[] 選択肢_バリエーション2軸 = reader.GetString("tbl_選択肢_商品バリエーション2軸").Split(' ');
                        if (選択肢_バリエーション2軸.Length > 1)
                        {
                            string[] 縦横スプリット = 選択肢_バリエーション2軸[0].Split('|');
                            row_rakuten_item["在庫タイプ"] = "2"; //0=在庫設定しない,1=通常在庫設定,2=項目選択肢別在庫設定
                            row_rakuten_item["在庫数"] = ""; //在庫タイプが1の時に指定
                            row_rakuten_item["在庫数表示"] = ""; //空(0=表示しない)
                            row_rakuten_item["項目選択肢別在庫用横軸項目名"] = 縦横スプリット[0]; //在庫タイプが2の時に指定 必須
                            row_rakuten_item["項目選択肢別在庫用縦軸項目名"] = 縦横スプリット[1]; //在庫タイプが2の時に指定 任意
                            row_rakuten_item["項目選択肢別在庫用残り表示閾値"] = ""; //在庫タイプが2の時に指定 -1=残り在庫数表示,0=表示しない,1～n=△で表示

                            for (int x = 1; x < 選択肢_バリエーション2軸.Length; x++)
                            {
                                Console.WriteLine(商品コード + " : " + x);
                                縦横スプリット = 選択肢_バリエーション2軸[x].Split('|');
                                row_rakuten_select = dt2_rakuten_select.NewRow();
                                row_rakuten_select["商品管理番号（商品URL）"] = 商品コード;
                                row_rakuten_select["項目選択肢用コントロールカラム"] = "n"; //空 n=新規登録,u=更新,d=削除
                                row_rakuten_select["Select/Checkbox用項目名"] = ""; //sかcの時入力必須
                                row_rakuten_select["Select/Checkbox用選択肢"] = ""; //sかcの時入力必須
                                row_rakuten_select["タグID"] = ""; //空
                                row_rakuten_select["画像URL"] = ""; //空
                                row_rakuten_select["選択肢タイプ"] = "i"; //s=セレクトボックス,c=チェックボックス,i=項目選択肢別在庫(チェックボックス)
                                row_rakuten_select["項目選択肢別在庫用横軸選択肢"] = 縦横スプリット[0]; //iの時入力必須
                                row_rakuten_select["項目選択肢別在庫用横軸選択肢子番号"] = ""; //空
                                row_rakuten_select["項目選択肢別在庫用縦軸選択肢"] = 縦横スプリット[1];  //iの時入力必須
                                row_rakuten_select["項目選択肢別在庫用縦軸選択肢子番号"] = ""; //空
                                row_rakuten_select["項目選択肢別在庫用取り寄せ可能表示"] = "0"; //iの時 0を入力
                                row_rakuten_select["項目選択肢別在庫用在庫数"] = "100"; //iの時入力必須 iの時100を入力する
                                row_rakuten_select["在庫戻しフラグ"] = "0"; //0=利用しない,1=利用する 0を入力
                                row_rakuten_select["在庫切れ時の注文受付"] = "0"; //0=受け付けない,1=受け付ける 0を入力
                                row_rakuten_select["在庫あり時納期管理番号"] = "★★※注意 : 手動で設定※★★"; //空
                                row_rakuten_select["在庫切れ時納期管理番号"] = "★★※注意 : 手動で設定※★★"; //空
                                dt2_rakuten_select.Rows.Add(row_rakuten_select);
                            }
                            val_flag = false;
                        }
                        if (val_flag)
                        {
                            //在庫タイプ
                            row_rakuten_item["在庫タイプ"] = "0"; //0=在庫設定しない,1=通常在庫設定,2=項目選択肢別在庫設定
                            row_rakuten_item["在庫数"] = ""; //在庫タイプが1の時に指定
                            row_rakuten_item["在庫数表示"] = ""; //空(0=表示しない)
                            row_rakuten_item["項目選択肢別在庫用横軸項目名"] = ""; //在庫タイプが2の時に指定 必須
                            row_rakuten_item["項目選択肢別在庫用縦軸項目名"] = ""; //在庫タイプが2の時に指定 任意
                            row_rakuten_item["項目選択肢別在庫用残り表示閾値"] = ""; //在庫タイプが2の時に指定 -1=残り在庫数表示,0=表示しない,1～n=△で表示
                        }
                        string[] 選択肢_バリエーション_プルダウン1 = reader.GetString("tbl_選択肢_プルダウン1").Split(' ');
                        if (選択肢_バリエーション_プルダウン1.Length > 1)
                        {
                            for (int x = 1; x < 選択肢_バリエーション_プルダウン1.Length; x++)
                            {
                                row_rakuten_select = dt2_rakuten_select.NewRow();
                                row_rakuten_select["商品管理番号（商品URL）"] = 商品コード;
                                row_rakuten_select["項目選択肢用コントロールカラム"] = "n"; //空 n=新規登録,u=更新,d=削除
                                row_rakuten_select["Select/Checkbox用項目名"] = 選択肢_バリエーション_プルダウン1[0]; //sかcの時入力必須
                                row_rakuten_select["Select/Checkbox用選択肢"] = 選択肢_バリエーション_プルダウン1[x];//sかcの時入力必須
                                row_rakuten_select["タグID"] = ""; //空
                                row_rakuten_select["画像URL"] = ""; //空
                                row_rakuten_select["選択肢タイプ"] = "s"; //s=セレクトボックス,c=チェックボックス,i=項目選択肢別在庫(チェックボックス)
                                row_rakuten_select["項目選択肢別在庫用横軸選択肢"] = ""; //sの時は空
                                row_rakuten_select["項目選択肢別在庫用横軸選択肢子番号"] = "";  //sの時は空
                                row_rakuten_select["項目選択肢別在庫用縦軸選択肢"] = ""; //sの時は空
                                row_rakuten_select["項目選択肢別在庫用縦軸選択肢子番号"] = "";  //sの時は空
                                row_rakuten_select["項目選択肢別在庫用取り寄せ可能表示"] = ""; //sの時は空
                                row_rakuten_select["項目選択肢別在庫用在庫数"] = ""; //sの時は空
                                row_rakuten_select["在庫戻しフラグ"] = ""; //sの時は空
                                row_rakuten_select["在庫切れ時の注文受付"] = ""; //sの時は空
                                row_rakuten_select["在庫あり時納期管理番号"] = ""; //sの時は空
                                row_rakuten_select["在庫切れ時納期管理番号"] = ""; //sの時は空
                                dt2_rakuten_select.Rows.Add(row_rakuten_select);
                            }
                        }
                        string[] 選択肢_バリエーション_プルダウン2 = reader.GetString("tbl_選択肢_プルダウン2").Split(' ');
                        if (選択肢_バリエーション_プルダウン2.Length > 1)
                        {
                            for (int x = 1; x < 選択肢_バリエーション_プルダウン2.Length; x++)
                            {
                                row_rakuten_select = dt2_rakuten_select.NewRow();
                                row_rakuten_select["商品管理番号（商品URL）"] = 商品コード;
                                row_rakuten_select["項目選択肢用コントロールカラム"] = "n"; //空 n=新規登録,u=更新,d=削除
                                row_rakuten_select["Select/Checkbox用項目名"] = 選択肢_バリエーション_プルダウン2[0]; //sかcの時入力必須
                                row_rakuten_select["Select/Checkbox用選択肢"] = 選択肢_バリエーション_プルダウン2[x];//sかcの時入力必須
                                row_rakuten_select["タグID"] = ""; //空
                                row_rakuten_select["画像URL"] = ""; //空
                                row_rakuten_select["選択肢タイプ"] = "s"; //s=セレクトボックス,c=チェックボックス,i=項目選択肢別在庫(チェックボックス)
                                row_rakuten_select["項目選択肢別在庫用横軸選択肢"] = ""; //sの時は空
                                row_rakuten_select["項目選択肢別在庫用横軸選択肢子番号"] = "";  //sの時は空
                                row_rakuten_select["項目選択肢別在庫用縦軸選択肢"] = ""; //sの時は空
                                row_rakuten_select["項目選択肢別在庫用縦軸選択肢子番号"] = "";  //sの時は空
                                row_rakuten_select["項目選択肢別在庫用取り寄せ可能表示"] = ""; //sの時は空
                                row_rakuten_select["項目選択肢別在庫用在庫数"] = ""; //sの時は空
                                row_rakuten_select["在庫戻しフラグ"] = ""; //sの時は空
                                row_rakuten_select["在庫切れ時の注文受付"] = ""; //sの時は空
                                row_rakuten_select["在庫あり時納期管理番号"] = ""; //sの時は空
                                row_rakuten_select["在庫切れ時納期管理番号"] = ""; //sの時は空
                                dt2_rakuten_select.Rows.Add(row_rakuten_select);
                            }
                        }

                        //category
                        string[] loader_category_arg_cat = { "cat_商品ジャンル1", "cat_商品ジャンル2", "cat_商品ジャンル3", "cat_全ジャンル_材料別", "cat_同シリーズ", "cat_同テイスト" };
                        for (int x = 0; x < loader_category_arg_cat.Length; x++)
                        {
                            if (reader.GetString(loader_category_arg_cat[x]).Length > 1)
                            {
                                row_rakuten_category = dt2_rakuten_category.NewRow();
                                row_rakuten_category["商品管理番号（商品URL）"] = 商品コード;
                                row_rakuten_category["表示先カテゴリ"] = reader.GetString(loader_category_arg_cat[x]).Replace(":", "\\");
                                if (reader.GetString("org_楽天_商品名").Length > 1)
                                    row_rakuten_category["商品名"] = reader.GetString("org_楽天_商品名");
                                else if (reader.GetString("gen_楽天_商品名").Length > 1)
                                    row_rakuten_category["商品名"] = reader.GetString("gen_楽天_商品名");
                                row_rakuten_category["コントロールカラム"] = "n"; //空 u=優先度の更新のみ,n=カテゴリ追加,d=削除
                                row_rakuten_category["URL"] = ""; //カテゴリURLが表示される。 入力不可
                                row_rakuten_category["1ページ複数形式"] = ""; //1を選択でメイン、2を選択でサブ
                                if (reader.GetString("sys_ソート") == "x")
                                    row_rakuten_category["優先度"] = ""; //店舗内カテゴリでの表示順位
                                else
                                    row_rakuten_category["優先度"] = reader.GetString("sys_ソート");
                                dt2_rakuten_category.Rows.Add(row_rakuten_category);
                            }
                        }
                        row_yahoo["path"] = loader_category;


                        //auction
                        row_rakuten_auction["商品状態"] = ""; //0=新品,1=中古
                        row_rakuten_auction["開始価格"] = ""; //必須
                        row_rakuten_auction["最高入札価格"] = ""; //空
                        row_rakuten_auction["即落価格"] = ""; //0=指定しない,1=指定する
                        row_rakuten_auction["決済方法"] = ""; //0=クレカ決済,1=銀行振込,99=全て
                        row_rakuten_auction["個別代引料"] = ""; //空
                        row_rakuten_auction["モバイル表示"] = "1"; //0=表示しない,1=表示する
                        row_rakuten_auction["商品アイコン"] = ""; //1=美品,2=限定,3=非売,4=新製,…,99=なし
                        row_rakuten_auction["入札期間設定（開始日時）"] = "★★※注意 : 手動で設定※★★"; //必須
                        row_rakuten_auction["入札期間設定（終了日時）"] = "★★※注意 : 手動で設定※★★"; //必須
                        row_rakuten_auction["最大入札個数"] = "★★※注意 : 手動で設定※★★"; //必須
                        row_rakuten_auction["出品個数"] = "★★※注意 : 手動で設定※★★"; //必須
                        row_rakuten_auction["入札履歴表示"] = "1"; //0=クローズ,1=オープン
                        row_rakuten_auction["自動再出品回数"] = "0"; //0~6
                        row_rakuten_auction["自動延長フラグ"] = "0"; //0=延長なし,1=延長する
                        row_rakuten_auction["早期終了フラグ"] = "0"; //0=早期終了しない,1=する
                        row_rakuten_auction["ヘッダー・フッター・レフトナビ"] = "自動選択"; //Fix
                        row_rakuten_auction["表示項目の並び順"] = "自動選択"; //Fix
                        row_rakuten_auction["共通説明文（小）"] = "自動選択"; //Fix
                        row_rakuten_auction["共通説明文（大）"] = "自動選択"; //Fix

                        dt2_yahoo.Rows.Add(row_yahoo);
                        dt2_rakuten_item.Rows.Add(row_rakuten_item);
                        dt2_rakuten_item_soukodasi.Rows.Add(row_rakuten_item_soukodasi);
                        dt2_rakuten_auction.Rows.Add(row_rakuten_auction);
                        dt2_rakuten_auction_soukodasi.Rows.Add(row_rakuten_auction_soukodasi);
                    }
                    reader.Close();
                    conn.Close();
                }
            }
            // DataSetにdtを追加します。
            ds2_yahoo.Tables.Add(dt2_yahoo);
            dt2_yahoo.TableName = "Table1";
            Console.WriteLine("Yahoo : " + dt2_yahoo.Rows.Count + "行見つかりました。");
            ds2_rakuten_item.Tables.Add(dt2_rakuten_item);
            dt2_rakuten_item.TableName = "Table1";
            Console.WriteLine("rakuten_item : " + dt2_rakuten_item.Rows.Count + "行見つかりました。");
            ds2_rakuten_item_soukodasi.Tables.Add(dt2_rakuten_item_soukodasi);
            dt2_rakuten_item_soukodasi.TableName = "Table1";
            Console.WriteLine("rakuten_item_soukodasi : " + dt2_rakuten_item_soukodasi.Rows.Count + "行見つかりました。");
            ds2_rakuten_select.Tables.Add(dt2_rakuten_select);
            dt2_rakuten_select.TableName = "Table1";
            Console.WriteLine("rakuten_select : " + dt2_rakuten_select.Rows.Count + "行見つかりました。");
            ds2_rakuten_category.Tables.Add(dt2_rakuten_category);
            dt2_rakuten_category.TableName = "Table1";
            Console.WriteLine("rakuten_category : " + dt2_rakuten_category.Rows.Count + "行見つかりました。");
            ds2_rakuten_auction.Tables.Add(dt2_rakuten_auction);
            dt2_rakuten_auction.TableName = "Table1";
            ds2_rakuten_auction_soukodasi.Tables.Add(dt2_rakuten_auction_soukodasi);
            dt2_rakuten_auction_soukodasi.TableName = "Table1";
            Console.WriteLine("rakuten_auction_soukodasi : " + dt2_rakuten_auction_soukodasi.Rows.Count + "行見つかりました。");
            Console.WriteLine("rakuten_auction : " + dt2_rakuten_auction.Rows.Count + "行見つかりました。");
            functionclass.DataTableToCsv(dt2_yahoo, outputfilename_yahoo, true);
            functionclass.DataTableToCsv(dt2_rakuten_item, outputfilename_rakuten_item, true);
            functionclass.DataTableToCsv(dt2_rakuten_item_soukodasi, outputfilename_rakuten_item_soukodasi, true);
            functionclass.DataTableToCsv(dt2_rakuten_select, outputfilename_rakuten_select, true);
            functionclass.DataTableToCsv(dt2_rakuten_category, outputfilename_rakuten_category, true);
            functionclass.DataTableToCsv(dt2_rakuten_auction, outputfilename_rakuten_auction, true);
            functionclass.DataTableToCsv(dt2_rakuten_auction_soukodasi, outputfilename_rakuten_auction_soukodasi, true);
        }
    }
}
