using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace csharp
{
    public class UtilityClass
    {
        /// <summary>
        /// テキストファイルを読み込み、中身のテキストを返す。
        /// </summary>
        /// <param name="inputTextFilePath"></param>
        /// <returns></returns>
        public static string ReadTextFile(string inputTextFilePath)
        {
            try
            {
                StreamReader sr = new StreamReader(inputTextFilePath, Encoding.GetEncoding("Shift_JIS"));
                string text = sr.ReadToEnd();
                sr.Close();
                return text;
            }
            catch (System.Exception)
            {
                MessageBox.Show($"エラー : 「{inputTextFilePath}」というファイルが見つかりませんでした。");
                return "";
            }
        }



        public static void OutputTextFile(string outputPath, string str, string str_enc = "UTF-8")
        {
            //テキストアウトプット
            //Shift_JIS , UTF-8
            var enc = Encoding.GetEncoding(str_enc);
            var writer = new StreamWriter(outputPath, false, enc);
            writer.WriteLine(str);
            writer.Close();
        }


        /// <summary>
        /// ランダムな数値を返します。
        /// </summary>
        /// <param name="iFirst"></param>
        /// <param name="iLast"></param>
        /// <returns></returns>
        public static int Random(int iFirst, int iLast)
        {
            Random cRandom = new System.Random();
            int iResult = cRandom.Next(iFirst, iLast);
            return iResult;
        }

        /// <summary>
        /// 現在時刻を数字のみで返します。string型。
        /// </summary>
        /// <returns></returns>
        public static string GetTime()
        {
            DateTime dt = DateTime.Now;
            int year = dt.Year;//年を取得する。「2000」となる。
            int month = dt.Month;//月を取得する。「9」となる。
            int day = dt.Day;//日を取得する。「30」となる。
            int hour = dt.Hour;//時間を取得する。「13」となる。
            int minute = dt.Minute;//分を取得する。「15」となる。
            int second = dt.Second;//秒を取得する。「30」となる。
            string System = ""
                + year
                + String.Format("{0:00}", month)
                + String.Format("{0:00}", day)
                + String.Format("{0:00}", hour)
                + String.Format("{0:00}", minute)
                + String.Format("{0:00}", second)
                ;
            return System;
        }

        /// <summary>
        /// DataTableからCSVへ出力する。
        /// </summary>
        /// <param name="dtDt"></param>
        /// <param name="sFilePath"></param>
        /// <param name="boolHeader"></param>
        public static void DataTableToCsv(DataTable dtDt, string sFilePath, bool boolHeader)
        {
            string sp = string.Empty;
            List<int> filterIndex = new List<int>();

            using (StreamWriter sw = new StreamWriter(sFilePath, false, Encoding.GetEncoding("Shift_JIS")))
            {
                //----------------------------------------------------------//
                // DataColumnの型から値を出力するかどうか判別します         //
                // 出力対象外となった項目は[データ]という形で出力します     //
                //----------------------------------------------------------//
                for (int i = 0; i < dtDt.Columns.Count; i++)
                {
                    switch (dtDt.Columns[i].DataType.ToString())
                    {
                        case "System.Boolean":
                        case "System.Byte":
                        case "System.Char":
                        case "System.DateTime":
                        case "System.Decimal":
                        case "System.Double":
                        case "System.Int16":
                        case "System.Int32":
                        case "System.Int64":
                        case "System.SByte":
                        case "System.Single":
                        case "System.String":
                        case "System.TimeSpan":
                        case "System.UInt16":
                        case "System.UInt32":
                        case "System.UInt64":
                            break;

                        default:
                            filterIndex.Add(i);
                            break;
                    }
                }
                //----------------------------------------------------------//
                // ヘッダーを出力します。                                   //
                //----------------------------------------------------------//
                if (boolHeader)
                {
                    foreach (DataColumn col in dtDt.Columns)
                    {
                        sw.Write(sp + "\"" + col.ToString().Replace("\"", "\"\"") + "\"");
                        sp = ",";
                    }
                    sw.WriteLine();
                }
                //----------------------------------------------------------//
                // 内容を出力します。                                       //
                //----------------------------------------------------------//
                foreach (DataRow row in dtDt.Rows)
                {
                    sp = string.Empty;
                    for (int i = 0; i < dtDt.Columns.Count; i++)
                    {
                        if (filterIndex.Contains(i))
                        {
                            sw.Write(sp + "\"[データ]\"");
                            sp = ",";
                        }
                        else
                        {
                            sw.Write(sp + "\"" + row[i].ToString().Replace("\"", "\"\"") + "\"");
                            sp = ",";
                        }
                    }
                    sw.WriteLine();
                }
            }
        }


        /// <summary>
        /// CSVファイルからDataTableを返す
        /// </summary>
        /// <param name="strFilePath">CSVファイルパス</param>
        /// <param name="isInHeader">1行目はヘッダー扱いとするか</param>
        /// <returns></returns>
        public static DataTable ConvertDataTableFromCsv(String strFilePath, Boolean isInHeader = true)
        {
            //StaticConvertDataTableToCsvは、改行が無視される。
            DataTable dt = new DataTable();
            String strInHeader = isInHeader ? "YES" : "NO";                // ヘッダー設定
            String strCon = "Provider=Microsoft.ACE.OLEDB.12.0;"      // プロバイダ設定
            //String strCon = "Provider=Microsoft.Jet.OLEDB.4.0;"     // Jetでやる場合
                                + "Data Source=" + Path.GetDirectoryName(strFilePath) + "\\; "          // ソースファイル指定
                                + "Extended Properties=\"Text;HDR=" + strInHeader + ";FMT=Delimited\"";
            OleDbConnection con = new OleDbConnection(strCon);
            String strCmd = "SELECT * FROM [" + Path.GetFileName(strFilePath) + "]";

            // 読み込み
            OleDbCommand cmd = new OleDbCommand(strCmd, con);
            OleDbDataAdapter adp = new OleDbDataAdapter(cmd);
            adp.Fill(dt);

            return dt;
        }

        public static string Left(string sStr, int iLen)
        {
            if (iLen < 0)
            {
                throw new ArgumentException("引数'len'は0以上でなければなりません。");
            }
            if (sStr == null)
            {
                return "";
            }
            if (sStr.Length <= iLen)
            {
                return sStr;
            }
            return sStr.Substring(0, iLen);
        }

        public static string Right(string sStr, int iLen)
        {
            if (iLen < 0)
            {
                throw new ArgumentException("引数'len'は0以上でなければなりません。");
            }
            if (sStr == null)
            {
                return "";
            }
            if (sStr.Length <= iLen)
            {
                return sStr;
            }
            return sStr.Substring(sStr.Length - iLen, iLen);
        }

        public static string Mid(string sStr, int iStart, int iLen)
        {
            if (iStart <= 0)
            {
                throw new ArgumentException("引数'start'は1以上でなければなりません。");
            }
            if (iLen < 0)
            {
                throw new ArgumentException("引数'len'は0以上でなければなりません。");
            }
            if (sStr == null || sStr.Length < iStart)
            {
                return "";
            }
            if (sStr.Length < (iStart + iLen))
            {
                return sStr.Substring(iStart - 1);
            }
            return sStr.Substring(iStart - 1, iLen);
        }



        public static string ListToString(List<string> list, string replaceString)
        {
            var str = "";
            for (var i = 0; i < list.Count(); i++)
            {
                if (i == 0)
                {
                    str = $"{list[i]}";
                }
                else
                {
                    str = $"{str}{replaceString}{list[i]}";
                }
            }
            return str;
        }

        public static List<string> StringToList(string str, string splitString)
        {
            string[] del = { splitString };
            var stringArray = str.Split(del, StringSplitOptions.None);
            List<string> stringList = new List<string>();
            stringList.AddRange(stringArray);
            return stringList;
        }

        public static void OutputHtml(string sFilename, string sStr)
        {
            Encoding sjisEnc = Encoding.GetEncoding("Shift_JIS");
            StreamWriter writer = new StreamWriter(sFilename + ".html", true, sjisEnc);
            writer.WriteLine(sStr);
            writer.Close();
        }

        public static DataTable GetDataTableFromExcelOfPureTable(String strFilePath, String strSheetName, Boolean isInHeader = true, Boolean isAllStrColum = true)
        {
            //try
            //{
            DataTable dt = new DataTable();
            String strInHeader = isInHeader ? "YES" : "NO";        // ヘッダー設定
            String strIMEX = isAllStrColum ? "IMEX=1;" : "";   // 文字列型設定
            String strFileEx = Path.GetExtension(strFilePath);   // ファイル拡張子
            String strExcelVer = "Excel ";                         // Excelファイルver確認

            if (strFileEx == ".xls")
            {
                strExcelVer += "8.0;";
            }
            else if (strFileEx == ".xlsx" || strFileEx == ".xlsm")
            {
                strExcelVer += "12.0;";
            }
            else
            {
                return null;
            }

            String strCon = "Provider=Microsoft.ACE.OLEDB.12.0;"      // プロバイダ設定
                                                                      //= "Provider=Microsoft.Jet.OLEDB.4.0;"     // Jetでやる場合（後で検証 xlsxでも使えるのか？）
                                + "Data Source=" + strFilePath + "; "       // ソースファイル指定
                                + "Extended Properties=\"" + strExcelVer    // Excelファイルver指定
                                + "HDR=" + strInHeader + ";"                // ヘッダー設定
                                + strIMEX                                   // フィールドの型を強制的にテキスト
                                + "\"";
            OleDbConnection con = new OleDbConnection(strCon);
            String strCmd = "SELECT * FROM [" + strSheetName + "$]";

            // 読み込み
            OleDbCommand cmd = new OleDbCommand(strCmd, con);
            OleDbDataAdapter adp = new OleDbDataAdapter(cmd);
            adp.Fill(dt);

            return dt;
            //}
            //catch
            //{
            //    MessageBox.Show("ERROR : ファイル名「item.xlsx」が見つからないか、シート名「Main」が見つかりません。");
            //    return null;
            //}
        }


        public static List<string> StaticGetFindcodeList一行(string sInputOriginalText, string sRegexPattern)
        {
            List<string> list = new List<string>();
            var reg = new Regex(sRegexPattern, RegexOptions.IgnoreCase | RegexOptions.None);
            for (Match m = reg.Match(sInputOriginalText); m.Success; m = m.NextMatch())
            {
                list.Add(m.Groups["findcode"].Value);
            }
            if (list.Count > 0)
            {
                return list;
            }
            else
            {
                return null;
            }
        }



        public static string 文字数を制限(string srcString, string splitCharactor, int maxByte)
        {
            var tmpStrList = new List<string>();
            foreach (var str in StringToList(srcString, splitCharactor))
            {
                var totalStr = $"{ListToString(tmpStrList, splitCharactor)} {str}";
                var totalByte = Encoding.GetEncoding("Shift_JIS").GetByteCount(totalStr);
                var maxLength = maxByte / 2;
                Debug.WriteLine($"maxByte:{maxByte}, maxLength:{maxLength}, {totalByte}byte {totalStr}");
                if (totalByte < maxByte && totalStr.Length < maxLength)
                {
                    tmpStrList.Add(str);
                }
                else
                {
                    Debug.WriteLine($"break : maxByte:{maxByte}, maxLength:{maxLength}, {totalByte}byte {totalStr}");
                    break;
                }
            }
            return ListToString(tmpStrList, splitCharactor);
        }














    }
}




