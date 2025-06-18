using System;
using System.Windows.Forms;

namespace csharp
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            var csvGenerater = new CsvGenerator();
            csvGenerater.CsvGenerate();
            csvGenerater.CsvCopyYahooItem2号店();
            csvGenerater.CsvCopyYahooOption2号店();
            MessageBox.Show("完了しました");
            csvGenerater.完了();
        }

    }

}