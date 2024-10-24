using Spire.Xls;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;

namespace ExcelConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            Program p = new Program();
            PathData pathData = new PathData();

            //CSVファイルパス読み込み
            Console.WriteLine("ファイルパスを指定してください");
            pathData.CsvPath = Console.ReadLine() ?? " ";

            Console.WriteLine("ファイル名を指定してください拡張子xlsxまで");
            pathData.FileName = Console.ReadLine() ?? " ";

            //Workbookクラスインスタンス生成
            Workbook workBook = new Workbook();

            //最新のワークシートを取得
            Worksheet sheet = p.FileRead(pathData.CsvPath,workBook).Worksheets[0];

            p.ExcelConvert(sheet,workBook,pathData.FileName);

            Console.WriteLine("保存されたファイルを指定してください");

            pathData.ExcelPath = Console.ReadLine() ?? " ";

            XLWorkbook wk = new XLWorkbook(pathData.ExcelPath);

            IXLWorksheet ixlWorkSheet = wk.Worksheet(1);

            int lastRow = ixlWorkSheet.LastRowUsed().RowNumber();

            p.CompensateData(lastRow, ixlWorkSheet);

            p.MovingAverage(lastRow, ixlWorkSheet);

            try
            {
                //データ保存
                wk.SaveAs(pathData.ExcelPath);
            }
            catch(Exception e)
            {
                Console.WriteLine(e + "正しくファイルを保存できませんでした");
                wk.Save();
            }
        }

        Workbook FileRead(string csvPath,Workbook book)
        {
            try
            {
                //CSVファイルをロード
                book.LoadFromFile(@csvPath, ",", 1, 1);
            }
            catch(Exception e)
            {
                Console.WriteLine(e.Message+"正しくファイルが読み込まれませんでした。");
                book.Save();
            }
            return book;
        }
        void ExcelConvert(Worksheet sheet,Workbook book,string fileName)
        {
            //ワークシート使用範囲にアクセス
            CellRange usedRange = sheet.AllocatedRange;

            //数値をテキストとして保存するときにエラーを無視
            usedRange.IgnoreErrorOptions = IgnoreErrorType.NumberAsText;

            //行の高さと列の幅を自動調整
            usedRange.AutoFitColumns();
            usedRange.AutoFitRows();

            try
            {
                book.SaveToFile(fileName, ExcelVersion.Version2016);
            }
            catch (Exception e)
            {
                Console.WriteLine(e + "正しくエクセルファイルに変換できませんでした");
                book.Save();
            }
        }
        void CompensateData(int lastRow, IXLWorksheet sheet)
        {
            //文字列から数値に変更
            for (int i = 2; i <= lastRow; i++)
            {
                string s = sheet.Cell(i, 2).Value.ToString();
                double d;
                if (double.TryParse(s, out d))
                {
                    sheet.Cell(i, 2).Value = d;
                }
                //データが入ってない場合0を代入
                else
                {
                    sheet.Cell(i, 2).Value = 0;
                }
            }
            //最初のセルが0の場合後ろの数値を取ってくる
            string firstCell = sheet.Cell(2, 2).Value.ToString();
            if (firstCell == "0")
            {
                for (int i = 3; i <= lastRow; i++)
                {
                    string s = sheet.Cell(i, 2).Value.ToString();
                    if (s != "0")
                    {
                        sheet.Cell(2, 2).Value = sheet.Cell(i, 2).Value;
                        break;
                    }
                }
            }
            //値が0の場合前の数値を取ってくる
            for (int i = 2; i <= lastRow; i++)
            {
                string s = sheet.Cell(i, 2).Value.ToString();
                if (s == "0" && i != 2)
                {
                    sheet.Cell(i, 2).Value = sheet.Cell(i - 1, 2).Value;
                }
            }
            //タイムスタンプを記入
            int timeStamp = 0;
            for (int i = 2; i <= lastRow; i++)
            {
                sheet.Cell(i, 3).Value = timeStamp;
                timeStamp += 2; //設定に応じて
            }
        }
        void MovingAverage(int lastRow,IXLWorksheet sheet)
        {
            List<double> cells = new List<double>();
            double avg = 0;
            int avgCell = 21;

            //移動平均を出した数値をエクセルに記入
            for (int i = 2; i <= lastRow; i++)
            {
                string s = sheet.Cell(i, 2).Value.ToString();
                cells.Add(double.Parse(s));
                if (cells.Count == 20)
                {
                    avg = cells.Average();
                    cells.RemoveAt(0);
                    sheet.Cell(avgCell, 4).Value = avg;
                    avgCell++;
                }
            }
        }
    }

    public class PathData
    {
        public string CsvPath { get; set; } = " ";

        public string FileName { get; set; } = " ";

        public string ExcelPath { get; set; } = " ";
    }
}