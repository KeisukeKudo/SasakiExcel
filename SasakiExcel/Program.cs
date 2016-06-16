using System.Drawing;
using System.IO;
using ClosedXML.Excel;

namespace SasakiExcel {
    class Program {

        /// <summary>
        /// 読み込む画像ファイル名
        /// </summary>
        private const string ImageFileName = "sasaki.jpg";

        /// <summary>
        /// 保存Excelファイル名
        /// </summary>
        private const string SaveFileName = "sasaki_nozomi.xlsx";

        /// <summary>
        /// Excelシート名
        /// </summary>
        private const string SheetName = "Nozomi Sheet";


        static void Main(string[] args) {
            //あったら消す
            if (File.Exists(SaveFileName)) {
                File.Delete(SaveFileName);
            }

            //画像をBitmapで取得
            //Excelファイルを新規作成､開く
            using (var bitmap = new Bitmap(ImageFileName))
            using (var book = new ClosedXML.Excel.XLWorkbook()) {
                //シートを作成
                var sheet = book.Worksheets.Add(SheetName);

                for (var y = 1; y <= bitmap.Height; y++) {
                    for (var x = 1; x <= bitmap.Width; x++) {
                        //ピクセルの色を取得
                        var color = bitmap.GetPixel(x - 1, y - 1);
                        //セルの背景色に設定
                        sheet.Cell(x, y).Style.Fill.SetBackgroundColor(XLColor.FromColor(color));
                    }
                }
                //高さ指定
                sheet.Rows(1, bitmap.Height).Height = 3.8;
                //横幅指定(効いてない?)
                sheet.Columns(1, bitmap.Width).Width = 0.1;
                //ルートに保存
                book.SaveAs(SaveFileName);
            }
        }
    }
}
