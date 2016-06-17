using System;
using System.Drawing;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace SasakiExcel {
    class Program {

        /// <summary>
        /// 読み込む画像ファイル名
        /// </summary>
        private const string ImageFileName = "sasaki.jpg";

        /// <summary>
        /// 保存Excelファイル名
        /// </summary>
        private const string ExcelFileName = "sasaki_nozomi.xlsx";

        /// <summary>
        /// Excelシート名
        /// </summary>
        private const string SheetName = "Nozomi Sheet";


        static void Main(string[] args) {

            if (!File.Exists(ImageFileName)) {
                Console.WriteLine(" {0}が存在しません", ImageFileName);
                Console.ReadLine();
                return;
            }
            
            try {
                Console.WriteLine(" 処理中です.........\r\n");
                ImageExcelToCopy();
                Console.WriteLine(" 正常終了しました");

            } catch (Exception e) {
                Console.WriteLine(" 異常終了しました");
                Console.WriteLine(e.Message);

            } finally {
                Console.WriteLine(" 何かキーを入力してください");
                Console.ReadLine();

            }
        }

        /// <summary>
        /// 対象画像のピクセルから色を取得し､Excelシートの各セルに設定､保存
        /// </summary>
        static void ImageExcelToCopy() {
            DeleteFile();

            //Excelファイル作成
            var outputFile = new FileInfo(ExcelFileName);

            //画像をBitmapで取得
            //Excelファイルを開く
            using (var bitmap = new Bitmap(ImageFileName))
            using (var book = new ExcelPackage(outputFile)) {
                //回転情報が無いのに回転してしまうので正しい位置に調整
                //時計回りに90度回転し､水平方向に反転
                bitmap.RotateFlip(RotateFlipType.Rotate90FlipX);

                //シートを作成
                var sheet = book.Workbook.Worksheets.Add(SheetName);
                for (var y = 1; y <= bitmap.Height; y++) {
                    //高さ指定
                    sheet.Row(y).Height = 3.8;
                    for (var x = 1; x <= bitmap.Width; x++) {
                        //ピクセルの色を取得
                        var color = bitmap.GetPixel(x - 1, y - 1);
                        //セルの背景色に設定
                        var cell = sheet.Cells[x, y];
                        SetBackgroundColor(cell, color);

                        //横幅指定
                        sheet.Column(x).Width = 0.7;
                    }
                }
                //ズームレベルの指定(25%)
                sheet.View.ZoomScale = 25;
                book.Save();
            }
        }

        /// <summary>
        /// セルの背景色を設定
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="color"></param>
        static void SetBackgroundColor(ExcelRange cell, Color color) {
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(color);
        }

        /// <summary>
        /// 既にファイルが存在している場合は削除
        /// </summary>
        static void DeleteFile() {
            if (File.Exists(ExcelFileName)) {
                File.Delete(ExcelFileName);
            }
        }

    }
}
