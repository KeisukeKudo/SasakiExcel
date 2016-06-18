using System;
using System.Drawing;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Diagnostics;

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
                Console.WriteLine("{0}が存在しません", ImageFileName);
                Console.ReadLine();
                return;
            }

            try {
                //Excelがインストールされていない場合は､以降の処理を続行するか聞く
                if (!IsExcelInstalled() && !IsProceed()) {
                    Console.WriteLine("処理を中断しました");
                    return;
                }

                var s = new Stopwatch();
                s.Start();
                Console.WriteLine("処理中です.........\r\n");
                ImageExcelToCopy();
                s.Stop();
                Console.WriteLine("正常終了しました({0:N0}ms)", s.ElapsedMilliseconds);

            } catch (Exception e) {
                Console.WriteLine("異常終了しました");
                Console.WriteLine(e.Message);

            } finally {
                Console.WriteLine("何かキーを入力してください");
                Console.ReadLine();

            }
        }

        /// <summary>
        /// 対象画像のピクセルから色を取得し､Excelシートの各セルに設定､保存
        /// </summary>
        static void ImageExcelToCopy() {
            if (File.Exists(ExcelFileName)) {
                File.Delete(ExcelFileName);
            }

            //Excelファイル作成
            var outputFile = new FileInfo(ExcelFileName);

            using (var bitmap = new Bitmap(ImageFileName))
            using (var book = new ExcelPackage(outputFile))
            using (var sheet = book.Workbook.Worksheets.Add(SheetName)) {
                //回転情報が無いのに回転してしまうので正しい位置に調整
                //時計回りに90度回転し､水平方向に反転
                bitmap.RotateFlip(RotateFlipType.Rotate90FlipX);
                for (var y = 1; y <= bitmap.Height; y++) {
                    //高さ指定
                    sheet.Row(y).Height = 3.8;
                    for (var x = 1; x <= bitmap.Width; x++) {
                        //セルの背景色に設定
                        var cell = sheet.Cells[x, y];
                        //ピクセルの色を取得
                        var color = bitmap.GetPixel(x - 1, y - 1);
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
        /// 実行環境にExcelがインストールされているか
        /// </summary>
        /// <returns></returns>
        static bool IsExcelInstalled() {
            return (Type.GetTypeFromProgID("Excel.Application") != null);
        }

        /// <summary>
        /// Excelがインストールされていない場合､処理を続行するか選択
        /// </summary>
        /// <returns></returns>
        static bool IsProceed() {
            Console.WriteLine("Excelがインストールされていません");
            Console.WriteLine("出力ファイルを閲覧できませんが､処理を続行しますか? Y/N");
            var response = Console.ReadLine().ToUpper();
            while (true) {
                switch (response) {
                    case "Y":
                        return true;
                    case "N":
                        return false;
                }
                //上記以外が入力された場合は再度入力を求める
                response = Console.ReadLine().ToUpper();
            }
        }
    }
}
