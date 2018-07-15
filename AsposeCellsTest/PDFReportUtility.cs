using Aspose.Cells;
using Aspose.Pdf;
using Aspose.Pdf.Facades;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;

namespace AsposeCellsTest
{
    public class PDFReportUtility
    {

        // 增加功能
        // 1.excel header 及 footer 及 可設定欄寬
        // 2.pdf watermark 整頁或是充滿

        /// <summary>
        /// DataTable 轉到 Excel 的參數
        /// </summary>
        public struct ExportDataTable2ExcelArg
        {
            public DataTable dataSource;
            //columnId, columnDisplayName, columnWidth
            public Dictionary<string, Tuple<string, double>> ColumnInfos;
            //是否允許換行
            public bool IsTextWrapped;
            //字型名稱，如果不允許換行的話，autoFitColumns 會依字型去算寬度
            public string FontName;
            //直/橫 印
            public PageOrientationType PageOrientation;
            //縮放比例
            public int PageScale; //10~400

            //表頭 左、中、右
            public string HeaderLeft;
            public string HeaderCenter;
            public string HeaderRight;
            //表尾 左、中、右
            public string FooterLeft;
            public string FooterCenter;
            public string FooterRight;

   
        }

        /// <summary>
        /// 浮水印設定參數
        /// </summary>
        public struct WatermarkArg
        {
            //浮水印字串
            public string Watermark;
            //浮水印 Stamp 的 Height
            public double WatermarkHeight;
            //浮水印 Stamp 的 Width
            public double WatermarkWidth;
            //浮水印 Stamp 的 水平間隔
            public double WatermarkHorizontalSpace;
            //浮水印 Stamp 的 垂直間隔
            public double WatermarkVerticalSpace;
            //浮水印 Stamp 貼上Style，目前有蓋滿一頁及水平蓋滿
            public WatermarkStyle WMStyle;
            //不透明度 0~ `
            public double Opacity;
            //旋轉角度
            public double RotateAngle;
        }

        public enum WatermarkStyle
        {
            //蓋滿一頁
            FitPage,
            //水平蓋滿
            RepeatHorizontal
        }

         

        /// <summary>
        /// 設定 DataTable 的 Column Name
        /// </summary>
        /// <param name="dt">datatable</param>
        /// <param name="columnInfos">item1:Column Name, item2:Column Width</param>
        private static void ChangedDataTableColumnName(DataTable dt, Dictionary<string, Tuple<string, double>> columnInfos)
        {
            //change columnName 
            if (columnInfos != null)
            {
                foreach (KeyValuePair<string, Tuple<string, double>> columnInfo in columnInfos)
                {
                    if (dt.Columns[columnInfo.Key] != null)
                        dt.Columns[columnInfo.Key].ColumnName = columnInfo.Value.Item1;
                }
            }
        }

        /// <summary>
        /// 設定 Worksheet 的 欄位寬度
        /// </summary>
        /// <param name="sheet">綁定 DataTable 的 Worksheet </param>
        /// <param name="columnInfos">item1:Column Name, item2:Column Width</param>
        private static void ChangedSheetColumnStyle(Worksheet sheet, Dictionary<string, Tuple<string, double>> columnInfos)
        {
            //change columnName 
            if (columnInfos != null)
            {
                var columnIndex = 0;
                foreach (KeyValuePair<string, Tuple<string, double>> columnInfo in columnInfos)
                {
                    if (columnInfo.Value.Item2 > -1)
                    {
                        sheet.Cells.SetColumnWidth(columnIndex, columnInfo.Value.Item2);
                    }
                    columnIndex++;
                }
            }
        }

        /// <summary>
        /// 將 DataTable 的資料轉到Excel處理並存成 PDF Stream
        /// </summary>
        /// <param name="arg"></param>
        /// <returns></returns>
        public static MemoryStream GenPDFFromDataTable(ExportDataTable2ExcelArg arg)
        {
            //Change DataTable's ColumnName
            ChangedDataTableColumnName(arg.dataSource, arg.ColumnInfos);

            //proc excel
            // Instantiating a Workbook object            
            var workbook = new Workbook();
            if (!string.IsNullOrWhiteSpace(arg.FontName))
            {
                var wbStyle = workbook.DefaultStyle;
                wbStyle.Font.Name = arg.FontName;
                workbook.DefaultStyle = wbStyle;
            }
            var worksheet = workbook.Worksheets[0];
            worksheet.Cells.ImportDataTable(arg.dataSource, true, "A1");


            
            //https://docs.aspose.com/display/cellsnet/Setting+Page+Options
            var pageSetup = workbook.Worksheets[0].PageSetup;
            pageSetup.PrintTitleRows = "$1:$1";
            pageSetup.IsPercentScale = true;
            pageSetup.Orientation = arg.PageOrientation;
            pageSetup.Zoom = arg.PageScale < 10 ? 100 : arg.PageScale;
            //https://docs.aspose.com/display/cellsnet/Setting+Headers+and+Footers
            if (!string.IsNullOrEmpty(arg.HeaderLeft))
            {
                pageSetup.SetHeader(0, arg.HeaderLeft);
            }
            if (!string.IsNullOrEmpty(arg.HeaderCenter))
            {
                pageSetup.SetHeader(1, arg.HeaderCenter);
            }
            if (!string.IsNullOrEmpty(arg.HeaderRight))
            {
                pageSetup.SetHeader(2, arg.HeaderRight);
            }

            if (!string.IsNullOrEmpty(arg.FooterLeft))
            {
                pageSetup.SetFooter(0, arg.FooterLeft);
            }
            if (!string.IsNullOrEmpty(arg.FooterCenter))
            {
                pageSetup.SetFooter(1, arg.FooterCenter);
            }
            if (!string.IsNullOrEmpty(arg.FooterRight))
            {
                pageSetup.SetFooter(2, arg.FooterRight);
            }

            var range = worksheet.Cells.MaxDisplayRange;
            //border
            //Setting border for each cell in the range
            var style = workbook.CreateStyle();
            var colorBlack = System.Drawing.Color.Black;
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Medium, colorBlack);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Medium, colorBlack);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Medium, colorBlack);
            style.SetBorder(BorderType.TopBorder, CellBorderType.Medium, colorBlack);
            style.IsTextWrapped = arg.IsTextWrapped;
            range.SetStyle(style);
            worksheet.AutoFitColumns();
            //adjust columns
            ChangedSheetColumnStyle(worksheet, arg.ColumnInfos);
            worksheet.AutoFitRows();
            //string xlsFile = Path.Combine(HttpContext.Current.Server.MapPath("./data"), $"test.xlsx");
            //workbook.Save(xlsFile, Aspose.Cells.SaveFormat.Xlsx);
            //save to stream
            var pdfStream = new MemoryStream();
            workbook.Save(pdfStream, Aspose.Cells.SaveFormat.Pdf);
            
            return pdfStream;
        }

        /// <summary>
        /// 將 PDF 加上 浮水印
        /// </summary>
        /// <param name="pdfStream"></param>
        /// <param name="arg"></param>
        /// <returns></returns>
        public static MemoryStream AddWatermark(MemoryStream pdfStream, WatermarkArg arg)
        {

            var pdfDocument = new Aspose.Pdf.Document(pdfStream);
            if (!string.IsNullOrWhiteSpace(arg.Watermark))
            {
                var text = new FormattedText(arg.Watermark);
                foreach (var page in pdfDocument.Pages)
                {
                    switch (arg.WMStyle)
                    {
                        case WatermarkStyle.FitPage:
                            AddWatermarkFitPage(page, arg);
                            break;
                        case WatermarkStyle.RepeatHorizontal:
                            AddWatermarkRepeatHorizontal(page, arg);
                            break;

                        default:
                            break;
                    }
                }
            }
            var newPdfStream = new MemoryStream();
            pdfDocument.Save(newPdfStream);
            return newPdfStream;
        }

        /// <summary>
        /// 浮水印跟頁面一樣大
        /// </summary>
        /// <param name="pdfPage"></param>
        /// <param name="arg"></param>
        private static void AddWatermarkFitPage(Aspose.Pdf.Page pdfPage, WatermarkArg arg)
        {
            var text = new FormattedText(arg.Watermark);
            var stamp = new TextStamp(text);
            stamp.RotateAngle = arg.RotateAngle;
            stamp.XIndent = arg.WatermarkHorizontalSpace;
            stamp.YIndent = arg.WatermarkVerticalSpace;
            stamp.Opacity = arg.Opacity;
            stamp.Width = pdfPage.CropBox.Width;
            stamp.Height = pdfPage.CropBox.Height;
            pdfPage.AddStamp(stamp);
        }

        //最小的 浮水印 長、寬
        const double minValue = 30;

        /// <summary>
        /// 依 浮水印 水平地蓋滿整個頁面
        /// </summary>
        /// <param name="pdfPage"></param>
        /// <param name="arg"></param>
        private static void AddWatermarkRepeatHorizontal(Aspose.Pdf.Page pdfPage, WatermarkArg arg)
        {

            if (arg.WatermarkHeight < minValue)
                throw new ArgumentException($"{nameof(arg.WatermarkHeight)} must greater than {minValue}");
            if (arg.WatermarkWidth < minValue)
                throw new ArgumentException($"{nameof(arg.WatermarkWidth)} must greater than {minValue}");

            var text = new FormattedText(arg.Watermark);
            var yIndent = pdfPage.CropBox.Height - arg.WatermarkHeight;
            var yLimit = 0 - (arg.WatermarkHeight + arg.WatermarkVerticalSpace);
            var pageWidth = pdfPage.CropBox.Width;
            var xIndent = 0d;
            while (yIndent > yLimit)
            {
                while (xIndent < pageWidth)
                {
                    var stamp = new TextStamp(text);
                    stamp.RotateAngle = arg.RotateAngle;
                    stamp.XIndent = xIndent;
                    stamp.YIndent = yIndent;
                    stamp.Opacity = arg.Opacity;
                    stamp.Width = arg.WatermarkWidth;
                    stamp.Height = arg.WatermarkHeight;
                    pdfPage.AddStamp(stamp);
                    xIndent += (arg.WatermarkWidth + arg.WatermarkHorizontalSpace);
                }
                xIndent = 0;
                var yIdentReduce = (arg.WatermarkHeight + arg.WatermarkVerticalSpace);

                yIndent -= yIdentReduce;
            }

        }

        /// <summary>
        /// 以角度線性期function來呈現
        /// </summary>
        /// <param name="pdfPage"></param>
        /// <param name="arg"></param>
        private static void AddWatermarkRepeatRotateAngle(Aspose.Pdf.Page pdfPage, WatermarkArg arg)
        {

            if (arg.WatermarkHeight < minValue)
                throw new ArgumentException($"{nameof(arg.WatermarkHeight)} must greater than {minValue}");
            if (arg.WatermarkWidth < minValue)
                throw new ArgumentException($"{nameof(arg.WatermarkWidth)} must greater than {minValue}");

            var text = new FormattedText(arg.Watermark);
            var yIndent = pdfPage.CropBox.Height - arg.WatermarkHeight;
            var yLimit = 0 - (arg.WatermarkHeight + arg.WatermarkVerticalSpace);
            var pageWidth = pdfPage.CropBox.Width;
            var pageHeight = pdfPage.CropBox.Height;
            var xIndent = 0d;
            while (yIndent > yLimit)
            {
                var y = yIndent;
                while (xIndent < pageWidth && y < pageHeight)
                {
                    var stamp = new TextStamp(text);
                    stamp.RotateAngle = arg.RotateAngle;
                    stamp.XIndent = xIndent;
                    stamp.YIndent = y;
                    stamp.Opacity = arg.Opacity;
                    stamp.Width = arg.WatermarkWidth;
                    stamp.Height = arg.WatermarkHeight;
                    pdfPage.AddStamp(stamp);
                    xIndent += (arg.WatermarkWidth + arg.WatermarkHorizontalSpace);
                    xIndent = xIndent + Math.Cos(30) * arg.WatermarkWidth;
                    y = y + Math.Sign(30) * (arg.WatermarkHeight + arg.WatermarkVerticalSpace);
                }
                xIndent = 0;
                var yIdentReduce = (arg.WatermarkHeight + arg.WatermarkVerticalSpace);
                yIndent -= yIdentReduce;
            }

            //到底了，要再連走 X
            var baseX = 0d;
            while (baseX < pageWidth)
            {
                var y = yIndent;
                xIndent = baseX;
                while (xIndent < pageWidth)
                {
                    var stamp = new TextStamp(text);
                    stamp.RotateAngle = arg.RotateAngle;
                    stamp.XIndent = xIndent;
                    stamp.YIndent = y;
                    stamp.Opacity = arg.Opacity;
                    stamp.Width = arg.WatermarkWidth;
                    stamp.Height = arg.WatermarkHeight;
                    pdfPage.AddStamp(stamp);
                    xIndent += (arg.WatermarkWidth + arg.WatermarkHorizontalSpace);
                    xIndent = xIndent + Math.Cos(30) * arg.WatermarkWidth;
                    y = y + Math.Sign(30) * (arg.WatermarkHeight + arg.WatermarkVerticalSpace);
                }
                baseX += (arg.WatermarkWidth + arg.WatermarkHorizontalSpace); ;
            }
        }


    }

}