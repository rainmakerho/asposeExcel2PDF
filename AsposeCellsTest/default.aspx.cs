using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Pdf;
using Aspose.Pdf.Facades;
using Aspose.Pdf.Text;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using static AsposeCellsTest.PDFReportUtility;

namespace AsposeCellsTest
{
    public partial class _default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if(Page.IsPostBack == false)
            {
                BindGridView();
                
            }
        }

        /// <summary>
        /// 輸入 DataTable 轉成 有浮水印 的 PDF
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="columnNameMappings"></param>
        /// <param name="folderName"></param>
        /// <param name="watermark"></param>
        /// <param name="pot"></param>
        /// <returns>產生的 pdf 檔名 (fullpath) </returns>
        private static string GenPDF(DataTable dt, Dictionary<string, string> columnNameMappings
            , string folderName, string watermark, PageOrientationType pot)
        {
            ChangeColumnDisplayName(dt, columnNameMappings);

            //output file name
            var fileNameWithoutExt = $"{Guid.NewGuid().ToString("N")}";
            var outputExcel = Path.Combine(folderName, $"{fileNameWithoutExt}_tmp.xlsx");
            var tempPdf = Path.Combine(folderName, $"{fileNameWithoutExt}_tmp.pdf");
            var outputPdf = Path.Combine(folderName, $"{fileNameWithoutExt}.pdf");

            //proc excel
            // Instantiating a Workbook object            
            var workbook = new Workbook();
            var worksheet = workbook.Worksheets[0];
            worksheet.Cells.ImportDataTable(dt, true, "A1");
            worksheet.AutoFitColumns();
            worksheet.AutoFitRows();

            var range = worksheet.Cells.MaxDisplayRange;
            var pageSetup = workbook.Worksheets[0].PageSetup;
            //var titleEndColumnName = CellsHelper.ColumnIndexToName(range.ColumnCount-1);
            //pageSetup.PrintTitleColumns = $"$A:${titleEndColumnName}";
            pageSetup.PrintTitleRows = "$1:$1";
            pageSetup.IsPercentScale = true;
            pageSetup.Orientation = pot;

            //border
            //Setting border for each cell in the range
            var style = workbook.CreateStyle();
            var colorBlack = System.Drawing.Color.Black;
            style.SetBorder(BorderType.BottomBorder, CellBorderType.Medium, colorBlack);
            style.SetBorder(BorderType.LeftBorder, CellBorderType.Medium, colorBlack);
            style.SetBorder(BorderType.RightBorder, CellBorderType.Medium, colorBlack);
            style.SetBorder(BorderType.TopBorder, CellBorderType.Medium, colorBlack);
            range.SetStyle(style);

            // Saving the Excel file
            //workbook.Save(outputExcel);
            //workbook.Save(tempPdf);
            //save to stream
            var pdfStream = new MemoryStream();
            workbook.Save(pdfStream, Aspose.Cells.SaveFormat.Pdf);

            var pdfDocument = new Aspose.Pdf.Document(pdfStream);
            if (!string.IsNullOrWhiteSpace(watermark))
            {
                //針對 PDF 加入 Watermark
                Aspose.Pdf.Facades.Stamp aStamp = new Aspose.Pdf.Facades.Stamp();
                aStamp.Rotation = 45;
                var textStamp = new TextStamp(watermark);
                //set properties of the stamp
                // textStamp.Background = true;
                textStamp.Opacity = 0.2;
                textStamp.TextState.FontSize = 60.0F;
                textStamp.HorizontalAlignment = HorizontalAlignment.Center;
                textStamp.VerticalAlignment = VerticalAlignment.Center;
                textStamp.RotateAngle = aStamp.Rotation;
                textStamp.TextState.Font = FontRepository.FindFont("Arial");
                textStamp.TextState.ForegroundColor = Aspose.Pdf.Color.Gray;
                foreach (var page in pdfDocument.Pages)
                {
                    page.AddStamp(textStamp);
                }
            }

            pdfDocument.Save(outputPdf);
            return outputPdf;
        }

        /// <summary>
        /// 將目前DataTable的資料依 Dictionary 來改它的 ColumnName
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="columnNameMappings"></param>
        private static void ChangeColumnDisplayName(DataTable dt, Dictionary<string, string> columnNameMappings)
        {
            //change columnName 
            if (columnNameMappings != null)
            {
                foreach (KeyValuePair<string, string> columnMapping in columnNameMappings)
                {
                    dt.Columns[columnMapping.Key].ColumnName = columnMapping.Value;
                }
            }
        }

        private DataTable GetDataSource()
        {
            // Instantiating a "Products" DataTable object
            var dataTable = new DataTable("Products");
            // Adding columns to the DataTable object
            dataTable.Columns.Add("ProductID", typeof(Int32));
            dataTable.Columns.Add("ProductName", typeof(string));
            dataTable.Columns.Add("ProductDesc", typeof(string));
            dataTable.Columns.Add("Units", typeof(Double));
            var rand = new Random();
            for (var i = 0; i < 90; i++)
            {
                dataTable.Rows.Add(i, $"產品名稱-{i} 讓產品名稱自己說話 abc {i} ...{i}", $"產品描述 -{i}塑造出帶有「情感」的品牌概念 -{i}產品描述 -產品描述 -產品描述 -", rand.NextDouble());
            }
            return dataTable;
        }

        protected void btnGenPDF_Click(object sender, EventArgs e)
        {
            var ds = GetDataSource();
            var columnMapping = new Dictionary<string, string>
            {
                {"ProductID", "產品代號" },
                {"ProductName", "產品名稱" },
                {"ProductDesc", "產品 描述" },
                 {"Units", "產品 庫存" }
            };
            var outFileName = GenPDF(ds, columnMapping, 
                Server.MapPath("./data"), "你好，我是亂馬客!!!"
                , PageOrientationType.Landscape);
            Response.Write($"<hr>Export to Pdf-1:{outFileName}");

            var ds2 = GetDataSource();
            var outFileName2 = GenPDF(ds2, null,
                Server.MapPath("./data"), ""
                , PageOrientationType.Portrait);
            Response.Write($"<hr>Export to Pdf-2:{outFileName2}");

        }


        private void BindGridView()
        {
            GridView1.DataSource = GetDataSource();
            GridView1.DataBind();
        }

        protected void GridView1_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            GridView1.PageIndex = e.NewPageIndex;
            BindGridView();
        }

        
        protected void btnGenPDFFitPage_Click(object sender, EventArgs e)
        {
            var excelArg = new ExportDataTable2ExcelArg
            {
                dataSource = GetDataSource(),
                HeaderCenter = "&24 This is Report Header ...",
                HeaderRight = $"&10 使用者:Rainmaker\r日期:{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}",
                FooterRight = "&10 &P/&N",
                ColumnInfos = new Dictionary<string, Tuple<string, double>>
                {
                    {"ProductID", new Tuple<string, double>($"產品\n代號", 5) },
                    {"ProductName", new Tuple<string, double>("產品名稱" , -1)},
                    {"ProductDesc", new Tuple<string, double>("產品 \n描述" , -1)},
                    {"Units", new Tuple<string, double>("產品 庫存" , -1) }
                },
                PageOrientation = PageOrientationType.Landscape,
                IsTextWrapped = true,

            };
            var pdfStream = GenPDFFromDataTable(excelArg);
            var fileNameWithoutExt = $"{Guid.NewGuid().ToString("N")}";
            //string pdfFileName = Path.Combine(Server.MapPath("./data"), $"{fileNameWithoutExt}_temp.pdf");
            //using (FileStream file = new FileStream(pdfFileName, FileMode.Create, System.IO.FileAccess.Write))
            //    pdfStream.CopyTo(file);

            var watermarkArg = new WatermarkArg
            {
                Watermark = $"* 使用者:亂馬客 *{Environment.NewLine}{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}",
                WMStyle = WatermarkStyle.FitPage,
                RotateAngle = 0,
                Opacity = .1

            };
            var waterStream = AddWatermark(pdfStream, watermarkArg);
            //另存檔案
            //string watermarkFileName = Path.Combine(Server.MapPath("./data"), $"{fileNameWithoutExt}.pdf");
            //using (FileStream file = new FileStream(watermarkFileName, FileMode.Create, System.IO.FileAccess.Write))
            //    waterStream.CopyTo(file);
            //直接給Client
            Response.ContentType = "application/pdf";
            Response.AddHeader("content-disposition", "attachment; filename=" + $"{fileNameWithoutExt}.pdf");
            var fileSize = waterStream.Length;
            byte[] pdfBuffer = new byte[(int)fileSize];
            waterStream.Read(pdfBuffer, 0, (int)fileSize);
            waterStream.Close();
            Response.BinaryWrite(pdfBuffer);
            Response.End();
        }

        protected void btnGenPDFRepeatHorizontal_Click(object sender, EventArgs e)
        {
            var excelArg = new ExportDataTable2ExcelArg
            {
                dataSource = GetDataSource(),
                HeaderCenter = "&24 This is Report Header ...",
                HeaderRight = $"&12 使用者:Rainmaker\r日期:{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}",
                FooterRight = "&10 &P/&N",
                ColumnInfos = new Dictionary<string, Tuple<string, double>>
                {
                    {"ProductID", new Tuple<string, double>($"產品代號", -1) },
                    {"ProductName", new Tuple<string, double>("產品名稱" , -1)},
                    {"ProductDesc", new Tuple<string, double>("產品描述" , -1)},
                    {"Units", new Tuple<string, double>("產品 庫存" , -1) }
                },
                PageOrientation = PageOrientationType.Landscape,
                IsTextWrapped = false,
                PageScale = 90,
                FontName = "Microsoft JhengHei Light"
            };
            var pdfStream = GenPDFFromDataTable(excelArg);
            var fileNameWithoutExt = $"{Guid.NewGuid().ToString("N")}";
            //string pdfFileName = Path.Combine(Server.MapPath("./data"), $"{fileNameWithoutExt}_temp.pdf");
            //using (FileStream file = new FileStream(pdfFileName, FileMode.Create, System.IO.FileAccess.Write))
            //    pdfStream.CopyTo(file);

            var watermarkArg = new WatermarkArg
            {
                Watermark = $"* 使用者:亂馬客  *{Environment.NewLine}{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}",
                WMStyle = WatermarkStyle.RepeatHorizontal,
                WatermarkHeight = 100,
                WatermarkWidth = 130,
                WatermarkHorizontalSpace = 50,
                WatermarkVerticalSpace = 30,
                RotateAngle = 30,
                Opacity = .1

            };
            var waterStream = AddWatermark(pdfStream, watermarkArg);
            //string watermarkFileName = Path.Combine(Server.MapPath("./data"), $"{fileNameWithoutExt}.pdf");
            //using (FileStream file = new FileStream(watermarkFileName, FileMode.Create, System.IO.FileAccess.Write))
            //    waterStream.CopyTo(file);
            Response.ContentType = "application/pdf";
            Response.AddHeader("content-disposition", "attachment; filename=" + $"{fileNameWithoutExt}.pdf");
            var fileSize = waterStream.Length;
            byte[] pdfBuffer = new byte[(int)fileSize];
            waterStream.Read(pdfBuffer, 0, (int)fileSize);
            waterStream.Close();
            Response.BinaryWrite(pdfBuffer);
            Response.End();
        }
    }
}