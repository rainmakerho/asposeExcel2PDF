# asposeExcel2PDF

### 前言

在 [透過 Aspose 將 datatable 的資料轉出成有浮水印的 PDF 檔](https://rainmakerho.github.io/2018/07/03/2018022/) 一文中，驗證可透過 Aspose 元件來達到將 DataTable 的資料轉成有浮水印的 PDF 檔案。但實際用在專案上還有段距離，例如，

1.  設定表頭/表尾
2.  欄位名稱可否換行，對齊方式
3.  可否設定欄位的寬度，格式(日期、數值)，對齊方式
4.  可否設定列印比例
5.  浮水印要有一定的覆蓋比率

### 實作

**DataTable => 換 ColumnName => Excel => Pdf => Pdf+浮水印**
以上面的需求，我們可以將程式切分成 2 段，一是 DataTable 轉資料到 Excel 轉存成 PDF，二再將該 PDF 加上浮水印。

1.  **設定表頭/表尾**
    可以透過 PageSetup.SetHeader(0, "表頭左邊"); 或是 PageSetup.SetFooter(2, "表尾右邊"); 它是使用 0 ~ 2 來表示 左、中、右 區段。這裡有一個部份要注意的是，如果在表頭、表尾區段要換行的話，也是使用 \n 嗎? 不可以哦~ 用 \n 中間會多一行哦! 所以這裡要改用 **\r** 才可以哦!

2.  **欄位名稱可否換行**
    在 ColumnName 中加入 \n 來換行，同時要設定 Cell 的 Stype IsTextWrapped 值為 true。如果 IsTextWrapped 設定為 fasle 的話，在 ColumnName 設定 \n 也不會換行的哦! 另外，程式中會呼叫

3.  **可否設定欄位的寬度**
    可以透過 Worksheet.Cells.SetColumnWidth 來設定欄寬 (0~255)[Cells.SetColumnWidth](https://apireference.aspose.com/net/cells/aspose.cells/cells/methods/setcolumnwidth)。

4.  **可否設定列印比例**
    可以設定 PageSetup.Zoom 的值 (10~400)，預設值是 100。有時如果我們的內容太長時，有可能會跨頁，所以我們可以設定「列印比例」來讓 Excel 印到 PDF 時，可以縮到一頁以內，這時就可以設定這個值了哦!

5.  **浮水印要有一定的覆蓋比率**
    在 PDF 要使用浮水印，我們可以透過 Stamp 物件來貼到 PDF 文件之中，因為我們範例是使用文字，所以是建立 TextStamp 物件，要針對文字做處理，可以透過 FormattedText 來處理。所以，如果浮水印要換行的話，可以使用 \n 或是加入 Environment.NewLine 即可。
    建立好 TextStamp 後，就要依需求來貼到 PDF 之中，如果是蓋滿整個 PDF 文件的話，就可以設定 TextStamp 物件的 Width, Height 跟 PDF Page 的 CropBox Width 及 Height 一樣的大小。
    如果是要依 TextStamp 物件去蓋滿 PDF Page 的話，就可以從 PDF Page 的 CropBox Width 及 Height 去 loop 貼到 PDF Page 上。

透過以上的分析後，我們就可以將它抽出來成為公用的 PDFReportUtility Class### 測試

### 測試，

_註: 因為共用的 function 都是回傳 Stream，所以這們可以將做好的 PDF 檔放到 Byte Array 之中，再透過 Response.BinaryWrite 傳給使用者下載檔案 ^\_^_

1.ColumnName 換行，IsTextWrapped = true, 設定寬度為 5，右表頭加入 2 列(換行)，設定日期格式，第一列表頭對齊方式為置中，浮水印蓋滿一頁。

```csharp
protected void btnGenPDFFitPage_Click(object sender, EventArgs e)
{

	var excelArg = new ExportDataTable2ExcelArg
	{
		dataSource = GetDataSource(),
		HeaderCenter = "&24 This is Report Header ...",
		HeaderRight = $"&10 使用者:Rainmaker\r日期:{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}",
		FooterRight = "&10 &P/&N",
		ColumnInfos = new Dictionary<string, Tuple<string, double, Aspose.Cells.Style>>
		{
			{"ProductID", new Tuple<string, double, Aspose.Cells.Style>($"產品\n代號", 5, null) },
			{"ProductName", new Tuple<string, double, Aspose.Cells.Style>("產品名稱" , -1, null) },
			{"ProductDesc", new Tuple<string, double, Aspose.Cells.Style>("產品 \n描述" , -1,null) },
			{"Units", new Tuple<string, double, Aspose.Cells.Style>("產品 庫存" , -1, new Aspose.Cells.Style{  HorizontalAlignment= TextAlignmentType.Center})},
			 {"CreDte", new Tuple<string, double, Aspose.Cells.Style>("日期" , 20, new Aspose.Cells.Style{ Number=22, Custom = "yyyy/mm/dd hh:mm:ss" , HorizontalAlignment= TextAlignmentType.Center}) }
		},
		PageOrientation = PageOrientationType.Landscape,
		IsTextWrapped = true,
		PageScale = 80,
		HeaderHorizontalAlignment = TextAlignmentType.Center
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
```

產生出來的 PDF 如下圖，
![浮水印蓋滿一頁的PDF](https://github.com/rainmakerho/asposeExcel2PDF/blob/master/onePage.png)

2.ColumnName 自動調整，IsTextWrapped = false，列印比例設定為 90，右表頭加入 2 列(換行)，Cell 字型設定為 "標楷體"(請檢查 Windows\fonts 中是否有該字型)，浮水印水平蓋滿一頁，旋轉角度 30 度。

```csharp
protected void btnGenPDFRepeatHorizontal_Click(object sender, EventArgs e)
{
	var excelArg = new ExportDataTable2ExcelArg
	{
		dataSource = GetDataSource(),
		HeaderCenter = "&24 This is Report Header ...",
		HeaderRight = $"&12 使用者:Rainmaker\r日期:{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}",
		FooterRight = "&10 &P/&N",
		ColumnInfos = new Dictionary<string, Tuple<string, double, Aspose.Cells.Style>>
		{
			{"ProductID", new Tuple<string, double, Aspose.Cells.Style>($"產品代號", -1, null) },
			{"ProductName", new Tuple<string, double, Aspose.Cells.Style>("產品名稱" , -1, null) },
			{"ProductDesc", new Tuple<string, double, Aspose.Cells.Style>("產品描述" , -1, null) },
			{"Units", new Tuple<string, double, Aspose.Cells.Style>("產品 庫存" , -1, null) },
			{"CreDte", new Tuple<string, double, Aspose.Cells.Style>("日期" , 10, new Aspose.Cells.Style{ Number = 14 }) }
		},
		PageOrientation = PageOrientationType.Landscape,
		IsTextWrapped = false,
		PageScale = 80,
		FontName = "標楷體",
		HeaderHorizontalAlignment = TextAlignmentType.Center
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
```


產生出來的 PDF 如下圖，
![浮水印水平蓋滿一頁的PDF](https://github.com/rainmakerho/asposeExcel2PDF/blob/master/repeatPage.png)

### 參考資料

[透過 Aspose 將 datatable 的資料轉出成有浮水印的 PDF 檔](https://rainmakerho.github.io/2018/07/03/2018022/) 

[Setting Page Options](https://docs.aspose.com/display/cellsnet/Setting+Page+Options) 

[Setting Headers and Footers](https://docs.aspose.com/display/cellsnet/Setting+Headers+and+Footers) 

[Aspose Stamps-Watermarks](https://github.com/aspose-pdf/Aspose.PDF-for-.NET/tree/master/Examples/CSharp/AsposePDF/Stamps-Watermarks) 

[Aspose Cells Style.Number Property](https://apireference.aspose.com/net/cells/aspose.cells/style/properties/number) 

[How to control and understand settings in the Format Cells dialog box in Excel](https://support.microsoft.com/en-us/help/264372/how-to-control-and-understand-settings-in-the-format-cells-dialog-box) 

[Column Number format to 4 decimal places](https://forum.aspose.com/t/column-number-format-to-4-decimal-places/41465) 


