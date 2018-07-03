# asposeExcel2PDF

透過 Aspose 將 datatable 的資料轉出成有浮水印的 PDF 檔

Aspose.Cells 可以將各種資料源轉成 Excel (可參考[cells methods](https://apireference.aspose.com/net/cells/aspose.cells/cells/methods/index)的 Import 相關 Methods).
所以我們想要的做法如下，

**DataTable => 換 ColumnName => Excel => Pdf => 浮水印**

測試的 DataTable 資料來測試看看，如下，

```csharp
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
		dataTable.Rows.Add(i, $"產品名稱-{i}", $"產品描述 -{i}", rand.NextDouble());
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
	//沒有 mapping 就用原生的 Column Name
	var outFileName2 = GenPDF(ds2, null,
		Server.MapPath("./data"), ""
		, PageOrientationType.Portrait);
	Response.Write($"<hr>Export to Pdf-2:{outFileName2}");

}
```
