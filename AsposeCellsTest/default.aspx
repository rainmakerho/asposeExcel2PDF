<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="default.aspx.cs" Inherits="AsposeCellsTest._default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <asp:Button ID="btnGenPDF" runat="server" Text="產生PDF" OnClick="btnGenPDF_Click" />
            <asp:GridView ID="GridView1" runat="server" AllowPaging="true" OnPageIndexChanging="GridView1_PageIndexChanging"></asp:GridView>
            <hr />
             <asp:Button ID="btnGenPDFFitPage" runat="server" Text="Watermark FitPage" OnClick="btnGenPDFFitPage_Click" />
            <asp:Button ID="btnGenPDFRepeatHorizontal" runat="server" Text="Watermark RepeatHorizontal" OnClick="btnGenPDFRepeatHorizontal_Click" />

        </div>
    </form>
</body>
</html>
