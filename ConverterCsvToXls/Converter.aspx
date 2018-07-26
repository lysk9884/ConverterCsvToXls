<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Converter.aspx.cs" Inherits="ConverterCsvToXls.Converter" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title></title>
    <style type="text/css">
        #File1 {
            height: 34px;
            width: 618px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <p>
                <asp:Label ID="Label1" runat="server" Text="Csv 에서 Xlsx 로 변환 하기"></asp:Label>
            </p>
            <input type="file" name="CsvFile" id="CsvFile" runat="server" multiple="multiple" accept=".csv" />
            <asp:Button ID="ToXlsBtn" runat="server" Text="변환" OnClick="ToXlsBtnClicked" />
            <input id="XlsDirInput" type="text" runat="server" />
            <asp:Label ID="Label3" runat="server" Text="변환된 Xlsx 파일이 저장될 위치"></asp:Label>
        </div>
        <div>
            <p>
                <asp:Label ID="Label2" runat="server" Text="Xls 에서 Csv 로 변환 하기"></asp:Label>
            </p>
            <input type="file" name="XlsFile" id="XlsFile" runat="server" multiple="multiple" accept=".xls, .xlsx, .xlsm" />
            <asp:Button ID="ToCsvBtn" runat="server" Text="변환" OnClick="ToCsvBtnClicked" />
            <input id="CsvDirInput" type="text" runat="server" />
            <asp:Label ID="Label4" runat="server" Text="변환된 csv 파일이 저장될 위치"></asp:Label>
        </div>
    </form>
</body>
</html>