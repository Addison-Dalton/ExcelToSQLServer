<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="ExcelToSQLServer.Default" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Excel to SQL Server</title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <div>
                <asp:GridView ID="dataGrid" runat="server"></asp:GridView>
            </div>
            <br />
            <div>
                <asp:Button ID="dataButton" Text="Update" OnClick="UpdateData_Click" runat="server" />
            </div>
            <br />
            <div>
                <asp:Label ID="resultLabel" Text="" runat="server"></asp:Label>
            </div>
        </div>
    </form>
</body>
</html>
