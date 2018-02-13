using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ExcelToSQLServer
{
    public partial class Default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void UpdateData_Click(object sender, EventArgs e)
        {
            GetExcelDatasource();
        }

        private void GetExcelDatasource()
        {
            string fileName = "dataSource.xlsx";
            string fileLocation = Server.MapPath("~/SpreadSheets/" + fileName);
            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
                fileLocation + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";

            OleDbConnection excelConnection = new OleDbConnection(connectionString);
            OleDbCommand excelCmd = new OleDbCommand();
            excelCmd.CommandType = System.Data.CommandType.Text;
            excelCmd.Connection = excelConnection;
            OleDbDataAdapter dataAdapter = new OleDbDataAdapter(excelCmd);
            DataTable excelRecords = new DataTable();
            excelConnection.Open();
            DataTable excelSheet = excelConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string getExcelSheet = excelSheet.Rows[0]["Table_Name"].ToString();
            excelCmd.CommandText = "SELECT * FROM [" + getExcelSheet + "]";
            dataAdapter.SelectCommand = excelCmd;
            dataAdapter.Fill(excelRecords);
            dataGrid.DataSource = excelRecords;
            dataGrid.DataBind();

        }
    }
}