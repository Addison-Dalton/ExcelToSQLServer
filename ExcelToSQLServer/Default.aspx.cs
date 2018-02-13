using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
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
            DataTable excelRecords = new DataTable();

            GetExcelDatasource(excelRecords);
            PopulateDatabase(excelRecords);
        }

        private void GetExcelDatasource(DataTable dataTable)
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
            excelConnection.Open();
            DataTable excelSheet = excelConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            string getExcelSheet = excelSheet.Rows[0]["Table_Name"].ToString();
            excelCmd.CommandText = "SELECT * FROM [" + getExcelSheet + "]";
            dataAdapter.SelectCommand = excelCmd;
            dataAdapter.Fill(dataTable);
            dataGrid.DataSource = dataTable;
            dataGrid.DataBind();
        }

        private void PopulateDatabase(DataTable dataTable)
        {
            String ConnectionString = "Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=C:\\Users\\Owner\\source\\repos\\ExcelToSQLServer\\ExcelToSQLServer\\App_Data\\SqlServerDataSource.mdf;Integrated Security=True";
            SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(ConnectionString);
            sqlBulkCopy.ColumnMappings.Add("ID", "Id");
            sqlBulkCopy.ColumnMappings.Add("First Name", "FirstName");
            sqlBulkCopy.ColumnMappings.Add("Last Name", "LastName");
            sqlBulkCopy.DestinationTableName = "Excel_Table";
            sqlBulkCopy.WriteToServer(dataTable);
        }
    }
}