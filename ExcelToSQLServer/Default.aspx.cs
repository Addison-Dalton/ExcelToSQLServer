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
            string dbConnectionString = "Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=C:\\Users\\Owner\\source\\repos\\ExcelToSQLServer\\ExcelToSQLServer\\App_Data\\SqlServerDataSource.mdf;Integrated Security=True;MultipleActiveResultSets=True";
            GetExcelDatasource(excelRecords);
            PopulateDatabase(excelRecords, dbConnectionString);
            DisplayGridView(dbConnectionString);
        }

        //gets the data from the excel spreadsheet and holds it within excelRecords.
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
        }

        //populates the database with data from the passed dataTable, which holds data from the excel spreadsheet.
        private void PopulateDatabase(DataTable dataTable, string connectionString)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                //remove all records from staging table to be sure it is empty before populating it
                SqlCommand deleteStagingRecords = new SqlCommand("DELETE FROM Staging_Table", connection);
                deleteStagingRecords.BeginExecuteNonQuery();

                //
                SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(connectionString);
                sqlBulkCopy.ColumnMappings.Add("ID", "Id");
                sqlBulkCopy.ColumnMappings.Add("First Name", "FirstName");
                sqlBulkCopy.ColumnMappings.Add("Last Name", "LastName");
                sqlBulkCopy.DestinationTableName = "Staging_Table";

                //catch for errors that may occur when writing to the staging table.
                try
                {
                    sqlBulkCopy.WriteToServer(dataTable);
                }
                catch (Exception ex)
                {
                    resultLabel.Text = ex.ToString();
                    //resultLabel.Text = "There was a problem writing to the database";
                }

                //merge the staging table into the primary Excel_Table
                SqlCommand mergeCommand = new SqlCommand("MergeTestProc", connection)
                {
                    CommandType = CommandType.StoredProcedure
                };
                mergeCommand.BeginExecuteNonQuery();
            }
        }

        //gets the data from database and displays it in the gridview.
        private void DisplayGridView(string connectionString)
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                SqlCommand sqlCommand = new SqlCommand("SELECT * FROM Excel_Table", connection);
                SqlDataReader dataReader = sqlCommand.ExecuteReader();
                dataGrid.DataSource = dataReader;
                dataGrid.DataBind();
            }
        }
    }
}