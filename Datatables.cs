using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
namespace Excel4SQL
{
    public class Datatables
    {
        public static DataTable Excel2Datatable()
        {
            string excelConnString = @"Provider=Microsoft.ACE.OLEDB.12.0;" +
                @"Data Source = D:\Workspace\Databases\Excel\OrgPE.xlsx;" +
                @"Extended Properties = ""EXCEL 12.0 XML;HDR=YES""";
            string excelSQL = "select * from [Sheet1$]";
            OleDbConnection excelConn = new OleDbConnection(excelConnString);
            excelConn.Open();
            //获取Excel中所有表的信息
            DataTable schemaTable = excelConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            //获取Excel的第一个Sheet表名
            string sheetName1 = schemaTable.Rows[0][2].ToString().Trim();
            //Console.WriteLine(sheetName1);
            //OleDbCommand excelCmd = new OleDbCommand(excelSQL, excelConn);
            OleDbDataAdapter excelAdapter = new OleDbDataAdapter(excelSQL, excelConn);
            DataSet dataSet = new DataSet();
            excelAdapter.Fill(dataSet, sheetName1);
            DataTable dataTable = new DataTable();
            dataTable = dataSet.Tables[0];
            foreach (DataRow dataRow in dataTable.Rows)
            {
                var str1 = dataRow[0].ToString();
                var str2 = dataRow[1].ToString();
                Console.WriteLine(str1 + "\t" + str2);
            }
            return dataTable;
        }

    }
}