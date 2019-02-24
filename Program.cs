using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace Excel4SQL
{
    class Program
    {
        static void Main(string[] args)
        {
            DataTable excelDT = Datatables.Excel2Datatable();
            Console.WriteLine( excelDT.GetType());
        }
    }
}
