using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Office.Interop.Excel;


namespace ANDPI_Sales_Analysis
{
    class Program
    {
        static void Main(string[] args)
        {
            //string[] sheets = {"2009.xls", "2010.xls", "2011.xls","2012.xls", "2013.xls"};
            string ANDPI_Properties = "input.txt";
            string properties = "allp.txt";
            ExcelManager eManager = new ExcelManager(ANDPI_Properties, properties);
            //eManager.generateExcel(false, "D:\\Documents/without");
            //eManager.generateExcel(true, "D:\\Documents/with");
            eManager.generateBulkExcel("BulkTest2");//"D:\\Documents/Bulk");

        }
    }
}
