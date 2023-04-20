using ExcelDataReader;
//using ExcelDataReader.DataSet;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

// ""

namespace LabWork3
{
    internal class Var14
    {
        public void meth()
        {
            /*            string filePath = @"D:\programming-technologies\cheats_prices 1.xlsx";

                        FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
                        IExcelDataReader excelReader;

                        //1. Reading Excel file
                        if (Path.GetExtension(filePath).ToUpper() == ".XLS")
                        {
                            //1.1 Reading from a binary Excel file ('97-2003 format; *.xls)
                            excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                        }
                        else
                        {
                            //1.2 Reading from a OpenXml Excel file (2007 format; *.xlsx)
                            excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                        }

                        //2. DataSet - The result of each spreadsheet will be created in the result.Tables
                        DataSet result = excelReader.AsDataSet();

                        //3. DataSet - Create column names from first row
                        excelReader.IsFirstRowAsColumnNames = false;


                        //reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);*/

            string filePath = @"D:\programming-technologies\cheats_prices 1.xlsx";
            
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {

                    // 2. Use the AsDataSet extension method
                    var result = reader.AsDataSet();

                    // The result of each spreadsheet is in result.Tables
                    DataTable dt = result.Tables[0];
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        Console.WriteLine();
                        for (int y = 0; y < dt.Columns.Count; y++)
                        {
                            Console.Write($"{dt.Rows[i][y].ToString()}\t");
                        }
                        Console.WriteLine();
                    }
                    //MessageBox.Show(dt.Rows[1][2].ToString());
                    //Console.WriteLine(dt.Rows[0][0]);
                }
            }
        }
    }
}
