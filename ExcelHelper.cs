using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace AutoUpDataBoss
{
    class ExcelHelper
    {
        /// <summary>
        /// 导出Excel文件
        /// </summary>
        /// <param name="filePath">导出并保存的路径</param>
        /// <param name="dataTable">数据集</param>
        /// <param name="sqlstring">第二张sheet中保存的脚本</param>
        /// <param name="columntxt">哪些列是文本的数组</param>
        /// <param name="columndate">哪些列是日期的数组</param>
        /// <returns></returns>
        public static bool DataTableToExcel(string filePath, System.Data.DataTable dataTable,string sqlstring, int[] columntxt, int[] columndate)
        {
            int rowNumber = dataTable.Rows.Count;
            int columnNumber = dataTable.Columns.Count;
            int colIndex = 0;
            if (rowNumber == 0)
            {
                return false;
            }
            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Add(true); //Excel.XlWBATemplate.xlWBATWorksheet
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];  //取出第一个工作表
            object missing = System.Reflection.Missing.Value;//添加一个sheet   
            worksheet = (Excel.Worksheet)workbook.Worksheets.Add(missing, missing, missing, missing);//添加一个sheet   
            Excel.Worksheet worksheet2 = (Excel.Worksheet)workbook.Worksheets[2];       //取出第二个工作表
            worksheet2.Name = "SQL Statement";
            worksheet2.Cells[1, 1] = sqlstring;
            excel.Visible = false;
            Excel.Range range;
            
            foreach (DataColumn col in dataTable.Columns)
            {
                colIndex++;
                excel.Cells[1, colIndex] = col.ColumnName;
            }

            object[,] objData = new object[rowNumber, columnNumber];

            for (int r = 0; r < rowNumber; r++) //循环添加数据行内容
            {
                for (int c = 0; c < columnNumber; c++)
                {
                    objData[r, c] = dataTable.Rows[r][c];
                }
            }

          //  range = worksheet.Range[excel.Cells[2, 1], excel.Cells[rowNumber + 1, 1]];  //Cells[行, 列]
         //   range.NumberFormat = "@";                                                   //针对某列进行格式化

            foreach (int role in columntxt)//遍历数组并定义日期格式,若为0则无文本格式
            {
                if (role != 0)
                {
                    range = worksheet.Range[excel.Cells[2, role], excel.Cells[rowNumber + 1, role]];
                    range.NumberFormat = "@";
                }
            }
            foreach (int role1 in columndate)//遍历数组并定义日期格式,若为0则无日期格式
            {
                if (role1 != 0)
                {
                    range = worksheet.Range[excel.Cells[2, role1], excel.Cells[rowNumber + 1, role1]];
                    range.NumberFormat = "yyyy-m-d hh:mm:ss";
                }
            }
            // range.NumberFormatLocal = "G/通用格式";
            // range.NumberFormat = "@";  //设置单元格格式为文本类型，文本类型可设置上下标
            //    range.NumberFormat = "0.00_ ";//设置单元格格式为数值类型，小数点后2位
            //   range.NumberFormat = "￥#,##0.00;￥-#,##0.00";//设设置单元格格式为货币类型，小数点后2位
            //   range.NumberForma = _"_ ￥*#,##0.00_;_ ￥*-#,##0.00_ ;_ ￥*""-""??_;_ @_ ";//置单元格格式为会计专用类型，小数点后2位
            //   range.NumberFormat = "yyyy-m-d";//设置单元格格式为日期类型
            //    range.NumberFormat = "[$-F400]h:mm:ss AM/PM";//设置单元格格式为时间类型
            //    range.NumberFormat = "0.00%";//设置单元格格式为百分比类型，小数点后2位
            //   range.NumberFormat = "# ?/?";// 设置单元格格式为分数类型，分母为一位数
            //    range.NumberFormat = "0.00E+00";//设置单元格格式为科学技术类型，小数位数为2
            //    range.NumberFormat = "000000";//设置单元格格式为特殊类型
         
            range = worksheet.Range[excel.Cells[2, 1], excel.Cells[rowNumber + 1, columnNumber]];
            range.Value2 = objData;
            range.Font.Size = 9;
            range.EntireColumn.AutoFit();
            range.Font.Name = "宋体";
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            string newPath = pathInfo.Parent.FullName;
            worksheet.SaveAs(newPath + filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            excel.Quit();
            excel = null;
            GC.Collect();
            return true;
        }


        /// <summary>
        /// 导出Excel文件
        /// </summary>
        /// <param name="filePath">导出并保存的路径</param>
        /// <param name="dataTable">数据集</param>
        /// <param name="sqlstring">第二张sheet中保存的脚本</param>
        /// <returns></returns>
        public static bool DataTableTowuxiExcel(string filePath, System.Data.DataTable dataTable1, System.Data.DataTable dataTable2, System.Data.DataTable dataTable3, System.Data.DataTable dataTable4, /*System.Data.DataTable dataTable5,*/ System.Data.DataTable dataTable6, System.Data.DataTable dataTable7, System.Data.DataTable dataTable8, System.Data.DataTable dataTable9, string sqlstring)
        {
            int rowNumber = 9;
            int columnNumber = 3;

            Excel.Application excel = new Excel.Application();
            Excel.Workbook workbook = excel.Workbooks.Add(true); //Excel.XlWBATemplate.xlWBATWorksheet
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Worksheets[1];  //取出第一个工作表
            object missing = System.Reflection.Missing.Value;//添加一个sheet   
            worksheet = (Excel.Worksheet)workbook.Worksheets.Add(missing, missing, missing, missing);//添加一个sheet   
            Excel.Worksheet worksheet2 = (Excel.Worksheet)workbook.Worksheets[2];       //取出第二个工作表
            worksheet2.Name = "SQL Statement";
            worksheet2.Cells[1, 1] = sqlstring;
            excel.Visible = false;
            Excel.Range range;

           // object[,] objData = new object[10, 3] { { "", "客户数", "用户数" }, { "好视乐", dataTable1.Rows[0]["好视乐客户数"].ToString(), dataTable1.Rows[0]["好视乐用户数"].ToString() }, { "看视界", dataTable2.Rows[0]["看视界客户数"].ToString(), dataTable2.Rows[0]["看视界用户数"].ToString() }, { "广联", dataTable3.Rows[0]["广联客户数"].ToString(), dataTable3.Rows[0]["广联用户数"].ToString() }, { "宽带缴费", "", dataTable4.Rows[0]["宽带缴费用户数"].ToString() }, { "留存缴费", dataTable5.Rows[0]["留存缴费客户数"].ToString(),"" }, { "江阴互动基本", "", dataTable6.Rows[0]["江阴互动基本缴费用户数"].ToString() }, { "周新增", dataTable7.Rows[0]["周新增客户数"].ToString(), "" }, { "年新增", dataTable8.Rows[0]["年新增客户数"].ToString(), "" }, { "数字电视缴费", Convert.ToInt32(dataTable9.Rows[0]["开通客户数"].ToString())+ Convert.ToInt32(dataTable9.Rows[0]["预开通客户数"].ToString()), "" } };
            object[,] objData = new object[9, 3] { { "", "客户数", "用户数" }, { "好视乐", dataTable1.Rows[0]["好视乐客户数"].ToString(), dataTable1.Rows[0]["好视乐用户数"].ToString() }, { "看视界", dataTable2.Rows[0]["看视界客户数"].ToString(), dataTable2.Rows[0]["看视界用户数"].ToString() }, { "广联", dataTable3.Rows[0]["广联客户数"].ToString(), dataTable3.Rows[0]["广联用户数"].ToString() }, { "宽带缴费", "", dataTable4.Rows[0]["宽带缴费用户数"].ToString() },  { "江阴互动基本", "", dataTable6.Rows[0]["江阴互动基本缴费用户数"].ToString() }, { "周新增", dataTable7.Rows[0]["周新增客户数"].ToString(), "" }, { "年新增", dataTable8.Rows[0]["年新增客户数"].ToString(), "" }, { "数字电视缴费", Convert.ToInt32(dataTable9.Rows[0]["开通客户数"].ToString()) + Convert.ToInt32(dataTable9.Rows[0]["预开通客户数"].ToString()), "" } };
            range = worksheet.Range[excel.Cells[2, 1], excel.Cells[rowNumber + 1, columnNumber]];
            range.Value2 = objData;
            range.Font.Size = 9;
            range.EntireColumn.AutoFit();
            range.Font.Name = "宋体";
            DirectoryInfo pathInfo = new DirectoryInfo(Environment.CurrentDirectory);
            string newPath = pathInfo.Parent.FullName;
            worksheet.SaveAs(newPath + filePath, Type.Missing, Type.Missing, Type.Missing,/* Type.Missing, */Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            excel.Quit();
            excel = null;
            GC.Collect();
            return true;
        }


        /// <summary>
        /// 将Excel文件导出至DataTable(第一行作为表头)
        /// </summary>
        /// <param name="ExcelFilePath">Excel文件路径</param>
        /// <param name="TableName">数据表名，如果数据表名错误，默认为第一个数据表名</param>
        public static DataTable InputFromExcel(string ExcelFilePath, string TableName)
        {
            if (!File.Exists(ExcelFilePath))
            {
                throw new Exception("Excel文件不存在！");
            }

            //如果数据表名不存在，则数据表名为Excel文件的第一个数据表
            ArrayList TableList = new ArrayList();
            TableList = GetExcelTables(ExcelFilePath);

            if (TableName.IndexOf(TableName) < 0)
            {
                TableName = TableList[0].ToString().Trim();
            }

            DataTable table = new DataTable();
            OleDbConnection dbcon = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + ExcelFilePath + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1';");
            OleDbCommand cmd = new OleDbCommand("select * from [" + TableName + "$]", dbcon);
            OleDbDataAdapter adapter = new OleDbDataAdapter(cmd);

            try
            {
                if (dbcon.State == ConnectionState.Closed)
                {
                    dbcon.Open();
                }
                adapter.Fill(table);
            }
            catch (Exception exp)
            {
                throw exp;
            }
            finally
            {
                if (dbcon.State == ConnectionState.Open)
                {
                    dbcon.Close();
                }
            }
            return table;
        }

        /// <summary>
        /// 获取Excel文件数据表列表
        /// </summary>
        public static ArrayList GetExcelTables(string ExcelFileName)
        {
            DataTable dt = new DataTable();
            ArrayList TablesList = new ArrayList();
            if (File.Exists(ExcelFileName))
            {
                using (OleDbConnection conn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties=Excel 8.0;Data Source=" + ExcelFileName))
                {
                    try
                    {
                        conn.Open();
                        dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                    }
                    catch (Exception exp)
                    {
                        throw exp;
                    }

                    //获取数据表个数
                    int tablecount = dt.Rows.Count;
                    for (int i = 0; i < tablecount; i++)
                    {
                        string tablename = dt.Rows[i][2].ToString().Trim().TrimEnd('$');
                        if (TablesList.IndexOf(tablename) < 0)
                        {
                            TablesList.Add(tablename);
                        }
                    }
                }
            }
            return TablesList;
        }

    }
}
