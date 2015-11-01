using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Microsoft.Office.Interop.Excel;

namespace Tools.Excel
{
    public class clsExcelReader
    {
        /// <summary>
        /// 获取或设置Excel文件名
        /// </summary>
        public string FileName { set; get; }

        /// <summary>
        /// 获取已用行数
        /// </summary>
        public int RowCount
        {
            get
            {
                return iRowCount;
            }
        }

        /// <summary>
        /// 获取已用列数
        /// </summary>
        public int ColumnCount
        {
            get
            {
                return iColumnCount;
            }
        }

        /// <summary>
        /// 获取错误信息
        /// </summary>
        public string ErrorString
        {
            get
            {
                return strErrorString;
            }
        }

        /// <summary>
        /// 获取或设置打开Excel文件的哪个工作簿
        /// </summary>
        public int SheetNumber { get; set; }

        //操作Excel的变量
        Microsoft.Office.Interop.Excel.Application app;
        Microsoft.Office.Interop.Excel.Sheets sheets;
        Microsoft.Office.Interop.Excel.Workbook workbook = null;
        Microsoft.Office.Interop.Excel.Worksheet worksheet;
        object oMissiong = System.Reflection.Missing.Value;
        //

        string strErrorString;
        int iRowCount = 0;
        int iColumnCount = 0;
        bool booFileOpenState = false;

        /// <summary>
        /// 持续打开Excel文件
        /// </summary>
        /// <returns></returns>
        public bool OpenFileContinuously()
        {
            try
            {
                app = new Microsoft.Office.Interop.Excel.Application();
                workbook = app.Workbooks.Open(FileName, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong, oMissiong);
                sheets = workbook.Worksheets;
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)sheets.get_Item(SheetNumber);//读取一张表
                iRowCount = worksheet.UsedRange.Rows.Count;
                iColumnCount = worksheet.UsedRange.Columns.Count;
                booFileOpenState = true;

                return true;
            }
            catch (Exception ex)
            {
                strErrorString = ex.Message;
                return false;
            }
        }

        /// <summary>
        /// 读取一个单元格的内容
        /// </summary>
        /// <param name="row">被读取单元格所在行</param>
        /// <param name="column">被读取单元格所在列</param>
        /// <returns></returns>
        public string getTextInOneCell(int row, int column)
        {
            try
            {
                Range range = (Range)worksheet.Cells[row, column];
                string strValue = range.Text;
                return strValue;
            }
            catch (Exception ex)
            {
                strErrorString = ex.Message;
                return null;
            }
        }

        /// <summary>
        /// 关闭Excel
        /// </summary>
        /// <returns></returns>
        public bool CloseFile()
        {
            try
            {
                workbook.Close(false, oMissiong, oMissiong);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                workbook = null;
                app.Workbooks.Close();
                app.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                app = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();

                return true;
            }
            catch (Exception ex)
            {
                strErrorString = ex.Message;
                return false;
            }
        }
    }
}
