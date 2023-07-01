using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ChartDemo.Controllers
{
    public class CellStyleController : Controller
    {
        // GET: CellStyle
        public ActionResult Export()
        {
            DataTable dt1 = CreateDataTable();
            string filename = "test1.xlsx";
            IWorkbook wb = DataTableToIWorkBook(dt1, filename);
            using (MemoryStream stream = new MemoryStream())
            {

                wb.Write(stream);

                //System.Web.MVC.File
                return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", filename);
            }
        }


        public static IWorkbook DataTableToIWorkBook(DataTable dt, string optionalstr = "xlsx")
        {
            IWorkbook wb;
            ISheet ws;
            if (optionalstr.EndsWith("s"))
            {
                wb = new HSSFWorkbook();
            }
            else
            {
                wb = new XSSFWorkbook();
            }

            if (dt.TableName!=string.Empty)
            {
                ws = wb.CreateSheet(dt.TableName);
            }
            else
            {
                ws = wb.CreateSheet("sheet1");
            }
            ws.CreateRow(0);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
            }

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                ws.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    ws.GetRow(i + 1).CreateCell(j).SetCellValue(dt.Rows[i][j].ToString());
                }
            }
            return wb;

        }

        private DataTable CreateDataTable()
        {
            DataTable dt = new DataTable();
            dt.Clear();
            dt.Columns.Add("UserID");
            dt.Columns.Add("UserName");

            DataRow _ravi = dt.NewRow();
            _ravi["UserID"] = "11111";
            _ravi["UserName"] = "Elvis";
            dt.Rows.Add(_ravi);

            _ravi = dt.NewRow();
            _ravi["UserID"] = "22222";
            _ravi["UserName"] = "John";
            dt.Rows.Add(_ravi);

            return dt;

        }

    }
}