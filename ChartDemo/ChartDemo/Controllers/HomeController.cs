using DocumentFormat.OpenXml.Drawing.Charts;
using Steema.TeeChart;
using Steema.TeeChart.Styles;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ChartDemo.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        public ActionResult GetChart()
        {
            int width = 800, height = 600;
            Steema.TeeChart.TChart mChart = new TChart();
            Steema.TeeChart.Styles.Bar mBar = new Bar();
            Steema.TeeChart.Styles.Line mLine = new Line();
            //Steema.TeeChart.Styles.Points points1 = new Steema.TeeChart.Styles.Points(mChart.Chart);

            //Color redColor = Color.Red;
            // Color redColor = Color.FromArgb(255, 0, 0); RGB寫法
            string hexColor = "#FF0000";
            Color redColor = ColorTranslator.FromHtml(hexColor);

            string barColor = "#FFFFE0";
            Color yellowColor = ColorTranslator.FromHtml(barColor);


            // 設定折線的寬度
            //mLine.LinePen.Width = 2; // 控制折線寬度，可以調整這個值

            mLine.Pointer.Style = Steema.TeeChart.Styles.PointerStyles.Circle;

            // 設定符號的大小
            //mLine.Pointer.SizeDouble = 1; // 控制符號大小，可以調整這個值

            // 設定符號是否顯示數據值
            mLine.Pointer.VertSize = 20; // 控制符號上方顯示數據值的高度
            mLine.Pointer.InflateMargins = true; // 設定為 true，使數據點數值不與折線重疊

            //折線圖的數據點是否顯示,沒有設定預設為false 不會顯示
            mLine.Pointer.Visible = true;


            //mLine.Pointer.Style = Steema.TeeChart.Styles.PointerStyles.Circle;
            //// 設定符號的水平和垂直大小
            mLine.Pointer.HorizSize = 2;
            mLine.Pointer.VertSize = 2;

            //points1.Color = redColor;

            // 設定 Bar 的邊框樣式、顏色和寬度
            mBar.Pen.Style = System.Drawing.Drawing2D.DashStyle.Solid; // 邊框樣式
            mBar.Pen.Color = Color.Black; // 邊框顏色
            mBar.Pen.Width = 1; // 邊框寬度

            mLine.Color = redColor;
            mBar.Color = yellowColor;
            mChart.Header.Text = "TeeChart via ImageShap PNG example";
            //將 Bar 加入 series
            //mChart.Series.Add(points1);
            mChart.Series.Add(mBar);
            mChart.Series.Add(mLine);
            // 內建帶入Sample 資料
            //points1.FillSampleValues();
            mBar.FillSampleValues();
            mLine.FillSampleValues();   
            mBar.XValues.DateTime = true; // mChart.Axes.Bottom..Lables.Angle = 90;
            mLine.XValues.DateTime = true;
            mChart.Axes.Bottom.Increment = Steema.TeeChart.Utils.GetDateTimeStep(DateTimeSteps.OneDay);
            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            mChart.Export.Image.JPEG.Width = width; 
            mChart.Export.Image.JPEG.Height = height; 

            // TODO Lock
            mChart.Export.Image.JPEG.Save(ms);
            ms.Position = 0; 

            ms.Flush();
            FileContentResult res = File(ms.ToArray(), "Image/PNG","");
            ms.Close();

            return res;
        }
    }
}