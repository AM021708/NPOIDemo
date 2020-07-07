using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using iTextSharp.text.pdf;
using Microsoft.Ajax.Utilities;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using NPOIDemo.Models;


namespace NPOIDemo.Controllers
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

        public ActionResult Download()
        {
            byte[] bytes = null;
            int rowCounter = 0;
            Rootobject robj = new Rootobject();
            string path = System.AppDomain.CurrentDomain.BaseDirectory;
            robj = JsonConvert.DeserializeObject<Rootobject>(System.IO.File.ReadAllText(path + "\\page1.json"));

            DateTime dateTime = Convert.ToDateTime(robj.data[0].date);
            var datetime = dateTime.ToFullTaiwanDate();

            //建立excel檔案物件
            HSSFWorkbook workbook = new HSSFWorkbook();

            //建立sheet
            //HSSFSheet sheet1 = (HSSFSheet)workbook.CreateSheet("MySheet"); ;// createSheet(workbook, "image query", srcTable);
            ISheet sheet1 = workbook.CreateSheet("MySheet");

            sheet1.SetMargin(MarginType.TopMargin, 0.4); // 0.4 = 1cm
            sheet1.SetMargin(MarginType.RightMargin, 0.4);
            sheet1.SetMargin(MarginType.BottomMargin, 0.4);
            sheet1.SetMargin(MarginType.LeftMargin, 0.4);
            sheet1.PrintSetup.FitWidth = 1;
            sheet1.PrintSetup.FitHeight = 0;
            //sheet1.PrintSetup.PaperSize = 55;
            //sheet1.FitToPage = false;

            //設定此excel的字體大小
            HSSFFont myFont = (HSSFFont)workbook.CreateFont();
            myFont.FontHeightInPoints = 14;
            myFont.FontName = "標楷體";

            HSSFFont Font12 = (HSSFFont)workbook.CreateFont();
            Font12.FontHeightInPoints = 12;
            Font12.FontName = "標楷體";

            HSSFFont Font18 = (HSSFFont)workbook.CreateFont();
            Font18.FontHeightInPoints = 18;
            Font18.FontName = "標楷體";

            HSSFFont Font16 = (HSSFFont)workbook.CreateFont();
            Font16.FontHeightInPoints = 16;
            Font16.FontName = "標楷體";

            HSSFFont Font12B = (HSSFFont)workbook.CreateFont();
            Font12B.FontHeightInPoints = 12;
            Font12B.IsBold = true;
            Font12B.FontName = "標楷體";

            HSSFFont Font12red = (HSSFFont)workbook.CreateFont();
            Font12red.FontHeightInPoints = 12;
            Font12red.IsBold = true;
            Font12red.FontName = "標楷體";
            Font12red.Color = HSSFColor.Red.Index;

            HSSFFont titleFont = (HSSFFont)workbook.CreateFont();
            titleFont.FontHeightInPoints = 20;
            titleFont.FontName = "標楷體";

            #region 建立header row的css樣式

            //框線、背景顏色、文字置中……等等
            HSSFCellStyle headcs = (HSSFCellStyle)workbook.CreateCellStyle();
            //啟動多行文字
            headcs.WrapText = true;
            //文字置中
            headcs.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            headcs.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            //框線樣式及顏色
            headcs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            headcs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            headcs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            headcs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            headcs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            headcs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            headcs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            headcs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            //headcs.FillBackgroundColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            headcs.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Grey25Percent.Index;
            headcs.FillPattern = NPOI.SS.UserModel.FillPattern.SolidForeground;
            headcs.SetFont(myFont);

            #endregion

            #region 建立非header row的css樣式

            HSSFCellStyle cs = (HSSFCellStyle)workbook.CreateCellStyle();
            //啟動多行文字
            cs.WrapText = true;
            //文字置中
            cs.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            cs.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            //框線樣式及顏色
            cs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            cs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            cs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            cs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            cs.SetFont(myFont);

            HSSFCellStyle PioTitlecs = (HSSFCellStyle)workbook.CreateCellStyle();
            PioTitlecs.WrapText = true;
            PioTitlecs.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            PioTitlecs.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Distributed;
            PioTitlecs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            PioTitlecs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            PioTitlecs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            PioTitlecs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            PioTitlecs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            PioTitlecs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            PioTitlecs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            PioTitlecs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            PioTitlecs.SetFont(Font18);


            #endregion

            #region boldcsLT
            HSSFCellStyle boldcsLT = (HSSFCellStyle)workbook.CreateCellStyle();
            boldcsLT.WrapText = true;
            boldcsLT.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            boldcsLT.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            boldcsLT.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            boldcsLT.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            boldcsLT.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            boldcsLT.BorderTop = NPOI.SS.UserModel.BorderStyle.Thick;
            boldcsLT.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsLT.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsLT.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsLT.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsLT.SetFont(Font12B);
            #endregion

            #region boldcsT
            HSSFCellStyle boldcsT = (HSSFCellStyle)workbook.CreateCellStyle();
            boldcsT.WrapText = true;
            boldcsT.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            boldcsT.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            boldcsT.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            boldcsT.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            boldcsT.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            boldcsT.BorderTop = NPOI.SS.UserModel.BorderStyle.Thick;
            boldcsT.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsT.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsT.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsT.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsT.SetFont(Font12B);
            #endregion

            #region boldcsRT
            HSSFCellStyle boldcsRT = (HSSFCellStyle)workbook.CreateCellStyle();
            boldcsRT.WrapText = true;
            boldcsRT.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            boldcsRT.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            boldcsRT.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            boldcsRT.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            boldcsRT.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            boldcsRT.BorderTop = NPOI.SS.UserModel.BorderStyle.Thick;
            boldcsRT.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsRT.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsRT.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsRT.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsRT.SetFont(Font12B);
            #endregion

            #region boldcsL
            HSSFCellStyle boldcsL = (HSSFCellStyle)workbook.CreateCellStyle();
            boldcsL.WrapText = true;
            boldcsL.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            boldcsL.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            boldcsL.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            boldcsL.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            boldcsL.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            boldcsL.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            boldcsL.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsL.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsL.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsL.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsL.SetFont(Font12);
            #endregion

            #region boldcsR
            HSSFCellStyle boldcsR = (HSSFCellStyle)workbook.CreateCellStyle();
            boldcsR.WrapText = true;
            boldcsR.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            boldcsR.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            boldcsR.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            boldcsR.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            boldcsR.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            boldcsR.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            boldcsR.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsR.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsR.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsR.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsR.SetFont(myFont);
            #endregion

            HSSFCellStyle Datecs = (HSSFCellStyle)workbook.CreateCellStyle();
            Datecs.WrapText = true;
            Datecs.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            Datecs.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            Datecs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            Datecs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            Datecs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            Datecs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            Datecs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            Datecs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            Datecs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            Datecs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            Datecs.SetFont(Font18);

            HSSFCellStyle DatecsL = (HSSFCellStyle)workbook.CreateCellStyle();
            DatecsL.WrapText = true;
            DatecsL.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            DatecsL.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            DatecsL.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            DatecsL.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            DatecsL.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            DatecsL.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            DatecsL.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            DatecsL.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            DatecsL.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            DatecsL.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            DatecsL.SetFont(Font18);

            HSSFCellStyle DatecsR = (HSSFCellStyle)workbook.CreateCellStyle();
            DatecsR.WrapText = true;
            DatecsR.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            DatecsR.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            DatecsR.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            DatecsR.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            DatecsR.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            DatecsR.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            DatecsR.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            DatecsR.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            DatecsR.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            DatecsR.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            DatecsR.SetFont(Font18);

            HSSFCellStyle Officercs = (HSSFCellStyle)workbook.CreateCellStyle();
            Officercs.WrapText = true;
            Officercs.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            Officercs.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            Officercs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            Officercs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            Officercs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            Officercs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            Officercs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            Officercs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            Officercs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            Officercs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            Officercs.SetFont(Font18);

            HSSFCellStyle Officercs16 = (HSSFCellStyle)workbook.CreateCellStyle();
            Officercs16.WrapText = true;
            Officercs16.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            Officercs16.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            Officercs16.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            Officercs16.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            Officercs16.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            Officercs16.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            Officercs16.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            Officercs16.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            Officercs16.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            Officercs16.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            Officercs16.SetFont(Font16);

            HSSFCellStyle OfficercsL = (HSSFCellStyle)workbook.CreateCellStyle();
            OfficercsL.WrapText = true;
            OfficercsL.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            OfficercsL.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            OfficercsL.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            OfficercsL.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            OfficercsL.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            OfficercsL.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            OfficercsL.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            OfficercsL.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            OfficercsL.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            OfficercsL.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            OfficercsL.SetFont(Font18);

            HSSFCellStyle OfficercsR = (HSSFCellStyle)workbook.CreateCellStyle();
            OfficercsR.WrapText = true;
            OfficercsR.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            OfficercsR.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            OfficercsR.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            OfficercsR.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            OfficercsR.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            OfficercsR.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            OfficercsR.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            OfficercsR.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            OfficercsR.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            OfficercsR.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            OfficercsR.SetFont(Font18);

            HSSFCellStyle Transcs = (HSSFCellStyle)workbook.CreateCellStyle();
            Transcs.WrapText = true;
            Transcs.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            Transcs.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            Transcs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            Transcs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            Transcs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            Transcs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            Transcs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            Transcs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            Transcs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            Transcs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            Transcs.SetFont(Font12);

            HSSFCellStyle TranscsL = (HSSFCellStyle)workbook.CreateCellStyle();
            TranscsL.WrapText = true;
            TranscsL.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            TranscsL.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            TranscsL.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            TranscsL.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            TranscsL.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            TranscsL.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            TranscsL.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            TranscsL.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            TranscsL.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            TranscsL.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            TranscsL.SetFont(Font12);

            HSSFCellStyle TranscsR = (HSSFCellStyle)workbook.CreateCellStyle();
            TranscsR.WrapText = true;
            TranscsR.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            TranscsR.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            TranscsR.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            TranscsR.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            TranscsR.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            TranscsR.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            TranscsR.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            TranscsR.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            TranscsR.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            TranscsR.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            TranscsR.SetFont(Font12);

            HSSFCellStyle InventoryTitlecs = (HSSFCellStyle)workbook.CreateCellStyle();
            InventoryTitlecs.WrapText = true;
            InventoryTitlecs.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            InventoryTitlecs.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            InventoryTitlecs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            InventoryTitlecs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            InventoryTitlecs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            InventoryTitlecs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            InventoryTitlecs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventoryTitlecs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventoryTitlecs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventoryTitlecs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventoryTitlecs.SetFont(Font12B);

            HSSFCellStyle Inventorycs = (HSSFCellStyle)workbook.CreateCellStyle();
            Inventorycs.WrapText = true;
            Inventorycs.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            Inventorycs.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            Inventorycs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            Inventorycs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            Inventorycs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            Inventorycs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            Inventorycs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            Inventorycs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            Inventorycs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            Inventorycs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            Inventorycs.SetFont(Font12);

            HSSFCellStyle InventorycsL = (HSSFCellStyle)workbook.CreateCellStyle();
            InventorycsL.WrapText = true;
            InventorycsL.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            InventorycsL.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            InventorycsL.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            InventorycsL.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            InventorycsL.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            InventorycsL.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            InventorycsL.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventorycsL.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventorycsL.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventorycsL.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventorycsL.SetFont(Font12);

            HSSFCellStyle InventorycsBL = (HSSFCellStyle)workbook.CreateCellStyle();
            InventorycsBL.WrapText = true;
            InventorycsBL.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            InventorycsBL.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            InventorycsBL.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            InventorycsBL.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            InventorycsBL.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            InventorycsBL.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            InventorycsBL.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventorycsBL.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventorycsBL.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventorycsBL.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventorycsBL.SetFont(Font12B);

            HSSFCellStyle InventorycsR = (HSSFCellStyle)workbook.CreateCellStyle();
            InventorycsR.WrapText = true;
            InventorycsR.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            InventorycsR.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            InventorycsR.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            InventorycsR.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            InventorycsR.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            InventorycsR.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            InventorycsR.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventorycsR.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventorycsR.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventorycsR.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventorycsR.SetFont(Font12);

            HSSFCellStyle otherscs = (HSSFCellStyle)workbook.CreateCellStyle();
            otherscs.WrapText = true;
            otherscs.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            otherscs.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            otherscs.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            otherscs.BorderLeft = NPOI.SS.UserModel.BorderStyle.None;
            otherscs.BorderRight = NPOI.SS.UserModel.BorderStyle.None;
            otherscs.BorderTop = NPOI.SS.UserModel.BorderStyle.None;
            otherscs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            otherscs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            otherscs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            otherscs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            otherscs.SetFont(Font12);

            HSSFCellStyle otherscsL = (HSSFCellStyle)workbook.CreateCellStyle();
            otherscsL.WrapText = true;
            otherscsL.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Top;
            otherscsL.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            otherscsL.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            otherscsL.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            otherscsL.BorderRight = NPOI.SS.UserModel.BorderStyle.None;
            otherscsL.BorderTop = NPOI.SS.UserModel.BorderStyle.None;
            otherscsL.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            otherscsL.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            otherscsL.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            otherscsL.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            otherscsL.SetFont(Font12);

            HSSFCellStyle otherscsR = (HSSFCellStyle)workbook.CreateCellStyle();
            otherscsR.WrapText = true;
            otherscsR.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            otherscsR.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            otherscsR.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            otherscsR.BorderLeft = NPOI.SS.UserModel.BorderStyle.None;
            otherscsR.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            otherscsR.BorderTop = NPOI.SS.UserModel.BorderStyle.None;
            otherscsR.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            otherscsR.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            otherscsR.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            otherscsR.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            otherscsR.SetFont(Font12);

            HSSFCellStyle othersTablecs = (HSSFCellStyle)workbook.CreateCellStyle();
            othersTablecs.WrapText = true;
            othersTablecs.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            othersTablecs.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Distributed;
            othersTablecs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            othersTablecs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            othersTablecs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            othersTablecs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            othersTablecs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            othersTablecs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            othersTablecs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            othersTablecs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            othersTablecs.SetFont(Font12);

            #region boldcsLB
            HSSFCellStyle boldcsLB = (HSSFCellStyle)workbook.CreateCellStyle();
            boldcsLB.WrapText = true;
            boldcsLB.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            boldcsLB.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            boldcsLB.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thick;
            boldcsLB.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            boldcsLB.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            boldcsLB.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            boldcsLB.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsLB.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsLB.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsLB.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsLB.SetFont(myFont);
            #endregion

            #region boldcsB
            HSSFCellStyle boldcsB = (HSSFCellStyle)workbook.CreateCellStyle();
            boldcsB.WrapText = true;
            boldcsB.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            boldcsB.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            boldcsB.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thick;
            boldcsB.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            boldcsB.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            boldcsB.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            boldcsB.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsB.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsB.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsB.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsB.SetFont(myFont);
            #endregion

            #region boldcsRB
            HSSFCellStyle boldcsRB = (HSSFCellStyle)workbook.CreateCellStyle();
            boldcsRB.WrapText = true;
            boldcsRB.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            boldcsRB.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            boldcsRB.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thick;
            boldcsRB.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            boldcsRB.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            boldcsRB.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            boldcsRB.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsRB.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsRB.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsRB.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            boldcsRB.SetFont(myFont);
            #endregion

            HSSFCellStyle InventorycsLB = (HSSFCellStyle)workbook.CreateCellStyle();
            InventorycsLB.WrapText = true;
            InventorycsLB.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            InventorycsLB.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            InventorycsLB.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thick;
            InventorycsLB.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            InventorycsLB.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            InventorycsLB.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            InventorycsLB.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventorycsLB.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventorycsLB.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventorycsLB.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventorycsLB.SetFont(Font12);

            HSSFCellStyle InventorycsBLB = (HSSFCellStyle)workbook.CreateCellStyle();
            InventorycsBLB.WrapText = true;
            InventorycsBLB.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            InventorycsBLB.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            InventorycsBLB.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thick;
            InventorycsBLB.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            InventorycsBLB.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            InventorycsBLB.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            InventorycsBLB.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventorycsBLB.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventorycsBLB.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventorycsBLB.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventorycsBLB.SetFont(Font12B);

            HSSFCellStyle InventorycsB = (HSSFCellStyle)workbook.CreateCellStyle();
            InventorycsB.WrapText = true;
            InventorycsB.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            InventorycsB.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            InventorycsB.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thick;
            InventorycsB.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            InventorycsB.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            InventorycsB.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            InventorycsB.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventorycsB.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventorycsB.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventorycsB.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventorycsB.SetFont(Font12);

            HSSFCellStyle InventorycsRB = (HSSFCellStyle)workbook.CreateCellStyle();
            InventorycsRB.WrapText = true;
            InventorycsRB.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            InventorycsRB.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            InventorycsRB.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thick;
            InventorycsRB.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            InventorycsRB.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            InventorycsRB.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            InventorycsRB.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventorycsRB.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventorycsRB.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventorycsRB.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            InventorycsRB.SetFont(Font12);

            HSSFCellStyle otherscsLB = (HSSFCellStyle)workbook.CreateCellStyle();
            otherscsLB.WrapText = true;
            otherscsLB.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            otherscsLB.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            otherscsLB.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thick;
            otherscsLB.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            otherscsLB.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            otherscsLB.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            otherscsLB.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            otherscsLB.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            otherscsLB.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            otherscsLB.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            otherscsLB.SetFont(Font12);

            HSSFCellStyle otherscsB = (HSSFCellStyle)workbook.CreateCellStyle();
            otherscsB.WrapText = true;
            otherscsB.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            otherscsB.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            otherscsB.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thick;
            otherscsB.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            otherscsB.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            otherscsB.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            otherscsB.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            otherscsB.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            otherscsB.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            otherscsB.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            otherscsB.SetFont(Font12);

            HSSFCellStyle otherscsRB = (HSSFCellStyle)workbook.CreateCellStyle();
            otherscsRB.WrapText = true;
            otherscsRB.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            otherscsRB.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            otherscsRB.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thick;
            otherscsRB.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            otherscsRB.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            otherscsRB.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            otherscsRB.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            otherscsRB.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            otherscsRB.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            otherscsRB.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            otherscsRB.SetFont(Font12);

            // P4
            HSSFCellStyle stampscsP4 = (HSSFCellStyle)workbook.CreateCellStyle();
            stampscsP4.WrapText = true;
            stampscsP4.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            stampscsP4.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            stampscsP4.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            stampscsP4.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            stampscsP4.BorderRight = NPOI.SS.UserModel.BorderStyle.None;
            stampscsP4.BorderTop = NPOI.SS.UserModel.BorderStyle.Thick;
            stampscsP4.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsP4.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsP4.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsP4.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsP4.SetFont(Font16);

            HSSFCellStyle stampscsLTforRead = (HSSFCellStyle)workbook.CreateCellStyle();
            stampscsLTforRead.WrapText = true;
            stampscsLTforRead.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Top;
            stampscsLTforRead.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            stampscsLTforRead.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            stampscsLTforRead.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            stampscsLTforRead.BorderRight = NPOI.SS.UserModel.BorderStyle.None;
            stampscsLTforRead.BorderTop = NPOI.SS.UserModel.BorderStyle.Thick;
            stampscsLTforRead.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsLTforRead.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsLTforRead.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsLTforRead.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsLTforRead.SetFont(Font16);

            HSSFCellStyle officerSigncs = (HSSFCellStyle)workbook.CreateCellStyle();
            officerSigncs.WrapText = true;
            officerSigncs.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            officerSigncs.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            officerSigncs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            officerSigncs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            officerSigncs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            officerSigncs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            officerSigncs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            officerSigncs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            officerSigncs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            officerSigncs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            officerSigncs.SetFont(Font12);

            // 1
            HSSFCellStyle stampscsLT = (HSSFCellStyle)workbook.CreateCellStyle();
            stampscsLT.WrapText = true;
            stampscsLT.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            stampscsLT.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            stampscsLT.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            stampscsLT.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            stampscsLT.BorderRight = NPOI.SS.UserModel.BorderStyle.None;
            stampscsLT.BorderTop = NPOI.SS.UserModel.BorderStyle.Thick;
            stampscsLT.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsLT.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsLT.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsLT.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsLT.SetFont(Font16);

            // 2
            HSSFCellStyle stampscsT = (HSSFCellStyle)workbook.CreateCellStyle();
            stampscsT.WrapText = true;
            stampscsT.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            stampscsT.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            stampscsT.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            stampscsT.BorderLeft = NPOI.SS.UserModel.BorderStyle.None;
            stampscsT.BorderRight = NPOI.SS.UserModel.BorderStyle.None;
            stampscsT.BorderTop = NPOI.SS.UserModel.BorderStyle.Thick;
            stampscsT.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsT.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsT.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsT.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsT.SetFont(Font12);
            // 3
            HSSFCellStyle stampscsRT = (HSSFCellStyle)workbook.CreateCellStyle();
            stampscsRT.WrapText = true;
            stampscsRT.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            stampscsRT.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            stampscsRT.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            stampscsRT.BorderLeft = NPOI.SS.UserModel.BorderStyle.None;
            stampscsRT.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            stampscsRT.BorderTop = NPOI.SS.UserModel.BorderStyle.Thick;
            stampscsRT.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsRT.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsRT.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsRT.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsRT.SetFont(Font12);

            // 4
            HSSFCellStyle stampscsL = (HSSFCellStyle)workbook.CreateCellStyle();
            stampscsL.WrapText = true;
            stampscsL.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            stampscsL.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            stampscsL.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            stampscsL.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            stampscsL.BorderRight = NPOI.SS.UserModel.BorderStyle.None;
            stampscsL.BorderTop = NPOI.SS.UserModel.BorderStyle.None;
            stampscsL.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsL.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsL.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsL.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsL.SetFont(Font12);

            HSSFCellStyle stampscsLrec = (HSSFCellStyle)workbook.CreateCellStyle(); // 給蓋印章的 L1
            stampscsLrec.WrapText = true;
            stampscsLrec.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            stampscsLrec.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            stampscsLrec.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            stampscsLrec.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            stampscsLrec.BorderRight = NPOI.SS.UserModel.BorderStyle.None;
            stampscsLrec.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            stampscsLrec.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsLrec.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsLrec.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsLrec.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsLrec.SetFont(Font12);

            HSSFCellStyle stampscsLrec2 = (HSSFCellStyle)workbook.CreateCellStyle(); // 給蓋印章的 L2
            stampscsLrec2.WrapText = true;
            stampscsLrec2.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            stampscsLrec2.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            stampscsLrec2.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            stampscsLrec2.BorderLeft = NPOI.SS.UserModel.BorderStyle.None;
            stampscsLrec2.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            stampscsLrec2.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            stampscsLrec2.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsLrec2.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsLrec2.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsLrec2.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsLrec2.SetFont(Font12);

            HSSFCellStyle stampscsLrec3 = (HSSFCellStyle)workbook.CreateCellStyle(); // 給蓋印章的 L3
            stampscsLrec3.WrapText = true;
            stampscsLrec3.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            stampscsLrec3.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            stampscsLrec3.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            stampscsLrec3.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            stampscsLrec3.BorderRight = NPOI.SS.UserModel.BorderStyle.None;
            stampscsLrec3.BorderTop = NPOI.SS.UserModel.BorderStyle.None;
            stampscsLrec3.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsLrec3.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsLrec3.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsLrec3.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsLrec3.SetFont(Font12);

            HSSFCellStyle stampscsLrec4 = (HSSFCellStyle)workbook.CreateCellStyle(); // 給蓋印章的 L4
            stampscsLrec4.WrapText = true;
            stampscsLrec4.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            stampscsLrec4.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            stampscsLrec4.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            stampscsLrec4.BorderLeft = NPOI.SS.UserModel.BorderStyle.None;
            stampscsLrec4.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            stampscsLrec4.BorderTop = NPOI.SS.UserModel.BorderStyle.None;
            stampscsLrec4.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsLrec4.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsLrec4.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsLrec4.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsLrec4.SetFont(Font12);

            HSSFCellStyle stampscsRrec = (HSSFCellStyle)workbook.CreateCellStyle(); // 給蓋印章的 R1
            stampscsRrec.WrapText = true;
            stampscsRrec.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            stampscsRrec.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            stampscsRrec.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            stampscsRrec.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            stampscsRrec.BorderRight = NPOI.SS.UserModel.BorderStyle.None;
            stampscsRrec.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            stampscsRrec.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsRrec.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsRrec.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsRrec.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsRrec.SetFont(Font12);

            HSSFCellStyle stampscsRrec2 = (HSSFCellStyle)workbook.CreateCellStyle(); // 給蓋印章的 R2
            stampscsRrec2.WrapText = true;
            stampscsRrec2.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            stampscsRrec2.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            stampscsRrec2.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            stampscsRrec2.BorderLeft = NPOI.SS.UserModel.BorderStyle.None;
            stampscsRrec2.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            stampscsRrec2.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            stampscsRrec2.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsRrec2.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsRrec2.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsRrec2.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsRrec2.SetFont(Font12);

            HSSFCellStyle stampscsRrec3 = (HSSFCellStyle)workbook.CreateCellStyle(); // 給蓋印章的 R3
            stampscsRrec3.WrapText = true;
            stampscsRrec3.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            stampscsRrec3.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            stampscsRrec3.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            stampscsRrec3.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            stampscsRrec3.BorderRight = NPOI.SS.UserModel.BorderStyle.None;
            stampscsRrec3.BorderTop = NPOI.SS.UserModel.BorderStyle.None;
            stampscsRrec3.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsRrec3.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsRrec3.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsRrec3.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsRrec3.SetFont(Font12);

            HSSFCellStyle stampscsRrec4 = (HSSFCellStyle)workbook.CreateCellStyle(); // 給蓋印章的 R4
            stampscsRrec4.WrapText = true;
            stampscsRrec4.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            stampscsRrec4.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            stampscsRrec4.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            stampscsRrec4.BorderLeft = NPOI.SS.UserModel.BorderStyle.None;
            stampscsRrec4.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            stampscsRrec4.BorderTop = NPOI.SS.UserModel.BorderStyle.None;
            stampscsRrec4.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsRrec4.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsRrec4.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsRrec4.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsRrec4.SetFont(Font12);

            // 5
            HSSFCellStyle stampscsNone = (HSSFCellStyle)workbook.CreateCellStyle();
            stampscsNone.WrapText = true;
            stampscsNone.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            stampscsNone.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            stampscsNone.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            stampscsNone.BorderLeft = NPOI.SS.UserModel.BorderStyle.None;
            stampscsNone.BorderRight = NPOI.SS.UserModel.BorderStyle.None;
            stampscsNone.BorderTop = NPOI.SS.UserModel.BorderStyle.None;
            stampscsNone.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsNone.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsNone.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsNone.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsNone.SetFont(Font12);

            // 6
            HSSFCellStyle stampscsR = (HSSFCellStyle)workbook.CreateCellStyle();
            stampscsR.WrapText = true;
            stampscsR.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            stampscsR.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            stampscsR.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            stampscsR.BorderLeft = NPOI.SS.UserModel.BorderStyle.None;
            stampscsR.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            stampscsR.BorderTop = NPOI.SS.UserModel.BorderStyle.None;
            stampscsR.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsR.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsR.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsR.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsR.SetFont(Font12);

            // 7
            HSSFCellStyle stampscsLB = (HSSFCellStyle)workbook.CreateCellStyle();
            stampscsLB.WrapText = true;
            stampscsLB.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            stampscsLB.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            stampscsLB.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thick;
            stampscsLB.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            stampscsLB.BorderRight = NPOI.SS.UserModel.BorderStyle.None;
            stampscsLB.BorderTop = NPOI.SS.UserModel.BorderStyle.None;
            stampscsLB.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsLB.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsLB.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsLB.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsLB.SetFont(Font12);

            // 8
            HSSFCellStyle stampscsB = (HSSFCellStyle)workbook.CreateCellStyle();
            stampscsB.WrapText = true;
            stampscsB.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            stampscsB.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            stampscsB.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thick;
            stampscsB.BorderLeft = NPOI.SS.UserModel.BorderStyle.None;
            stampscsB.BorderRight = NPOI.SS.UserModel.BorderStyle.None;
            stampscsB.BorderTop = NPOI.SS.UserModel.BorderStyle.None;
            stampscsB.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsB.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsB.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsB.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsB.SetFont(Font12);

            // 9
            HSSFCellStyle stampscsRB = (HSSFCellStyle)workbook.CreateCellStyle();
            stampscsRB.WrapText = true;
            stampscsRB.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            stampscsRB.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            stampscsRB.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thick;
            stampscsRB.BorderLeft = NPOI.SS.UserModel.BorderStyle.None;
            stampscsRB.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            stampscsRB.BorderTop = NPOI.SS.UserModel.BorderStyle.None;
            stampscsRB.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsRB.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsRB.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsRB.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscsRB.SetFont(Font12);

            HSSFCellStyle stampscs = (HSSFCellStyle)workbook.CreateCellStyle();
            stampscs.WrapText = true;
            stampscs.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            stampscs.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            stampscs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            stampscs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            stampscs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            stampscs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            stampscs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampscs.SetFont(Font12);

            HSSFCellStyle tiltecs = (HSSFCellStyle)workbook.CreateCellStyle();
            tiltecs.WrapText = true;
            tiltecs.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            tiltecs.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Distributed;
            tiltecs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            tiltecs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            tiltecs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            tiltecs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thick;
            tiltecs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            tiltecs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            tiltecs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            tiltecs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            tiltecs.SetFont(myFont);

            HSSFCellStyle tilte20 = (HSSFCellStyle)workbook.CreateCellStyle();
            tilte20.WrapText = true;
            tilte20.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            tilte20.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Distributed;
            tilte20.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            tilte20.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            tilte20.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            tilte20.BorderTop = NPOI.SS.UserModel.BorderStyle.Thick;
            tilte20.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            tilte20.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            tilte20.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            tilte20.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            tilte20.SetFont(titleFont);

            HSSFCellStyle tiltecs20 = (HSSFCellStyle)workbook.CreateCellStyle();
            tiltecs20.WrapText = true;
            tiltecs20.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            tiltecs20.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Distributed;
            tiltecs20.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            tiltecs20.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            tiltecs20.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            tiltecs20.BorderTop = NPOI.SS.UserModel.BorderStyle.None;
            tiltecs20.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            tiltecs20.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            tiltecs20.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            tiltecs20.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            tiltecs20.SetFont(titleFont);

            HSSFCellStyle tilteBottomcs20 = (HSSFCellStyle)workbook.CreateCellStyle();
            tilteBottomcs20.WrapText = true;
            tilteBottomcs20.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            tilteBottomcs20.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Distributed;
            tilteBottomcs20.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            tilteBottomcs20.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            tilteBottomcs20.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            tilteBottomcs20.BorderTop = NPOI.SS.UserModel.BorderStyle.None;
            tilteBottomcs20.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            tilteBottomcs20.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            tilteBottomcs20.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            tilteBottomcs20.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            tilteBottomcs20.SetFont(titleFont);



            #region 合併儲存格









            /*
            sheet1.AddMergedRegion(new CellRangeAddress(6, 6, 1, 6)); // transactionItem
            sheet1.AddMergedRegion(new CellRangeAddress(6, 6, 7, 8)); // transactionAmount

            sheet1.AddMergedRegion(new CellRangeAddress(7, 7, 1, 6)); // transactionItem
            sheet1.AddMergedRegion(new CellRangeAddress(7, 7, 7, 8)); // transactionAmount

            sheet1.AddMergedRegion(new CellRangeAddress(8, 8, 1, 6)); // transactionItem
            sheet1.AddMergedRegion(new CellRangeAddress(8, 8, 7, 8)); // transactionAmount

            sheet1.AddMergedRegion(new CellRangeAddress(9, 9, 1, 6)); // transactionItem
            sheet1.AddMergedRegion(new CellRangeAddress(9, 9, 7, 8)); // transactionAmount

            sheet1.AddMergedRegion(new CellRangeAddress(10, 10, 1, 6)); // transactionItem
            sheet1.AddMergedRegion(new CellRangeAddress(10, 10, 7, 8)); // transactionAmount

            sheet1.AddMergedRegion(new CellRangeAddress(11, 11, 1, 6)); // transactionItem
            sheet1.AddMergedRegion(new CellRangeAddress(11, 11, 7, 8)); // transactionAmount

            sheet1.AddMergedRegion(new CellRangeAddress(12, 12, 1, 6)); // transactionItem
            sheet1.AddMergedRegion(new CellRangeAddress(12, 12, 7, 8)); // transactionAmount

            sheet1.AddMergedRegion(new CellRangeAddress(13, 13, 1, 6)); // transactionItem
            sheet1.AddMergedRegion(new CellRangeAddress(13, 13, 7, 8)); // transactionAmount

            sheet1.AddMergedRegion(new CellRangeAddress(14, 14, 1, 6)); // transactionItem
            sheet1.AddMergedRegion(new CellRangeAddress(14, 14, 7, 8)); // transactionAmount

            sheet1.AddMergedRegion(new CellRangeAddress(15, 15, 1, 6)); // transactionItem
            sheet1.AddMergedRegion(new CellRangeAddress(15, 15, 7, 8)); // transactionAmount
            
            
            

            sheet1.AddMergedRegion(new CellRangeAddress(20, 20, 0, 10)); // inventoryTitle
            sheet1.AddMergedRegion(new CellRangeAddress(23, 23, 1, 3)); // inventoryLiaison
            sheet1.AddMergedRegion(new CellRangeAddress(23, 23, 4, 5)); // inventoryLiaison
            sheet1.AddMergedRegion(new CellRangeAddress(23, 23, 6, 8)); // inventoryLiaison
            sheet1.AddMergedRegion(new CellRangeAddress(23, 23, 9, 10)); // inventoryLiaison
            */









            sheet1.AddMergedRegion(new CellRangeAddress(0, 0, 11, 12));

            sheet1.AddMergedRegion(new CellRangeAddress(2, 4, 11, 11));
            sheet1.AddMergedRegion(new CellRangeAddress(2, 4, 12, 12));
            sheet1.AddMergedRegion(new CellRangeAddress(6, 8, 11, 11));
            sheet1.AddMergedRegion(new CellRangeAddress(6, 8, 12, 12));
            sheet1.AddMergedRegion(new CellRangeAddress(10, 12, 11, 11));
            sheet1.AddMergedRegion(new CellRangeAddress(10, 12, 12, 12));
            #endregion




            //建立title row
            //前兩個是縱向範圍，後兩個為橫向範圍
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter + 1, 0, 10)); // title (0, 0, 0, 10)
            HSSFRow titleRow = (HSSFRow)sheet1.CreateRow(rowCounter);
            titleRow.CreateCell(rowCounter).SetCellValue("國防部參謀本部資通電軍指揮部");//國防部參謀本部資通電軍指揮部忠信營區總值日官紀事簿
            titleRow.GetCell(rowCounter).CellStyle = tilte20;
            titleRow.Height = 20 * 20;
            for (int c = 1; c < 11; c++)
            {
                if (c == 10)
                {
                    titleRow.CreateCell(c).SetCellValue("");
                    titleRow.GetCell(c).CellStyle = tilte20;
                }
                titleRow.CreateCell(c).SetCellValue("");
                titleRow.GetCell(c).CellStyle = tilte20;
            }
            titleRow.CreateCell(11).SetCellValue("簽核");
            titleRow.GetCell(11).CellStyle = officerSigncs;
            titleRow.CreateCell(12).SetCellValue("");
            titleRow.GetCell(12).CellStyle = officerSigncs;


            HSSFRow titleRow2 = (HSSFRow)sheet1.CreateRow(rowCounter + 1);
            titleRow2.CreateCell(rowCounter).SetCellValue("");//國防部參謀本部資通電軍指揮部忠信營區總值日官紀事簿
            titleRow2.GetCell(rowCounter).CellStyle = tiltecs20;
            titleRow2.Height = 20 * 20;
            for (int c = 1; c < 11; c++)
            {
                if (c == 10)
                {
                    titleRow2.CreateCell(c).SetCellValue("");
                    titleRow2.GetCell(c).CellStyle = tiltecs20;
                }
                titleRow2.CreateCell(c).SetCellValue("");
                titleRow2.GetCell(c).CellStyle = tiltecs20;
            }
            if (robj.data[0].next_audit == 0)
            {
                titleRow2.CreateCell(11).SetCellValue("V");
                titleRow2.GetCell(11).CellStyle = officerSigncs;
                titleRow2.CreateCell(12).SetCellValue("");
                titleRow2.GetCell(12).CellStyle = officerSigncs;
            }
            else if (robj.data[0].next_audit == 1)
            {
                titleRow2.CreateCell(11).SetCellValue("");
                titleRow2.GetCell(11).CellStyle = officerSigncs;
                titleRow2.CreateCell(12).SetCellValue("V");
                titleRow2.GetCell(12).CellStyle = officerSigncs;
            }
            else
            {
                titleRow2.CreateCell(11).SetCellValue("");
                titleRow2.GetCell(11).CellStyle = officerSigncs;
                titleRow2.CreateCell(12).SetCellValue("");
                titleRow2.GetCell(12).CellStyle = officerSigncs;
            }


            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter + 2, rowCounter + 3, 0, 10)); // title (0, 0, 0, 10)
            HSSFRow titleRow3 = (HSSFRow)sheet1.CreateRow(rowCounter + 2);
            titleRow3.CreateCell(rowCounter).SetCellValue("忠信營區總值日官紀事簿");//國防部參謀本部資通電軍指揮部忠信營區總值日官紀事簿
            titleRow3.GetCell(rowCounter).CellStyle = tiltecs20;
            //titleRow3.Height = 20 * 20;
            for (int c = 1; c < 11; c++)
            {
                if (c == 10)
                {
                    titleRow3.CreateCell(c).SetCellValue("");
                    titleRow3.GetCell(c).CellStyle = tiltecs20;
                }
                titleRow3.CreateCell(c).SetCellValue("");
                titleRow3.GetCell(c).CellStyle = tiltecs20;
            }
            titleRow3.CreateCell(11).SetCellValue("指揮官");
            titleRow3.GetCell(11).CellStyle = officerSigncs;
            titleRow3.CreateCell(12).SetCellValue("副指揮官");
            titleRow3.GetCell(12).CellStyle = officerSigncs;

            HSSFRow titleRow4 = (HSSFRow)sheet1.CreateRow(rowCounter + 3);
            titleRow4.CreateCell(rowCounter).SetCellValue("");//國防部參謀本部資通電軍指揮部忠信營區總值日官紀事簿
            titleRow4.GetCell(rowCounter).CellStyle = tilteBottomcs20;
            //titleRow4.Height = 20 * 20;
            for (int c = 1; c < 11; c++)
            {
                if (c == 10)
                {
                    titleRow4.CreateCell(c).SetCellValue("");
                    titleRow4.GetCell(c).CellStyle = tilteBottomcs20;
                }
                titleRow4.CreateCell(c).SetCellValue("");
                titleRow4.GetCell(c).CellStyle = tilteBottomcs20;
            }
            titleRow4.CreateCell(11).SetCellValue("");
            titleRow4.GetCell(11).CellStyle = officerSigncs;
            titleRow4.CreateCell(12).SetCellValue("");
            titleRow4.GetCell(12).CellStyle = officerSigncs;


            //用header的文字長度寬度來調整欄位寬度比較準確
            //不要最後才用整個完成的excel資料去調整
            //圖片才不會扭曲失真
            //自動調整每個欄位的大小
            //sheet1.AutoSizeColumn(rowCounter);
            rowCounter = rowCounter + 4;

            #region Header
            //前兩個是縱向範圍，後兩個為橫向範圍
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 0, 4)); // Date (4, 4, 0, 4)
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 5, 6)); // status (4, 4, 5, 6)
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 8, 9)); // weather (4, 4, 8, 9)
            HSSFRow headerRow = (HSSFRow)sheet1.CreateRow(rowCounter);
            headerRow.CreateCell(11).SetCellValue("");
            headerRow.GetCell(11).CellStyle = officerSigncs;
            headerRow.CreateCell(12).SetCellValue("");
            headerRow.GetCell(12).CellStyle = officerSigncs;
            headerRow.Height = 30 * 20;
            for (int i = 0; i < 11; i++)
            {
                if (i == 0)
                {
                    headerRow.CreateCell(i).SetCellValue("時間：民國" + datetime);
                    headerRow.GetCell(i).CellStyle = DatecsL;
                    //sheet1.SetColumnWidth(i, 300);
                    //sheet1.AutoSizeColumn(i);
                }
                else if (i == 5)
                {
                    headerRow.CreateCell(i).SetCellValue("星期");
                    headerRow.GetCell(i).CellStyle = Datecs;
                    //sheet1.AutoSizeColumn(i);
                }
                else if (i == 7)
                {
                    if (robj.data[0].status == 0)
                    {
                        headerRow.CreateCell(i).SetCellValue("日");
                        headerRow.GetCell(i).CellStyle = Datecs;
                        //sheet1.AutoSizeColumn(i);
                    }
                    else if (robj.data[0].status == 1)
                    {
                        headerRow.CreateCell(i).SetCellValue("一");
                        headerRow.GetCell(i).CellStyle = Datecs;
                        //sheet1.AutoSizeColumn(i);
                    }
                    else if (robj.data[0].status == 2)
                    {
                        headerRow.CreateCell(i).SetCellValue("二");
                        headerRow.GetCell(i).CellStyle = Datecs;
                        //sheet1.AutoSizeColumn(i);
                    }
                    else if (robj.data[0].status == 3)
                    {
                        headerRow.CreateCell(i).SetCellValue("三");
                        headerRow.GetCell(i).CellStyle = Datecs;
                        //sheet1.AutoSizeColumn(i);
                    }
                    else if (robj.data[0].status == 4)
                    {
                        headerRow.CreateCell(i).SetCellValue("四");
                        headerRow.GetCell(i).CellStyle = Datecs;
                        //sheet1.AutoSizeColumn(i);
                    }
                    else if (robj.data[0].status == 5)
                    {
                        headerRow.CreateCell(i).SetCellValue("五");
                        headerRow.GetCell(i).CellStyle = Datecs;
                        //sheet1.AutoSizeColumn(i);
                    }
                    else if (robj.data[0].status == 6)
                    {
                        headerRow.CreateCell(i).SetCellValue("六");
                        headerRow.GetCell(i).CellStyle = Datecs;
                        //sheet1.AutoSizeColumn(i);
                    }
                    else
                    {
                        headerRow.CreateCell(i).SetCellValue("");
                        headerRow.GetCell(i).CellStyle = Datecs;
                        //sheet1.AutoSizeColumn(i);
                    }

                }
                else if (i == 8)
                {
                    headerRow.CreateCell(i).SetCellValue("天候");
                    headerRow.GetCell(i).CellStyle = Datecs;
                    //sheet1.AutoSizeColumn(i);
                }
                else if (i == 10)
                {
                    if (robj.data[0].weather == 0)
                    {
                        headerRow.CreateCell(i).SetCellValue("晴");
                        headerRow.GetCell(i).CellStyle = DatecsR;
                        //sheet1.AutoSizeColumn(i);
                    }
                    else if (robj.data[0].weather == 1)
                    {
                        headerRow.CreateCell(i).SetCellValue("陰");
                        headerRow.GetCell(i).CellStyle = DatecsR;
                        //sheet1.AutoSizeColumn(i);
                    }
                    else if (robj.data[0].weather == 2)
                    {
                        headerRow.CreateCell(i).SetCellValue("雨");
                        headerRow.GetCell(i).CellStyle = DatecsR;
                        //sheet1.AutoSizeColumn(i);
                    }


                }
                else
                {
                    headerRow.CreateCell(i).SetCellValue("");
                    headerRow.GetCell(i).CellStyle = Datecs;
                    //sheet1.AutoSizeColumn(i);
                }
            }
            rowCounter++;
            #endregion

            #region Officer
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 0, 2)); // mainSign (5, 5, 0, 2)          
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 4, 6)); // mainGive (5, 5, 4, 6)
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 8, 10)); // mainRecieve (5, 5, 8, 10)
            HSSFRow mainOfficerRow = (HSSFRow)sheet1.CreateRow(rowCounter);
            // 打勾
            if (robj.data[0].next_audit == 2)
            {
                mainOfficerRow.CreateCell(11).SetCellValue("V");
                mainOfficerRow.GetCell(11).CellStyle = officerSigncs;
                mainOfficerRow.CreateCell(12).SetCellValue("");
                mainOfficerRow.GetCell(12).CellStyle = officerSigncs;
            }
            if (robj.data[0].next_audit == 3)
            {
                mainOfficerRow.CreateCell(11).SetCellValue("");
                mainOfficerRow.GetCell(11).CellStyle = officerSigncs;
                mainOfficerRow.CreateCell(12).SetCellValue("V");
                mainOfficerRow.GetCell(12).CellStyle = officerSigncs;
            }
            else
            {
                mainOfficerRow.CreateCell(11).SetCellValue("");
                mainOfficerRow.GetCell(11).CellStyle = officerSigncs;
                mainOfficerRow.CreateCell(12).SetCellValue("");
                mainOfficerRow.GetCell(12).CellStyle = officerSigncs;
            }

            //mainOfficerRow.Height = 10 * 20;
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter + 1, rowCounter + 1, 0, 2)); // subSign (6, 6, 0, 2)        
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter + 1, rowCounter + 1, 4, 6)); // subGive (6, 6, 4, 6)
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter + 1, rowCounter + 1, 8, 10)); // subRecieve (6, 6, 8, 10)
            HSSFRow subOfficerRow = (HSSFRow)sheet1.CreateRow(rowCounter + 1);
            subOfficerRow.CreateCell(11).SetCellValue("參謀長");
            subOfficerRow.GetCell(11).CellStyle = officerSigncs;
            subOfficerRow.CreateCell(12).SetCellValue("政戰主任");
            subOfficerRow.GetCell(12).CellStyle = officerSigncs;
            //subOfficerRow.Height = 10 * 20;
            for (int i = 0; i < 2; i++)
            {
                for (int j = 0; j < 11; j++)
                {
                    if (i == 0) // 總值日官
                    {
                        if (j == 0)
                        {
                            mainOfficerRow.CreateCell(j).SetCellValue("總值日官簽名");
                            mainOfficerRow.GetCell(j).CellStyle = OfficercsL;
                            //sheet1.AutoSizeColumn(j);
                        }
                        else if (j == 3)
                        {
                            mainOfficerRow.CreateCell(j).SetCellValue("交");
                            mainOfficerRow.GetCell(j).CellStyle = Officercs;
                            //sheet1.AutoSizeColumn(j);
                        }
                        else if (j == 4)
                        {
                            mainOfficerRow.CreateCell(j).SetCellValue(robj.data[0].officer.main.give);
                            mainOfficerRow.GetCell(j).CellStyle = Officercs16;
                            //sheet1.AutoSizeColumn(j);
                        }
                        else if (j == 7)
                        {
                            mainOfficerRow.CreateCell(j).SetCellValue("接");
                            mainOfficerRow.GetCell(j).CellStyle = Officercs;
                            //sheet1.AutoSizeColumn(j);
                        }
                        else if (j == 8)
                        {
                            mainOfficerRow.CreateCell(j).SetCellValue(robj.data[0].officer.main.recieve);
                            mainOfficerRow.GetCell(j).CellStyle = Officercs16;
                            //sheet1.AutoSizeColumn(j);
                        }
                        else if (j == 10)
                        {
                            mainOfficerRow.CreateCell(j).SetCellValue("");
                            mainOfficerRow.GetCell(j).CellStyle = OfficercsR;
                        }
                        else
                        {
                            mainOfficerRow.CreateCell(j).SetCellValue("");
                            mainOfficerRow.GetCell(j).CellStyle = Officercs;
                        }
                    }

                    if (i == 1)
                    {

                        if (j == 0)
                        {
                            subOfficerRow.CreateCell(j).SetCellValue("副總值日官簽名");
                            subOfficerRow.GetCell(j).CellStyle = OfficercsL;
                            //sheet1.AutoSizeColumn(j);
                        }
                        else if (j == 3)
                        {
                            subOfficerRow.CreateCell(j).SetCellValue("交");
                            subOfficerRow.GetCell(j).CellStyle = Officercs;
                            //sheet1.AutoSizeColumn(j);
                        }
                        else if (j == 4)
                        {
                            subOfficerRow.CreateCell(j).SetCellValue(robj.data[0].officer.sub.give);
                            subOfficerRow.GetCell(j).CellStyle = Officercs16;
                            //sheet1.AutoSizeColumn(j);
                        }
                        else if (j == 7)
                        {
                            subOfficerRow.CreateCell(j).SetCellValue("接");
                            subOfficerRow.GetCell(j).CellStyle = Officercs;
                            //sheet1.AutoSizeColumn(j);
                        }
                        else if (j == 8)
                        {
                            subOfficerRow.CreateCell(j).SetCellValue(robj.data[0].officer.sub.recieve);
                            subOfficerRow.GetCell(j).CellStyle = Officercs16;
                            //sheet1.AutoSizeColumn(j);
                        }
                        else if (j == 10)
                        {
                            subOfficerRow.CreateCell(j).SetCellValue("");
                            subOfficerRow.GetCell(j).CellStyle = OfficercsR;
                        }
                        else
                        {
                            mainOfficerRow.CreateCell(j).SetCellValue("");
                            mainOfficerRow.GetCell(j).CellStyle = Officercs;
                        }
                    }

                }
                rowCounter++;

            }
            #endregion

            #region Transaction
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 1, 6)); // transactionItem (7, 7, 1, 6)
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 7, 8)); // transactionAmount (7, 7, 7, 8)
            HSSFRow transactionRow = (HSSFRow)sheet1.CreateRow(rowCounter);
            transactionRow.CreateCell(11).SetCellValue("");
            transactionRow.GetCell(11).CellStyle = officerSigncs;
            transactionRow.CreateCell(12).SetCellValue("");
            transactionRow.GetCell(12).CellStyle = officerSigncs;
            string[] columnTransaction = new string[5] { "項次", "移交清冊及設備",
                "數量", "交", "接" };
            for (int c = 0; c < 11; c++)
            {
                if (c == 0)
                {
                    transactionRow.CreateCell(c).SetCellValue("項次");
                    transactionRow.GetCell(c).CellStyle = boldcsL;
                }
                else if (c == 1)
                {
                    transactionRow.CreateCell(c).SetCellValue("移交清冊及設備");
                    transactionRow.GetCell(c).CellStyle = cs;
                }
                else if (c == 7)
                {
                    transactionRow.CreateCell(c).SetCellValue("數量");
                    transactionRow.GetCell(c).CellStyle = cs;
                }
                else if (c == 9)
                {
                    transactionRow.CreateCell(c).SetCellValue("交");
                    transactionRow.GetCell(c).CellStyle = cs;
                }
                else if (c == 10)
                {
                    transactionRow.CreateCell(c).SetCellValue("接");
                    transactionRow.GetCell(c).CellStyle = boldcsR;
                }
                else
                {
                    transactionRow.CreateCell(c).SetCellValue("");
                    transactionRow.GetCell(c).CellStyle = cs;
                }
            }
            rowCounter++;

            for (int i = 0; i < robj.data[0].transaction.Length; i++)
            {
                sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 1, 6)); // 從8到17
                sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 7, 8));
                HSSFRow dataRow = (HSSFRow)sheet1.CreateRow(rowCounter); // 定義標題欄位的位置，雖然在for迴圈裡面，名稱可以都一樣但還是更改成不同的比較好看的懂
                if (i == 0)
                {
                    dataRow.CreateCell(11).SetCellValue("");
                    dataRow.GetCell(11).CellStyle = officerSigncs;
                    dataRow.CreateCell(12).SetCellValue("");
                    dataRow.GetCell(12).CellStyle = officerSigncs;
                }
                if (i == 1)
                { //打勾
                    if (robj.data[0].next_audit == 4)
                    {
                        dataRow.CreateCell(11).SetCellValue("V");
                        dataRow.GetCell(11).CellStyle = officerSigncs;
                        dataRow.CreateCell(12).SetCellValue("");
                        dataRow.GetCell(12).CellStyle = officerSigncs;
                    }
                    else if (robj.data[0].next_audit == 5)
                    {
                        dataRow.CreateCell(11).SetCellValue("");
                        dataRow.GetCell(11).CellStyle = officerSigncs;
                        dataRow.CreateCell(12).SetCellValue("V");
                        dataRow.GetCell(12).CellStyle = officerSigncs;
                    }
                    else
                    {
                        dataRow.CreateCell(11).SetCellValue("");
                        dataRow.GetCell(11).CellStyle = officerSigncs;
                        dataRow.CreateCell(12).SetCellValue("");
                        dataRow.GetCell(12).CellStyle = officerSigncs;
                    }

                }
                if (i == 2)
                {
                    dataRow.CreateCell(11).SetCellValue("副參謀長");
                    dataRow.GetCell(11).CellStyle = officerSigncs;
                    dataRow.CreateCell(12).SetCellValue("單位主管");
                    dataRow.GetCell(12).CellStyle = officerSigncs;
                }
                if (i == 3)
                {
                    dataRow.CreateCell(11).SetCellValue("");
                    dataRow.GetCell(11).CellStyle = officerSigncs;
                    dataRow.CreateCell(12).SetCellValue("");
                    dataRow.GetCell(12).CellStyle = officerSigncs;
                }
                if (i == 4)
                {
                    dataRow.CreateCell(11).SetCellValue("");
                    dataRow.GetCell(11).CellStyle = officerSigncs;
                    dataRow.CreateCell(12).SetCellValue("");
                    dataRow.GetCell(12).CellStyle = officerSigncs;
                }
                for (int j = 0; j < 11; j++)
                {
                    if (j == 0)
                    {
                        int num = i + 1;
                        dataRow.CreateCell(j).SetCellValue(num);
                        dataRow.GetCell(j).CellStyle = TranscsL;
                    }
                    else if (j == 1)
                    {
                        dataRow.CreateCell(j).SetCellValue(robj.data[0].transaction[i].item);
                        dataRow.GetCell(j).CellStyle = Transcs;
                    }
                    else if (j == 7)
                    {
                        dataRow.CreateCell(j).SetCellValue(robj.data[0].transaction[i].amount);
                        dataRow.GetCell(j).CellStyle = cs;
                    }
                    else if (j == 9)
                    {
                        if (robj.data[0].transaction[i].give)
                        {
                            dataRow.CreateCell(j).SetCellValue("V");
                            dataRow.GetCell(j).CellStyle = cs;
                        }
                        else
                        {
                            dataRow.CreateCell(j).SetCellValue("");
                            dataRow.GetCell(j).CellStyle = cs;
                        }

                    }
                    else if (j == 10)
                    {
                        if (robj.data[0].transaction[i].recieve)
                        {
                            dataRow.CreateCell(j).SetCellValue("V");
                            dataRow.GetCell(j).CellStyle = TranscsR;
                        }
                        else
                        {
                            dataRow.CreateCell(j).SetCellValue("");
                            dataRow.GetCell(j).CellStyle = TranscsR;
                        }
                    }
                    else
                    {
                        dataRow.CreateCell(j).SetCellValue("");
                        dataRow.GetCell(j).CellStyle = Transcs;
                    }
                }
                rowCounter++;
            }
            #endregion

            #region Inventory
            /*先判別inventory.length有無>10 V
             * 如果正好=10則就照原本的第一筆資料那樣處理 V
             * 如果>10則先填滿一row以後重新判斷一次 V
             * 如果<10則再繼續判斷有無>5 V
             * 如果正好=5則資料隔一欄填入並每一欄都向右合併一格 V
             * 如果>5則先宣告一計數器來計inventory.length-5並依據這數字決定從第一欄位開始要合併多少格(這個最難) V
             * 如果<5則先照=5的情況先合併右邊一格再依據需要填入的欄位數再合併一次
             * 
             * 
            */

            for (int i = 0; i < robj.data[0].inventory.Length; i++) // 總歸還是從有幾個inventory項目而來
            {
                sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 0, 10)); // inventoryTitle (18, 18, 0, 10)
                HSSFRow inventoryRow = (HSSFRow)sheet1.CreateRow(rowCounter); // start with 18
                inventoryRow.CreateCell(0).SetCellValue(robj.data[0].inventory[i].title);
                inventoryRow.GetCell(0).CellStyle = boldcsLT;
                rowCounter++;
                for (int c = 1; c < 11; c++)
                {
                    if (c == 10)
                    {
                        inventoryRow.CreateCell(c).SetCellValue("");
                        inventoryRow.GetCell(c).CellStyle = boldcsRT;
                    }
                    else
                    {
                        inventoryRow.CreateCell(c).SetCellValue("");
                        inventoryRow.GetCell(c).CellStyle = boldcsT;
                    }
                }
                HSSFRow dataItemRow = (HSSFRow)sheet1.CreateRow(rowCounter);
                for (int c = 0; c < 11; c++)
                {
                    if (c == 0)
                    {
                        dataItemRow.CreateCell(c).SetCellValue("");
                        dataItemRow.GetCell(c).CellStyle = InventorycsBL;
                    }
                    else if (c == 10)
                    {
                        dataItemRow.CreateCell(c).SetCellValue("");
                        dataItemRow.GetCell(c).CellStyle = InventorycsR;
                    }
                    dataItemRow.CreateCell(c).SetCellValue("");
                    dataItemRow.GetCell(c).CellStyle = Inventorycs;
                }
                int invenItemCount = 0;
                invenItemCount = robj.data[0].inventory[i].items.Length;
                if (robj.data[0].inventory[i].items.Length >= 10) //大於或正好=10
                {
                    for (int j = 0; j < 10; j++)
                    {
                        if (j == 0)
                        {
                            dataItemRow.CreateCell(j).SetCellValue("品名");
                            dataItemRow.GetCell(j).CellStyle = InventorycsBL;
                            dataItemRow.CreateCell(j + 1).SetCellValue(robj.data[0].inventory[i].items[j].item);
                            dataItemRow.GetCell(j + 1).CellStyle = Inventorycs;
                            //sheet1.AutoSizeColumn(j);
                        }
                        else if (j == robj.data[0].inventory[i].items.Length - 1)
                        {
                            dataItemRow.CreateCell(j + 1).SetCellValue(robj.data[0].inventory[i].items[j].item);
                            dataItemRow.GetCell(j + 1).CellStyle = InventorycsR;
                        }
                        else
                        {
                            dataItemRow.CreateCell(j + 1).SetCellValue(robj.data[0].inventory[i].items[j].item);
                            dataItemRow.GetCell(j + 1).CellStyle = Inventorycs;
                        }
                    }
                    rowCounter++;
                    invenItemCount = invenItemCount - 10;
                }
                else
                {
                    if (robj.data[0].inventory[i].items.Length == 5)
                    {
                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 1, 2)); // temp for inventory
                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 3, 4)); // temp for inventory
                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 5, 6)); // temp for inventory
                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 7, 8)); // temp for inventory
                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 9, 10)); // temp for inventory

                        for (int j = 0; j < robj.data[0].inventory[i].items.Length; j++)
                        {
                            if (j == 0)
                            {
                                dataItemRow.CreateCell(j).SetCellValue("品名");
                                dataItemRow.GetCell(j).CellStyle = InventorycsBL;
                                dataItemRow.CreateCell(j + 1).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                dataItemRow.GetCell(j + 1).CellStyle = Inventorycs;
                                //sheet1.AutoSizeColumn(j);
                            }
                            else if (j == 1)
                            {
                                dataItemRow.CreateCell(j + 2).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                dataItemRow.GetCell(j + 2).CellStyle = Inventorycs;
                            }
                            else if (j == 2)
                            {
                                dataItemRow.CreateCell(j + 3).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                dataItemRow.GetCell(j + 3).CellStyle = Inventorycs;
                            }
                            else if (j == 3)
                            {
                                dataItemRow.CreateCell(j + 4).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                dataItemRow.GetCell(j + 4).CellStyle = Inventorycs;
                            }
                            else if (j == 4)
                            {
                                dataItemRow.CreateCell(j + 5).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                dataItemRow.GetCell(j + 5).CellStyle = Inventorycs;
                            }
                        }
                        rowCounter++;
                    }
                    else if (robj.data[0].inventory[i].items.Length > 5)
                    {
                        if (robj.data[0].inventory[i].items.Length == 6)
                        {
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 1, 2)); // temp for inventory
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 3, 4)); // temp for inventory
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 5, 6)); // temp for inventory
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 7, 8)); // temp for inventory


                            for (int j = 0; j < robj.data[0].inventory[i].items.Length; j++)
                            {
                                if (j == 0)
                                {
                                    dataItemRow.CreateCell(j).SetCellValue("品名");
                                    dataItemRow.GetCell(j).CellStyle = InventorycsBL;
                                    dataItemRow.CreateCell(j + 1).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                    dataItemRow.GetCell(j + 1).CellStyle = Inventorycs;
                                    //sheet1.AutoSizeColumn(j);
                                }
                                else if (j == 1)
                                {
                                    dataItemRow.CreateCell(j + 2).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                    dataItemRow.GetCell(j + 2).CellStyle = Inventorycs;
                                }
                                else if (j == 2)
                                {
                                    dataItemRow.CreateCell(j + 3).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                    dataItemRow.GetCell(j + 3).CellStyle = Inventorycs;
                                }
                                else if (j == 3)
                                {
                                    dataItemRow.CreateCell(j + 4).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                    dataItemRow.GetCell(j + 4).CellStyle = Inventorycs;
                                }
                                else if (j == 4)
                                {
                                    dataItemRow.CreateCell(j + 5).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                    dataItemRow.GetCell(j + 5).CellStyle = Inventorycs;
                                }
                                else
                                {
                                    dataItemRow.CreateCell(j + 5).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                    dataItemRow.GetCell(j + 5).CellStyle = Inventorycs;
                                }
                                dataItemRow.GetCell(10).CellStyle = InventorycsR;
                            }
                            rowCounter++;
                        }
                        else if (robj.data[0].inventory[i].items.Length == 7)
                        {
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 1, 2)); // temp for inventory
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 3, 4)); // temp for inventory
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 5, 6)); // temp for inventory                                

                            for (int j = 0; j < robj.data[0].inventory[i].items.Length; j++)
                            {
                                if (j == 0)
                                {
                                    dataItemRow.CreateCell(j).SetCellValue("品名");
                                    dataItemRow.GetCell(j).CellStyle = InventorycsBL;
                                    dataItemRow.CreateCell(j + 1).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                    dataItemRow.GetCell(j + 1).CellStyle = Inventorycs;
                                    //sheet1.AutoSizeColumn(j);
                                }
                                else if (j == 1)
                                {
                                    dataItemRow.CreateCell(j + 2).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                    dataItemRow.GetCell(j + 2).CellStyle = Inventorycs;
                                }
                                else if (j == 2)
                                {
                                    dataItemRow.CreateCell(j + 3).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                    dataItemRow.GetCell(j + 3).CellStyle = Inventorycs;
                                }
                                else if (j == 3)
                                {
                                    dataItemRow.CreateCell(j + 4).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                    dataItemRow.GetCell(j + 4).CellStyle = Inventorycs;
                                }
                                else
                                {
                                    dataItemRow.CreateCell(j + 4).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                    dataItemRow.GetCell(j + 4).CellStyle = Inventorycs;
                                }
                                dataItemRow.GetCell(10).CellStyle = InventorycsR;
                            }
                            rowCounter++;
                        }
                        else if (robj.data[0].inventory[i].items.Length == 8)
                        {
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 1, 2)); // temp for inventory
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 3, 4)); // temp for inventory                              

                            for (int j = 0; j < robj.data[0].inventory[i].items.Length; j++)
                            {
                                if (j == 0)
                                {
                                    dataItemRow.CreateCell(j).SetCellValue("品名");
                                    dataItemRow.GetCell(j).CellStyle = InventorycsBL;
                                    dataItemRow.CreateCell(j + 1).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                    dataItemRow.GetCell(j + 1).CellStyle = Inventorycs;
                                    //sheet1.AutoSizeColumn(j);
                                }
                                else if (j == 1)
                                {
                                    dataItemRow.CreateCell(j + 2).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                    dataItemRow.GetCell(j + 2).CellStyle = Inventorycs;
                                }
                                else if (j == 2)
                                {
                                    dataItemRow.CreateCell(j + 3).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                    dataItemRow.GetCell(j + 3).CellStyle = Inventorycs;
                                }
                                else
                                {
                                    dataItemRow.CreateCell(j + 3).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                    dataItemRow.GetCell(j + 3).CellStyle = Inventorycs;
                                }
                                dataItemRow.GetCell(10).CellStyle = InventorycsR;
                            }
                            rowCounter++;
                        }
                        else if (robj.data[0].inventory[i].items.Length == 9)
                        {
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 1, 2)); // temp for inventory                                

                            for (int j = 0; j < robj.data[0].inventory[i].items.Length; j++)
                            {
                                if (j == 0)
                                {
                                    dataItemRow.CreateCell(j).SetCellValue("品名");
                                    dataItemRow.GetCell(j).CellStyle = InventorycsBL;
                                    dataItemRow.CreateCell(j + 1).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                    dataItemRow.GetCell(j + 1).CellStyle = Inventorycs;
                                    //sheet1.AutoSizeColumn(j);
                                }
                                else if (j == 1)
                                {
                                    dataItemRow.CreateCell(j + 2).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                    dataItemRow.GetCell(j + 2).CellStyle = Inventorycs;
                                }
                                else
                                {
                                    dataItemRow.CreateCell(j + 2).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                    dataItemRow.GetCell(j + 2).CellStyle = Inventorycs;
                                }
                                dataItemRow.GetCell(10).CellStyle = InventorycsR;
                            }
                            rowCounter++;
                        }
                    }
                    else // < 5
                    {
                        if (robj.data[0].inventory[i].items.Length == 1)
                        {
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 1, 10)); // temp for inventory

                            for (int j = 0; j < robj.data[0].inventory[i].items.Length; j++)
                            {
                                if (j == 0)
                                {
                                    dataItemRow.CreateCell(j).SetCellValue("品名");
                                    dataItemRow.GetCell(j).CellStyle = InventorycsBL;
                                    dataItemRow.CreateCell(j + 1).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                    dataItemRow.GetCell(j + 1).CellStyle = Inventorycs;
                                    //sheet1.AutoSizeColumn(j);
                                }
                            }
                            rowCounter++;
                        }
                        else if (robj.data[0].inventory[i].items.Length == 2)
                        {
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 1, 5)); // temp for inventory
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 6, 10)); // temp for inventory

                            for (int j = 0; j < robj.data[0].inventory[i].items.Length; j++)
                            {
                                if (j == 0)
                                {
                                    dataItemRow.CreateCell(j).SetCellValue("品名");
                                    dataItemRow.GetCell(j).CellStyle = InventorycsBL;
                                    dataItemRow.CreateCell(j + 1).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                    dataItemRow.GetCell(j + 1).CellStyle = Inventorycs;
                                    //sheet1.AutoSizeColumn(j);
                                }
                                else if (j == 1)
                                {
                                    dataItemRow.CreateCell(j + 5).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                    dataItemRow.GetCell(j + 5).CellStyle = Inventorycs;
                                }
                                dataItemRow.GetCell(10).CellStyle = InventorycsR;
                            }
                            rowCounter++;
                        }
                        else if (robj.data[0].inventory[i].items.Length == 3)
                        {
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 1, 4)); // temp for inventory
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 5, 7)); // temp for inventory
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 8, 10)); // temp for inventory

                            for (int j = 0; j < robj.data[0].inventory[i].items.Length; j++)
                            {
                                if (j == 0)
                                {
                                    dataItemRow.CreateCell(j).SetCellValue("品名");
                                    dataItemRow.GetCell(j).CellStyle = InventorycsBL;
                                    dataItemRow.CreateCell(j + 1).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                    dataItemRow.GetCell(j + 1).CellStyle = Inventorycs;
                                    //sheet1.AutoSizeColumn(j);
                                }
                                else if (j == 1)
                                {
                                    dataItemRow.CreateCell(j + 4).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                    dataItemRow.GetCell(j + 4).CellStyle = Inventorycs;
                                }
                                else if (j == 2)
                                {
                                    dataItemRow.CreateCell(j + 6).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                    dataItemRow.GetCell(j + 6).CellStyle = Inventorycs;
                                }
                                dataItemRow.GetCell(10).CellStyle = InventorycsR;
                            }
                            rowCounter++;
                        }
                        else if (robj.data[0].inventory[i].items.Length == 4)
                        {
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 1, 3)); // temp for inventory
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 4, 5)); // temp for inventory
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 6, 8)); // temp for inventory
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 9, 10)); // temp for inventory

                            for (int j = 0; j < robj.data[0].inventory[i].items.Length; j++)
                            {
                                if (j == 0)
                                {
                                    dataItemRow.CreateCell(j).SetCellValue("品名");
                                    dataItemRow.GetCell(j).CellStyle = InventorycsBL;
                                    dataItemRow.CreateCell(j + 1).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                    dataItemRow.GetCell(j + 1).CellStyle = Inventorycs;
                                    //sheet1.AutoSizeColumn(j);
                                }
                                else if (j == 1)
                                {
                                    dataItemRow.CreateCell(j + 3).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                    dataItemRow.GetCell(j + 3).CellStyle = Inventorycs;
                                }
                                else if (j == 2)
                                {
                                    dataItemRow.CreateCell(j + 4).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                    dataItemRow.GetCell(j + 4).CellStyle = Inventorycs;
                                }
                                else if (j == 3)
                                {
                                    dataItemRow.CreateCell(j + 6).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                    dataItemRow.GetCell(j + 6).CellStyle = Inventorycs;
                                }
                                dataItemRow.GetCell(10).CellStyle = InventorycsR;
                            }
                            rowCounter++;
                        }
                    }
                }
                HSSFRow dataCheckRow = (HSSFRow)sheet1.CreateRow(rowCounter);
                for (int c = 0; c < 11; c++)
                {
                    if (c == 0)
                    {
                        dataCheckRow.CreateCell(c).SetCellValue("");
                        dataCheckRow.GetCell(c).CellStyle = InventorycsBL;
                    }
                    else if (c == 10)
                    {
                        dataCheckRow.CreateCell(c).SetCellValue("");
                        dataCheckRow.GetCell(c).CellStyle = InventorycsR;
                    }
                    else
                    {
                        dataCheckRow.CreateCell(c).SetCellValue("");
                        dataCheckRow.GetCell(c).CellStyle = Inventorycs;
                    }
                }
                invenItemCount = 0;
                invenItemCount = robj.data[0].inventory[i].items.Length;
                if (robj.data[0].inventory[i].items.Length >= 10) //大於或正好=10
                {
                    for (int j = 0; j < 10; j++)
                    {
                        if (j == 0)
                        {
                            dataCheckRow.CreateCell(j).SetCellValue("清點");
                            dataCheckRow.GetCell(j).CellStyle = InventorycsBL;
                            if (robj.data[0].inventory[i].items[j].check)
                            {
                                dataCheckRow.CreateCell(j + 1).SetCellValue("V");
                                dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                            }
                            else
                            {
                                dataCheckRow.CreateCell(j + 1).SetCellValue("");
                                dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                            }

                            sheet1.SetColumnWidth(j + 1, 10 * 256);
                            //sheet1.AutoSizeColumn(j);
                        }
                        else if (j == robj.data[0].inventory[i].items.Length - 1)
                        {
                            if (robj.data[0].inventory[i].items[j].check)
                            {
                                dataCheckRow.CreateCell(j + 1).SetCellValue("V");
                                dataCheckRow.GetCell(j + 1).CellStyle = InventorycsR;
                            }
                            else
                            {
                                dataCheckRow.CreateCell(j + 1).SetCellValue("");
                                dataCheckRow.GetCell(j + 1).CellStyle = InventorycsR;
                            }
                            sheet1.SetColumnWidth(j + 1, 10 * 256);
                        }
                        else
                        {
                            if (robj.data[0].inventory[i].items[j].check)
                            {
                                dataCheckRow.CreateCell(j + 1).SetCellValue("V");
                                dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                            }
                            else
                            {
                                dataCheckRow.CreateCell(j + 1).SetCellValue("");
                                dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                            }
                            sheet1.SetColumnWidth(j + 1, 10 * 256);
                        }
                    }
                    rowCounter++;
                    invenItemCount = invenItemCount - 10;
                }
                else
                {
                    if (robj.data[0].inventory[i].items.Length == 5)
                    {
                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 1, 2)); // temp for inventory
                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 3, 4)); // temp for inventory
                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 5, 6)); // temp for inventory
                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 7, 8)); // temp for inventory
                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 9, 10)); // temp for inventory

                        for (int j = 0; j < robj.data[0].inventory[i].items.Length; j++)
                        {
                            if (j == 0)
                            {
                                dataCheckRow.CreateCell(j).SetCellValue("清點");
                                dataCheckRow.GetCell(j).CellStyle = InventorycsBL;
                                if (robj.data[0].inventory[i].items[j].check)
                                {
                                    dataCheckRow.CreateCell(j + 1).SetCellValue("V");
                                    dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                }
                                else
                                {
                                    dataCheckRow.CreateCell(j + 1).SetCellValue("");
                                    dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                }
                                sheet1.SetColumnWidth(j + 1, 10 * 256);
                                //sheet1.AutoSizeColumn(j);
                            }
                            else if (j == 1)
                            {
                                if (robj.data[0].inventory[i].items[j].check)
                                {
                                    dataCheckRow.CreateCell(j + 2).SetCellValue("V");
                                    dataCheckRow.GetCell(j + 2).CellStyle = Inventorycs;
                                }
                                else
                                {
                                    dataCheckRow.CreateCell(j + 2).SetCellValue("");
                                    dataCheckRow.GetCell(j + 2).CellStyle = Inventorycs;
                                }
                                sheet1.SetColumnWidth(j + 1, 10 * 256);
                            }
                            else if (j == 2)
                            {
                                if (robj.data[0].inventory[i].items[j].check)
                                {
                                    dataCheckRow.CreateCell(j + 3).SetCellValue("V");
                                    dataCheckRow.GetCell(j + 3).CellStyle = Inventorycs;
                                }
                                else
                                {
                                    dataCheckRow.CreateCell(j + 3).SetCellValue("");
                                    dataCheckRow.GetCell(j + 3).CellStyle = Inventorycs;
                                }
                                sheet1.SetColumnWidth(j + 1, 10 * 256);
                            }
                            else if (j == 3)
                            {
                                if (robj.data[0].inventory[i].items[j].check)
                                {
                                    dataCheckRow.CreateCell(j + 4).SetCellValue("V");
                                    dataCheckRow.GetCell(j + 4).CellStyle = Inventorycs;
                                }
                                else
                                {
                                    dataCheckRow.CreateCell(j + 4).SetCellValue("");
                                    dataCheckRow.GetCell(j + 4).CellStyle = Inventorycs;
                                }
                                sheet1.SetColumnWidth(j + 1, 10 * 256);
                            }
                            else if (j == 4)
                            {
                                if (robj.data[0].inventory[i].items[j].check)
                                {
                                    dataCheckRow.CreateCell(j + 5).SetCellValue("V");
                                    dataCheckRow.GetCell(j + 5).CellStyle = Inventorycs;
                                }
                                else
                                {
                                    dataCheckRow.CreateCell(j + 5).SetCellValue("");
                                    dataCheckRow.GetCell(j + 5).CellStyle = Inventorycs;
                                }
                                sheet1.SetColumnWidth(j + 1, 10 * 256);
                            }
                            dataCheckRow.GetCell(10).CellStyle = InventorycsR;
                        }
                        rowCounter++;
                    }
                    else if (robj.data[0].inventory[i].items.Length > 5)
                    {
                        if (robj.data[0].inventory[i].items.Length == 6)
                        {
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 1, 2)); // temp for inventory
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 3, 4)); // temp for inventory
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 5, 6)); // temp for inventory
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 7, 8)); // temp for inventory


                            for (int j = 0; j < robj.data[0].inventory[i].items.Length; j++)
                            {
                                if (j == 0)
                                {
                                    dataCheckRow.CreateCell(j).SetCellValue("清點");
                                    dataCheckRow.GetCell(j).CellStyle = InventorycsBL;
                                    if (robj.data[0].inventory[i].items[j].check)
                                    {
                                        dataCheckRow.CreateCell(j + 1).SetCellValue("V");
                                        dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 1).SetCellValue("");
                                        dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                    }
                                    sheet1.SetColumnWidth(j + 1, 10 * 256);
                                    //sheet1.AutoSizeColumn(j);
                                }
                                else if (j == 1)
                                {
                                    if (robj.data[0].inventory[i].items[j].check)
                                    {
                                        dataCheckRow.CreateCell(j + 2).SetCellValue("V");
                                        dataCheckRow.GetCell(j + 2).CellStyle = Inventorycs;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 2).SetCellValue("");
                                        dataCheckRow.GetCell(j + 2).CellStyle = Inventorycs;
                                    }
                                    sheet1.SetColumnWidth(j + 1, 10 * 256);
                                }
                                else if (j == 2)
                                {
                                    if (robj.data[0].inventory[i].items[j].check)
                                    {
                                        dataCheckRow.CreateCell(j + 3).SetCellValue("V");
                                        dataCheckRow.GetCell(j + 3).CellStyle = Inventorycs;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 3).SetCellValue("");
                                        dataCheckRow.GetCell(j + 3).CellStyle = Inventorycs;
                                    }
                                    sheet1.SetColumnWidth(j + 1, 10 * 256);
                                }
                                else if (j == 3)
                                {
                                    if (robj.data[0].inventory[i].items[j].check)
                                    {
                                        dataCheckRow.CreateCell(j + 4).SetCellValue("V");
                                        dataCheckRow.GetCell(j + 4).CellStyle = Inventorycs;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 4).SetCellValue("");
                                        dataCheckRow.GetCell(j + 4).CellStyle = Inventorycs;
                                    }
                                    sheet1.SetColumnWidth(j + 1, 10 * 256);
                                }
                                else if (j == 4)
                                {
                                    if (robj.data[0].inventory[i].items[j].check)
                                    {
                                        dataCheckRow.CreateCell(j + 5).SetCellValue("V");
                                        dataCheckRow.GetCell(j + 5).CellStyle = Inventorycs;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 5).SetCellValue("");
                                        dataCheckRow.GetCell(j + 5).CellStyle = Inventorycs;
                                    }
                                    sheet1.SetColumnWidth(j + 1, 10 * 256);
                                }
                                else
                                {
                                    if (robj.data[0].inventory[i].items[j].check)
                                    {
                                        dataCheckRow.CreateCell(j + 5).SetCellValue("V");
                                        dataCheckRow.GetCell(j + 5).CellStyle = Inventorycs;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 5).SetCellValue("");
                                        dataCheckRow.GetCell(j + 5).CellStyle = Inventorycs;
                                    }
                                    sheet1.SetColumnWidth(j + 1, 10 * 256);
                                }
                                dataCheckRow.GetCell(10).CellStyle = InventorycsR;
                            }
                            rowCounter++;
                        }
                        else if (robj.data[0].inventory[i].items.Length == 7)
                        {
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 1, 2)); // temp for inventory
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 3, 4)); // temp for inventory
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 5, 6)); // temp for inventory                                

                            for (int j = 0; j < robj.data[0].inventory[i].items.Length; j++)
                            {
                                if (j == 0)
                                {
                                    dataCheckRow.CreateCell(j).SetCellValue("清點");
                                    dataCheckRow.GetCell(j).CellStyle = InventorycsBL;
                                    if (robj.data[0].inventory[i].items[j].check)
                                    {
                                        dataCheckRow.CreateCell(j + 1).SetCellValue("V");
                                        dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 1).SetCellValue("");
                                        dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                    }
                                    sheet1.SetColumnWidth(j + 1, 10 * 256);
                                    //sheet1.AutoSizeColumn(j);
                                }
                                else if (j == 1)
                                {
                                    if (robj.data[0].inventory[i].items[j].check)
                                    {
                                        dataCheckRow.CreateCell(j + 2).SetCellValue("V");
                                        dataCheckRow.GetCell(j + 2).CellStyle = Inventorycs;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 2).SetCellValue("");
                                        dataCheckRow.GetCell(j + 2).CellStyle = Inventorycs;
                                    }
                                    sheet1.SetColumnWidth(j + 1, 10 * 256);
                                }
                                else if (j == 2)
                                {
                                    if (robj.data[0].inventory[i].items[j].check)
                                    {
                                        dataCheckRow.CreateCell(j + 3).SetCellValue("V");
                                        dataCheckRow.GetCell(j + 3).CellStyle = Inventorycs;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 3).SetCellValue("");
                                        dataCheckRow.GetCell(j + 3).CellStyle = Inventorycs;
                                    }
                                    sheet1.SetColumnWidth(j + 1, 10 * 256);
                                }
                                else if (j == 3)
                                {
                                    if (robj.data[0].inventory[i].items[j].check)
                                    {
                                        dataCheckRow.CreateCell(j + 4).SetCellValue("V");
                                        dataCheckRow.GetCell(j + 4).CellStyle = Inventorycs;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 4).SetCellValue("");
                                        dataCheckRow.GetCell(j + 4).CellStyle = Inventorycs;
                                    }
                                    sheet1.SetColumnWidth(j + 1, 10 * 256);
                                }
                                else
                                {
                                    if (robj.data[0].inventory[i].items[j].check)
                                    {
                                        dataCheckRow.CreateCell(j + 4).SetCellValue("V");
                                        dataCheckRow.GetCell(j + 4).CellStyle = Inventorycs;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 4).SetCellValue("");
                                        dataCheckRow.GetCell(j + 4).CellStyle = Inventorycs;
                                    }
                                    sheet1.SetColumnWidth(j + 1, 10 * 256);
                                }
                                dataCheckRow.GetCell(10).CellStyle = InventorycsR;
                            }
                            rowCounter++;
                        }
                        else if (robj.data[0].inventory[i].items.Length == 8)
                        {
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 1, 2)); // temp for inventory
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 3, 4)); // temp for inventory                              

                            for (int j = 0; j < robj.data[0].inventory[i].items.Length; j++)
                            {
                                if (j == 0)
                                {
                                    dataCheckRow.CreateCell(j).SetCellValue("清點");
                                    dataCheckRow.GetCell(j).CellStyle = InventorycsBL;
                                    if (robj.data[0].inventory[i].items[j].check)
                                    {
                                        dataCheckRow.CreateCell(j + 1).SetCellValue("V");
                                        dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 1).SetCellValue("");
                                        dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                    }
                                    sheet1.SetColumnWidth(j + 1, 10 * 256);
                                    //sheet1.AutoSizeColumn(j);
                                }
                                else if (j == 1)
                                {
                                    if (robj.data[0].inventory[i].items[j].check)
                                    {
                                        dataCheckRow.CreateCell(j + 2).SetCellValue("V");
                                        dataCheckRow.GetCell(j + 2).CellStyle = Inventorycs;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 2).SetCellValue("");
                                        dataCheckRow.GetCell(j + 2).CellStyle = Inventorycs;
                                    }
                                    sheet1.SetColumnWidth(j + 1, 10 * 256);
                                }
                                else if (j == 2)
                                {
                                    if (robj.data[0].inventory[i].items[j].check)
                                    {
                                        dataCheckRow.CreateCell(j + 3).SetCellValue("V");
                                        dataCheckRow.GetCell(j + 3).CellStyle = Inventorycs;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 3).SetCellValue("");
                                        dataCheckRow.GetCell(j + 3).CellStyle = Inventorycs;
                                    }
                                    sheet1.SetColumnWidth(j + 1, 10 * 256);
                                }
                                else
                                {
                                    if (robj.data[0].inventory[i].items[j].check)
                                    {
                                        dataCheckRow.CreateCell(j + 3).SetCellValue("V");
                                        dataCheckRow.GetCell(j + 3).CellStyle = Inventorycs;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 3).SetCellValue("");
                                        dataCheckRow.GetCell(j + 3).CellStyle = Inventorycs;
                                    }
                                    sheet1.SetColumnWidth(j + 1, 10 * 256);
                                }
                                dataCheckRow.GetCell(10).CellStyle = InventorycsR;
                            }
                            rowCounter++;
                        }
                        else if (robj.data[0].inventory[i].items.Length == 9)
                        {
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 1, 2)); // temp for inventory                                

                            for (int j = 0; j < robj.data[0].inventory[i].items.Length; j++)
                            {
                                if (j == 0)
                                {
                                    dataCheckRow.CreateCell(j).SetCellValue("清點");
                                    dataCheckRow.GetCell(j).CellStyle = InventorycsBL;
                                    if (robj.data[0].inventory[i].items[j].check)
                                    {
                                        dataCheckRow.CreateCell(j + 1).SetCellValue("V");
                                        dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 1).SetCellValue("");
                                        dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                    }
                                    sheet1.SetColumnWidth(j + 1, 10 * 256);
                                    //sheet1.AutoSizeColumn(j);
                                }
                                else if (j == 1)
                                {
                                    if (robj.data[0].inventory[i].items[j].check)
                                    {
                                        dataCheckRow.CreateCell(j + 2).SetCellValue("V");
                                        dataCheckRow.GetCell(j + 2).CellStyle = Inventorycs;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 2).SetCellValue("");
                                        dataCheckRow.GetCell(j + 2).CellStyle = Inventorycs;
                                    }
                                    sheet1.SetColumnWidth(j + 1, 10 * 256);
                                }
                                else
                                {
                                    if (robj.data[0].inventory[i].items[j].check)
                                    {
                                        dataCheckRow.CreateCell(j + 2).SetCellValue("V");
                                        dataCheckRow.GetCell(j + 2).CellStyle = Inventorycs;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 2).SetCellValue("");
                                        dataCheckRow.GetCell(j + 2).CellStyle = Inventorycs;
                                    }
                                    sheet1.SetColumnWidth(j + 1, 10 * 256);
                                }
                                dataCheckRow.GetCell(10).CellStyle = InventorycsR;
                            }
                            rowCounter++;
                        }
                    }
                    else
                    {
                        if (robj.data[0].inventory[i].items.Length == 1)
                        {
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 1, 10)); // temp for inventory

                            for (int j = 0; j < robj.data[0].inventory[i].items.Length; j++)
                            {
                                if (j == 0)
                                {
                                    dataCheckRow.CreateCell(j).SetCellValue("清點");
                                    dataCheckRow.GetCell(j).CellStyle = InventorycsBL;
                                    if (robj.data[0].inventory[i].items[j].check)
                                    {
                                        dataCheckRow.CreateCell(j + 1).SetCellValue("V");
                                        dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 1).SetCellValue("");
                                        dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                    }
                                    sheet1.SetColumnWidth(j + 1, 10 * 256);
                                    //sheet1.AutoSizeColumn(j);
                                }
                                dataCheckRow.GetCell(10).CellStyle = InventorycsR;
                            }
                            rowCounter++;
                        }
                        else if (robj.data[0].inventory[i].items.Length == 2)
                        {
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 1, 5)); // temp for inventory
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 6, 10)); // temp for inventory

                            for (int j = 0; j < robj.data[0].inventory[i].items.Length; j++)
                            {
                                if (j == 0)
                                {
                                    dataCheckRow.CreateCell(j).SetCellValue("清點");
                                    dataCheckRow.GetCell(j).CellStyle = InventorycsBL;
                                    if (robj.data[0].inventory[i].items[j].check)
                                    {
                                        dataCheckRow.CreateCell(j + 1).SetCellValue("V");
                                        dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 1).SetCellValue("");
                                        dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                    }
                                    sheet1.SetColumnWidth(j + 1, 10 * 256);
                                    //sheet1.AutoSizeColumn(j);
                                }
                                else if (j == 1)
                                {
                                    if (robj.data[0].inventory[i].items[j].check)
                                    {
                                        dataCheckRow.CreateCell(j + 5).SetCellValue("V");
                                        dataCheckRow.GetCell(j + 5).CellStyle = Inventorycs;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 5).SetCellValue("");
                                        dataCheckRow.GetCell(j + 5).CellStyle = Inventorycs;
                                    }
                                    sheet1.SetColumnWidth(j + 1, 10 * 256);
                                }
                                dataCheckRow.GetCell(10).CellStyle = InventorycsR;
                            }
                            rowCounter++;
                        }
                        else if (robj.data[0].inventory[i].items.Length == 3)
                        {
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 1, 4)); // temp for inventory
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 5, 7)); // temp for inventory
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 8, 10)); // temp for inventory

                            for (int j = 0; j < robj.data[0].inventory[i].items.Length; j++)
                            {
                                if (j == 0)
                                {
                                    dataCheckRow.CreateCell(j).SetCellValue("清點");
                                    dataCheckRow.GetCell(j).CellStyle = InventorycsBL;
                                    if (robj.data[0].inventory[i].items[j].check)
                                    {
                                        dataCheckRow.CreateCell(j + 1).SetCellValue("V");
                                        dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 1).SetCellValue("");
                                        dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                    }
                                    sheet1.SetColumnWidth(j + 1, 10 * 256);
                                    //sheet1.AutoSizeColumn(j);
                                }
                                else if (j == 1)
                                {
                                    if (robj.data[0].inventory[i].items[j].check)
                                    {
                                        dataCheckRow.CreateCell(j + 4).SetCellValue("V");
                                        dataCheckRow.GetCell(j + 4).CellStyle = Inventorycs;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 4).SetCellValue("");
                                        dataCheckRow.GetCell(j + 4).CellStyle = Inventorycs;
                                    }
                                    sheet1.SetColumnWidth(j + 1, 10 * 256);
                                }
                                else if (j == 2)
                                {
                                    if (robj.data[0].inventory[i].items[j].check)
                                    {
                                        dataCheckRow.CreateCell(j + 6).SetCellValue("V");
                                        dataCheckRow.GetCell(j + 6).CellStyle = Inventorycs;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 6).SetCellValue("");
                                        dataCheckRow.GetCell(j + 6).CellStyle = Inventorycs;
                                    }
                                    sheet1.SetColumnWidth(j + 1, 10 * 256);
                                }
                                dataCheckRow.GetCell(10).CellStyle = InventorycsR;
                            }
                            rowCounter++;
                        }
                        else if (robj.data[0].inventory[i].items.Length == 4)
                        {
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 1, 3)); // temp for inventory
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 4, 5)); // temp for inventory
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 6, 8)); // temp for inventory
                            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 9, 10)); // temp for inventory

                            for (int j = 0; j < robj.data[0].inventory[i].items.Length; j++)
                            {
                                if (j == 0)
                                {
                                    dataCheckRow.CreateCell(j).SetCellValue("清點");
                                    dataCheckRow.GetCell(j).CellStyle = InventorycsBL;
                                    if (robj.data[0].inventory[i].items[j].check)
                                    {
                                        dataCheckRow.CreateCell(j + 1).SetCellValue("V");
                                        dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 1).SetCellValue("");
                                        dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                    }
                                    sheet1.SetColumnWidth(j + 1, 10 * 256);
                                    //sheet1.AutoSizeColumn(j);
                                }
                                else if (j == 1)
                                {
                                    if (robj.data[0].inventory[i].items[j].check)
                                    {
                                        dataCheckRow.CreateCell(j + 3).SetCellValue("V");
                                        dataCheckRow.GetCell(j + 3).CellStyle = Inventorycs;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 3).SetCellValue("");
                                        dataCheckRow.GetCell(j + 3).CellStyle = Inventorycs;
                                    }
                                    sheet1.SetColumnWidth(j + 1, 10 * 256);
                                }
                                else if (j == 2)
                                {
                                    if (robj.data[0].inventory[i].items[j].check)
                                    {
                                        dataCheckRow.CreateCell(j + 4).SetCellValue("V");
                                        dataCheckRow.GetCell(j + 4).CellStyle = Inventorycs;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 4).SetCellValue("");
                                        dataCheckRow.GetCell(j + 4).CellStyle = Inventorycs;
                                    }
                                    sheet1.SetColumnWidth(j + 1, 10 * 256);
                                }
                                else if (j == 3)
                                {
                                    if (robj.data[0].inventory[i].items[j].check)
                                    {
                                        dataCheckRow.CreateCell(j + 6).SetCellValue("V");
                                        dataCheckRow.GetCell(j + 6).CellStyle = Inventorycs;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 6).SetCellValue("");
                                        dataCheckRow.GetCell(j + 6).CellStyle = Inventorycs;
                                    }
                                    sheet1.SetColumnWidth(j + 1, 10 * 256);
                                }
                                dataCheckRow.GetCell(10).CellStyle = InventorycsR;
                            }
                            rowCounter++;
                        }
                    }
                }
                sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 1, 3)); // inventoryLiaison (21, 21, 1, 3)
                sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 4, 5)); // inventoryLiaison (21, 21, 4, 5)
                sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 6, 8)); // inventoryLiaison (21, 21, 6, 8)
                sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 9, 10)); // inventoryLiaison (21, 21, 9, 10)
                HSSFRow dataLiaison = (HSSFRow)sheet1.CreateRow(rowCounter);
                //dataLiaison.Height = 10 * 20;
                for (int j = 0; j < 11; j++)
                {
                    if (j == 0)
                    {
                        dataLiaison.CreateCell(j).SetCellValue("通聯");
                        dataLiaison.GetCell(j).CellStyle = InventorycsBLB;
                    }
                    else if (j == 1)
                    {
                        dataLiaison.CreateCell(j).SetCellValue(robj.data[0].inventory[i].liaison.morning);
                        dataLiaison.GetCell(j).CellStyle = InventorycsB;
                    }
                    else if (j == 4)
                    {
                        dataLiaison.CreateCell(j).SetCellValue(robj.data[0].inventory[i].liaison.afternoon);
                        dataLiaison.GetCell(j).CellStyle = InventorycsB;
                    }
                    else if (j == 6)
                    {
                        dataLiaison.CreateCell(j).SetCellValue(robj.data[0].inventory[i].liaison.evening);
                        dataLiaison.GetCell(j).CellStyle = InventorycsB;
                    }
                    else if (j == 9)
                    {
                        dataLiaison.CreateCell(j).SetCellValue(robj.data[0].inventory[i].liaison.midnight);
                        dataLiaison.GetCell(j).CellStyle = InventorycsB;
                    }
                    else if (j == 10)
                    {
                        dataLiaison.CreateCell(j).SetCellValue("");
                        dataLiaison.GetCell(j).CellStyle = InventorycsRB;
                    }
                    else
                    {
                        dataLiaison.CreateCell(j).SetCellValue("");
                        dataLiaison.GetCell(j).CellStyle = InventorycsB;
                    }
                }
                rowCounter++;
            }

            /*
            for (int i = 0; i < robj.data[0].inventory.Length; i++)
            {
                sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 0, 10)); // inventoryTitle (18, 18, 0, 10)
                HSSFRow inventoryRow = (HSSFRow)sheet1.CreateRow(rowCounter); // start with 18
                                                                              //inventoryRow.Height = 10 * 20;
                inventoryRow.CreateCell(0).SetCellValue(robj.data[0].inventory[i].title);
                inventoryRow.GetCell(0).CellStyle = boldcsLT;
                rowCounter++;
                for (int c = 1; c < 11; c++)
                {
                    if (c == 10)
                    {
                        inventoryRow.CreateCell(c).SetCellValue("");
                        inventoryRow.GetCell(c).CellStyle = boldcsRT;
                    }
                    else
                    {
                        inventoryRow.CreateCell(c).SetCellValue("");
                        inventoryRow.GetCell(c).CellStyle = boldcsT;
                    }
                }

                HSSFRow dataItemRow = (HSSFRow)sheet1.CreateRow(rowCounter);
                //dataItemRow.Height = 10 * 20;
                for (int c = 0; c < 11; c++)
                {
                    if (c == 0)
                    {
                        dataItemRow.CreateCell(c).SetCellValue("");
                        dataItemRow.GetCell(c).CellStyle = InventorycsBL;
                    }
                    else if (c == 10)
                    {
                        dataItemRow.CreateCell(c).SetCellValue("");
                        dataItemRow.GetCell(c).CellStyle = InventorycsR;
                    }
                    dataItemRow.CreateCell(c).SetCellValue("");
                    dataItemRow.GetCell(c).CellStyle = Inventorycs;
                }
                if (i == 1)
                {
                    sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 1, 2)); // temp for inventory
                    sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 3, 4)); // temp for inventory
                    sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 5, 6)); // temp for inventory
                    sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 7, 8)); // temp for inventory

                    for (int j = 0; j < robj.data[0].inventory[i].items.Length; j++)
                    {
                        if (j == 0)
                        {
                            dataItemRow.CreateCell(j).SetCellValue("品名");
                            dataItemRow.GetCell(j).CellStyle = InventorycsBL;
                            dataItemRow.CreateCell(j + 1).SetCellValue(robj.data[0].inventory[i].items[j].item);
                            dataItemRow.GetCell(j + 1).CellStyle = Inventorycs;
                            //sheet1.AutoSizeColumn(j);
                        }
                        else if (j == 1)
                        {
                            dataItemRow.CreateCell(j + 2).SetCellValue(robj.data[0].inventory[i].items[j].item);
                            dataItemRow.GetCell(j + 2).CellStyle = Inventorycs;
                        }
                        else if (j == 2)
                        {
                            dataItemRow.CreateCell(j + 3).SetCellValue(robj.data[0].inventory[i].items[j].item);
                            dataItemRow.GetCell(j + 3).CellStyle = Inventorycs;
                        }
                        else if (j == 3)
                        {
                            dataItemRow.CreateCell(j + 4).SetCellValue(robj.data[0].inventory[i].items[j].item);
                            dataItemRow.GetCell(j + 4).CellStyle = Inventorycs;
                        }
                        else if (j == 4)
                        {
                            dataItemRow.CreateCell(j + 5).SetCellValue(robj.data[0].inventory[i].items[j].item);
                            dataItemRow.GetCell(j + 5).CellStyle = Inventorycs;
                        }
                        else if (j == robj.data[0].inventory[i].items.Length - 1)
                        {
                            dataItemRow.CreateCell(j + 5).SetCellValue(robj.data[0].inventory[i].items[j].item);
                            dataItemRow.GetCell(j + 5).CellStyle = InventorycsR;
                        }
                        /*
                        else
                        {
                            dataItemRow.CreateCell(j + 1).SetCellValue(robj.data[0].inventory[i].items[j].item);
                            dataItemRow.GetCell(j + 1).CellStyle = Inventorycs;
                        }

                    }
                    rowCounter++;
                }
                else
                {
                    for (int j = 0; j < robj.data[0].inventory[i].items.Length; j++)
                    {
                        if (j == 0)
                        {
                            dataItemRow.CreateCell(j).SetCellValue("品名");
                            dataItemRow.GetCell(j).CellStyle = InventorycsBL;
                            dataItemRow.CreateCell(j + 1).SetCellValue(robj.data[0].inventory[i].items[j].item);
                            dataItemRow.GetCell(j + 1).CellStyle = Inventorycs;
                            //sheet1.AutoSizeColumn(j);
                        }

                        else if (j == robj.data[0].inventory[i].items.Length - 1 && j != 9)
                        {
                            dataItemRow.CreateCell(j + 1).SetCellValue(robj.data[0].inventory[i].items[j].item);
                            dataItemRow.GetCell(j + 1).CellStyle = Inventorycs;
                            j++;
                            while (j < 10)
                            {
                                if (j == 9)
                                {
                                    dataItemRow.CreateCell(j + 1).SetCellValue("");
                                    dataItemRow.GetCell(j + 1).CellStyle = InventorycsR;
                                }
                                else
                                {
                                    dataItemRow.CreateCell(j + 1).SetCellValue("");
                                    dataItemRow.GetCell(j + 1).CellStyle = Inventorycs;
                                }
                                j++;
                            }
                        }
                        else if (j == robj.data[0].inventory[i].items.Length - 1)
                        {
                            dataItemRow.CreateCell(j + 1).SetCellValue(robj.data[0].inventory[i].items[j].item);
                            dataItemRow.GetCell(j + 1).CellStyle = InventorycsR;
                        }
                        else
                        {
                            dataItemRow.CreateCell(j + 1).SetCellValue(robj.data[0].inventory[i].items[j].item);
                            dataItemRow.GetCell(j + 1).CellStyle = Inventorycs;
                        }

                    }
                    rowCounter++;
                }
                

                HSSFRow dataCheckRow = (HSSFRow)sheet1.CreateRow(rowCounter);
                //dataCheckRow.Height = 10 * 20;
                for (int c = 0; c < 11; c++)
                {
                    if (c == 0)
                    {
                        dataCheckRow.CreateCell(c).SetCellValue("");
                        dataCheckRow.GetCell(c).CellStyle = InventorycsBL;
                    }
                    else if (c == 10)
                    {
                        dataCheckRow.CreateCell(c).SetCellValue("");
                        dataCheckRow.GetCell(c).CellStyle = InventorycsR;
                    }
                    else
                    {
                        dataCheckRow.CreateCell(c).SetCellValue("");
                        dataCheckRow.GetCell(c).CellStyle = Inventorycs;
                    }
                }
                if (i == 1)
                {
                    sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 1, 2)); // temp for inventory
                    sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 3, 4)); // temp for inventory
                    sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 5, 6)); // temp for inventory
                    sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 7, 8)); // temp for inventory

                    for (int j = 0; j < robj.data[0].inventory[i].items.Length; j++)
                    {
                        if (j == 0)
                        {
                            dataCheckRow.CreateCell(j).SetCellValue("清點");
                            dataCheckRow.GetCell(j).CellStyle = InventorycsBL;
                            if (robj.data[0].inventory[i].items[j].check)
                            {
                                dataCheckRow.CreateCell(j + 1).SetCellValue("V");
                                dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                //sheet1.SetColumnWidth(j, 10 * 256);
                            }
                            else
                            {
                                dataCheckRow.CreateCell(j + 1).SetCellValue("");
                                dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                //sheet1.SetColumnWidth(j, 10 * 256);
                            }
                        }
                        else if (j == 1)
                        {
                            if (robj.data[0].inventory[i].items[j].check)
                            {
                                dataCheckRow.CreateCell(j + 2).SetCellValue("V");
                                dataCheckRow.GetCell(j + 2).CellStyle = Inventorycs;
                                //sheet1.SetColumnWidth(j, 10 * 256);
                            }
                            else
                            {
                                dataCheckRow.CreateCell(j + 2).SetCellValue("");
                                dataCheckRow.GetCell(j + 2).CellStyle = Inventorycs;
                                //sheet1.SetColumnWidth(j, 10 * 256);
                            }
                        }
                        else if (j == 2)
                        {
                            if (robj.data[0].inventory[i].items[j].check)
                            {
                                dataCheckRow.CreateCell(j + 3).SetCellValue("V");
                                dataCheckRow.GetCell(j + 3).CellStyle = Inventorycs;
                                //sheet1.SetColumnWidth(j, 10 * 256);
                            }
                            else
                            {
                                dataCheckRow.CreateCell(j + 3).SetCellValue("");
                                dataCheckRow.GetCell(j + 3).CellStyle = Inventorycs;
                                //sheet1.SetColumnWidth(j, 10 * 256);
                            }
                        }
                        else if (j == 3)
                        {
                            if (robj.data[0].inventory[i].items[j].check)
                            {
                                dataCheckRow.CreateCell(j + 4).SetCellValue("V");
                                dataCheckRow.GetCell(j + 4).CellStyle = Inventorycs;
                                sheet1.SetColumnWidth(j, 10 * 256);
                            }
                            else
                            {
                                dataCheckRow.CreateCell(j + 4).SetCellValue("");
                                dataCheckRow.GetCell(j + 4).CellStyle = Inventorycs;
                                sheet1.SetColumnWidth(j, 10 * 256);
                            }
                        }
                        else if (j == 4)
                        {
                            if (robj.data[0].inventory[i].items[j].check)
                            {
                                dataCheckRow.CreateCell(j + 5).SetCellValue("V");
                                dataCheckRow.GetCell(j + 5).CellStyle = Inventorycs;
                                sheet1.SetColumnWidth(j, 10 * 256);
                            }
                            else
                            {
                                dataCheckRow.CreateCell(j + 5).SetCellValue("");
                                dataCheckRow.GetCell(j + 5).CellStyle = Inventorycs;
                                sheet1.SetColumnWidth(j, 10 * 256);
                            }
                        }
                        else if (j == robj.data[0].inventory[i].items.Length - 1)
                        {
                            if (robj.data[0].inventory[i].items[j].check)
                            {
                                dataCheckRow.CreateCell(j + 5).SetCellValue("V");
                                dataCheckRow.GetCell(j + 5).CellStyle = InventorycsR;
                                sheet1.SetColumnWidth(j, 10 * 256);
                            }
                            else
                            {
                                dataCheckRow.CreateCell(j + 5).SetCellValue("");
                                dataCheckRow.GetCell(j + 5).CellStyle = InventorycsR;
                                sheet1.SetColumnWidth(j, 10 * 256);
                            }
                        }
                        /*
                        else
                        {
                            dataCheckRow.CreateCell(j + 1).SetCellValue(robj.data[0].inventory[i].items[j].check);
                            dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                            sheet1.SetColumnWidth(j, 10 * 256);
                        }
                    }
                    rowCounter++;
                }
                else
                {
                    for (int j = 0; j < robj.data[0].inventory[i].items.Length; j++)
                    {
                        if (j == 0)
                        {
                            dataCheckRow.CreateCell(j).SetCellValue("清點");
                            dataCheckRow.GetCell(j).CellStyle = InventorycsBL;
                            if (robj.data[0].inventory[i].items[j].check)
                            {
                                dataCheckRow.CreateCell(j + 1).SetCellValue("V");
                                dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                sheet1.SetColumnWidth(j, 10 * 256);
                            }
                            else
                            {
                                dataCheckRow.CreateCell(j + 1).SetCellValue("");
                                dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                sheet1.SetColumnWidth(j, 10 * 256);
                            }
                            
                        }
                        else if (j == robj.data[0].inventory[i].items.Length - 1 && j != 9)
                        {
                            if (robj.data[0].inventory[i].items[j].check)
                            {
                                dataCheckRow.CreateCell(j + 1).SetCellValue("V");
                                dataCheckRow.GetCell(j + 1).CellStyle = InventorycsR;
                                sheet1.SetColumnWidth(j + 1, 10 * 256);
                            }
                            else
                            {
                                dataCheckRow.CreateCell(j + 1).SetCellValue("");
                                dataCheckRow.GetCell(j + 1).CellStyle = InventorycsR;
                                sheet1.SetColumnWidth(j + 1, 10 * 256);
                            }
                            
                            while (j < 10)
                            {
                                if (j == 9)
                                {
                                    dataCheckRow.CreateCell(j + 1).SetCellValue("");
                                    dataCheckRow.GetCell(j + 1).CellStyle = InventorycsR;
                                }
                                else
                                {
                                    dataCheckRow.CreateCell(j + 1).SetCellValue("");
                                    dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                }
                                j++;
                            }
                        }
                        else if (j == robj.data[0].inventory[i].items.Length - 1)
                        {
                            if (robj.data[0].inventory[i].items[j].check)
                            {
                                dataCheckRow.CreateCell(j + 1).SetCellValue("V");
                                dataCheckRow.GetCell(j + 1).CellStyle = InventorycsR;
                                sheet1.SetColumnWidth(j, 10 * 256);
                            }
                            else
                            {
                                dataCheckRow.CreateCell(j + 1).SetCellValue("");
                                dataCheckRow.GetCell(j + 1).CellStyle = InventorycsR;
                                sheet1.SetColumnWidth(j, 10 * 256);
                            }
                        }
                        else
                        {
                            if (robj.data[0].inventory[i].items[j].check)
                            {
                                dataCheckRow.CreateCell(j + 1).SetCellValue("V");
                                dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                sheet1.SetColumnWidth(j, 10 * 256);
                            }
                            else
                            {
                                dataCheckRow.CreateCell(j + 1).SetCellValue("");
                                dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                sheet1.SetColumnWidth(j, 10 * 256);
                            }
                        }
                    }
                    rowCounter++;
                }
                

                sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 1, 3)); // inventoryLiaison (21, 21, 1, 3)
                sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 4, 5)); // inventoryLiaison (21, 21, 4, 5)
                sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 6, 8)); // inventoryLiaison (21, 21, 6, 8)
                sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 9, 10)); // inventoryLiaison (21, 21, 9, 10)
                HSSFRow dataLiaison = (HSSFRow)sheet1.CreateRow(rowCounter);
                //dataLiaison.Height = 10 * 20;
                for (int j = 0; j < 11; j++)
                {
                    if (j == 0)
                    {
                        dataLiaison.CreateCell(j).SetCellValue("通聯");
                        dataLiaison.GetCell(j).CellStyle = InventorycsBLB;
                    }
                    else if (j == 1)
                    {
                        dataLiaison.CreateCell(j).SetCellValue(robj.data[0].inventory[i].liaison.morning);
                        dataLiaison.GetCell(j).CellStyle = InventorycsB;
                    }
                    else if (j == 4)
                    {
                        dataLiaison.CreateCell(j).SetCellValue(robj.data[0].inventory[i].liaison.afternoon);
                        dataLiaison.GetCell(j).CellStyle = InventorycsB;
                    }
                    else if (j == 6)
                    {
                        dataLiaison.CreateCell(j).SetCellValue(robj.data[0].inventory[i].liaison.evening);
                        dataLiaison.GetCell(j).CellStyle = InventorycsB;
                    }
                    else if (j == 9)
                    {
                        dataLiaison.CreateCell(j).SetCellValue(robj.data[0].inventory[i].liaison.midnight);
                        dataLiaison.GetCell(j).CellStyle = InventorycsB;
                    }
                    else if (j == 10)
                    {
                        dataLiaison.CreateCell(j).SetCellValue("");
                        dataLiaison.GetCell(j).CellStyle = InventorycsRB;
                    }
                    else
                    {
                        dataLiaison.CreateCell(j).SetCellValue("");
                        dataLiaison.GetCell(j).CellStyle = InventorycsB;
                    }
                }
                rowCounter++;
            }
            */
            /*
            if (robj.data[0].inventory.Length == 10)
            {
            }
            else if (robj.data[0].inventory.Length < 10)
            {
                int itemLength = robj.data[0].inventory.Length;
                if (itemLength % 5 == 0)
                {
                    int merge = 1;
                    for (int i = 0; i < robj.data[0].inventory.Length; i++)
                    {
                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 0, 10)); // inventoryTitle (18, 18, 0, 10)
                        HSSFRow inventoryRow = (HSSFRow)sheet1.CreateRow(rowCounter); // start with 18
                                                                                      //inventoryRow.Height = 10 * 20;
                        inventoryRow.CreateCell(0).SetCellValue(robj.data[0].inventory[i].title);
                        inventoryRow.GetCell(0).CellStyle = boldcsLT;
                        rowCounter++;
                        for (int c = 1; c < 11; c++)
                        {
                            if (c == 10)
                            {
                                inventoryRow.CreateCell(c).SetCellValue("");
                                inventoryRow.GetCell(c).CellStyle = boldcsRT;
                            }
                            else
                            {
                                inventoryRow.CreateCell(c).SetCellValue("");
                                inventoryRow.GetCell(c).CellStyle = boldcsT;
                            }
                        }

                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, merge, merge + 1));
                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, merge + 2, merge + 3));
                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, merge + 4, merge + 5));
                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, merge + 6, merge + 7));
                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, merge + 8, merge + 9));
                        HSSFRow dataItemRow = (HSSFRow)sheet1.CreateRow(rowCounter);
                        //dataItemRow.Height = 10 * 20;
                        for (int c = 0; c < 11; c++)
                        {
                            if (c == 0)
                            {
                                dataItemRow.CreateCell(c).SetCellValue("");
                                dataItemRow.GetCell(c).CellStyle = boldcsL;
                            }
                            else if (c == 10)
                            {
                                dataItemRow.CreateCell(c).SetCellValue("");
                                dataItemRow.GetCell(c).CellStyle = boldcsR;
                            }
                            dataItemRow.CreateCell(c).SetCellValue("");
                            dataItemRow.GetCell(c).CellStyle = cs;
                        }
                        for (int j = 0; j < robj.data[0].inventory[i].items.Length; j++)
                        {
                            if (j == 0)
                            {
                                dataItemRow.CreateCell(j).SetCellValue("品名");
                                dataItemRow.GetCell(j).CellStyle = InventorycsBL;
                                dataItemRow.CreateCell(j + 1).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                dataItemRow.GetCell(j + 1).CellStyle = Inventorycs;
                                //sheet1.AutoSizeColumn(j);
                            }
                            else if (j == 1)
                            {
                                dataItemRow.CreateCell(j + 2).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                dataItemRow.GetCell(j + 2).CellStyle = Inventorycs;
                            }
                            else if (j == 2)
                            {
                                dataItemRow.CreateCell(j + 3).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                dataItemRow.GetCell(j + 3).CellStyle = Inventorycs;
                            }
                            else if (j == 3)
                            {
                                dataItemRow.CreateCell(j + 4).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                dataItemRow.GetCell(j + 4).CellStyle = Inventorycs;
                            }
                            else if (j == 4)
                            {
                                dataItemRow.CreateCell(j + 5).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                dataItemRow.GetCell(j + 5).CellStyle = Inventorycs;
                            }
                            else
                            {
                                dataItemRow.CreateCell(10).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                dataItemRow.GetCell(10).CellStyle = InventorycsR;
                            }

                        }
                        rowCounter++;

                        HSSFRow dataCheckRow = (HSSFRow)sheet1.CreateRow(rowCounter);
                        //dataCheckRow.Height = 10 * 20;
                        for (int c = 0; c < 11; c++)
                        {
                            if (c == 0)
                            {
                                dataCheckRow.CreateCell(c).SetCellValue("");
                                dataCheckRow.GetCell(c).CellStyle = boldcsL;
                            }
                            else if (c == 10)
                            {
                                dataCheckRow.CreateCell(c).SetCellValue("");
                                dataCheckRow.GetCell(c).CellStyle = boldcsR;
                            }
                            else
                            {
                                dataCheckRow.CreateCell(c).SetCellValue("");
                                dataCheckRow.GetCell(c).CellStyle = cs;
                            }
                        }
                        for (int j = 0; j < robj.data[0].inventory[i].items.Length; j++)
                        {
                            if (j == 0)
                            {
                                dataCheckRow.CreateCell(j).SetCellValue("清點");
                                dataCheckRow.GetCell(j).CellStyle = InventorycsBL;
                                dataCheckRow.CreateCell(j + 1).SetCellValue(robj.data[0].inventory[i].items[j].check);
                                dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                sheet1.SetColumnWidth(j, 10 * 256);
                            }
                            else if (j == robj.data[0].inventory[i].items.Length - 1 && j != 9)
                            {
                                dataCheckRow.CreateCell(j + 1).SetCellValue(robj.data[0].inventory[i].items[j].check);
                                dataCheckRow.GetCell(j + 1).CellStyle = InventorycsR;
                                sheet1.SetColumnWidth(j + 1, 10 * 256);
                                while (j < 10)
                                {
                                    if (j == 9)
                                    {
                                        dataCheckRow.CreateCell(j + 1).SetCellValue("");
                                        dataCheckRow.GetCell(j + 1).CellStyle = InventorycsR;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 1).SetCellValue("");
                                        dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                    }
                                    j++;
                                }
                            }
                            else if (j == robj.data[0].inventory[i].items.Length - 1)
                            {
                                dataCheckRow.CreateCell(j + 1).SetCellValue(robj.data[0].inventory[i].items[j].check);
                                dataCheckRow.GetCell(j + 1).CellStyle = InventorycsR;
                                sheet1.SetColumnWidth(j, 10 * 256);
                            }
                            else
                            {
                                dataCheckRow.CreateCell(j + 1).SetCellValue(robj.data[0].inventory[i].items[j].check);
                                dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                sheet1.SetColumnWidth(j, 10 * 256);
                            }
                        }
                        rowCounter++;

                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 1, 3)); // inventoryLiaison (21, 21, 1, 3)
                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 4, 5)); // inventoryLiaison (21, 21, 4, 5)
                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 6, 8)); // inventoryLiaison (21, 21, 6, 8)
                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 9, 10)); // inventoryLiaison (21, 21, 9, 10)
                        HSSFRow dataLiaison = (HSSFRow)sheet1.CreateRow(rowCounter);
                        //dataLiaison.Height = 10 * 20;
                        for (int j = 0; j < 11; j++)
                        {
                            if (j == 0)
                            {
                                dataLiaison.CreateCell(j).SetCellValue("通聯");
                                dataLiaison.GetCell(j).CellStyle = InventorycsBLB;
                            }
                            else if (j == 1)
                            {
                                dataLiaison.CreateCell(j).SetCellValue(robj.data[0].inventory[i].liaison.morning);
                                dataLiaison.GetCell(j).CellStyle = InventorycsB;
                            }
                            else if (j == 4)
                            {
                                dataLiaison.CreateCell(j).SetCellValue(robj.data[0].inventory[i].liaison.afternoon);
                                dataLiaison.GetCell(j).CellStyle = InventorycsB;
                            }
                            else if (j == 6)
                            {
                                dataLiaison.CreateCell(j).SetCellValue(robj.data[0].inventory[i].liaison.evening);
                                dataLiaison.GetCell(j).CellStyle = InventorycsB;
                            }
                            else if (j == 9)
                            {
                                dataLiaison.CreateCell(j).SetCellValue(robj.data[0].inventory[i].liaison.midnight);
                                dataLiaison.GetCell(j).CellStyle = InventorycsB;
                            }
                            else if (j == 10)
                            {
                                dataLiaison.CreateCell(j).SetCellValue("");
                                dataLiaison.GetCell(j).CellStyle = InventorycsRB;
                            }
                            else
                            {
                                dataLiaison.CreateCell(j).SetCellValue("");
                                dataLiaison.GetCell(j).CellStyle = InventorycsB;
                            }
                        }
                        rowCounter++;
                    }
                }
                else if(itemLength > 5)
                {
                    int merge = 1;
                    for (int i = 0; i < robj.data[0].inventory.Length; i++)
                    {
                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 0, 10)); // inventoryTitle (18, 18, 0, 10)
                        HSSFRow inventoryRow = (HSSFRow)sheet1.CreateRow(rowCounter); // start with 18
                                                                                      //inventoryRow.Height = 10 * 20;
                        inventoryRow.CreateCell(0).SetCellValue(robj.data[0].inventory[i].title);
                        inventoryRow.GetCell(0).CellStyle = boldcsLT;
                        rowCounter++;
                        for (int c = 1; c < 11; c++)
                        {
                            if (c == 10)
                            {
                                inventoryRow.CreateCell(c).SetCellValue("");
                                inventoryRow.GetCell(c).CellStyle = boldcsRT;
                            }
                            else
                            {
                                inventoryRow.CreateCell(c).SetCellValue("");
                                inventoryRow.GetCell(c).CellStyle = boldcsT;
                            }
                        }

                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, merge, merge + 1));
                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, merge + 2, merge + 3));
                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, merge + 4, merge + 5));
                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, merge + 6, merge + 7));
                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, merge + 8, merge + 9));
                        HSSFRow dataItemRow = (HSSFRow)sheet1.CreateRow(rowCounter);
                        //dataItemRow.Height = 10 * 20;
                        for (int c = 0; c < 11; c++)
                        {
                            if (c == 0)
                            {
                                dataItemRow.CreateCell(c).SetCellValue("");
                                dataItemRow.GetCell(c).CellStyle = boldcsL;
                            }
                            else if (c == 10)
                            {
                                dataItemRow.CreateCell(c).SetCellValue("");
                                dataItemRow.GetCell(c).CellStyle = boldcsR;
                            }
                            dataItemRow.CreateCell(c).SetCellValue("");
                            dataItemRow.GetCell(c).CellStyle = cs;
                        }
                        for (int j = 0; j < robj.data[0].inventory[i].items.Length; j++)
                        {
                            if (j == 0)
                            {
                                dataItemRow.CreateCell(j).SetCellValue("品名");
                                dataItemRow.GetCell(j).CellStyle = InventorycsBL;
                                dataItemRow.CreateCell(j + 1).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                dataItemRow.GetCell(j + 1).CellStyle = Inventorycs;
                                //sheet1.AutoSizeColumn(j);
                            }
                            else if (j == 1)
                            {
                                dataItemRow.CreateCell(j + 2).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                dataItemRow.GetCell(j + 2).CellStyle = Inventorycs;
                            }
                            else if (j == 2)
                            {
                                dataItemRow.CreateCell(j + 3).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                dataItemRow.GetCell(j + 3).CellStyle = Inventorycs;
                            }
                            else if (j == 3)
                            {
                                dataItemRow.CreateCell(j + 4).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                dataItemRow.GetCell(j + 4).CellStyle = Inventorycs;
                            }
                            else if (j == 4)
                            {
                                dataItemRow.CreateCell(j + 5).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                dataItemRow.GetCell(j + 5).CellStyle = Inventorycs;
                            }
                            else
                            {
                                dataItemRow.CreateCell(10).SetCellValue(robj.data[0].inventory[i].items[j].item);
                                dataItemRow.GetCell(10).CellStyle = InventorycsR;
                            }

                        }
                        rowCounter++;

                        HSSFRow dataCheckRow = (HSSFRow)sheet1.CreateRow(rowCounter);
                        //dataCheckRow.Height = 10 * 20;
                        for (int c = 0; c < 11; c++)
                        {
                            if (c == 0)
                            {
                                dataCheckRow.CreateCell(c).SetCellValue("");
                                dataCheckRow.GetCell(c).CellStyle = boldcsL;
                            }
                            else if (c == 10)
                            {
                                dataCheckRow.CreateCell(c).SetCellValue("");
                                dataCheckRow.GetCell(c).CellStyle = boldcsR;
                            }
                            else
                            {
                                dataCheckRow.CreateCell(c).SetCellValue("");
                                dataCheckRow.GetCell(c).CellStyle = cs;
                            }
                        }
                        for (int j = 0; j < robj.data[0].inventory[i].items.Length; j++)
                        {
                            if (j == 0)
                            {
                                dataCheckRow.CreateCell(j).SetCellValue("清點");
                                dataCheckRow.GetCell(j).CellStyle = InventorycsBL;
                                dataCheckRow.CreateCell(j + 1).SetCellValue(robj.data[0].inventory[i].items[j].check);
                                dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                sheet1.SetColumnWidth(j, 10 * 256);
                            }
                            else if (j == robj.data[0].inventory[i].items.Length - 1 && j != 9)
                            {
                                dataCheckRow.CreateCell(j + 1).SetCellValue(robj.data[0].inventory[i].items[j].check);
                                dataCheckRow.GetCell(j + 1).CellStyle = InventorycsR;
                                sheet1.SetColumnWidth(j + 1, 10 * 256);
                                while (j < 10)
                                {
                                    if (j == 9)
                                    {
                                        dataCheckRow.CreateCell(j + 1).SetCellValue("");
                                        dataCheckRow.GetCell(j + 1).CellStyle = InventorycsR;
                                    }
                                    else
                                    {
                                        dataCheckRow.CreateCell(j + 1).SetCellValue("");
                                        dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                    }
                                    j++;
                                }
                            }
                            else if (j == robj.data[0].inventory[i].items.Length - 1)
                            {
                                dataCheckRow.CreateCell(j + 1).SetCellValue(robj.data[0].inventory[i].items[j].check);
                                dataCheckRow.GetCell(j + 1).CellStyle = InventorycsR;
                                sheet1.SetColumnWidth(j, 10 * 256);
                            }
                            else
                            {
                                dataCheckRow.CreateCell(j + 1).SetCellValue(robj.data[0].inventory[i].items[j].check);
                                dataCheckRow.GetCell(j + 1).CellStyle = Inventorycs;
                                sheet1.SetColumnWidth(j, 10 * 256);
                            }
                        }
                        rowCounter++;

                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 1, 3)); // inventoryLiaison (21, 21, 1, 3)
                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 4, 5)); // inventoryLiaison (21, 21, 4, 5)
                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 6, 8)); // inventoryLiaison (21, 21, 6, 8)
                        sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 9, 10)); // inventoryLiaison (21, 21, 9, 10)
                        HSSFRow dataLiaison = (HSSFRow)sheet1.CreateRow(rowCounter);
                        //dataLiaison.Height = 10 * 20;
                        for (int j = 0; j < 11; j++)
                        {
                            if (j == 0)
                            {
                                dataLiaison.CreateCell(j).SetCellValue("通聯");
                                dataLiaison.GetCell(j).CellStyle = InventorycsBLB;
                            }
                            else if (j == 1)
                            {
                                dataLiaison.CreateCell(j).SetCellValue(robj.data[0].inventory[i].liaison.morning);
                                dataLiaison.GetCell(j).CellStyle = InventorycsB;
                            }
                            else if (j == 4)
                            {
                                dataLiaison.CreateCell(j).SetCellValue(robj.data[0].inventory[i].liaison.afternoon);
                                dataLiaison.GetCell(j).CellStyle = InventorycsB;
                            }
                            else if (j == 6)
                            {
                                dataLiaison.CreateCell(j).SetCellValue(robj.data[0].inventory[i].liaison.evening);
                                dataLiaison.GetCell(j).CellStyle = InventorycsB;
                            }
                            else if (j == 9)
                            {
                                dataLiaison.CreateCell(j).SetCellValue(robj.data[0].inventory[i].liaison.midnight);
                                dataLiaison.GetCell(j).CellStyle = InventorycsB;
                            }
                            else if (j == 10)
                            {
                                dataLiaison.CreateCell(j).SetCellValue("");
                                dataLiaison.GetCell(j).CellStyle = InventorycsRB;
                            }
                            else
                            {
                                dataLiaison.CreateCell(j).SetCellValue("");
                                dataLiaison.GetCell(j).CellStyle = InventorycsB;
                            }
                        }
                        rowCounter++;
                    }
                }
                else if(itemLength < 5)
                {

                }
            }
            */
            #endregion
            sheet1.SetColumnWidth(11, 3 * 256);
            sheet1.SetColumnWidth(12, 3 * 256);
            #region others
            for (int i = 0; i < robj.data[0].others.Length; i++)
            {
                sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 0, 10)); // others (26, 26, 0, 10)
                HSSFRow othersTitleRow = (HSSFRow)sheet1.CreateRow(rowCounter);
                othersTitleRow.Height = 15 * 20;
                othersTitleRow.CreateCell(i).SetCellValue(robj.data[0].others[i].title);
                othersTitleRow.GetCell(i).CellStyle = boldcsLT;
                rowCounter++;
                for (int c = 1; c < 11; c++)
                {
                    if (c == 10)
                    {
                        othersTitleRow.CreateCell(c).SetCellValue("");
                        othersTitleRow.GetCell(c).CellStyle = boldcsRT;
                    }
                    else
                    {
                        othersTitleRow.CreateCell(c).SetCellValue("");
                        othersTitleRow.GetCell(c).CellStyle = boldcsT;
                    }

                }

                sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 0, 10)); // others (27, 27, 0, 10)
                HSSFRow othersDescriptionRow = (HSSFRow)sheet1.CreateRow(rowCounter);
                othersDescriptionRow.Height = 15 * 20;
                othersDescriptionRow.CreateCell(i).SetCellValue(robj.data[0].others[i].description);
                othersDescriptionRow.GetCell(i).CellStyle = otherscsL;
                rowCounter++;
                for (int c = 1; c < 11; c++)
                {
                    if (c == 10)
                    {
                        othersDescriptionRow.CreateCell(c).SetCellValue("");
                        othersDescriptionRow.GetCell(c).CellStyle = otherscsR;
                    }
                    else
                    {
                        othersDescriptionRow.CreateCell(c).SetCellValue("");
                        othersDescriptionRow.GetCell(c).CellStyle = otherscs;
                    }

                }

                sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 0, 10)); // others (28, 28, 0, 10)
                HSSFRow othersRemarkRow = (HSSFRow)sheet1.CreateRow(rowCounter);
                //othersRemarkRow.Height = 10 * 20;
                othersRemarkRow.Height = 15 * 20;
                othersRemarkRow.CreateCell(i).SetCellValue(robj.data[0].others[i].remark);
                othersRemarkRow.GetCell(i).CellStyle = otherscsLB;
                rowCounter++;
                for (int c = 1; c < 11; c++)
                {
                    if (c == 10)
                    {
                        othersRemarkRow.CreateCell(c).SetCellValue("");
                        othersRemarkRow.GetCell(c).CellStyle = otherscsRB;
                    }
                    else
                    {
                        othersRemarkRow.CreateCell(c).SetCellValue("");
                        othersRemarkRow.GetCell(c).CellStyle = otherscsB;
                    }

                }
            }
            #endregion

            /* for(int i = 0; i < robj.data[0].note.Length; i++)
            {
                HSSFRow noteTitleRow = (HSSFRow)sheet1.CreateRow(rowCounter); // 26
                noteTitleRow.CreateCell(i).SetCellValue(robj.data[0].note[i].title);
                noteTitleRow.GetCell(i).CellStyle = cs;
                rowCounter++;
                HSSFRow noteContentRow = (HSSFRow)sheet1.CreateRow(rowCounter); //27
                noteContentRow.CreateCell(i).SetCellValue(robj.data[0].note[i].content);
                noteContentRow.GetCell(i).CellStyle = cs;
                rowCounter++;
            }*/




            //sheet1.AddMergedRegion(new CellRangeAddress(96, 107, 0, 10)); // 臨時交辦事項
            /*for(int i = 29; i < 108; i++)
            {
                sheet1.AddMergedRegion(new CellRangeAddress(i, i, 0, 10));
            }*/

            string T = "<p><strong>壹、每日記事</strong>(摘述人員、時間、地點及執行概況）</p><p style=\"margin-left:20px;\">0200、0400、0600時監督衛兵副哨換哨交接監督，狀況良好。</p><p style=\"margin-left:20px;\">0600：升旗。</p><p style=\"margin-left:20px;\">0800：正、副總值日官交接。</p><p style=\"margin-left:20px;\">0800、1000、1200、1400、1600、1800、2000、2200、2400時監督衛兵副哨換哨交接，狀況良好。</p><p style=\"margin-left:20px;\">0730、0900、1030、1130、1400、1630、1730時段副總值日官實施營區巡查及檢視警監系統狀況良好。</p><p style=\"margin-left:20px;\">1500：電話會談，狀況均正常(啟動碼*62*2261#)&nbsp;</p><figure class=\"table\"><table><tbody><tr><td><p style=\"text-align:center;\">總值日官室</p><p style=\"text-align:center;\">林中校</p></td><td><p style=\"text-align:center;\">營門</p><p style=\"text-align:center;\">周下士</p></td><td><p style=\"text-align:center;\">待命班</p><p style=\"text-align:center;\">吳下士</p></td><td><p style=\"text-align:center;\">戰情中心</p><p style=\"text-align:center;\">王少校</p></td></tr><tr><td><p style=\"text-align:center;\">訓測中心</p><p style=\"text-align:center;\">許少校</p></td><td><p style=\"text-align:center;\">資通聯隊</p><p style=\"text-align:center;\">黃士官長</p></td><td><p style=\"text-align:center;\">研析中心</p><p style=\"text-align:center;\">張中尉</p></td><td><p style=\"text-align:center;\">戰具維管課</p><p style=\"text-align:center;\">游中尉</p></td></tr></tbody></table></figure><p style=\"margin-left:20px;\">1950：高勤官林輝隆中校實施營區巡查，狀況良好。</p><p style=\"margin-left:20px;\">2100：夜間巡查人員實施勤前教育。</p><p style=\"margin-left:20px;\">5月5日忠信營區人數統計：</p><p style=\"margin-left:20px;\">應到：543［指揮部217員、資通聯隊61員、網戰聯隊105員(環境研析中心70技術支援中心35)、訓中76員及支援隊84員］。</p><p style=\"margin-left:20px;\">實到：68［指揮部14員、資通聯隊4員、網戰聯隊14員(環境研析中心7技術支援中心7)、訓中9員及支援隊27員］。</p><p><strong>貳、注意事項</strong></p><p style=\"margin-left:20px;\">一、<span style=\"background-color:hsl(0,0%,90%);\">正副總值日官</span>:</p><p style=\"margin-left:40px;\">1. 熟記報告詞及掌握營區狀況，監看警監畫面及熟稔各類狀況應處作為，參謀長將不定期實施抽點。</p><p style=\"margin-left:40px;\">2. 於平日部隊運動時間前，督導醫務所醫護人員及救護車於16時就位，並完成相關醫護器材準備。</p><p style=\"margin-left:40px;\">3. 每日針對出勤車輛，依「軍車出勤行車安全風險因子檢查表」完成檢核並再次對車長及駕駛予以任務提示，始可放行。</p><p style=\"margin-left:40px;\">4. 負責值日官室周邊環境、設施維護，確維環境整潔、線路整齊、設施妥善，遇損壞狀況立即與業管單位回報。</p><p style=\"margin-left:40px;\">5. 督導衛哨兵上、下哨交接及清槍時，須配戴大盤帽，上哨前檢查衛哨兵服裝儀容及是否按照標準作業程序完成清槍動作，並須全程在旁監督確維安全。</p><p style=\"margin-left:40px;\">6. 各巡查人員查察水塔加壓馬達抽水情形，如有異常以電話通知支援隊駐隊官協處。</p><p style=\"margin-left:40px;\">7. 管制警衛組於平日將升旗臺後方走道LED燈開啟(開啟時間同環場燈時間)、2330時關閉(<span style=\"color:hsl(0,75%,60%);\"><strong>指揮官交辦事項</strong></span>)。</p><p style=\"margin-left:40px;\">8. 督導副控室(225807)，2/1-3/31日1830時、4/1-9/30日1900時、10/1-10/31日1800時、11/1-1/31日1730時開啟環場燈，統一關閉時間2130時，以節約能源(<span style=\"color:hsl(0,75%,60%);\"><strong>指揮官交辦事項</strong></span>)。</p><p style=\"margin-left:40px;\">9. 每日無線電機37C波頻切換依戰情中心通信戰情官通知，候令切換。</p><p style=\"margin-left:40px;\">10. 要求衛兵針對營門進出精神狀態異常者(酒駕、吸毒等)立即回報總值日官及其部隊長掌握管制(<span style=\"color:hsl(0,75%,60%);\"><strong>參謀長交辦事項</strong></span>)。</p><p style=\"margin-left:20px;\">二、<span style=\"background-color:hsl(0,0%,90%);\">總值日官</span>:</p><p style=\"margin-left:40px;\">1. 每日0600、1810時帶隊前往升旗台實施升(降)旗，完畢後整隊帶回；另於0645時至餐廳用餐。</p><p style=\"margin-left:40px;\">2. 掌握營區全般狀況後，遇狀況立即通報高勤官於JUIKER群組回報各級長官知悉。</p><p style=\"margin-left:40px;\">3. 登載高勤官每日不定時巡視營區狀況。</p><p style=\"margin-left:40px;\">4. 確實掌握當日及隔日進入營區施工廠商人數(由戰情官提供)並不定期抽查，於早上回報高勤官(<span style=\"color:hsl(0,75%,60%);\"><strong>指揮官交辦事項</strong></span>)。</p><p style=\"margin-left:40px;\">5. 遇上級督導應立即通報高勤官、戰情官及駐隊官。</p><p style=\"margin-left:20px;\">三、<span style=\"background-color:hsl(0,0%,90%);\">副總值日官</span>:</p><p style=\"margin-left:40px;\">每日0730、1130及1730時巡視營區、抽查部隊集合狀況及人員服裝儀容後，將問題回報總值日官掌握、記錄於交接事項。</p><p><strong>參、臨時交辦事項</strong></p><p style=\"margin-left:40px;\">1. 正、副總值日官人員應落實督促總日官室環境清潔，副總日官於每日下班前完成會客室清潔檢查。</p>";
            string TInfo = T;
            //以防有不正常結尾-->EX:"文字 <a styl"
            //T = Regex.Replace(T, "<[^>]*", "", RegexOptions.IgnoreCase);
            // 先拆字
            List<string> otherContent = new List<string>();
            // paragraph 1
            int e = TInfo.IndexOf("<figure");
            string temp = T.Substring(0, e);
            TInfo = TInfo.Remove(0, e);
            otherContent.Add(temp);
            // Table
            e = TInfo.IndexOf("</figure>");
            temp = TInfo.Substring(0, e);
            TInfo = TInfo.Remove(0, e);
            otherContent.Add(temp);
            // paragraph 2
            e = TInfo.IndexOf("<p><strong>貳、注意事項");
            temp = TInfo.Substring(0, e);
            TInfo = TInfo.Remove(0, e);
            otherContent.Add(temp);
            // paragraph 3
            e = TInfo.IndexOf("<p><strong>參、臨時交辦事項");
            temp = TInfo.Substring(0, e);
            TInfo = TInfo.Remove(0, e);
            otherContent.Add(temp);
            // paragraph 4
            otherContent.Add(TInfo);
            // 區分有無表格，如果有需要將表格拆字
            /*foreach(string temp2 in otherContent)
            {
                temp = temp2;
            }*/
            string tableString = otherContent[1]; // 拉一份表格部分的HTML
            string[] tableArray; // 表格內的字 倆倆一組
            tableString = Regex.Replace(tableString, "</p>", ",", RegexOptions.IgnoreCase);
            int eTable = 0;
            int tableRow = 0;
            int tablecol = 0;
            while (eTable != -1) //eTable < tableString.Length
            {
                eTable = tableString.IndexOf("<tr>", eTable, tableString.Length - eTable);
                if (eTable != -1)
                {
                    tableRow++;
                    eTable++;
                }
            }
            eTable = 0;
            while (eTable != -1)
            {
                eTable = tableString.IndexOf("<td>", eTable, tableString.Length - eTable);
                if (eTable != -1)
                {
                    tablecol++;
                    eTable++;
                }
            }
            tableString = Regex.Replace(tableString, "<[^>]*>|&nbsp;", "", RegexOptions.IgnoreCase);
            List<string> tableList = new List<string>();
            tableArray = Regex.Split(tableString, ",");

            // 每個段落分別去TAG
            for (int i = 0; i < otherContent.Count; i++)
            {
                if (i == 1)
                {
                    continue;
                }
                else
                {
                    otherContent[i] = Regex.Replace(otherContent[i], "</p>", "\n", RegexOptions.IgnoreCase);
                    otherContent[i] = Regex.Replace(otherContent[i], "<p style=\"margin-left:20px;\">", "\t", RegexOptions.IgnoreCase);
                    otherContent[i] = Regex.Replace(otherContent[i], "<p style=\"margin-left:40px;\">", "\t\t", RegexOptions.IgnoreCase);
                    otherContent[i] = Regex.Replace(otherContent[i], "&nbsp;", "\t", RegexOptions.IgnoreCase);
                    otherContent[i] = Regex.Replace(otherContent[i], "<[^>]*>|", "", RegexOptions.IgnoreCase);
                }
            }

            // 幫特定字上紅色
            HSSFRichTextString richText3 = new HSSFRichTextString(otherContent[3]);
            richText3.ApplyFont(0, richText3.Length, Font12);
            eTable = 0;
            while (eTable != -1)
            {
                eTable = otherContent[3].IndexOf("(指揮官交辦事項)", eTable, otherContent[3].Length - eTable);
                if (eTable != -1)
                {
                    richText3.ApplyFont(eTable + 1, eTable + 8, Font12red);
                    eTable++;
                }
            }
            eTable = 0;
            while (eTable != -1)
            {
                eTable = otherContent[3].IndexOf("(參謀長交辦事項)", eTable, otherContent[3].Length - eTable);
                if (eTable != -1)
                {
                    richText3.ApplyFont(eTable + 1, eTable + 8, Font12red);
                    eTable++;
                }
            }
            eTable = 0;
            while (eTable != -1)
            {
                eTable = otherContent[3].IndexOf("(副指揮官交辦事項)", eTable, otherContent[3].Length - eTable);
                if (eTable != -1)
                {
                    richText3.ApplyFont(eTable + 1, eTable + 9, Font12red);
                    eTable++;
                }
            }
            eTable = 0;
            while (eTable != -1)
            {
                eTable = otherContent[3].IndexOf("(政戰主任交辦事項)", eTable, otherContent[3].Length - eTable);
                if (eTable != -1)
                {
                    richText3.ApplyFont(eTable + 1, eTable + 9, Font12red);
                    eTable++;
                }
            }
            eTable = 0;
            while (eTable != -1)
            {
                eTable = otherContent[3].IndexOf("(副參謀長交辦事項)", eTable, otherContent[3].Length - eTable);
                if (eTable != -1)
                {
                    richText3.ApplyFont(eTable + 1, eTable + 9, Font12red);
                    eTable++;
                }
            }
            eTable = 0;
            while (eTable != -1)
            {
                eTable = otherContent[3].IndexOf("(單位主管交辦事項)", eTable, otherContent[3].Length - eTable);
                if (eTable != -1)
                {
                    richText3.ApplyFont(eTable + 1, eTable + 9, Font12red);
                    eTable++;
                }
            }
            // 每項標題上粗體
            HSSFRichTextString richText = new HSSFRichTextString(otherContent[0]);
            richText.ApplyFont(0, richText.Length, Font12);
            eTable = 0;
            int eBoldstart = 0;
            int eBoldend = 0;
            while (eTable != -1)
            {
                eTable = otherContent[0].IndexOf("壹、每日記事", eTable, otherContent[0].Length - eTable);
                //eBoldstart = otherContent[0].IndexOf("<strong>", eBoldstart, otherContent[0].Length - eBoldstart);
                //eBoldend = otherContent[0].IndexOf("</strong>", eBoldend, otherContent[0].Length - eBoldend);
                if (eTable != -1)
                {
                    richText.ApplyFont(eTable, eTable + 6, Font12B);
                    eTable++;
                }
            }
            eTable = 0;
            while (eTable != -1)
            {
                eTable = otherContent[3].IndexOf("貳、注意事項", eTable, otherContent[3].Length - eTable);
                if (eTable != -1)
                {
                    richText3.ApplyFont(eTable, eTable + 6, Font12B);
                    eTable++;
                }
            }
            HSSFRichTextString richText4 = new HSSFRichTextString(otherContent[4]);
            richText4.ApplyFont(0, richText4.Length, Font12);
            eTable = 0;
            while (eTable != -1)
            {
                eTable = otherContent[4].IndexOf("參、臨時交辦事項", eTable, otherContent[4].Length - eTable);
                if (eTable != -1)
                {
                    richText4.ApplyFont(eTable, eTable + 8, Font12B);
                    eTable++;
                }
            }



            HSSFRow noteTitleRow = (HSSFRow)sheet1.CreateRow(rowCounter); // 29
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 0, 10)); // 每日記事
            noteTitleRow.Height = 120 * 20;
            noteTitleRow.CreateCell(0).SetCellValue(richText);
            noteTitleRow.GetCell(0).CellStyle = otherscsL;




            rowCounter++;
            for (int c = 1; c < 11; c++)
            {
                if (c == 10)
                {
                    noteTitleRow.CreateCell(c).SetCellValue("");
                    noteTitleRow.GetCell(c).CellStyle = otherscsR;
                }
                else
                {
                    noteTitleRow.CreateCell(c).SetCellValue("");
                    noteTitleRow.GetCell(c).CellStyle = otherscs;
                }
            }

            //noteContentRow.Height = 10 * 20;
            int jj = 0;
            for (int r = 0; r < tableRow; r++)
            {
                HSSFRow noteContentRow = (HSSFRow)sheet1.CreateRow(rowCounter); //30開始 看表格多長
                noteContentRow.Height = 45 * 20;
                sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 1, 2));
                sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 3, 4));
                sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 5, 6));
                sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 7, 8));
                sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 9, 10));
                noteContentRow.CreateCell(0).SetCellValue("");
                noteContentRow.GetCell(0).CellStyle = otherscsL;
                for (int c = 1; c < 11; c++)
                {
                    if (c == 10)
                    {
                        noteContentRow.CreateCell(c).SetCellValue("");
                        noteContentRow.GetCell(c).CellStyle = otherscsR;
                    }
                    else if (c == 1)
                    {

                        noteContentRow.CreateCell(c).SetCellValue(tableArray[jj] + "\n" + tableArray[jj + 1]);
                        noteContentRow.GetCell(c).CellStyle = othersTablecs;
                        jj += 2;
                    }
                    else if (c == 2)
                    {
                        noteContentRow.CreateCell(c).SetCellValue("");
                        noteContentRow.GetCell(c).CellStyle = othersTablecs;
                    }
                    else if (c == 3)
                    {

                        noteContentRow.CreateCell(c).SetCellValue(tableArray[jj] + "\n" + tableArray[jj + 1]);
                        noteContentRow.GetCell(c).CellStyle = othersTablecs;
                        jj += 2;
                    }
                    else if (c == 4)
                    {
                        noteContentRow.CreateCell(c).SetCellValue("");
                        noteContentRow.GetCell(c).CellStyle = othersTablecs;
                    }
                    else if (c == 5)
                    {

                        noteContentRow.CreateCell(c).SetCellValue(tableArray[jj] + "\n" + tableArray[jj + 1]);
                        noteContentRow.GetCell(c).CellStyle = othersTablecs;
                        jj += 2;
                    }
                    else if (c == 6)
                    {
                        noteContentRow.CreateCell(c).SetCellValue("");
                        noteContentRow.GetCell(c).CellStyle = othersTablecs;
                    }
                    else if (c == 7)
                    {

                        noteContentRow.CreateCell(c).SetCellValue(tableArray[jj] + "\n" + tableArray[jj + 1]);
                        noteContentRow.GetCell(c).CellStyle = othersTablecs;
                        jj += 2;
                    }
                    else if (c == 8)
                    {
                        noteContentRow.CreateCell(c).SetCellValue("");
                        noteContentRow.GetCell(c).CellStyle = othersTablecs;
                    }
                    else
                    {
                        noteContentRow.CreateCell(c).SetCellValue("");
                        noteContentRow.GetCell(c).CellStyle = otherscs;
                    }

                }
                rowCounter++;
            }

            HSSFRow noteContentRowSpare = (HSSFRow)sheet1.CreateRow(rowCounter); // 32
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 0, 10)); // 每日記事
            noteContentRowSpare.Height = 150 * 20;
            noteContentRowSpare.CreateCell(0).SetCellValue(otherContent[2]); //otherContent[2]
            noteContentRowSpare.GetCell(0).CellStyle = otherscsL;
            rowCounter++; // ->33
            for (int c = 1; c < 11; c++)
            {
                if (c == 10)
                {
                    noteContentRowSpare.CreateCell(c).SetCellValue("");
                    noteContentRowSpare.GetCell(c).CellStyle = otherscsR;
                }
                else
                {
                    noteContentRowSpare.CreateCell(c).SetCellValue("");
                    noteContentRowSpare.GetCell(c).CellStyle = otherscs;
                }
            }
            HSSFRow noteTitleRow2 = (HSSFRow)sheet1.CreateRow(rowCounter); // 33
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 0, 10)); // 注意事項
            noteTitleRow2.Height = 440 * 20;
            noteTitleRow2.CreateCell(0).SetCellValue(richText3);
            noteTitleRow2.GetCell(0).CellStyle = otherscsL;
            rowCounter++;
            for (int c = 1; c < 11; c++)
            {
                if (c == 10)
                {
                    noteTitleRow2.CreateCell(c).SetCellValue("");
                    noteTitleRow2.GetCell(c).CellStyle = otherscsR;
                }
                else
                {
                    noteTitleRow2.CreateCell(c).SetCellValue("");
                    noteTitleRow2.GetCell(c).CellStyle = otherscs;
                }

            }

            HSSFRow noteContentRow2 = (HSSFRow)sheet1.CreateRow(rowCounter); //58
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 0, 10)); // 注意事項
            noteContentRow2.Height = 60 * 20;
            noteContentRow2.CreateCell(0).SetCellValue(richText4);
            noteContentRow2.GetCell(0).CellStyle = otherscsL;
            rowCounter++;
            for (int c = 1; c < 11; c++)
            {
                if (c == 10)
                {
                    noteContentRow2.CreateCell(c).SetCellValue("");
                    noteContentRow2.GetCell(c).CellStyle = otherscsR;
                }
                else
                {
                    noteContentRow2.CreateCell(c).SetCellValue("");
                    noteContentRow2.GetCell(c).CellStyle = otherscsL;
                }

            }
            //sheet1.AddMergedRegion(new CellRangeAddress(35, 46, 0, 10));
            // rowCounter = 35;
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 0, 4)); // stamps 擬辦...108-47=61 47-35=12
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter + 1, rowCounter + 1, 0, 1)); // stamps 副總值日官...
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter + 1, rowCounter + 1, 3, 4)); // stamps 總值日官...
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter + 2, rowCounter + 2, 0, 1)); // stamps 蓋印章
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter + 2, rowCounter + 2, 3, 4)); // stamps 蓋印章
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter + 3, rowCounter + 3, 0, 1)); // stamps 蓋印章
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter + 3, rowCounter + 3, 3, 4)); // stamps 蓋印章

            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter + 4, rowCounter + 4, 0, 1)); // stamps 空行->蓋印章
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter + 4, rowCounter + 4, 3, 4)); // stamps 空行->蓋印章

            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter + 5, rowCounter + 5, 0, 1)); // stamps 蓋印章
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter + 5, rowCounter + 5, 3, 4)); // stamps 蓋印章
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter + 6, rowCounter + 6, 0, 1)); // stamps 蓋印章
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter + 6, rowCounter + 6, 3, 4)); // stamps 蓋印章

            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter + 7, rowCounter + 9, 0, 4)); // stamps 
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter + 9, 5, 5)); // stamps 批示
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter + 9, 6, 10)); // stamps 右半邊
            for (int i = rowCounter; i < rowCounter + 10; i++) //已修正+1 97->35
            {
                HSSFRow notedataRow = (HSSFRow)sheet1.CreateRow(i);

                //notedataRow.Height = 10 * 20;
                if (i == 44)
                {
                    for (int c = 0; c < 11; c++)
                    {
                        if (c == 0)
                        {
                            notedataRow.CreateCell(c).SetCellValue("");
                            notedataRow.GetCell(c).CellStyle = otherscsLB;
                        }
                        else if (c == 10)
                        {
                            notedataRow.CreateCell(c).SetCellValue("");
                            notedataRow.GetCell(c).CellStyle = otherscsRB;
                        }
                        else
                        {
                            notedataRow.CreateCell(c).SetCellValue("");
                            notedataRow.GetCell(c).CellStyle = otherscsB;
                        }
                    }
                }
                else
                {
                    for (int c = 0; c < 11; c++)
                    {
                        if (c == 0)
                        {
                            notedataRow.CreateCell(c).SetCellValue("");
                            notedataRow.GetCell(c).CellStyle = boldcsL;
                        }
                        else if (c == 10)
                        {
                            notedataRow.CreateCell(c).SetCellValue("");
                            notedataRow.GetCell(c).CellStyle = boldcsR;
                        }
                        else
                        {
                            notedataRow.CreateCell(c).SetCellValue("");
                            notedataRow.GetCell(c).CellStyle = cs;
                        }
                    }
                }
            }

            HSSFRow stampsRow = (HSSFRow)sheet1.CreateRow(rowCounter); // 108->47->35
            //stampsRow.Height = 10 * 20;
            stampsRow.CreateCell(0).SetCellValue("擬辦：呈請鈞閱");
            stampsRow.GetCell(0).CellStyle = stampscsLT;
            stampsRow.Height = 15 * 20;
            //rowCounter = rowCounter + 10;
            for (int c = 1; c < 11; c++)
            {
                if (c == 10)
                {
                    stampsRow.CreateCell(c).SetCellValue("");
                    stampsRow.GetCell(c).CellStyle = stampscsRT;
                }
                else if (c == 6)
                {
                    stampsRow.CreateCell(c).SetCellValue("閱");
                    stampsRow.GetCell(c).CellStyle = stampscsLTforRead;
                }
                else if (c == 5)
                {
                    stampsRow.CreateCell(c).SetCellValue("批示");
                    stampsRow.GetCell(c).CellStyle = stampscsP4;
                }
                else
                {
                    stampsRow.CreateCell(c).SetCellValue("");
                    stampsRow.GetCell(c).CellStyle = stampscsT;
                }
            }
            rowCounter++;

            HSSFRow stampsRow104 = (HSSFRow)sheet1.CreateRow(rowCounter); // 36
            stampsRow104.CreateCell(0).SetCellValue("副總值日官(交)：");
            stampsRow104.GetCell(0).CellStyle = stampscsL;
            for (int c = 1; c < 11; c++)
            {
                if (c == 5)
                {
                    stampsRow104.CreateCell(c).SetCellValue("");
                    stampsRow104.GetCell(c).CellStyle = stampscsL;
                }
                else if (c == 3)
                {
                    stampsRow104.CreateCell(c).SetCellValue("總值日官(交)：");
                    stampsRow104.GetCell(c).CellStyle = stampscsNone;
                }
                else if (c == 6)
                {
                    stampsRow104.CreateCell(c).SetCellValue("");
                    stampsRow104.GetCell(c).CellStyle = stampscsL;
                }
                else if (c == 10)
                {
                    stampsRow104.CreateCell(c).SetCellValue("");
                    stampsRow104.GetCell(c).CellStyle = stampscsR;
                }
                else
                {
                    stampsRow104.CreateCell(c).SetCellValue("");
                    stampsRow104.GetCell(c).CellStyle = stampscsNone;
                }
            }
            rowCounter++;

            HSSFRow stampsRow105 = (HSSFRow)sheet1.CreateRow(rowCounter); // 37
            stampsRow105.CreateCell(0).SetCellValue("");
            stampsRow105.GetCell(0).CellStyle = stampscsL;
            byte[] imagebytesSubGive = System.IO.File.ReadAllBytes(path + "\\images\\sample1.jpg");
            WriteImageToCell(ref workbook, ref sheet1, ref stampsRow105, imagebytesSubGive, rowCounter, 0);
            for (int c = 1; c < 11; c++)
            {
                if (c == 1)
                {
                    stampsRow105.CreateCell(c).SetCellValue("");
                    stampsRow105.GetCell(c).CellStyle = stampscsNone;

                }
                else if (c == 2)
                {
                    stampsRow105.CreateCell(c).SetCellValue("");
                    stampsRow105.GetCell(c).CellStyle = stampscsNone;
                    byte[] imagebytesMainGive = System.IO.File.ReadAllBytes(path + "\\images\\sample2.jpg");
                    WriteImageToCell2(ref workbook, ref sheet1, ref stampsRow105, imagebytesMainGive, rowCounter, c);
                }
                else if (c == 4)
                {
                    stampsRow105.CreateCell(c).SetCellValue("");
                    stampsRow105.GetCell(c).CellStyle = stampscsNone;
                }
                else if (c == 5)
                {
                    stampsRow105.CreateCell(c).SetCellValue("");
                    stampsRow105.GetCell(c).CellStyle = stampscsL;
                }
                else if (c == 6)
                {
                    stampsRow105.CreateCell(c).SetCellValue("");
                    stampsRow105.GetCell(c).CellStyle = stampscsL;
                }
                else if (c == 10)
                {
                    stampsRow105.CreateCell(c).SetCellValue("");
                    stampsRow105.GetCell(c).CellStyle = stampscsR;
                }
                else
                {
                    stampsRow105.CreateCell(c).SetCellValue("");
                    stampsRow105.GetCell(c).CellStyle = stampscsNone;
                }
            }
            rowCounter++;

            HSSFRow stampsRow107 = (HSSFRow)sheet1.CreateRow(rowCounter); // 38
            stampsRow107.CreateCell(0).SetCellValue("");
            stampsRow107.GetCell(0).CellStyle = stampscsL;
            for (int c = 1; c < 11; c++)
            {
                if (c == 5)
                {
                    stampsRow107.CreateCell(c).SetCellValue("");
                    stampsRow107.GetCell(c).CellStyle = stampscsL;
                }
                else if (c == 6)
                {
                    stampsRow107.CreateCell(c).SetCellValue("");
                    stampsRow107.GetCell(c).CellStyle = stampscsL;

                }
                else if (c == 10)
                {
                    stampsRow107.CreateCell(c).SetCellValue("");
                    stampsRow107.GetCell(c).CellStyle = stampscsR;
                }
                else
                {
                    stampsRow107.CreateCell(c).SetCellValue("");
                    stampsRow107.GetCell(c).CellStyle = stampscsNone;
                }
            }
            rowCounter++;

            HSSFRow stampsRow108 = (HSSFRow)sheet1.CreateRow(rowCounter); // 39 main/sub revieve
            stampsRow108.CreateCell(0).SetCellValue("副總值日官(接)：");
            stampsRow108.GetCell(0).CellStyle = stampscsL;
            for (int c = 1; c < 11; c++)
            {
                if (c == 1)
                {
                    stampsRow108.CreateCell(c).SetCellValue("");
                    stampsRow108.GetCell(c).CellStyle = stampscsL;
                }
                else if (c == 3)
                {
                    stampsRow108.CreateCell(c).SetCellValue("總值日官(接)：");
                    stampsRow108.GetCell(c).CellStyle = stampscsNone;
                }
                else if (c == 4)
                {
                    stampsRow108.CreateCell(c).SetCellValue("");
                    stampsRow108.GetCell(c).CellStyle = stampscsL;
                }
                else if (c == 5)
                {
                    stampsRow108.CreateCell(c).SetCellValue("");
                    stampsRow108.GetCell(c).CellStyle = stampscsL;
                }
                else if (c == 6)
                {
                    stampsRow108.CreateCell(c).SetCellValue("");
                    stampsRow108.GetCell(c).CellStyle = stampscsL;
                }
                else if (c == 7)
                {
                    stampsRow108.CreateCell(c).SetCellValue("");
                    stampsRow108.GetCell(c).CellStyle = stampscsL;
                    byte[] imagebytesRead = System.IO.File.ReadAllBytes(path + "\\images\\sample4.jpg");
                    WriteImageToCellRead(ref workbook, ref sheet1, ref stampsRow108, imagebytesRead, rowCounter, c);
                }
                else if (c == 10)
                {
                    stampsRow108.CreateCell(c).SetCellValue("");
                    stampsRow108.GetCell(c).CellStyle = stampscsR;
                }
                else
                {
                    stampsRow108.CreateCell(c).SetCellValue("");
                    stampsRow108.GetCell(c).CellStyle = stampscsNone;
                }
            }
            rowCounter++;

            HSSFRow stampsRow109 = (HSSFRow)sheet1.CreateRow(rowCounter); // 40
            stampsRow109.CreateCell(0).SetCellValue("");
            stampsRow109.GetCell(0).CellStyle = stampscsL;
            byte[] imagebytesSubecieve = System.IO.File.ReadAllBytes(path + "\\images\\sample3.jpg");
            WriteImageToCell(ref workbook, ref sheet1, ref stampsRow109, imagebytesSubecieve, rowCounter, 0);
            for (int c = 1; c < 11; c++)
            {
                if (c == 1)
                {
                    stampsRow109.CreateCell(c).SetCellValue("");
                    stampsRow109.GetCell(c).CellStyle = stampscsNone;
                }
                else if (c == 2)
                {
                    stampsRow109.CreateCell(c).SetCellValue("");
                    stampsRow109.GetCell(c).CellStyle = stampscsNone;
                    byte[] imagebytesMainGive = System.IO.File.ReadAllBytes(path + "\\images\\sample4.jpg");
                    WriteImageToCell2(ref workbook, ref sheet1, ref stampsRow109, imagebytesMainGive, rowCounter, c);
                }
                else if (c == 4)
                {
                    stampsRow109.CreateCell(c).SetCellValue("");
                    stampsRow109.GetCell(c).CellStyle = stampscsNone;
                }
                else if (c == 5)
                {
                    stampsRow109.CreateCell(c).SetCellValue("");
                    stampsRow109.GetCell(c).CellStyle = stampscsL;
                }
                else if (c == 6)
                {
                    stampsRow109.CreateCell(c).SetCellValue("");
                    stampsRow109.GetCell(c).CellStyle = stampscsL;
                }

                else if (c == 10)
                {
                    stampsRow109.CreateCell(c).SetCellValue("");
                    stampsRow109.GetCell(c).CellStyle = stampscsR;
                }
                else
                {
                    stampsRow109.CreateCell(c).SetCellValue("");
                    stampsRow109.GetCell(c).CellStyle = stampscsNone;
                }
            }
            rowCounter++;

            HSSFRow stampsRow111 = (HSSFRow)sheet1.CreateRow(rowCounter); // 41
            stampsRow111.CreateCell(0).SetCellValue("");
            stampsRow111.GetCell(0).CellStyle = stampscsL;
            for (int c = 1; c < 11; c++)
            {
                if (c == 1)
                {
                    stampsRow111.CreateCell(c).SetCellValue("");
                    stampsRow111.GetCell(c).CellStyle = stampscsNone;
                }
                else if (c == 3)
                {
                    stampsRow111.CreateCell(c).SetCellValue("");
                    stampsRow111.GetCell(c).CellStyle = stampscsNone;
                }
                else if (c == 4)
                {
                    stampsRow111.CreateCell(c).SetCellValue("");
                    stampsRow111.GetCell(c).CellStyle = stampscsNone;
                }
                else if (c == 5)
                {
                    stampsRow111.CreateCell(c).SetCellValue("");
                    stampsRow111.GetCell(c).CellStyle = stampscsL;
                }
                else if (c == 6)
                {
                    stampsRow111.CreateCell(c).SetCellValue("");
                    stampsRow111.GetCell(c).CellStyle = stampscsL;
                }
                else if (c == 10)
                {
                    stampsRow111.CreateCell(c).SetCellValue("");
                    stampsRow111.GetCell(c).CellStyle = stampscsR;
                }
                else
                {
                    stampsRow111.CreateCell(c).SetCellValue("");
                    stampsRow111.GetCell(c).CellStyle = stampscsNone;
                }
            }
            rowCounter++;

            HSSFRow stampsRow112 = (HSSFRow)sheet1.CreateRow(rowCounter); // 42
            stampsRow112.CreateCell(0).SetCellValue("");
            stampsRow112.GetCell(0).CellStyle = stampscsL;
            for (int c = 1; c < 11; c++)
            {
                if (c == 1)
                {
                    stampsRow112.CreateCell(c).SetCellValue("");
                    stampsRow112.GetCell(c).CellStyle = stampscsNone;
                }
                else if (c == 3)
                {
                    stampsRow112.CreateCell(c).SetCellValue("");
                    stampsRow112.GetCell(c).CellStyle = stampscsNone;
                }
                else if (c == 4)
                {
                    stampsRow112.CreateCell(c).SetCellValue("");
                    stampsRow112.GetCell(c).CellStyle = stampscsNone;
                }
                else if (c == 5)
                {
                    stampsRow112.CreateCell(c).SetCellValue("");
                    stampsRow112.GetCell(c).CellStyle = stampscsL;
                }
                else if (c == 6)
                {
                    stampsRow112.CreateCell(c).SetCellValue("");
                    stampsRow112.GetCell(c).CellStyle = stampscsL;
                }
                else if (c == 10)
                {
                    stampsRow112.CreateCell(c).SetCellValue("");
                    stampsRow112.GetCell(c).CellStyle = stampscsR;
                }
                else
                {
                    stampsRow112.CreateCell(c).SetCellValue("");
                    stampsRow112.GetCell(c).CellStyle = stampscsNone;
                }
            }
            rowCounter++;

            HSSFRow stampsRow116 = (HSSFRow)sheet1.CreateRow(rowCounter); // 43
            stampsRow116.CreateCell(0).SetCellValue("");
            stampsRow116.GetCell(0).CellStyle = stampscsL;
            for (int c = 1; c < 11; c++)
            {
                if (c == 1)
                {
                    stampsRow116.CreateCell(c).SetCellValue("");
                    stampsRow116.GetCell(c).CellStyle = stampscsNone;
                }
                else if (c == 3)
                {
                    stampsRow116.CreateCell(c).SetCellValue("");
                    stampsRow116.GetCell(c).CellStyle = stampscsNone;
                }
                else if (c == 4)
                {
                    stampsRow116.CreateCell(c).SetCellValue("");
                    stampsRow116.GetCell(c).CellStyle = stampscsNone;
                }
                else if (c == 5)
                {
                    stampsRow116.CreateCell(c).SetCellValue("");
                    stampsRow116.GetCell(c).CellStyle = stampscsL;
                }
                else if (c == 6)
                {
                    stampsRow116.CreateCell(c).SetCellValue("");
                    stampsRow116.GetCell(c).CellStyle = stampscsL;
                }
                else if (c == 10)
                {
                    stampsRow116.CreateCell(c).SetCellValue("");
                    stampsRow116.GetCell(c).CellStyle = stampscsR;
                }
                else
                {
                    stampsRow116.CreateCell(c).SetCellValue("");
                    stampsRow116.GetCell(c).CellStyle = stampscsNone;
                }
            }
            rowCounter++;

            HSSFRow stampsRow113 = (HSSFRow)sheet1.CreateRow(rowCounter); // 44
            stampsRow113.CreateCell(0).SetCellValue("");
            stampsRow113.GetCell(0).CellStyle = stampscsLB;
            for (int c = 1; c < 11; c++)
            {
                if (c == 5)
                {
                    stampsRow113.CreateCell(c).SetCellValue("");
                    stampsRow113.GetCell(c).CellStyle = stampscsLB;
                }
                else if (c == 6)
                {
                    stampsRow113.CreateCell(c).SetCellValue("");
                    stampsRow113.GetCell(c).CellStyle = stampscsLB;
                }
                else if (c == 10)
                {
                    stampsRow113.CreateCell(c).SetCellValue("");
                    stampsRow113.GetCell(c).CellStyle = stampscsRB;
                }
                else
                {
                    stampsRow113.CreateCell(c).SetCellValue("");
                    stampsRow113.GetCell(c).CellStyle = stampscsB;
                }
            }
            rowCounter++;

            /*
            for (int i = 104; i < 113; i++)
            {
                HSSFRow stampsDataRow = (HSSFRow)sheet1.CreateRow(i);
                //stampsDataRow.Height = 10 * 20;
                if (i == 104)
                {
                    for (int c = 0; c < 11; c++)
                    {
                        if (c == 0)
                        {
                            stampsDataRow.CreateCell(c).SetCellValue("副總值日官：");
                            stampsDataRow.GetCell(c).CellStyle = stampscsL;
                        }
                        else if (c == 5)
                        {
                            stampsDataRow.CreateCell(c).SetCellValue("");
                            stampsDataRow.GetCell(c).CellStyle = stampscsL;
                        }
                        else if (c == 6)
                        {
                            stampsDataRow.CreateCell(c).SetCellValue("");
                            stampsDataRow.GetCell(c).CellStyle = stampscsL;
                        }
                        else if (c == 10)
                        {
                            stampsDataRow.CreateCell(c).SetCellValue("");
                            stampsDataRow.GetCell(c).CellStyle = stampscsR;
                        }
                        else
                        {
                            stampsDataRow.CreateCell(c).SetCellValue("");
                            stampsDataRow.GetCell(c).CellStyle = stampscsNone;
                        }
                    }
                }
                else if (i == 105)
                {
                    for (int c = 0; c < 11; c++)
                    {
                        if (c == 0)
                        {
                            stampsDataRow.CreateCell(c).SetCellValue("");
                            stampsDataRow.GetCell(c).CellStyle = stampscsLrec;
                        }
                        else if (c == 1)
                        {
                            stampsDataRow.CreateCell(c).SetCellValue("");
                            stampsDataRow.GetCell(c).CellStyle = stampscsLrec2;
                        }
                        else if (c == 2)
                        {
                            stampsDataRow.CreateCell(c).SetCellValue("");
                            stampsDataRow.GetCell(c).CellStyle = stampscsNone; 
                        }
                        else if (c == 3)
                        {
                            stampsDataRow.CreateCell(c).SetCellValue("");
                            stampsDataRow.GetCell(c).CellStyle = stampscsRrec;
                        }
                        else if (c == 4)
                        {
                            stampsDataRow.CreateCell(c).SetCellValue("");
                            stampsDataRow.GetCell(c).CellStyle = stampscsRrec2;
                        }
                        else if (c == 5)
                        {
                            stampsDataRow.CreateCell(c).SetCellValue("");
                            stampsDataRow.GetCell(c).CellStyle = stampscsL;
                        }
                        else if (c == 6)
                        {
                            stampsDataRow.CreateCell(c).SetCellValue("");
                            stampsDataRow.GetCell(c).CellStyle = stampscsL;
                        }
                        else if (c == 10)
                        {
                            stampsDataRow.CreateCell(c).SetCellValue("");
                            stampsDataRow.GetCell(c).CellStyle = stampscsR;
                        }
                        else
                        {
                            stampsDataRow.CreateCell(c).SetCellValue("");
                            stampsDataRow.GetCell(c).CellStyle = stampscsNone;
                        }
                    }
                }
                else if (i == 112)
                {
                    for (int c = 0; c < 11; c++)
                    {
                        if (c == 0)
                        {
                            stampsDataRow.CreateCell(c).SetCellValue("");
                            stampsDataRow.GetCell(c).CellStyle = stampscsLB;
                        }
                        else if (c == 6)
                        {
                            stampsDataRow.CreateCell(c).SetCellValue("");
                            stampsDataRow.GetCell(c).CellStyle = stampscsLB;
                        }
                        else if (c == 5)
                        {
                            stampsDataRow.CreateCell(c).SetCellValue("");
                            stampsDataRow.GetCell(c).CellStyle = stampscsLB;
                        }
                        else if (c == 10)
                        {
                            stampsDataRow.CreateCell(c).SetCellValue("");
                            stampsDataRow.GetCell(c).CellStyle = stampscsRB;
                        }
                        else
                        {
                            stampsDataRow.CreateCell(c).SetCellValue("");
                            stampsDataRow.GetCell(c).CellStyle = stampscsB;
                        }
                    }
                }
                else
                {
                    for (int c = 0; c < 11; c++)
                    {
                        if (c == 0)
                        {
                            stampsDataRow.CreateCell(c).SetCellValue("");
                            stampsDataRow.GetCell(c).CellStyle = stampscsL;
                        }
                        if (c == 5)
                        {
                            stampsDataRow.CreateCell(c).SetCellValue("");
                            stampsDataRow.GetCell(c).CellStyle = stampscsL;
                        }
                        if (c == 6)
                        {
                            stampsDataRow.CreateCell(c).SetCellValue("");
                            stampsDataRow.GetCell(c).CellStyle = stampscsL;
                        }
                        if (c == 10)
                        {
                            stampsDataRow.CreateCell(c).SetCellValue("");
                            stampsDataRow.GetCell(c).CellStyle = stampscsR;
                        }
                        else
                        {
                            stampsDataRow.CreateCell(c).SetCellValue("");
                            stampsDataRow.GetCell(c).CellStyle = stampscsNone;
                        }
                    }
                } 
            }
            */

            rowCounter = rowCounter + 18;
            // rowCounter = 63;
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 0, 10)); // pioTitle
            //sheet1.AddMergedRegion(new CellRangeAddress(123, 124, 0, 10)); // pioTitle
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter + 1, rowCounter + 1, 0, 2)); // "日期"
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter + 1, rowCounter + 1, 3, 10)); // pioDate

            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter + 2, rowCounter + 2, 3, 4)); // pioFirmLeader&Tel
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter + 2, rowCounter + 2, 6, 7)); // pioTitle
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter + 2, rowCounter + 2, 8, 9)); // pioOversee&Tel

            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter + 3, rowCounter + 3, 3, 4)); // pioFirmLeader&Tel
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter + 3, rowCounter + 3, 6, 7)); // pioTitle
            sheet1.AddMergedRegion(new CellRangeAddress(rowCounter + 3, rowCounter + 3, 8, 9)); // pioOversee&Tel

            HSSFRow passitoutRow = (HSSFRow)sheet1.CreateRow(rowCounter); // 121->126
            passitoutRow.Height = 45 * 20;
            passitoutRow.CreateCell(0).SetCellValue("國防部參謀本部資通電軍指揮部\n忠信營區每日廠商進出營區管制紀錄簿");
            passitoutRow.GetCell(0).CellStyle = PioTitlecs;
            for (int c = 1; c < 11; c++)
            {
                passitoutRow.CreateCell(c).SetCellValue("");
                passitoutRow.GetCell(c).CellStyle = cs;
            }
            rowCounter++;
            /*
            HSSFRow passitoutRow2 = (HSSFRow)sheet1.CreateRow(rowCounter + 1);
            passitoutRow2.CreateCell(rowCounter).SetCellValue("");//國防部參謀本部資通電軍指揮部忠信營區總值日官紀事簿
            passitoutRow2.GetCell(rowCounter).CellStyle = cs;
            //titleRow.Height = 20 * 20;
            for (int c = 1; c < 11; c++)
            {
                if (c == 10)
                {
                    passitoutRow2.CreateCell(c).SetCellValue("");
                    passitoutRow2.GetCell(c).CellStyle = cs;
                }
                passitoutRow2.CreateCell(c).SetCellValue("");
                passitoutRow2.GetCell(c).CellStyle = cs;
            }

            HSSFRow passitoutRow3 = (HSSFRow)sheet1.CreateRow(rowCounter + 2);
            passitoutRow3.CreateCell(rowCounter).SetCellValue("");//國防部參謀本部資通電軍指揮部忠信營區總值日官紀事簿
            passitoutRow3.GetCell(rowCounter).CellStyle = cs;
            //titleRow.Height = 20 * 20;
            for (int c = 1; c < 11; c++)
            {
                if (c == 10)
                {
                    passitoutRow3.CreateCell(c).SetCellValue("");
                    passitoutRow3.GetCell(c).CellStyle = cs;
                }
                passitoutRow3.CreateCell(c).SetCellValue("");
                passitoutRow3.GetCell(c).CellStyle = cs;
            }

            HSSFRow passitoutRow4 = (HSSFRow)sheet1.CreateRow(rowCounter + 3);
            passitoutRow4.CreateCell(rowCounter).SetCellValue("");//國防部參謀本部資通電軍指揮部忠信營區總值日官紀事簿
            passitoutRow4.GetCell(rowCounter).CellStyle = cs;
            //titleRow.Height = 20 * 20;
            for (int c = 1; c < 11; c++)
            {
                if (c == 10)
                {
                    passitoutRow4.CreateCell(c).SetCellValue("");
                    passitoutRow4.GetCell(c).CellStyle = cs;
                }
                passitoutRow4.CreateCell(c).SetCellValue("");
                passitoutRow4.GetCell(c).CellStyle = cs;
            }
            rowCounter = rowCounter + 4;
            */

            HSSFRow pioDateRow = (HSSFRow)sheet1.CreateRow(rowCounter); // 125->127
            pioDateRow.Height = 20 * 20;
            for (int i = 0; i < 11; i++)
            {
                if (i == 0)
                {
                    pioDateRow.CreateCell(i).SetCellValue("日期");
                    pioDateRow.GetCell(i).CellStyle = cs;
                    //sheet1.AutoSizeColumn(i);
                }
                else if (i == 3)
                {
                    pioDateRow.CreateCell(i).SetCellValue(datetime);
                    pioDateRow.GetCell(i).CellStyle = cs;
                    //sheet1.AutoSizeColumn(i);
                }
                else
                {
                    pioDateRow.CreateCell(i).SetCellValue("");
                    pioDateRow.GetCell(i).CellStyle = cs;
                }

            }
            rowCounter++;
            HSSFRow pioDataRow = (HSSFRow)sheet1.CreateRow(rowCounter); // 128
            pioDataRow.Height = 40 * 20;

            for (int i = 0; i < 11; i++)
            {
                if (i == 0)
                {
                    pioDataRow.CreateCell(i).SetCellValue("項次");
                    pioDataRow.GetCell(i).CellStyle = cs;
                    //sheet1.AutoSizeColumn(i);
                }
                if (i == 1)
                {
                    pioDataRow.CreateCell(i).SetCellValue("單位");
                    pioDataRow.GetCell(i).CellStyle = cs;
                    //sheet1.AutoSizeColumn(i);
                }
                if (i == 2)
                {
                    pioDataRow.CreateCell(i).SetCellValue("地點");
                    pioDataRow.GetCell(i).CellStyle = cs;
                    //sheet1.AutoSizeColumn(i);
                }
                if (i == 3)
                {
                    pioDataRow.CreateCell(i).SetCellValue("廠商帶隊人員\n(電話)");
                    pioDataRow.GetCell(i).CellStyle = cs;
                    //sheet1.AutoSizeColumn(i);
                }
                if (i == 4)
                {
                    pioDataRow.CreateCell(i).SetCellValue("");
                    pioDataRow.GetCell(i).CellStyle = cs;
                    //sheet1.AutoSizeColumn(i);
                }
                if (i == 5)
                {
                    pioDataRow.CreateCell(i).SetCellValue("人數");
                    pioDataRow.GetCell(i).CellStyle = cs;
                    //sheet1.AutoSizeColumn(i);
                }
                if (i == 6)
                {
                    pioDataRow.CreateCell(i).SetCellValue("工程項目");
                    pioDataRow.GetCell(i).CellStyle = cs;
                    //sheet1.AutoSizeColumn(i);
                }
                if (i == 7)
                {
                    pioDataRow.CreateCell(i).SetCellValue("");
                    pioDataRow.GetCell(i).CellStyle = cs;
                    //sheet1.AutoSizeColumn(i);
                }
                if (i == 8)
                {
                    pioDataRow.CreateCell(i).SetCellValue("監工人員\n(電話)");
                    pioDataRow.GetCell(i).CellStyle = cs;
                    //sheet1.AutoSizeColumn(i);
                }
                if (i == 9)
                {
                    pioDataRow.CreateCell(i).SetCellValue("");
                    pioDataRow.GetCell(i).CellStyle = cs;
                    //sheet1.AutoSizeColumn(i);
                }
                if (i == 10)
                {
                    pioDataRow.CreateCell(i).SetCellValue("備考");
                    pioDataRow.GetCell(i).CellStyle = cs;
                    //sheet1.AutoSizeColumn(i);
                }
            }
            rowCounter++;

            for (int i = 0; i < robj.data[0].pass_in_out.Length; i++)
            {
                HSSFRow dataRow = (HSSFRow)sheet1.CreateRow(i + rowCounter); // start from 129
                //dataRow.Height = 10 * 20;
                for (int j = 0; j < 11; j++)
                {
                    if (j == 0)
                    {
                        int num = i + 1;
                        dataRow.CreateCell(j).SetCellValue(num);
                        dataRow.GetCell(j).CellStyle = cs;
                    }
                    else if (j == 1)
                    {
                        dataRow.CreateCell(j).SetCellValue(robj.data[0].pass_in_out[i].unit);
                        dataRow.GetCell(j).CellStyle = cs;
                    }
                    else if (j == 2)
                    {
                        dataRow.CreateCell(j).SetCellValue(robj.data[0].pass_in_out[i].place);
                        dataRow.GetCell(j).CellStyle = cs;
                    }
                    else if (j == 3)
                    {
                        dataRow.CreateCell(j).SetCellValue(robj.data[0].pass_in_out[i].firm_leader.name + robj.data[0].pass_in_out[i].firm_leader.tel);
                        dataRow.GetCell(j).CellStyle = cs;

                    }
                    else if (j == 4)
                    {
                        dataRow.CreateCell(j).SetCellValue("");
                        dataRow.GetCell(j).CellStyle = cs;

                    }
                    else if (j == 5)
                    {
                        dataRow.CreateCell(j).SetCellValue(robj.data[0].pass_in_out[i].amount);
                        dataRow.GetCell(j).CellStyle = cs;
                    }
                    else if (j == 6)
                    {
                        dataRow.CreateCell(j).SetCellValue(robj.data[0].pass_in_out[i].work);
                        dataRow.GetCell(j).CellStyle = cs;
                    }
                    else if (j == 7)
                    {
                        dataRow.CreateCell(j).SetCellValue("");
                        dataRow.GetCell(j).CellStyle = cs;

                    }
                    else if (j == 8)
                    {
                        dataRow.CreateCell(j).SetCellValue(robj.data[0].pass_in_out[i].oversee.name + robj.data[0].pass_in_out[i].oversee.tel);
                        dataRow.GetCell(j).CellStyle = cs;
                    }
                    else if (j == 9)
                    {
                        dataRow.CreateCell(j).SetCellValue("");
                        dataRow.GetCell(j).CellStyle = cs;

                    }
                    else if (j == 10)
                    {
                        dataRow.CreateCell(j).SetCellValue(robj.data[0].pass_in_out[i].remark);
                        dataRow.GetCell(j).CellStyle = cs;
                    }
                }
                rowCounter++;
            }





            //輸出excel到串流
            MemoryStream stream = new MemoryStream();
            workbook.Write(stream);

            /*
            FileStream file = new FileStream(@"C:\temp.xlsx", FileMode.Create);//產生檔案
            workbook.Write(file);
            file.Close();
            */
            stream.Flush();
            stream.Position = 0;
            bytes = StreamToBytes(stream);

            //Write the stream data of workbook to the root directory   
            string filename = path + "\\test.xls";
            FileStream file = new FileStream(filename, FileMode.Create);
            workbook.Write(file);
            file.Close();

            // Excel 檔案位置
            string sourcexlsx = path + "\\test.xls";
            // PDF 儲存位置
            string targetpdf = path + "\\result.pdf";
            //建立 Excel application instance
            Application appExcel = new Application();
            //開啟 Excel 檔案
            var xlsxDocument = appExcel.Workbooks.Open(filename);
            //匯出為 pdf
            xlsxDocument.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, targetpdf);
            //關閉 Excel 檔
            xlsxDocument.Close();
            //結束 Excel
            appExcel.Quit();

            //釋放資源
            sheet1 = null;
            workbook = null;

            //我要下載的檔案位置
            //string filepath = Server.MapPath("~/123.zip");
            //取得檔案名稱
            string pdfFilename = System.IO.Path.GetFileName(targetpdf);
            //讀成串流
            Stream iStream = new FileStream(targetpdf, FileMode.Open, FileAccess.Read, FileShare.Read);
            //回傳出檔案
            return File(iStream, "application/pdf", DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf");

            /* string contentType = "";
            //下載excel的content type
            contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            //string cTypePDF = "application/pdf";
            return File(bytes, contentType, DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls"); //xls */

        }

        private Image ConvertByteToImage(byte[] Buffer)
        {
            if (Buffer == null || Buffer.Length == 0) { return null; }
            Image oImage = null;
            try
            {
                MemoryStream oMemoryStream = new MemoryStream(Buffer);
                //設定資料流位置
                //oMemoryStream.Position = 0;
                oImage = System.Drawing.Image.FromStream(oMemoryStream);
            }
            catch
            {
                throw;
            }
            return oImage;
        }

        private void WriteImageToCell(ref HSSFWorkbook workbook, ref ISheet sheet,
            ref HSSFRow dataRow
            , byte[] imageBytes, int rowIndex, int columnIndex)
        {
            try
            {
                //建立繪圖物件
                var patriarch = sheet.CreateDrawingPatriarch();

                var image = ConvertByteToImage(imageBytes);
                if (image != null)
                {
                    // 產生縮圖
                    decimal sizeRatio = ((decimal)image.Height / image.Width);
                    int thumbWidth = 140;//設定縮圖的寬度上限
                    int thumbHeight = decimal.ToInt32(sizeRatio * thumbWidth);
                    var thumbStream = image.GetThumbnailImage(thumbWidth, thumbHeight, () => false, IntPtr.Zero);
                    var memoryStream = new MemoryStream();
                    thumbStream.Save(memoryStream, ImageFormat.Jpeg);

                    // 將縮圖加入到 workbook 中
                    //在網路上找不太到這個pictureIndex的意義
                    int pictureIndex = 0;
                    pictureIndex = workbook.AddPicture(memoryStream.ToArray(), NPOI.SS.UserModel.PictureType.JPEG);

                    // 將縮圖定位到 worksheet 中
                    var anchor = new HSSFClientAnchor(50, 50, 0, 0, columnIndex,
                        rowIndex, columnIndex, rowIndex);
                    var picture = patriarch.CreatePicture(anchor, pictureIndex);
                    var size = picture.GetImageDimension();
                    dataRow.HeightInPoints = size.Height;
                    picture.Resize();
                }
            }
            catch (Exception ex)
            {
                // 圖片載入失敗，顯示錯誤訊息
                dataRow.GetCell(columnIndex).SetCellValue(ex.Message);
            }
        }

        private void WriteImageToCell2(ref HSSFWorkbook workbook, ref ISheet sheet,
            ref HSSFRow dataRow
            , byte[] imageBytes, int rowIndex, int columnIndex)
        {
            try
            {
                //建立繪圖物件
                var patriarch = sheet.CreateDrawingPatriarch();

                var image = ConvertByteToImage(imageBytes);
                if (image != null)
                {
                    // 產生縮圖
                    decimal sizeRatio = ((decimal)image.Height / image.Width);
                    int thumbWidth = 140;//設定縮圖的寬度上限
                    int thumbHeight = decimal.ToInt32(sizeRatio * thumbWidth);
                    var thumbStream = image.GetThumbnailImage(thumbWidth, thumbHeight, () => false, IntPtr.Zero);
                    var memoryStream = new MemoryStream();
                    thumbStream.Save(memoryStream, ImageFormat.Jpeg);

                    // 將縮圖加入到 workbook 中
                    //在網路上找不太到這個pictureIndex的意義
                    int pictureIndex = 0;
                    pictureIndex = workbook.AddPicture(memoryStream.ToArray(), NPOI.SS.UserModel.PictureType.JPEG);

                    // 將縮圖定位到 worksheet 中
                    var anchor = new HSSFClientAnchor(450, 50, 0, 0, columnIndex,
                        rowIndex, columnIndex, rowIndex);
                    var picture = patriarch.CreatePicture(anchor, pictureIndex);
                    var size = picture.GetImageDimension();
                    dataRow.HeightInPoints = size.Height;
                    picture.Resize();
                }
            }
            catch (Exception ex)
            {
                // 圖片載入失敗，顯示錯誤訊息
                dataRow.GetCell(columnIndex).SetCellValue(ex.Message);
            }
        }

        private void WriteImageToCellRead(ref HSSFWorkbook workbook, ref ISheet sheet,
            ref HSSFRow dataRow
            , byte[] imageBytes, int rowIndex, int columnIndex)
        {
            try
            {
                //建立繪圖物件
                var patriarch = sheet.CreateDrawingPatriarch();

                var image = ConvertByteToImage(imageBytes);
                if (image != null)
                {
                    // 產生縮圖
                    decimal sizeRatio = ((decimal)image.Height / image.Width);
                    int thumbWidth = 120;//設定縮圖的寬度上限
                    int thumbHeight = decimal.ToInt32(sizeRatio * thumbWidth);
                    var thumbStream = image.GetThumbnailImage(thumbWidth, thumbHeight, () => false, IntPtr.Zero);
                    var memoryStream = new MemoryStream();
                    thumbStream.Save(memoryStream, ImageFormat.Jpeg);

                    // 將縮圖加入到 workbook 中
                    //在網路上找不太到這個pictureIndex的意義
                    int pictureIndex = 0;
                    pictureIndex = workbook.AddPicture(memoryStream.ToArray(), NPOI.SS.UserModel.PictureType.JPEG);

                    // 將縮圖定位到 worksheet 中
                    var anchor = new HSSFClientAnchor(500, 20, 0, 0, columnIndex,
                        rowIndex, columnIndex, rowIndex);
                    var picture = patriarch.CreatePicture(anchor, pictureIndex);
                    var size = picture.GetImageDimension();
                    dataRow.HeightInPoints = size.Height;
                    picture.Resize();
                }
            }
            catch (Exception ex)
            {
                // 圖片載入失敗，顯示錯誤訊息
                dataRow.GetCell(columnIndex).SetCellValue(ex.Message);
            }
        }

        private void WriteImageToCellMonitor(ref HSSFWorkbook workbook, ref ISheet sheet,
            ref HSSFRow dataRow
            , byte[] imageBytes, int rowIndex, int columnIndex)
        {
            try
            {
                //建立繪圖物件
                var patriarch = sheet.CreateDrawingPatriarch();

                var image = ConvertByteToImage(imageBytes);
                if (image != null)
                {
                    // 產生縮圖
                    decimal sizeRatio = ((decimal)image.Height / image.Width);
                    int thumbWidth = 180;//設定縮圖的寬度上限
                    int thumbHeight = decimal.ToInt32(sizeRatio * thumbWidth);
                    var thumbStream = image.GetThumbnailImage(thumbWidth, thumbHeight, () => false, IntPtr.Zero);
                    var memoryStream = new MemoryStream();
                    thumbStream.Save(memoryStream, ImageFormat.Jpeg);

                    // 將縮圖加入到 workbook 中
                    //在網路上找不太到這個pictureIndex的意義
                    int pictureIndex = 0;
                    pictureIndex = workbook.AddPicture(memoryStream.ToArray(), NPOI.SS.UserModel.PictureType.JPEG);

                    // 將縮圖定位到 worksheet 中
                    var anchor = new HSSFClientAnchor(50, 10, 0, 0, columnIndex,
                        rowIndex, columnIndex, rowIndex);
                    var picture = patriarch.CreatePicture(anchor, pictureIndex);
                    var size = picture.GetImageDimension();
                    dataRow.HeightInPoints = size.Height;
                    picture.Resize();


                }
            }
            catch (Exception ex)
            {
                // 圖片載入失敗，顯示錯誤訊息
                dataRow.GetCell(columnIndex).SetCellValue(ex.Message);
            }
        }
        private byte[] StreamToBytes(Stream stream)
        {
            byte[] bytes = new byte[stream.Length];
            stream.Read(bytes, 0, bytes.Length);

            // 設置當前流的位置為流的開始
            stream.Seek(0, SeekOrigin.Begin);
            return bytes;
        }
        public ActionResult Monitor()
        {
            byte[] bytes = null;
            int rowCounter = 0;
            Rootobject2 robj2 = new Rootobject2();
            string path = System.AppDomain.CurrentDomain.BaseDirectory;
            robj2 = JsonConvert.DeserializeObject<Rootobject2>(System.IO.File.ReadAllText(path + "\\page2.json"));

            HSSFWorkbook workbookforMonitor = new HSSFWorkbook();
            ISheet sheet2 = workbookforMonitor.CreateSheet("MySheet");
            //sheet2.PrintSetup.ValidSettings = true;
            sheet2.SetMargin(MarginType.TopMargin, 0.4); // 0.4 = 1cm
            sheet2.SetMargin(MarginType.RightMargin, 0.4);
            sheet2.SetMargin(MarginType.BottomMargin, 0.4);
            sheet2.SetMargin(MarginType.LeftMargin, 0.4);
            sheet2.PrintSetup.FitWidth = 1;
            sheet2.PrintSetup.FitHeight = 0;
            //sheet2.PrintSetup.PaperSize = 55;

            #region FontConfig
            HSSFFont Font12 = (HSSFFont)workbookforMonitor.CreateFont();
            Font12.FontHeightInPoints = 12;
            Font12.FontName = "標楷體";

            HSSFFont Font14 = (HSSFFont)workbookforMonitor.CreateFont();
            Font14.FontHeightInPoints = 14;
            Font14.FontName = "標楷體";

            HSSFFont Font16 = (HSSFFont)workbookforMonitor.CreateFont();
            Font16.FontHeightInPoints = 16;
            Font16.FontName = "標楷體";

            HSSFFont Font18 = (HSSFFont)workbookforMonitor.CreateFont();
            Font18.FontHeightInPoints = 18;
            Font18.FontName = "標楷體";

            HSSFFont Font20 = (HSSFFont)workbookforMonitor.CreateFont();
            Font20.FontHeightInPoints = 20;
            Font20.FontName = "標楷體";


            HSSFFont Font12B = (HSSFFont)workbookforMonitor.CreateFont();
            Font12B.FontHeightInPoints = 12;
            Font12B.IsBold = true;
            Font12B.FontName = "標楷體";

            HSSFFont Font14UB = (HSSFFont)workbookforMonitor.CreateFont();
            Font14UB.FontHeightInPoints = 14;
            Font14UB.Underline = FontUnderlineType.Single;
            Font14UB.IsBold = true;
            Font14UB.FontName = "標楷體";

            HSSFFont Font20B = (HSSFFont)workbookforMonitor.CreateFont();
            Font20B.FontHeightInPoints = 20;
            Font20B.IsBold = true;
            Font20B.FontName = "標楷體";
            #endregion

            #region HSSFCellStyle
            HSSFCellStyle titleRowcsLT = (HSSFCellStyle)workbookforMonitor.CreateCellStyle();
            titleRowcsLT.WrapText = true;
            titleRowcsLT.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            titleRowcsLT.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Distributed;
            titleRowcsLT.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            titleRowcsLT.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            titleRowcsLT.BorderRight = NPOI.SS.UserModel.BorderStyle.None;
            titleRowcsLT.BorderTop = NPOI.SS.UserModel.BorderStyle.Thick;
            titleRowcsLT.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            titleRowcsLT.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            titleRowcsLT.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            titleRowcsLT.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            titleRowcsLT.SetFont(Font20B);

            HSSFCellStyle titleRowcsT = (HSSFCellStyle)workbookforMonitor.CreateCellStyle();
            titleRowcsT.WrapText = true;
            titleRowcsT.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            titleRowcsT.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            titleRowcsT.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            titleRowcsT.BorderLeft = NPOI.SS.UserModel.BorderStyle.None;
            titleRowcsT.BorderRight = NPOI.SS.UserModel.BorderStyle.None;
            titleRowcsT.BorderTop = NPOI.SS.UserModel.BorderStyle.Thick;
            titleRowcsT.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            titleRowcsT.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            titleRowcsT.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            titleRowcsT.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            titleRowcsT.SetFont(Font20B);

            HSSFCellStyle titleRowcsRT = (HSSFCellStyle)workbookforMonitor.CreateCellStyle();
            titleRowcsRT.WrapText = true;
            titleRowcsRT.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            titleRowcsRT.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            titleRowcsRT.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            titleRowcsRT.BorderLeft = NPOI.SS.UserModel.BorderStyle.None;
            titleRowcsRT.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            titleRowcsRT.BorderTop = NPOI.SS.UserModel.BorderStyle.Thick;
            titleRowcsRT.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            titleRowcsRT.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            titleRowcsRT.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            titleRowcsRT.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            titleRowcsRT.SetFont(Font20B);

            HSSFCellStyle subtitlecsLT = (HSSFCellStyle)workbookforMonitor.CreateCellStyle();
            subtitlecsLT.WrapText = true;
            subtitlecsLT.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            subtitlecsLT.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            subtitlecsLT.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            subtitlecsLT.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            subtitlecsLT.BorderRight = NPOI.SS.UserModel.BorderStyle.None;
            subtitlecsLT.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            subtitlecsLT.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsLT.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsLT.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsLT.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsLT.SetFont(Font12);

            HSSFCellStyle subtitlecsT = (HSSFCellStyle)workbookforMonitor.CreateCellStyle();
            subtitlecsT.WrapText = true;
            subtitlecsT.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            subtitlecsT.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            subtitlecsT.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            subtitlecsT.BorderLeft = NPOI.SS.UserModel.BorderStyle.None;
            subtitlecsT.BorderRight = NPOI.SS.UserModel.BorderStyle.None;
            subtitlecsT.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            subtitlecsT.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsT.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsT.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsT.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsT.SetFont(Font12);

            HSSFCellStyle subtitlecsRT = (HSSFCellStyle)workbookforMonitor.CreateCellStyle();
            subtitlecsRT.WrapText = true;
            subtitlecsRT.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            subtitlecsRT.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            subtitlecsRT.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            subtitlecsRT.BorderLeft = NPOI.SS.UserModel.BorderStyle.None;
            subtitlecsRT.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            subtitlecsRT.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            subtitlecsRT.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsRT.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsRT.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsRT.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsRT.SetFont(Font12);

            HSSFCellStyle subtitlecsL = (HSSFCellStyle)workbookforMonitor.CreateCellStyle();
            subtitlecsL.WrapText = true;
            subtitlecsL.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            subtitlecsL.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            subtitlecsL.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            subtitlecsL.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            subtitlecsL.BorderRight = NPOI.SS.UserModel.BorderStyle.None;
            subtitlecsL.BorderTop = NPOI.SS.UserModel.BorderStyle.None;
            subtitlecsL.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsL.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsL.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsL.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsL.SetFont(Font12);

            HSSFCellStyle subtitlecsNone = (HSSFCellStyle)workbookforMonitor.CreateCellStyle();
            subtitlecsNone.WrapText = true;
            subtitlecsNone.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            subtitlecsNone.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            subtitlecsNone.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            subtitlecsNone.BorderLeft = NPOI.SS.UserModel.BorderStyle.None;
            subtitlecsNone.BorderRight = NPOI.SS.UserModel.BorderStyle.None;
            subtitlecsNone.BorderTop = NPOI.SS.UserModel.BorderStyle.None;
            subtitlecsNone.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsNone.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsNone.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsNone.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsNone.SetFont(Font12);

            HSSFCellStyle subtitlecs = (HSSFCellStyle)workbookforMonitor.CreateCellStyle();
            subtitlecs.WrapText = true;
            subtitlecs.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            subtitlecs.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            subtitlecs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            subtitlecs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            subtitlecs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            subtitlecs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            subtitlecs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecs.SetFont(Font12);

            HSSFCellStyle subtitlecsR = (HSSFCellStyle)workbookforMonitor.CreateCellStyle();
            subtitlecsR.WrapText = true;
            subtitlecsR.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            subtitlecsR.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            subtitlecsR.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            subtitlecsR.BorderLeft = NPOI.SS.UserModel.BorderStyle.None;
            subtitlecsR.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            subtitlecsR.BorderTop = NPOI.SS.UserModel.BorderStyle.None;
            subtitlecsR.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsR.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsR.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsR.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsR.SetFont(Font12);

            HSSFCellStyle subtitlecsLB = (HSSFCellStyle)workbookforMonitor.CreateCellStyle();
            subtitlecsLB.WrapText = true;
            subtitlecsLB.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            subtitlecsLB.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            subtitlecsLB.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            subtitlecsLB.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            subtitlecsLB.BorderRight = NPOI.SS.UserModel.BorderStyle.None;
            subtitlecsLB.BorderTop = NPOI.SS.UserModel.BorderStyle.None;
            subtitlecsLB.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsLB.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsLB.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsLB.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsLB.SetFont(Font12);

            HSSFCellStyle subtitlecsB = (HSSFCellStyle)workbookforMonitor.CreateCellStyle();
            subtitlecsB.WrapText = true;
            subtitlecsB.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            subtitlecsB.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            subtitlecsB.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            subtitlecsB.BorderLeft = NPOI.SS.UserModel.BorderStyle.None;
            subtitlecsB.BorderRight = NPOI.SS.UserModel.BorderStyle.None;
            subtitlecsB.BorderTop = NPOI.SS.UserModel.BorderStyle.None;
            subtitlecsB.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsB.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsB.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsB.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsB.SetFont(Font12);

            HSSFCellStyle subtitlecsRB = (HSSFCellStyle)workbookforMonitor.CreateCellStyle();
            subtitlecsRB.WrapText = true;
            subtitlecsRB.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            subtitlecsRB.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            subtitlecsRB.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            subtitlecsRB.BorderLeft = NPOI.SS.UserModel.BorderStyle.None;
            subtitlecsRB.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            subtitlecsRB.BorderTop = NPOI.SS.UserModel.BorderStyle.None;
            subtitlecsRB.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsRB.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsRB.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsRB.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            subtitlecsRB.SetFont(Font12);

            HSSFCellStyle contentcsLT = (HSSFCellStyle)workbookforMonitor.CreateCellStyle();
            contentcsLT.WrapText = true;
            contentcsLT.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            contentcsLT.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            contentcsLT.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            contentcsLT.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            contentcsLT.BorderRight = NPOI.SS.UserModel.BorderStyle.None;
            contentcsLT.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            contentcsLT.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            contentcsLT.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            contentcsLT.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            contentcsLT.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            contentcsLT.SetFont(Font12);

            HSSFCellStyle contentcs = (HSSFCellStyle)workbookforMonitor.CreateCellStyle();
            contentcs.WrapText = true;
            contentcs.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            contentcs.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            contentcs.BorderBottom = NPOI.SS.UserModel.BorderStyle.None;
            contentcs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            contentcs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            contentcs.BorderTop = NPOI.SS.UserModel.BorderStyle.None;
            contentcs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            contentcs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            contentcs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            contentcs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            contentcs.SetFont(Font12);

            HSSFCellStyle contentcsLB = (HSSFCellStyle)workbookforMonitor.CreateCellStyle();
            contentcsLB.WrapText = true;
            contentcsLB.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            contentcsLB.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            contentcsLB.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            contentcsLB.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            contentcsLB.BorderRight = NPOI.SS.UserModel.BorderStyle.None;
            contentcsLB.BorderTop = NPOI.SS.UserModel.BorderStyle.None;
            contentcsLB.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            contentcsLB.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            contentcsLB.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            contentcsLB.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            contentcsLB.SetFont(Font12);

            HSSFCellStyle stampssigncsL = (HSSFCellStyle)workbookforMonitor.CreateCellStyle();
            stampssigncsL.WrapText = true;
            stampssigncsL.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            stampssigncsL.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            stampssigncsL.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thick;
            stampssigncsL.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            stampssigncsL.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            stampssigncsL.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            stampssigncsL.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampssigncsL.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampssigncsL.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampssigncsL.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampssigncsL.SetFont(Font12);

            HSSFCellStyle stampssigncs = (HSSFCellStyle)workbookforMonitor.CreateCellStyle();
            stampssigncs.WrapText = true;
            stampssigncs.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            stampssigncs.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            stampssigncs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thick;
            stampssigncs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            stampssigncs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            stampssigncs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            stampssigncs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampssigncs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampssigncs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampssigncs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampssigncs.SetFont(Font12);

            HSSFCellStyle stampssigncsR = (HSSFCellStyle)workbookforMonitor.CreateCellStyle();
            stampssigncsR.WrapText = true;
            stampssigncsR.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            stampssigncsR.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            stampssigncsR.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thick;
            stampssigncsR.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            stampssigncsR.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            stampssigncsR.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            stampssigncsR.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampssigncsR.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampssigncsR.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampssigncsR.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            stampssigncsR.SetFont(Font12);

            HSSFCellStyle passtosigncsL = (HSSFCellStyle)workbookforMonitor.CreateCellStyle();
            passtosigncsL.WrapText = true;
            passtosigncsL.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            passtosigncsL.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            passtosigncsL.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            passtosigncsL.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thick;
            passtosigncsL.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            passtosigncsL.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            passtosigncsL.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            passtosigncsL.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            passtosigncsL.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            passtosigncsL.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            passtosigncsL.SetFont(Font20);

            HSSFCellStyle passtosigncs = (HSSFCellStyle)workbookforMonitor.CreateCellStyle();
            passtosigncs.WrapText = true;
            passtosigncs.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            passtosigncs.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            passtosigncs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            passtosigncs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            passtosigncs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            passtosigncs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            passtosigncs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            passtosigncs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            passtosigncs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            passtosigncs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            passtosigncs.SetFont(Font20);

            HSSFCellStyle passtosigncsR = (HSSFCellStyle)workbookforMonitor.CreateCellStyle();
            passtosigncsR.WrapText = true;
            passtosigncsR.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            passtosigncsR.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            passtosigncsR.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            passtosigncsR.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            passtosigncsR.BorderRight = NPOI.SS.UserModel.BorderStyle.Thick;
            passtosigncsR.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            passtosigncsR.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            passtosigncsR.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            passtosigncsR.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            passtosigncsR.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            passtosigncsR.SetFont(Font20);

            HSSFCellStyle contentRow3cs = (HSSFCellStyle)workbookforMonitor.CreateCellStyle();
            contentRow3cs.WrapText = true;
            contentRow3cs.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            contentRow3cs.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            contentRow3cs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            contentRow3cs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            contentRow3cs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            contentRow3cs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            contentRow3cs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            contentRow3cs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            contentRow3cs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            contentRow3cs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            contentRow3cs.SetFont(Font12);

            HSSFCellStyle contentRow2cs = (HSSFCellStyle)workbookforMonitor.CreateCellStyle();
            contentRow2cs.WrapText = true;
            contentRow2cs.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            contentRow2cs.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            contentRow2cs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            contentRow2cs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            contentRow2cs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            contentRow2cs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            contentRow2cs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            contentRow2cs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            contentRow2cs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            contentRow2cs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            contentRow2cs.SetFont(Font14UB);

            HSSFCellStyle contentRow2csL = (HSSFCellStyle)workbookforMonitor.CreateCellStyle();
            contentRow2csL.WrapText = true;
            contentRow2csL.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            contentRow2csL.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            contentRow2csL.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            contentRow2csL.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            contentRow2csL.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            contentRow2csL.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            contentRow2csL.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            contentRow2csL.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            contentRow2csL.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            contentRow2csL.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Black.Index;
            contentRow2csL.SetFont(Font12);
            #endregion


            sheet2.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 0, 9));
            HSSFRow titleRow = (HSSFRow)sheet2.CreateRow(rowCounter);
            titleRow.Height = 45 * 20;
            titleRow.CreateCell(0).SetCellValue("國防部參謀本部資通電軍指揮部忠信營區總值日官室\n智慧型警監系統行動準據");
            titleRow.GetCell(0).CellStyle = titleRowcsLT;
            for (int c = 1; c < 10; c++)
            {
                if (c == 9)
                {
                    titleRow.CreateCell(c).SetCellValue("");
                    titleRow.GetCell(c).CellStyle = titleRowcsRT;
                }
                else
                {
                    titleRow.CreateCell(c).SetCellValue("");
                    titleRow.GetCell(c).CellStyle = titleRowcsT;
                }
            }
            rowCounter++;

            sheet2.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 2, 5)); // 工作事項
            sheet2.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 7, 9)); // 備考
            HSSFRow subtitleRow = (HSSFRow)sheet2.CreateRow(rowCounter);
            subtitleRow.Height = 25 * 20;
            for (int c = 0; c < 10; c++)
            {
                if (c == 0)
                {
                    subtitleRow.CreateCell(c).SetCellValue("項次");
                    subtitleRow.GetCell(c).CellStyle = subtitlecsLT;
                    sheet2.SetColumnWidth(c, 10 * 256);
                }
                else if (c == 1)
                {
                    subtitleRow.CreateCell(c).SetCellValue("時間");
                    subtitleRow.GetCell(c).CellStyle = contentcsLT;
                }
                else if (c == 2)
                {
                    subtitleRow.CreateCell(c).SetCellValue("工作事項");
                    subtitleRow.GetCell(c).CellStyle = contentcsLT;
                }
                else if (c == 6)
                {
                    subtitleRow.CreateCell(c).SetCellValue("執行情形");
                    subtitleRow.GetCell(c).CellStyle = contentcsLT;
                }
                else if (c == 7)
                {
                    subtitleRow.CreateCell(c).SetCellValue("備考");
                    subtitleRow.GetCell(c).CellStyle = contentcsLT;
                }
                else if (c == 9)
                {
                    subtitleRow.CreateCell(c).SetCellValue("");
                    subtitleRow.GetCell(c).CellStyle = subtitlecsRT;
                }
                else
                {
                    subtitleRow.CreateCell(c).SetCellValue("");
                    subtitleRow.GetCell(c).CellStyle = subtitlecsT;
                }
            }
            rowCounter++;

            sheet2.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter + 21, 0, 0));
            sheet2.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter + 21, 1, 1));
            sheet2.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter + 21, 2, 5));
            sheet2.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter + 21, 6, 6));
            //sheet2.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter + 38, 7, 9));
            //sheet2.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 11, 15));
            int num = 1;
            for (int i = 0; i < 22; i++)
            {
                HSSFRow contentRow = (HSSFRow)sheet2.CreateRow(rowCounter);
                sheet2.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 7, 9));
                if (i == 0)
                {
                    for (int c = 0; c < 10; c++)
                    {
                        if (c == 0)
                        {
                            contentRow.CreateCell(c).SetCellValue(num);
                            contentRow.GetCell(c).CellStyle = subtitlecsLT;
                        }
                        else if (c == 1)
                        {
                            contentRow.CreateCell(c).SetCellValue(robj2.data[0].duty.morning.time + "時");
                            contentRow.GetCell(c).CellStyle = contentcsLT;
                        }
                        else if (c == 2)
                        {
                            contentRow.CreateCell(c).SetCellValue(robj2.data[0].duty.morning.work);
                            contentRow.GetCell(c).CellStyle = contentcsLT;
                        }
                        else if (c == 6)
                        {
                            if (robj2.data[0].duty.morning.execute)
                            {
                                contentRow.CreateCell(c).SetCellValue("正常☑\n異常☐");
                                contentRow.GetCell(c).CellStyle = contentcsLT;
                            }
                            else
                            {
                                contentRow.CreateCell(c).SetCellValue("正常☐\n異常☑");
                                contentRow.GetCell(c).CellStyle = contentcsLT;
                            }

                        }
                        else if (c == 7)
                        {
                            contentRow.CreateCell(c).SetCellValue(robj2.data[0].duty.morning.remark[0].label);
                            contentRow.GetCell(c).CellStyle = contentcsLT;
                        }
                        else if (c == 9)
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = subtitlecsRT;
                        }
                        else
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = subtitlecsT;
                        }
                    }
                    num++;
                }
                else if (i == 1)
                {
                    for (int c = 0; c < 10; c++)
                    {
                        if (c == 0)
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = subtitlecsL;
                        }
                        else if (c == 7)
                        {
                            contentRow.CreateCell(c);
                            contentRow.GetCell(c).CellStyle = contentcs;
                            byte[] imagebytes = System.IO.File.ReadAllBytes(path + "\\images\\sample6.jpg");
                            WriteImageToCellMonitor(ref workbookforMonitor, ref sheet2, ref contentRow, imagebytes, rowCounter, c);
                        }
                        else if (c == 9)
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = subtitlecsR;
                        }
                        else
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = contentcs;
                        }
                    }

                }
                else if (i == 2)
                {
                    for (int c = 0; c < 10; c++)
                    {
                        if (c == 0)
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = subtitlecsL;
                        }
                        else if (c == 7)
                        {
                            contentRow.CreateCell(c).SetCellValue(robj2.data[0].duty.morning.remark[1].label);
                            contentRow.GetCell(c).CellStyle = contentcs;
                        }
                        else if (c == 9)
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = subtitlecsR;
                        }
                        else
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = contentcs;
                        }
                    }
                }
                else if (i == 3)
                {
                    for (int c = 0; c < 10; c++)
                    {
                        if (c == 0)
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = subtitlecsL;
                        }
                        else if (c == 7)
                        {
                            contentRow.CreateCell(c);
                            contentRow.GetCell(c).CellStyle = contentcs;
                            byte[] imagebytes = System.IO.File.ReadAllBytes(path + "\\images\\sample7.jpg");
                            WriteImageToCellMonitor(ref workbookforMonitor, ref sheet2, ref contentRow, imagebytes, rowCounter, c);
                        }
                        else if (c == 9)
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = subtitlecsR;
                        }
                        else
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = contentcs;
                        }
                    }

                }
                else if (i == 7)
                {
                    for (int c = 0; c < 10; c++)
                    {
                        if (c == 0)
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = subtitlecsL;
                        }
                        else if (c == 7)
                        {
                            contentRow.CreateCell(c).SetCellValue(robj2.data[0].duty.morning.remark[2].label);
                            contentRow.GetCell(c).CellStyle = contentcs;
                        }
                        else if (c == 9)
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = subtitlecsR;
                        }
                        else
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = contentcs;
                        }
                    }
                }
                else if (i == 12)
                {
                    for (int c = 0; c < 10; c++)
                    {
                        if (c == 0)
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = subtitlecsL;
                        }
                        else if (c == 7)
                        {
                            contentRow.CreateCell(c).SetCellValue(robj2.data[0].duty.morning.remark[3].label);
                            contentRow.GetCell(c).CellStyle = contentcs;
                        }
                        else if (c == 9)
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = subtitlecsR;
                        }
                        else
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = contentcs;
                        }
                    }
                }
                else if (i == 17)
                {
                    for (int c = 0; c < 10; c++)
                    {
                        if (c == 0)
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = subtitlecsL;
                        }
                        else if (c == 7)
                        {
                            contentRow.CreateCell(c).SetCellValue(robj2.data[0].duty.morning.remark[4].label);
                            contentRow.GetCell(c).CellStyle = contentcs;
                        }
                        else if (c == 9)
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = subtitlecsR;
                        }
                        else
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = contentcs;
                        }
                    }
                }
                else if (i == 21)
                {
                    for (int c = 0; c < 10; c++)
                    {
                        if (c == 0)
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = subtitlecsLB;
                        }
                        else if (c == 1)
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = contentcsLB;
                        }
                        else if (c == 2)
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = contentcsLB;
                        }
                        else if (c == 6)
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = contentcsLB;
                        }
                        else if (c == 7)
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = contentcsLB;
                        }
                        else if (c == 9)
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = subtitlecsRB;
                        }
                        else
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = subtitlecsB;
                        }
                    }

                }
                else
                {
                    for (int c = 0; c < 10; c++)
                    {
                        if (c == 0)
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = subtitlecsL;
                        }
                        else if (c == 9)
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = subtitlecsR;
                        }
                        else
                        {
                            contentRow.CreateCell(c).SetCellValue("");
                            contentRow.GetCell(c).CellStyle = contentcs;
                        }
                    }

                }
                rowCounter++;
            }
            //rowCounter = rowCounter + 39;
            sheet2.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 2, 5));
            sheet2.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 7, 9)); // 備考
            HSSFRow contentRow2 = (HSSFRow)sheet2.CreateRow(rowCounter);
            contentRow2.Height = 45 * 20;
            for (int c = 0; c < 10; c++)
            {
                if (c == 0)
                {
                    contentRow2.CreateCell(c).SetCellValue(num);
                    contentRow2.GetCell(c).CellStyle = subtitlecsLT;
                }
                else if (c == 1)
                {
                    contentRow2.CreateCell(c).SetCellValue(robj2.data[0].duty.evening.time + "時");
                    contentRow2.GetCell(c).CellStyle = contentRow2csL;
                }
                else if (c == 2)
                {
                    contentRow2.CreateCell(c).SetCellValue("記錄警監系統運作狀況：");
                    contentRow2.GetCell(c).CellStyle = contentRow2cs;
                }
                else if (c == 6)
                {
                    if (robj2.data[0].duty.evening.execute)
                    {
                        contentRow2.CreateCell(c).SetCellValue("已完成☑");
                        contentRow2.GetCell(c).CellStyle = contentRow2csL;
                    }
                    else
                    {
                        contentRow2.CreateCell(c).SetCellValue("已完成☐");
                        contentRow2.GetCell(c).CellStyle = contentRow2csL;
                    }

                }
                else if (c == 7)
                {
                    contentRow2.CreateCell(c).SetCellValue("");
                    contentRow2.GetCell(c).CellStyle = contentRow2cs;
                }
                else if (c == 9)
                {
                    contentRow2.CreateCell(c).SetCellValue("");
                    contentRow2.GetCell(c).CellStyle = subtitlecsRT;
                }
                else
                {
                    contentRow2.CreateCell(c).SetCellValue("");
                    contentRow2.GetCell(c).CellStyle = subtitlecsT;
                }
            }
            rowCounter++;

            sheet2.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 1, 9)); // 備考
            HSSFRow contentRow3 = (HSSFRow)sheet2.CreateRow(rowCounter);
            contentRow3.Height = 45 * 20;
            for (int c = 0; c < 10; c++)
            {
                if (c == 0)
                {
                    contentRow3.CreateCell(c).SetCellValue("備考");
                    contentRow3.GetCell(c).CellStyle = subtitlecsLT;
                }
                else if (c == 1)
                {
                    contentRow3.CreateCell(c).SetCellValue(robj2.data[0].remark);
                    contentRow3.GetCell(c).CellStyle = contentRow3cs;
                }
                else if (c == 9)
                {
                    contentRow3.CreateCell(c).SetCellValue("");
                    contentRow3.GetCell(c).CellStyle = subtitlecsRT;
                }
                else
                {
                    contentRow3.CreateCell(c).SetCellValue("");
                    contentRow3.GetCell(c).CellStyle = subtitlecsT;
                }
            }
            rowCounter++;

            sheet2.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 0, 4)); // 呈核
            sheet2.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 5, 9)); // 批示
            HSSFRow passtosign = (HSSFRow)sheet2.CreateRow(rowCounter);
            passtosign.Height = 25 * 20;
            for (int c = 0; c < 10; c++)
            {
                if (c == 0)
                {
                    passtosign.CreateCell(c).SetCellValue("呈核");
                    passtosign.GetCell(c).CellStyle = passtosigncsL;
                }
                else if (c == 5)
                {
                    passtosign.CreateCell(c).SetCellValue("批示");
                    passtosign.GetCell(c).CellStyle = passtosigncs;
                }
                else if (c == 9)
                {
                    passtosign.CreateCell(c).SetCellValue("");
                    passtosign.GetCell(c).CellStyle = passtosigncsR;
                }
                else
                {
                    passtosign.CreateCell(c).SetCellValue("");
                    passtosign.GetCell(c).CellStyle = passtosigncs;
                }
            }
            rowCounter++;

            sheet2.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 0, 4)); // 呈核
            sheet2.AddMergedRegion(new CellRangeAddress(rowCounter, rowCounter, 5, 9)); // 批示
            HSSFRow stampssign = (HSSFRow)sheet2.CreateRow(rowCounter);
            stampssign.Height = 65 * 20;
            for (int c = 0; c < 10; c++)
            {
                if (c == 0)
                {
                    stampssign.CreateCell(c).SetCellValue("");
                    stampssign.GetCell(c).CellStyle = stampssigncsL;
                }
                else if (c == 5)
                {
                    stampssign.CreateCell(c).SetCellValue("");
                    stampssign.GetCell(c).CellStyle = stampssigncs;
                }
                else if (c == 9)
                {
                    stampssign.CreateCell(c).SetCellValue("");
                    stampssign.GetCell(c).CellStyle = stampssigncsR;
                }
                else
                {
                    stampssign.CreateCell(c).SetCellValue("");
                    stampssign.GetCell(c).CellStyle = stampssigncs;
                }
            }
            rowCounter++;

            string filename = path + "\\test2.xls";
            FileStream file = new FileStream(filename, FileMode.Create);
            workbookforMonitor.Write(file);
            file.Close();

            // Excel 檔案位置
            string sourcexlsx = path + "\\test2.xls";
            // PDF 儲存位置
            string targetpdf = path + "\\result2.pdf";
            //建立 Excel application instance
            Application appExcel = new Application();
            //開啟 Excel 檔案
            var xlsxDocument = appExcel.Workbooks.Open(sourcexlsx);
            //匯出為 pdf
            xlsxDocument.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, targetpdf);
            //關閉 Excel 檔
            xlsxDocument.Close();
            //結束 Excel
            appExcel.Quit();

            //釋放資源
            sheet2 = null;
            workbookforMonitor = null;

            Stream iStream = new FileStream(targetpdf, FileMode.Open, FileAccess.Read, FileShare.Read);
            return File(iStream, "application/pdf", DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf");
        }
    }
}