using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelProjectMvc.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using System.Web.UI;

namespace ExcelProjectMvc.Controllers
{
    public class HomeController : Controller
    {
        // GET: Home

        public ActionResult Index()
        {

            return View(GetPerson());
        }

        [HttpPost]
        public PartialViewResult AddPerson(PersonModel model)
        {
            GetPerson().AddPerson(model);
            var list = GetPerson();

            return PartialView("~/Views/Home/_PartialListView.cshtml", list);


        }

        public ListPerson GetPerson()
        {
            var person = (ListPerson)Session["Person"];
            if (person == null)
            {
                person = new ListPerson();
                Session["Person"] = person;
            }
            return person;
        }

        public ActionResult ConvertToExcel()
        {
            MemoryStream ms = new MemoryStream();
            SpreadsheetDocument xl = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook);
            WorkbookPart wbp = xl.AddWorkbookPart();
            WorksheetPart wsp = wbp.AddNewPart<WorksheetPart>();
            Workbook wb = new Workbook();
            FileVersion fv = new FileVersion();
            fv.ApplicationName = "Microsoft Office Excel";
            Worksheet ws = new Worksheet();

            SheetData sd = new SheetData();
            var list = GetPerson();

            Row r1 = new Row();
            Cell c1 = new Cell();
            Cell c2 = new Cell();
            Cell c3 = new Cell();
            Cell c4 = new Cell();


            c1.DataType = CellValues.String;
            c2.DataType = CellValues.String;
            c3.DataType = CellValues.String;
            c4.DataType = CellValues.String;

            c1.CellValue = new CellValue("Ad");
            c2.CellValue = new CellValue("Soyad");
            c3.CellValue = new CellValue("Adres");
            c4.CellValue = new CellValue("Email");
            r1.Append(c1);
            r1.Append(c2);
            r1.Append(c3);
            r1.Append(c4);

            sd.Append(r1);

            foreach (var item in list.Persons)
            {
                Row r2 = new Row();
                Cell c5 = new Cell();
                Cell c6 = new Cell();
                Cell c7 = new Cell();
                Cell c8 = new Cell();

                c5.DataType = CellValues.String;
                c6.DataType = CellValues.String;
                c7.DataType = CellValues.String;
                c8.DataType = CellValues.String;

                c5.CellValue = new CellValue(item.Name);
                c6.CellValue = new CellValue(item.Surname);
                c7.CellValue = new CellValue(item.Address);
                c8.CellValue = new CellValue(item.Email);

                r2.Append(c5);
                r2.Append(c6);
                r2.Append(c7);
                r2.Append(c8);

                sd.Append(r2);
            }


            ws.Append(sd);
            wsp.Worksheet = ws;
            wsp.Worksheet.Save();
            Sheets sheets = new Sheets();
            Sheet sheet = new Sheet();
            sheet.Name = "first sheet";
            sheet.SheetId = 1;
            sheet.Id = wbp.GetIdOfPart(wsp);
            sheets.Append(sheet);
            wb.Append(fv);
            wb.Append(sheets);

            xl.WorkbookPart.Workbook = wb;
            xl.WorkbookPart.Workbook.Save();
            xl.Close();

            string fileName = "testOpenXml.xlsx";
            Response.Clear();

            byte[] dt = ms.ToArray();

            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.AddHeader("Content-Disposition", string.Format("attachment; filename={0}", fileName));
            Response.BinaryWrite(dt);
            Response.Flush();
            Response.End();

            return View();
        }

    }
}