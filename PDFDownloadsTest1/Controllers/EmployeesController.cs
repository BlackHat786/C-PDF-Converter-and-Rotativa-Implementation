using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using PDFDownloadsTest1.Models;
using System.IO;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.html.simpleparser;
using System.Diagnostics;
using System.Drawing.Imaging;
using Rotativa.MVC;

namespace PDFDownloadsTest1.Controllers
{
    public class EmployeesController : Controller
    {
        private PDFDownloadsTest1Context db = new PDFDownloadsTest1Context();

        // GET: Employees
        public ActionResult Index()
        {
            return View(db.Employees.ToList());
        }
        public ActionResult ExportPDF()//Rotativa Implementation
        {
            return new ActionAsPdf("Index");
            
        }
        // GET: Employees/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Employee employee = db.Employees.Find(id);
            if (employee == null)
            {
                return HttpNotFound();
            }
            return View(employee);
        }

        public FileResult createPDF()//iTextSharp Implementation
        {
            MemoryStream workstream = new MemoryStream();
            StringBuilder status = new StringBuilder("");
            DateTime time = DateTime.Now;
            string PDFFileName = string.Format("SamplePdf" + time.ToString("yyyyMMdd") + "-" + ".pdf");//Saving the PDF Name as
            Document Doc = new Document();
            Doc.SetMargins(0f, 0f, 0f,0f);
            PdfPTable tableLayout = new PdfPTable(5);
            Doc.SetMargins(0f, 0f, 0f, 0f);
            
            string strAttachment = Server.MapPath("~/Users/Dynamoe/Documents/visual studio 2017/Projects/PDFDownloadsTest1/PDFDownloadsTest1/App_Data" + PDFFileName);//change this to the file path as to where your project is situated.

            PdfWriter.GetInstance(Doc, workstream).CloseStream = false;
            Doc.Open();
            Doc.Add(Add_Content_To_PDF(tableLayout));
            Doc.Close();

            byte[] byteInfo = workstream.ToArray();
            workstream.Write(byteInfo, 0, byteInfo.Length);
            workstream.Position = 0;

            return File(workstream, "application/pdf", PDFFileName);

        }

        protected PdfPTable Add_Content_To_PDF(PdfPTable tableLayout)//Creates the PDF Table
        {
            float[] headers = { 50, 24, 45, 35, 50 };   
            tableLayout.SetWidths(headers); 
            tableLayout.WidthPercentage = 100;   
            tableLayout.HeaderRows = 1;

            List<Employee> employees = db.Employees.ToList<Employee>();
            //Heading of the Cell
            tableLayout.AddCell(new PdfPCell(new Phrase("Employee Details by Mo'Ameen", new Font(Font.FontFamily.HELVETICA, 25, 2, new iTextSharp.text.BaseColor(0, 0, 0))))
                {
                Colspan = 12, Border = 0, PaddingBottom = 5, HorizontalAlignment = Element.ALIGN_CENTER, BackgroundColor= BaseColor.DARK_GRAY 
        });
            //Column Name
            tableLayout.AddCell(new PdfPCell(new Phrase("Employee Name", new Font(Font.FontFamily.HELVETICA, 8, 1, iTextSharp.text.BaseColor.BLACK)))
            {
                HorizontalAlignment = Element.ALIGN_LEFT,
                Padding = 5,
                BackgroundColor = new BaseColor(255, 255, 255)
            });
            tableLayout.AddCell(new PdfPCell(new Phrase("Surname", new Font(Font.FontFamily.HELVETICA, 8, 1, BaseColor.BLACK)))
            {
                HorizontalAlignment = Element.ALIGN_LEFT,
                Padding = 5,
                BackgroundColor = new BaseColor(255, 255, 255)
            });
            tableLayout.AddCell(new PdfPCell(new Phrase("Gender", new Font(Font.FontFamily.HELVETICA, 8, 1, BaseColor.BLACK)))
            {
                HorizontalAlignment = Element.ALIGN_LEFT,
                Padding = 5,
                BackgroundColor = new BaseColor(255, 255, 255)
            });
            tableLayout.AddCell(new PdfPCell(new Phrase("City", new Font(Font.FontFamily.HELVETICA, 8, 1, BaseColor.BLACK)))
            {
                HorizontalAlignment = Element.ALIGN_LEFT,
                Padding = 5,
                BackgroundColor = new BaseColor(255, 255, 255)
            });
            tableLayout.AddCell(new PdfPCell(new Phrase("Date Hired", new Font(Font.FontFamily.HELVETICA, 8, 1, BaseColor.BLACK)))
            {
                HorizontalAlignment = Element.ALIGN_LEFT,
                Padding = 5,
                BackgroundColor = new BaseColor(255, 255, 255)
            });
            foreach (var emp in employees)
            {    //loops through each record and adds it to the pdf table          
                tableLayout.AddCell(new PdfPCell(new Phrase(emp.EmpName.ToString())));
                tableLayout.AddCell(new PdfPCell(new Phrase(emp.EmpSName)));
                tableLayout.AddCell(new PdfPCell(new Phrase(emp.Gender)));
                tableLayout.AddCell(new PdfPCell(new Phrase(emp.City)));
                tableLayout.AddCell(new PdfPCell(new Phrase(emp.DateHired.ToString())));

            }
            return (tableLayout);//returns the table in a pdf format
        }

        private static void AddCellToHeader(PdfPTable tableLayout, string cellText)
        {

            tableLayout.AddCell(new PdfPCell(new Phrase(cellText, new Font(Font.FontFamily.HELVETICA, 8, 1, BaseColor.YELLOW)))
    {
                HorizontalAlignment = Element.ALIGN_LEFT, Padding = 5, BackgroundColor = new BaseColor(128, 0, 0)
    });
        }

        private static void AddCellToBody(PdfPTable tableLayout, string cellText)
        {
            tableLayout.AddCell(new PdfPCell(new Phrase(cellText, new Font(Font.FontFamily.HELVETICA, 8, 1, BaseColor.BLACK)))
             {
                HorizontalAlignment = Element.ALIGN_LEFT, Padding = 5, BackgroundColor = new BaseColor(255, 255, 255)
    });
        }
        // GET: Employees/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: Employees/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "EmployeeID,EmpName,EmpSName,Gender,City,DateHired")] Employee employee)
        {
            if (ModelState.IsValid)
            {
                db.Employees.Add(employee);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(employee);
        }

        // GET: Employees/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Employee employee = db.Employees.Find(id);
            if (employee == null)
            {
                return HttpNotFound();
            }
            return View(employee);
        }

        // POST: Employees/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "EmployeeID,EmpName,EmpSName,Gender,City,DateHired")] Employee employee)
        {
            if (ModelState.IsValid)
            {
                db.Entry(employee).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(employee);
        }

        // GET: Employees/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Employee employee = db.Employees.Find(id);
            if (employee == null)
            {
                return HttpNotFound();
            }
            return View(employee);
        }

        // POST: Employees/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            Employee employee = db.Employees.Find(id);
            db.Employees.Remove(employee);
            db.SaveChanges();
            return RedirectToAction("Index");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
