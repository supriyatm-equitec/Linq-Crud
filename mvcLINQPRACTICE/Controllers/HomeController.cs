using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using mvcLINQPRACTICE.Models;
using PagedList;

using System;
using System.Collections.Generic;
using System.Drawing.Printing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web;
using System.Web.Mvc;

namespace mvcLINQPRACTICE.Controllers
{
    public class HomeController : Controller
    {
        DataClasses1DataContext db = new DataClasses1DataContext();
      

        //insert data in db
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Index(employee emp)
        {
            if (ModelState.IsValid)
            {
                db.employees.InsertOnSubmit(emp);
                ModelState.Clear();
                db.SubmitChanges();

            }
            return View();
        }

        //details





        public ActionResult Show(string searchString, string sortOrder, int? page, int? pageSize)
        {
            // Your existing code...

            var data = db.employees.Where(e => !e.IsDeleted);



            // Apply search
            if (!string.IsNullOrEmpty(searchString))
            {
                data = data.Where(e => e.E_Name.Contains(searchString) ||
                                       e.Gender.StartsWith(searchString)||
                                       //e.Mobile_no.Contains(searchString) ||
                                       e.Email.Contains(searchString) ||
                                       e.Age.ToString().Contains(searchString));
            }

            // Apply sorting
            ViewBag.CurrentSort = sortOrder;
            ViewBag.NameSortParm = sortOrder == "name_desc" ? "name_asc" : "name_desc";
            ViewBag.GenderSortParm = sortOrder == "gender_desc" ? "gender_asc" : "gender_desc";
            //ViewBag.MobileSortParm = sortOrder == "mobile_desc" ? "mobile_asc" : "mobile_desc";
            ViewBag.EmailSortParm = sortOrder == "email_desc" ? "email_asc" : "email_desc";
            ViewBag.AgeSortParm = sortOrder == "age_desc" ? "age_asc" : "age_desc";

            switch (sortOrder)
            {
                case "name_desc":
                    data = data.OrderByDescending(e => e.E_Name);
                    break;
                case "gender_desc":
                    data = data.OrderByDescending(e => e.Gender);
                    break;
                case "gender_asc":
                    data = data.OrderBy(e => e.Gender);
                    break;
                case "mobile_desc":
                    data = data.OrderByDescending(e => e.Mobile_no);
                    break;
                case "mobile_asc":
                    data = data.OrderBy(e => e.Mobile_no);
                    break;
                case "email_desc":
                    data = data.OrderByDescending(e => e.Email);
                    break;
                case "email_asc":
                    data = data.OrderBy(e => e.Email);
                    break;
                case "age_desc":
                    data = data.OrderByDescending(e => e.Age);
                    break;
                case "age_asc":
                    data = data.OrderBy(e => e.Age);
                    break;
                default:
                    data = data.OrderBy(e => e.E_Name);
                    break;
            }

            //// Apply pagination
            //int pageSize = 10;
            //int pageNumber = (page ?? 1);

            //var pagedData = data.ToPagedList(pageNumber, pageSize);

            //return View(pagedData);


            // Apply pagination




            int pageNumber = (page ?? 1);
            int pageSizeValue = pageSize ?? 0;

            // Set pageSize to int.MaxValue if 0 is selected (representing "Show All")
            if (pageSizeValue <= 0)
            {
                pageSizeValue = 10;
            }

            // Populate the dropdown with available page size options
            ViewBag.PageSize = new SelectList(new List<int> { 5, 10, 15, 0 });

            // Assuming 'data' is an IQueryable or IEnumerable, use ToPagedList for pagination
            var pagedData = data.ToPagedList(pageNumber, pageSizeValue);

            return View(pagedData);



        }




        // GET: Home/Details/5
        public ActionResult Details(int id)
        {
            var getempdeatils = db.employees.Where(x => x.id == id).FirstOrDefault();
            if (getempdeatils == null)
            {
                return HttpNotFound();
            }
            return View(getempdeatils);
        }


        //edit data
        public ActionResult Edit(int id)
        {
            var row = db.employees.Where(x => x.id == id).FirstOrDefault();
            if (row == null)
            {
                return HttpNotFound();
            }
            return View(row);
        }
        [HttpPost]

        public ActionResult Edit(int id, employee emp)
        {
            if (ModelState.IsValid)
            {
                employee emp2 = db.employees.Where(x => x.id == id).FirstOrDefault();
                emp2.E_Name = emp.E_Name;
                emp2.Age = emp.Age;
                emp2.Email = emp.Email;
                emp2.Gender = emp.Gender;
                emp2.Mobile_no = emp.Mobile_no;
                db.SubmitChanges();
                return RedirectToAction("show");


            }
            return View();
        }


        //DELETE
        public ActionResult Delete(int id)
        {
            var employee = db.employees.FirstOrDefault(x => x.id == id && !x.IsDeleted);

            if (employee == null)
            {
                return HttpNotFound();
            }

            return View(employee);
        }

        [HttpPost]
        public ActionResult Delete(int id, employee emp)
        {
            try
            {
                var employeeToDelete = db.employees.FirstOrDefault(x => x.id == id && x.IsDeleted==false);

                if (employeeToDelete != null)
                {
                    employeeToDelete.IsDeleted = true;
                    db.SubmitChanges();
                }

                return RedirectToAction("Show");
            }
            catch
            {
                return View();
            }
        }

        //exporttoexcel
        [HttpPost]


        public ActionResult ExportToExcel()
        {
         List<employee> data=db.employees.Where(x => x.IsDeleted==true).ToList();

            using(var workbook =new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("EmpList");

                //add headers to the worksheet
                worksheet.Cell(1, 1).Value = "id";
                worksheet.Cell(1, 2).Value = "Name";
                worksheet.Cell(1, 3).Value = "Gender";
                worksheet.Cell(1, 4).Value = "Mobile_No";
                worksheet.Cell(1, 5).Value = "Email";
                worksheet.Cell(1, 6).Value = "Age";

                //populate the data of db from ro 2
                int row = 2;
                foreach(var employee in data)
                {
                    worksheet.Cell(row, 1).Value = employee.id;
                    worksheet.Cell(row, 2).Value = employee.E_Name;
                    worksheet.Cell(row, 3).Value = employee.Gender;
                    worksheet.Cell(row, 4).Value = employee.Mobile_no;
                    worksheet.Cell(row, 5).Value = employee.Email;
                    worksheet.Cell(row, 6).Value = employee.Age;
                    row++;
                }

                // Save the workbook to a memory stream
                using (var stream = new System.IO.MemoryStream())
                {
                    workbook.SaveAs(stream);
                    stream.Position = 0;

                    // Set the file name for the response
                    var fileName = "EmpList.xlsx";

                    // Return the file as a FileContentResult
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
                }

        



            }
            
           
        }



    }
}

