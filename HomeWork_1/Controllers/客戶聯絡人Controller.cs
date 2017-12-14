using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using HomeWork_1.Models;
using ClosedXML.Excel;
using System.Data;
namespace HomeWork_1.Controllers
{
    public class 客戶聯絡人Controller : Controller
    {
        // GET: 客戶聯絡人
       客戶聯絡人Repository repo = RepositoryHelper.Get客戶聯絡人Repository();

        public ActionResult Index(string 姓名, string 職稱)
        {
            var data = repo.All();

            if (!String.IsNullOrEmpty(姓名))
            {
                data = data.Where(p => p.姓名.IndexOf(姓名) > -1);
            }
            if (!String.IsNullOrEmpty(職稱))
            {
                data = data.Where(p => p.職稱.Equals(職稱));
            }
            List<SelectListItem> list = new List<SelectListItem>();
            list.Add(new SelectListItem() { Text = "董事長", Value = "董事長" });
            list.Add(new SelectListItem() { Text = "總經理", Value = "總經理" });
            list.Add(new SelectListItem() { Text = "助理", Value = "助理" });
            ViewData["List"] = list;
            return View(data);
        }

        public ActionResult Delete(int Id)
        {
            var item = repo.Find(Id);
            item.是否刪除 = true;
            repo.UnitOfWork.Commit();
            return RedirectToAction("Index");
        }

        public ActionResult Edit(int Id)
        {
            var items = RepositoryHelper.Get客戶資料Repository().All().ToList().Select(n => new SelectListItem { Text = n.Id.ToString(), Value = n.Id.ToString() });
            ViewData["IdList"] = items;
            var item = repo.Find(Id);
            return View(item);
        }

        [HttpPost]
        //[ValidateAntiForgeryToken]
        public ActionResult Edit(int id, 客戶聯絡人 data)
        {
            var items = RepositoryHelper.Get客戶資料Repository().All().ToList().Select(n => new SelectListItem { Text = n.Id.ToString(), Value = n.Id.ToString() });
            ViewData["IdList"] = items;

            if (ModelState.IsValid)
            {
                var item = repo.Find(id);

                item.客戶Id = data.客戶Id;
                item.職稱 = data.職稱;
                item.姓名 = data.姓名;
                item.Email = data.Email;
                item.手機 = data.手機;
                item.電話 = data.電話;
                repo.UnitOfWork.Commit();

                return RedirectToAction("Index");
            }

            return View(data);
        }


        public ActionResult Details(int Id)
        {
            var item = repo.Find(Id);
            return View(item);
        }

        public ActionResult Create()
        {
            Dictionary<int ,string>dic = RepositoryHelper.Get客戶資料Repository().All().ToDictionary(n=>n.Id,n=>n.Id.ToString());
            //var items= new SelectList(dic, "Value", "Key");
            var items = RepositoryHelper.Get客戶資料Repository().All().ToList().Select(n => new SelectListItem { Text = n.Id.ToString(), Value = n.Id.ToString() });
            ViewData["IdList"]= items;
            return View();
        }
        [HttpPost]
        //[ValidateAntiForgeryToken]
        public ActionResult Create(客戶聯絡人 data)
        {
            var items = RepositoryHelper.Get客戶資料Repository().All().ToList().Select(n => new SelectListItem { Text = n.Id.ToString(), Value = n.Id.ToString() });
            ViewData["IdList"] = items;

            if (ModelState.IsValid)
            {
                repo.Add(data);
                repo.UnitOfWork.Commit();
                return RedirectToAction("Index");
            }

            return View(data);
        }
        public ActionResult FileDW()
        {
            var data = repo.All();
            XLWorkbook workbook = new XLWorkbook();
            workbook.AddWorksheet(ToDataTable.CreateDataTable(data), "Sheet1");
            workbook.SaveAs(@"C:\Users\user\Downloads\客戶聯絡人.xlsx");
            workbook.Dispose();
            return File(@"C:\Users\user\Downloads\客戶聯絡人.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "客戶聯絡人.xlsx");
        }
    }
}