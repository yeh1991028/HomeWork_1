using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using HomeWork_1.Models;
using ClosedXML.Excel;

namespace HomeWork_1.Controllers
{
    public class 客戶銀行資訊Controller : Controller
    {
        客戶銀行資訊Repository repo = RepositoryHelper.Get客戶銀行資訊Repository();

        public ActionResult Index(int? Id)
        {
            var data = repo.All();

            if (Id!=null)
            {
                data = data.Where(p => p.客戶Id == Id);
            }
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
        public ActionResult Edit(int id, 客戶銀行資訊 data)
        {
            var items = RepositoryHelper.Get客戶資料Repository().All().ToList().Select(n => new SelectListItem { Text = n.Id.ToString(), Value = n.Id.ToString() });
            ViewData["IdList"] = items;

            if (ModelState.IsValid)
            {
                var item = repo.Find(id);

                item.Id = data.Id;
                item.客戶Id = data.客戶Id;
                item.銀行名稱 = data.銀行名稱;
                item.銀行代碼 = data.銀行代碼;
                item.分行代碼 = data.分行代碼;
                item.帳戶名稱 = data.帳戶名稱;
                item.帳戶號碼 = data.帳戶號碼;
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
            var items = RepositoryHelper.Get客戶資料Repository().All().ToList().Select(n => new SelectListItem { Text = n.Id.ToString(), Value = n.Id.ToString() });
            ViewData["IdList"] = items;

            return View();
        }
        [HttpPost]
        //[ValidateAntiForgeryToken]
        public ActionResult Create(客戶銀行資訊 data)
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
            workbook.SaveAs(@"C:\Users\user\Downloads\客戶資料.xlsx");
            workbook.Dispose();
            return File(@"C:\Users\user\Downloads\客戶資料.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "客戶資料.xlsx");
        }
    }
}
