using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using HomeWork_1.Models;
using ClosedXML.Excel;

namespace HomeWork_1.Controllers
{
    public class 客戶資料Controller : Controller
    {
        客戶資料Repository repo = RepositoryHelper.Get客戶資料Repository();

        public ActionResult Index(string strName,string strType)
        {
            var data = repo.All();

            if (!String.IsNullOrEmpty(strName))
            {
                data = data.Where(p=>p.客戶名稱.IndexOf(strName)>-1);
            }
            if (!String.IsNullOrEmpty(strType))
            {
                data = data.Where(p => p.客戶分類.Equals(strType));
            }
            List<SelectListItem> list = new List<SelectListItem>();
            list.Add(new SelectListItem() { Text = "一般客戶", Value = "一般客戶" });
            list.Add(new SelectListItem() { Text = "公司客戶", Value = "公司客戶" });
            list.Add(new SelectListItem() { Text = "政府機構", Value = "政府機構" });
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
            var item = repo.Find(Id);
            return View(item);
        }

        [HttpPost]
        //[ValidateAntiForgeryToken]
        public ActionResult Edit(int id,客戶資料 data)
        {
            if (ModelState.IsValid)
            {
                var item = repo.Find(id);

                item.客戶名稱 = data.客戶名稱;
                item.統一編號 = data.統一編號;
                item.電話 = data.電話;
                item.傳真 = data.傳真;
                item.地址 = data.地址;
                item.Email = data.Email;
                item.客戶分類 = data.客戶分類;
                repo.UnitOfWork.Commit();

                return RedirectToAction("Index");
            }

            return View(data);
        }


        public ActionResult Detail(int Id)
        {
            var item = repo.Find(Id);
            return View(item);
        }

        public ActionResult Create()
        {
            return View();
        }
        [HttpPost]
        //[ValidateAntiForgeryToken]
        public ActionResult Create(客戶資料 data)
        {
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