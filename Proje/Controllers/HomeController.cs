using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ClosedXML.Excel;
using Proje.Models;

namespace Proje.Controllers
{
    public class HomeController : Controller
    {

        private DatabaseLogicLayer dbLogicLayer;

        public HomeController()
        {
            dbLogicLayer = DatabaseLogicLayer.Instance;
        }

        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Upload(HttpPostedFileBase file)
        {
            if (file != null && file.ContentLength > 0)
            {
                try
                {
                    using (XLWorkbook workbook = new XLWorkbook(file.InputStream))
                    {
                        IXLWorksheet worksheet = workbook.Worksheet(1);

                        var firstCell = worksheet.FirstCellUsed();
                        var lastCell = worksheet.LastCellUsed();
                        var range = worksheet.Range(firstCell.Address, lastCell.Address);

                        int rowCount = worksheet.RowsUsed().Count();
                        int columnCount = worksheet.ColumnsUsed().Count();
                        dbLogicLayer.CreateTableWithColumnCount(columnCount);

                        foreach (var row in range.RowsUsed())
                        {
                            List<String> values = new List<String>();
                            foreach (var cell in row.Cells())
                            {
                                values.Add(cell.Value.ToString());
                            }
                            dbLogicLayer.InsertDataToDatabase(values);
                        }
                    }

                    TempData["Message"] = "Excel dosyası başarıyla yüklendi ve okundu!";
                }
                catch (Exception ex)
                {
                    TempData["Message"] = "Excel dosyası okunurken bir hata oluştu: " + ex.Message;
                }
            }
            else
            {
                TempData["Message"] = "Lütfen bir Excel dosyası seçin!";
            }

            return RedirectToAction("Index", "Home");
        }

        [HttpGet]
        public ActionResult listsOfData(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return RedirectToAction("Index", "Home");
            }
            ViewBag.nameOfTable = name;
            List<List<string>> tableData = dbLogicLayer.GetTableData(name);
            return View(tableData);
        }

        public ActionResult DownloadExcel(string tableName)
        {
            FileContentResult fileResult = dbLogicLayer.DownloadTableDataAsExcel(tableName);
            return fileResult;
        }

        [HttpGet]
        public ActionResult updateData(string tableName, int? index)
        {
            if (string.IsNullOrEmpty(tableName) || !index.HasValue)
            {
                return RedirectToAction("Index", "Home");
            }
            List<string> tableData = dbLogicLayer.GetItemById(tableName, (int)index);
            ViewBag.tableName = tableName;
            ViewBag.index = index;
            return View(tableData);
        }

        [HttpPost]
        public ActionResult updateData(List<string> data, string tableName, string index)
        {
            dbLogicLayer.UpdateItemById(tableName, int.Parse(index), data);
            return RedirectToAction("listsOfData", "Home", new { name = tableName });
        }

        [HttpGet]
        public ActionResult deleteData(string tableName, int? index)
        {
            if (string.IsNullOrEmpty(tableName) || !index.HasValue)
            {
                return RedirectToAction("Index", "Home");
            }
            dbLogicLayer.DeleteItemById(tableName, (int)index);
            return RedirectToAction("listsOfData", "Home", new { name = tableName });
        }




    }
}