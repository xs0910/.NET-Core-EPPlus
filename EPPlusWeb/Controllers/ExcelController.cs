using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using EPPlusCommon;
using Microsoft.AspNetCore.Hosting;
using System.IO;
using EPPlusWeb.Models;

namespace EPPlusWeb.Controllers
{
    public class ExcelController : Controller
    {
        private readonly IHostingEnvironment _hosting;
        public ExcelController(IHostingEnvironment hosting)
        {
            _hosting = hosting;
        }
        public IActionResult Import()
        {
            string folder = _hosting.WebRootPath;
            string fileName = Path.Combine(folder, "Excel", "Test.xlsx");
            bool result = EPPlusHelper.ImportExcel(fileName, ExcelData.GetExcelData());
            string str = result ? "导入Excel成功:" + fileName : "导入失败";
            return Content(str);
        }
        public IActionResult Read()
        {
            string folder = _hosting.WebRootPath;
            string fileName = Path.Combine(folder, "Excel", "Test.xlsx");
            string result = EPPlusHelper.ReadExcel(fileName);
            return Content(result);
        }
    }
}