using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.AspNetCore.Mvc;
using System.Web;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Hosting.Internal;
using Microsoft.AspNetCore.Hosting;
using System.IO;
using JiangLiQuery.Library;
using JiangLiQuery.Model.View.Response;
using JiangLiQuery.FileUpload;
using JiangLiQuery.DataConversion;
using System.Data;

namespace JiangLiQuery.Controllers
{
    public class PayrollsController :Controller
    {
        private readonly IHostingEnvironment _hostingEnvironment;

        public PayrollsController(IHostingEnvironment hostingEnvironment) {
            _hostingEnvironment = hostingEnvironment;
        }
        public IActionResult Index() {
           
            return View();
        }


        [HttpGet]
        public IActionResult Import() {
            ResultModel result = new ResultModel();
            result.Code = 0;
            result.Message = "Excel Import";
            return View(result);
        }

        [HttpPost]
        public IActionResult Import(IFormFile file) {
            ResultModel result = new ResultModel();
            if (file.Length > 0)
            {
                string fileName = string.Format("{0}_{1}", DateTime.Now.ToString("MMddHHmmss"), file.FileName);
                string SavePath = Path.Combine(_hostingEnvironment.WebRootPath, Path.Combine("importfiles", fileName));
                var fi = new FileImport(file, SavePath);
                result = fi.Import();
                if (result.Code == 200) {
                    PayrollsConversion pc = new PayrollsConversion(SavePath);
                    DataTable dt= pc.Conversion();
                }
            }
            else {
                result.Code = 0;
                result.Message = "请选择要上传的文件！";
            }

           return View(result);
        }
    }
}
