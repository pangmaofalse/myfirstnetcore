using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using System;
using JiangLiQuery.Library;
using System.IO;
using JiangLiQuery.Model.View.Response;

namespace JiangLiQuery.FileUpload
{
    public class FileImport
    {
        private IFormFile _iformFile;
        private string _savePath;

        private string[] _arrformat ={ ".xls", ".et", ".cvs"} ;

        public FileImport() {

        }
        public FileImport(IFormFile iFormFile,string savePath) {
            _iformFile = iFormFile;
            _savePath = savePath;
        }

        public ResultModel Import() {
            return Import(_iformFile, _savePath);
        }

        public ResultModel Import(IFormFile formFile, string savePath) {
            if (formFile != null)
            {
                
                if (formFile.Length > 0)
                {
                    string filename = formFile.FileName;
                    string fileExt = FileHelper.GetExtension(filename);
                    if (IsExtension(fileExt))
                    {
                        //string fileName = string.Format("{0}_{1}", DateTime.Now.ToString("MMddHHmmss"), _iformFile.FileName);

                        //string fileNamePath = Path.Combine("importfiles", fileName);

                        // string SavePath = Path.Combine(_webRootPath, fileNamePath);

                        using (FileStream fs = new FileStream(savePath, FileMode.CreateNew))
                        {
                            formFile.CopyTo(fs);
                            fs.Flush();
                        }
                        return new ResultModel(200, "上传成功！");
                    }
                    else
                    {
                        return new ResultModel(0, "上传失败！请检查上传文件格式是否正确！");
                    }
                }
                return new ResultModel(0, "请选择上传文件！");
            }
            else {
                return new ResultModel(0, "请传入文件！");
            }
        }

        private bool IsExtension(string fileExt) {
            foreach (string f in _arrformat) {
                if (f.Equals(fileExt)) {
                    return true;
                }
            }
            return false;
        }
    }
}
