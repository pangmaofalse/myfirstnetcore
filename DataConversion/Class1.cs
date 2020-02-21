using System;
using JiangLiQuery.Model.Entity;
using JiangLiQuery.ExcelHelper;
using System.Data;

namespace JiangLiQuery.DataConversion
{
    public class PayrollsConversion
    {
        ExcelRender excelRender = new ExcelRender();
        private string _savepath;

        public PayrollsConversion(string savepath) {
            _savepath = savepath;
        }

        public DataTable Conversion() {
           DataTable dt=excelRender.ExcelToDataTable(_savepath);

           return dt;
        }
    }
}
