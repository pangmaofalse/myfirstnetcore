using System;
using System.Collections.Generic;
using System.Text;

namespace JiangLiQuery.Model.View.Response
{
    public class ResultModel
    {
        public ResultModel() {

        }


        public ResultModel(int code, string message) {
            Code = code;
            Message = message;
        }

        public ResultModel(int code,string message,object data) {
            Code = code;
            Message = message;
            Data = data;
        }

        public int Code { get; set; }
        public string Message { get; set; }
        public object Data { get; set; }
    }
}
