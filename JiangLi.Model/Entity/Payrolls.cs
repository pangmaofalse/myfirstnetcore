using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel.DataAnnotations;

namespace JiangLiQuery.Model.Entity
{
    public class Payrolls
    {
        [Display(Name="Id")]
        public int Id { get; set; }

        [Display(Name="人员编号")]
        public int IdCard { get; set; }

        [Display(Name="姓名")]
        public string Name { get; set; }

        [Display(Name="所属公司")]
        public string Company { get; set; }

        [Display(Name="所属部门")]
        public string Department { get; set; }

        [Display(Name="所属岗位")]
        public string Position { get; set; }

        [Display(Name="人员类别")]
        public string Category { get; set; }

        [Display(Name="发放日期")]
        public DateTime OutTime { get; set; }

        [Display(Name="工资内容")]
        public string PayrollContext { get; set; }

        [Display(Name="导入日期")]
        public DateTime AddTime { get; set; }
    }
}
