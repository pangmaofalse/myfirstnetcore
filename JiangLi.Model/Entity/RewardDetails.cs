using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel.DataAnnotations;

namespace JiangLiQuery.Model.Entity
{
    public class RewardDetails
    {
        [Display(Name ="Id")]
        public int Id { get; set; }

        [Display(Name ="人员编码")]
        public int IdCard { get; set; }

        [Display(Name = "部门名称1")]
        public string DepartmentOne { get; set; }

        [Display(Name = "部门名称2")]
        public string DepartmentTwo { get; set; }

        [Display(Name = "姓名")]
        public string Name { get; set; }

        [Display(Name = "发放日期")]
        public DateTime OutTime { get; set; }

        [Display(Name = "奖励明细")]
        public string Reward { get; set; }

        [Display(Name = "添加时间")]
        public DateTime AddTime { get; set; }
    }
}
