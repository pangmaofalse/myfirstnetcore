using System;
using System.Collections.Generic;
using System.Text;

namespace JiangLiQuery.Library
{
    public class StringPlus
    {
        public static string ReplaceTrim(string val)
        {
            string result = val.ToString().Replace("_", "").Replace(",", "").Trim();
            return result.Equals("") ? "0" : val;
        }
    }
}
