﻿using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Text;

namespace JiangLiQuery.Controllers
{
    public class HomeController:Controller
    {
        public IActionResult Index() {
            return View();
        }
    }
}
