﻿using System;
using System.Collections.Generic;
using System.Text;
using JiangLiQuery.IServices;
using JiangLiQuery.Model.Entity;
using JiangLiQuery.Data;
using System.Linq;
using System.Threading.Tasks;

namespace JiangLiQuery.Services
{
    public class PayrollsService : IService<Payrolls>
    {
        private readonly DataContext _context;

        public PayrollsService(DataContext context) {
            _context = context;
        }

        public Payrolls Delete(Payrolls newModel)
        {
             _context.Payrolls.Remove(newModel);
            _context.SaveChanges();
            return newModel;
        }

        public Payrolls Install(Payrolls newModel)
        {
            _context.Payrolls.Add(newModel);
            _context.SaveChanges();
            return newModel;
        }

        public IEnumerable<Payrolls> Install(IEnumerable<Payrolls> list) {

            foreach (Payrolls p in list) {
                _context.Payrolls.Add(p);
            }
            _context.SaveChanges();

            return list;
        }

        public Payrolls Modify(Payrolls newModel)
        {
            _context.Payrolls.Update(newModel);
            _context.SaveChanges();
            return newModel;
        }

        public IEnumerable<Payrolls> Query()
        {
            return _context.Payrolls.ToList();
        }

        public Payrolls Query(int id)
        {
            var result = from payroll in _context.Payrolls
                         select new
                         {
                             id = payroll.IdCard
                         };


            //return _context.Payrolls.Find(id);

           return _context.Payrolls.FirstOrDefault(p => p.IdCard == id);
        }
    }
}
