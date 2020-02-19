using System;
using System.Collections.Generic;
using System.Text;
using JiangLiQuery.Model.Entity;
using Microsoft.EntityFrameworkCore;

namespace JiangLiQuery.Data
{
    public class DataContext :DbContext
    {
        private readonly string defaultConnection="Data Source=.;Initial Catalog=JiangLiSystemDB;Integrated Security=True;Connect Timeout=30;Encrypt=False;TrustServerCertificate=False;ApplicationIntent=ReadWrite;MultiSubnetFailover=False";
        public DataContext(DbContextOptions<DataContext> options) {

        }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlServer(defaultConnection);
        }

        public DbSet<Payrolls> Payrolls { get; set; }
        public DbSet<RewardDetails> RewardDetails { get; set; }
        public DbSet<AnnualBonus> AnnualBonus { get; set; }
    }
}
