﻿// <auto-generated />
using System;
using JiangLiQuery.Data;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Infrastructure;
using Microsoft.EntityFrameworkCore.Metadata;
using Microsoft.EntityFrameworkCore.Migrations;
using Microsoft.EntityFrameworkCore.Storage.ValueConversion;

namespace JiangLiQuery.Data.Migrations
{
    [DbContext(typeof(DataContext))]
    [Migration("20200219080937_InistallDBContext")]
    partial class InistallDBContext
    {
        protected override void BuildTargetModel(ModelBuilder modelBuilder)
        {
#pragma warning disable 612, 618
            modelBuilder
                .HasAnnotation("ProductVersion", "2.2.6-servicing-10079")
                .HasAnnotation("Relational:MaxIdentifierLength", 128)
                .HasAnnotation("SqlServer:ValueGenerationStrategy", SqlServerValueGenerationStrategy.IdentityColumn);

            modelBuilder.Entity("JiangLiQuery.Model.Entity.AnnualBonus", b =>
                {
                    b.Property<int>("Id")
                        .ValueGeneratedOnAdd()
                        .HasAnnotation("SqlServer:ValueGenerationStrategy", SqlServerValueGenerationStrategy.IdentityColumn);

                    b.Property<DateTime>("AddTime");

                    b.Property<string>("Category");

                    b.Property<string>("Company");

                    b.Property<string>("Department");

                    b.Property<int>("IdCard");

                    b.Property<string>("Name");

                    b.Property<DateTime>("OutTime");

                    b.Property<string>("Position");

                    b.HasKey("Id");

                    b.ToTable("AnnualBonus");
                });

            modelBuilder.Entity("JiangLiQuery.Model.Entity.Payrolls", b =>
                {
                    b.Property<int>("Id")
                        .ValueGeneratedOnAdd()
                        .HasAnnotation("SqlServer:ValueGenerationStrategy", SqlServerValueGenerationStrategy.IdentityColumn);

                    b.Property<DateTime>("AddTime");

                    b.Property<string>("Category");

                    b.Property<string>("Company");

                    b.Property<string>("Department");

                    b.Property<int>("IdCard");

                    b.Property<string>("Name");

                    b.Property<DateTime>("OutTime");

                    b.Property<string>("PayrollContext");

                    b.Property<string>("Position");

                    b.HasKey("Id");

                    b.ToTable("Payrolls");
                });

            modelBuilder.Entity("JiangLiQuery.Model.Entity.RewardDetails", b =>
                {
                    b.Property<int>("Id")
                        .ValueGeneratedOnAdd()
                        .HasAnnotation("SqlServer:ValueGenerationStrategy", SqlServerValueGenerationStrategy.IdentityColumn);

                    b.Property<DateTime>("AddTime");

                    b.Property<string>("DepartmentOne");

                    b.Property<string>("DepartmentTwo");

                    b.Property<int>("IdCard");

                    b.Property<string>("Name");

                    b.Property<DateTime>("OutTime");

                    b.Property<string>("Reward");

                    b.HasKey("Id");

                    b.ToTable("RewardDetails");
                });
#pragma warning restore 612, 618
        }
    }
}
