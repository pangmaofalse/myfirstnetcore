using System;
using Microsoft.EntityFrameworkCore.Metadata;
using Microsoft.EntityFrameworkCore.Migrations;

namespace JiangLiQuery.Data.Migrations
{
    public partial class InistallDBContext : Migration
    {
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "AnnualBonus",
                columns: table => new
                {
                    Id = table.Column<int>(nullable: false)
                        .Annotation("SqlServer:ValueGenerationStrategy", SqlServerValueGenerationStrategy.IdentityColumn),
                    IdCard = table.Column<int>(nullable: false),
                    Name = table.Column<string>(nullable: true),
                    Company = table.Column<string>(nullable: true),
                    Department = table.Column<string>(nullable: true),
                    Position = table.Column<string>(nullable: true),
                    Category = table.Column<string>(nullable: true),
                    OutTime = table.Column<DateTime>(nullable: false),
                    AddTime = table.Column<DateTime>(nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_AnnualBonus", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "Payrolls",
                columns: table => new
                {
                    Id = table.Column<int>(nullable: false)
                        .Annotation("SqlServer:ValueGenerationStrategy", SqlServerValueGenerationStrategy.IdentityColumn),
                    IdCard = table.Column<int>(nullable: false),
                    Name = table.Column<string>(nullable: true),
                    Company = table.Column<string>(nullable: true),
                    Department = table.Column<string>(nullable: true),
                    Position = table.Column<string>(nullable: true),
                    Category = table.Column<string>(nullable: true),
                    OutTime = table.Column<DateTime>(nullable: false),
                    PayrollContext = table.Column<string>(nullable: true),
                    AddTime = table.Column<DateTime>(nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_Payrolls", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "RewardDetails",
                columns: table => new
                {
                    Id = table.Column<int>(nullable: false)
                        .Annotation("SqlServer:ValueGenerationStrategy", SqlServerValueGenerationStrategy.IdentityColumn),
                    IdCard = table.Column<int>(nullable: false),
                    DepartmentOne = table.Column<string>(nullable: true),
                    DepartmentTwo = table.Column<string>(nullable: true),
                    Name = table.Column<string>(nullable: true),
                    OutTime = table.Column<DateTime>(nullable: false),
                    Reward = table.Column<string>(nullable: true),
                    AddTime = table.Column<DateTime>(nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_RewardDetails", x => x.Id);
                });
        }

        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "AnnualBonus");

            migrationBuilder.DropTable(
                name: "Payrolls");

            migrationBuilder.DropTable(
                name: "RewardDetails");
        }
    }
}
