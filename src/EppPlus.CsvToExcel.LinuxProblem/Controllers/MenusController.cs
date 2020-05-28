using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;

namespace EppPlus.CsvToExcel.LinuxProblem.Controllers
{
    [ApiController]
    [Route("api/v1/menus")]
    public class MenusController : ControllerBase
    {
        private static readonly IReadOnlyList<Menu> Menus = new List<Menu>
        {
            new Menu("Cheesburger", "Vanilla ice cream and fries"),
            new Menu("Ripeye Steak", "Ice cream sandwich"),
        };

        private const string Delimiter = ";";

        [HttpGet("csv")]
        public IActionResult ExportCsv()
        {
            var csv = GenerateCsv();

            return this.File(Encoding.UTF8.GetBytes(csv), "text/csv", "menus.csv", true);
        }

        [HttpGet("csv-to-excel")]
        public IActionResult ExportExcelFromCsv()
        {
            var csv = GenerateCsv();

            using var pck = new ExcelPackage();

            var ws = pck.Workbook.Worksheets.Add("Sheet1");

            var format = new ExcelTextFormat
            {
                DataTypes = new[]
                {
                    eDataTypes.String, eDataTypes.String
                },
                Delimiter = Delimiter.First(),
                Encoding = new UTF8Encoding(),
            };

            using var range = ws.Cells[1, 1];

            range.LoadFromText(csv, format);

            var bytes = pck.GetAsByteArray();

            return this.File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "menus.xlsx",
                true);
        }

        [HttpGet("list-to-excel")]
        public IActionResult ExportExcelFromList()
        {
            using var pck = new ExcelPackage();

            var ws = pck.Workbook.Worksheets.Add("Sheet1");

            using var range = ws.Cells[1, 1];

            range.LoadFromCollection(Menus, true);

            var bytes = pck.GetAsByteArray();

            return this.File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "menus.xlsx",
                true);
        }

        private static string GenerateCsv()
        {
            var header = $"{nameof(Menu.Main)}{Delimiter}{nameof(Menu.Desert)}";
            var builder = new StringBuilder();
            builder.AppendLine(header);

            foreach (var menu in Menus)
            {
                var row = $"{menu.Main}{Delimiter}{menu.Desert}";
                builder.AppendLine(row);
            }

            return builder.ToString();
        }
    }
}