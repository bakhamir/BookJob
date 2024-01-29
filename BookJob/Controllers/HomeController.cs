using BookJob.Models;
using ClosedXML.Excel;
using CsvHelper;
using CsvHelper.Configuration;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using System.Formats.Asn1;
using System.Globalization;
using System.Text;

namespace BookJob.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index(string query)
        {
            return View();
        }
        public IActionResult Book()
        {
            BookJob.Models.book book = new BookJob.Models.book();
            IEnumerable<book> model = book.GetBookByName("Dance");
            return View(model);
        }

        [HttpGet, Route("JsonIndex")]
        public ActionResult JsonIndex()
        {
            BookJob.Models.book book = new BookJob.Models.book();
            IEnumerable<book> model = book.GetBookByName("Dance");
            return Json(model);
        }

        [HttpGet, Route("GetExcel")]
        public ActionResult GetExcel()
        {
            using (var ms = new MemoryStream())
            {
                using (XLWorkbook wb = new XLWorkbook())
                {
                    var ws = wb.AddWorksheet("report");
                    BookJob.Models.book book = new BookJob.Models.book();
                    IEnumerable<book> lst = book.GetBookByName("Dance");
          

                    ws.Cell(2, 1).InsertData(lst);
                    ws.RangeUsed().SetAutoFilter();
                    ws.Columns("A", "B").AdjustToContents();

                    ws.SheetView.FreezeRows(1);
                    wb.SaveAs(ms);
                    ms.Position = 0;
                    ms.Flush();
                    var bytes = ms.ToArray();

                    return File(bytes,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    "report____" + DateTime.Now.ToString("ddMMyyyy_hhmmss") + ".xlsx");
                }
            }

        }
        [HttpGet, Route("GetCsv")]
        public ActionResult GetCsv()
        {
            var csvConfig = new CsvConfiguration(CultureInfo.CurrentCulture)
            {
                HasHeaderRecord = true,
                Delimiter = ",",
                Encoding = Encoding.UTF8
            };

            BookJob.Models.book book = new BookJob.Models.book();
            IEnumerable<book> lst = book.GetBookByName("Dance");
            MemoryStream memoryStream = new();

            // We *do* want to dispose of the StreamWriter
            using StreamWriter streamWriter = new(memoryStream, leaveOpen: true);

            // I assume that CsvWriter implements IDisposable too
            using CsvWriter csvWriter = new(streamWriter, CultureInfo.InvariantCulture);
            {
                csvWriter.WriteRecords(lst);
                streamWriter.Flush();
                    memoryStream.Position = 0;
                    return File(memoryStream, "text/csv", "exported_items.csv");
                }
        }
        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}