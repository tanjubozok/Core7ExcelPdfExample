using Core7ExcelPdfExample.Business;
using Core7ExcelPdfExample.Dtos.Categories;
using Core7ExcelPdfExample.Helpers;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Globalization;
using System.Text;

namespace Core7ExcelPdfExample.Controllers;

public class ExcelController : Controller
{
    private readonly ICategoryService _categoryService;

    public ExcelController(ICategoryService categoryService)
    {
        _categoryService = categoryService;
    }

    public IActionResult Index()
    {
        return View();
    }

    [HttpPost]
    public IActionResult ExportExcel()
    {
        //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        string reportname = $"CategoryList_{Guid.NewGuid():N}.xlsx";
        var list = _categoryService.CategoryList();
        if (list.Count > 0)
        {
            var exportBytes = Export.Excel<CategoryListDto>(list, reportname);
            return File(exportBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", reportname);
        }
        else
        {
            TempData["Message"] = "No Data to Export";
            return View();
        }
    }

    [HttpPost]
    public IActionResult CreateExcel()
    {
        using var package = new ExcelPackage();
        
        var worksheet = package.Workbook.Worksheets.Add("ExchangeRate");
        
        //add headers
        worksheet.Cells[1, 1].Value = "Amount";
        worksheet.Cells[1, 2].Value = "Currency";
        worksheet.Cells[1, 3].Value = "Amount";
        worksheet.Cells[1, 4].Value = "Currency";
        
        //Add some items...
        worksheet.Cells["A2"].Value = 1;
        worksheet.Cells["B2"].Value = "USD";
        worksheet.Cells["C2"].Value = 7.8;
        worksheet.Cells["D2"].Value = "HKD";
        
        //Add some items...
        worksheet.Cells["A3"].Value = 1;
        worksheet.Cells["B3"].Value = "GDB";
        worksheet.Cells["C3"].Value = 10.3;
        worksheet.Cells["D3"].Value = "HKD";
        worksheet.Cells["C2:C3"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        worksheet.Cells["C2:C3"].Style.Font.Bold = true;
        
        var xlFile = new FileInfo(@"C:\Temp\CreateExcel.xlsx");
        package.SaveAs(xlFile);//save the Excel file
        
        return View();
    }

    [HttpPost]
    public async Task<IActionResult> ExportCsv()
    {
        var filePath = new FileInfo(@"C:\Temp\ExportCsv.xlsx");
        if (filePath != null)
        {
            using var package = new ExcelPackage(filePath);

            var firstSheet = package.Workbook.Worksheets.First(); //first worksheet
            var endRow = firstSheet.Dimension.End.Row; //total no of Rows
            var endCol = firstSheet.Dimension.End.Column; //total no of Column​

            var endColLetter = ((char)(endCol + 'A' - 1)).ToString() + endRow.ToString();

            var format = new ExcelOutputTextFormat
            {
                Delimiter = ';',
                Culture = new CultureInfo("en-GB"),
                Encoding = new UTF8Encoding(),
            };

            var csvFile = new FileInfo(@"C:\Temp\exportToCSV.csv");
            await firstSheet.Cells["A1:" + endColLetter].SaveToTextAsync(csvFile, format);
        }
        return View();
    }

    [HttpPost]
    public IActionResult ReadExcel()
    {
        var filePath = new FileInfo(@"C:\Temp\ReadExcel.xlsx");

        if (filePath != null)
        {
            using var package = new ExcelPackage(filePath);

            var firstSheet = package.Workbook.Worksheets.First(); //first worksheet
            var endRow = firstSheet.Dimension.End.Row; //total no of Rows
            var endCol = firstSheet.Dimension.End.Column; //total no of Column

            //first row is the header so we skip that and start from 2
            for (int i = 2; i <= endRow; i++)
            {
                for (int j = 1; j <= endCol; j++)
                {
                    var value = firstSheet.Cells[i, j].Value.ToString();
                }
            }
        }
        return View();
    }
}