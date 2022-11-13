using Microsoft.AspNetCore.Mvc;

namespace Core7ExcelPdfExample.Controllers;

public class PdfController : Controller
{
    public IActionResult Index()
    {
        return View();
    }
}
