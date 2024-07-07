using BusinessLayer.Interfaces;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

public class ExcelController : Controller
{
    private readonly IExcelServices _excelSheet;

    // Dependency injection
    public ExcelController(IExcelServices excelSheet)
    {
        _excelSheet = excelSheet;
    }
    public IActionResult Index()
    {
        return View();
    }

    [HttpPost]
    public IActionResult Import(IFormFile file)
    {
        // Check If the File Empty 
        if (file == null || file.Length == 0)
        {
            ViewBag.Message = "Invalid file.";
            return View("Index");
        }
        // Using The Services In Dependency injection
        var result = _excelSheet.ProcessExcelFile(file);
        var fileResult = File(result, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ModifiedTaxes.xlsx");
        return fileResult;
        
    }
}