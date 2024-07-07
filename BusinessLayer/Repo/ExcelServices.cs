using BusinessLayer.Interfaces;
using ExcelWebApp.Models;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessLayer.Repo
{
    public class ExcelServices : IExcelServices
    {
        public byte[] ProcessExcelFile(IFormFile file)
        {
            var taxRecords = new List<TaxRecord>();
            dynamic streamReturn ;

            using (var stream = new MemoryStream())
            {
                file.CopyTo(stream);
                using (var package = new ExcelPackage(stream))
                {
                    var worksheet = package.Workbook.Worksheets[0];
                    int rowCount = worksheet.Dimension.Rows;
                    for (int row = 2; row <= rowCount; row++)
                    {
                        var totalValueAfterTaxing = decimal.Parse(worksheet.Cells[row, 7].Value.ToString().Trim());
                        var taxingValue = decimal.Parse(worksheet.Cells[row, 8].Value.ToString().Trim());
                        taxRecords.Add(new TaxRecord
                        {
                            InvNo = int.Parse(worksheet.Cells[row, 1].Value.ToString().Trim()),
                            InvCURNo = worksheet.Cells[row, 2].Value.ToString().Trim(),
                            InvDate = DateTime.Parse(worksheet.Cells[row, 3].Value.ToString().Trim()),
                            CustomerCode = int.Parse(worksheet.Cells[row, 4].Value.ToString().Trim()),
                            CustomerName = worksheet.Cells[row, 5].Value.ToString().Trim(),
                            RegCountryAprev = worksheet.Cells[row, 6].Value.ToString().Trim(),
                            TotalValueAfterTaxing = totalValueAfterTaxing,
                            TaxingValue = taxingValue,
                            TotalValueBeforeTaxing = totalValueAfterTaxing / (1 + taxingValue / 100)
                        });
                    }
                }
            }

            // Calculate the sum of "Total Value After Taxing"
            var totalValueAfterTax = taxRecords.Sum(r => r.TotalValueAfterTaxing);

            // Add a new row for the total value
            var totalRecord = new TaxRecord
            {
                InvNo = 0,
                CustomerName = "Total",
                TotalValueAfterTaxing = totalValueAfterTax
            };
            taxRecords.Add(totalRecord);

            // Save the modified sheet
            using (var stream = new MemoryStream())
            {
                using (var package = new ExcelPackage(stream))
                {
                    var worksheet = package.Workbook.Worksheets.Add("Taxes");
                    worksheet.Cells[1, 1].Value = "Inv_No";
                    worksheet.Cells[1, 2].Value = "Inv_CURNo";
                    worksheet.Cells[1, 3].Value = "Inv_Date";
                    worksheet.Cells[1, 4].Value = "Customer Code";
                    worksheet.Cells[1, 5].Value = "Customer Name";
                    worksheet.Cells[1, 6].Value = "REG_COUNTRY_APREV";
                    worksheet.Cells[1, 7].Value = "Total Value After Taxing";
                    worksheet.Cells[1, 8].Value = "Taxing Value";
                    worksheet.Cells[1, 9].Value = "Total Value Before Taxing";

                    for (int row = 0; row < taxRecords.Count; row++)
                    {
                        worksheet.Cells[row + 2, 1].Value = taxRecords[row].InvNo;
                        worksheet.Cells[row + 2, 2].Value = taxRecords[row].InvCURNo;
                        worksheet.Cells[row + 2, 3].Value = taxRecords[row].InvDate.ToString("dd/MM/yyyy");
                        worksheet.Cells[row + 2, 4].Value = taxRecords[row].CustomerCode;
                        worksheet.Cells[row + 2, 5].Value = taxRecords[row].CustomerName;
                        worksheet.Cells[row + 2, 6].Value = taxRecords[row].RegCountryAprev;
                        worksheet.Cells[row + 2, 7].Value = taxRecords[row].TotalValueAfterTaxing;
                        worksheet.Cells[row + 2, 8].Value = taxRecords[row].TaxingValue;
                        worksheet.Cells[row + 2, 9].Value = taxRecords[row].TotalValueBeforeTaxing;
                    }

                    package.Save();
                }
                streamReturn = stream.ToArray();

            }
            return streamReturn;
        }
    }
}
