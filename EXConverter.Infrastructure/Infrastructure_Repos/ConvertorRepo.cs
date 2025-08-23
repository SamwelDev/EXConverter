using ClosedXML.Excel;
using DinkToPdf;
using DinkToPdf.Contracts;
using EXConverter.Application.Application_Serv;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static EXConverter.Infrastructure.Infrastructure_Repos.ConvertorRepo;
using static System.Net.Mime.MediaTypeNames;

namespace EXConverter.Infrastructure.Infrastructure_Repos
{
    public class ConvertorRepo : IConService
    {
        private readonly IConverter _converter;

        public ConvertorRepo(IConverter converter)
        {
            _converter = converter;
        }

        public string ConvertExcelToCsv(Stream excelStream, string delimiter = ",")
        {
            using var workbook = new XLWorkbook(excelStream);
            var worksheet = workbook.Worksheets.First();

            var sb = new StringBuilder();
            int columnCount = worksheet.LastColumnUsed().ColumnNumber();

            foreach (var row in worksheet.RowsUsed())
            {
                var cells = Enumerable.Range(1, columnCount)
                                      .Select(i => EscapeCsvValue(row.Cell(i).GetValue<string>()))
                                      .ToArray();
                sb.AppendLine(string.Join(delimiter, cells));
            }

            return sb.ToString();
        }

        public byte[] GetCsvBytes(Stream excelStream, string delimiter = ",")
        {
            var csvContent = ConvertExcelToCsv(excelStream, delimiter);
            var bom = new byte[] { 0xEF, 0xBB, 0xBF }; // ut-8
            var csvBytes = Encoding.UTF8.GetBytes(csvContent);
            return bom.Concat(csvBytes).ToArray();
        }

        private string EscapeCsvValue(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";

            bool mustQuote = value.Contains(",") || value.Contains("\"") || value.Contains("\n");
            if (mustQuote)
            {
                value = value.Replace("\"", "\"\"");
                return $"\"{value}\"";
            }
            return value;
        }
        public byte[] ConvertExcelToPdf(Stream excelStream)
        {
            using var workbook = new XLWorkbook(excelStream);
            var worksheet = workbook.Worksheets.First();

            var sb = new StringBuilder();
            sb.Append("<html><head><style>");
            sb.Append("table { border-collapse: collapse; width: 100%; font-family: Arial; }");
            sb.Append("td, th { border: 1px solid #ccc; padding: 6px; font-size: 12px; }");
            sb.Append("th { background-color: #f2f2f2; }");
            sb.Append("</style></head><body><table>");

            foreach (var row in worksheet.RowsUsed())
            {
                sb.Append("<tr>");
                foreach (var cell in row.Cells(1, worksheet.LastColumnUsed().ColumnNumber()))
                {
                    sb.Append($"<td>{System.Net.WebUtility.HtmlEncode(cell.GetValue<string>())}</td>");
                }
                sb.Append("</tr>");
            }

            sb.Append("</table></body></html>");
            string htmlContent = sb.ToString();

            var doc = new HtmlToPdfDocument()
            {
                GlobalSettings = new GlobalSettings
                {
                    PaperSize = PaperKind.A4,
                    Orientation = Orientation.Portrait
                },
                Objects = {
                new ObjectSettings
                {
                    HtmlContent = htmlContent,
                    WebSettings = { DefaultEncoding = "utf-8" }
                }
            }
            };

            return _converter.Convert(doc);
        }
    }
}
