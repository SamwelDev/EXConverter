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

        public byte[] ConvertExcelToPdf(Stream excelStream)
        {
            using var workbook = new XLWorkbook(excelStream);
            var worksheet = workbook.Worksheets.First();

            var sb = new StringBuilder();
            sb.Append("<html><head><style>");
            sb.Append("table { border-collapse: collapse; width: 100%; font-family: Arial; }");
            sb.Append("td, th { border: 1px solid #ccc; padding: 6px; font-size: 12px; }");
            sb.Append("th { background-color: #f2f2f2; }");
            sb.Append("</style></head><body>");
            sb.Append("<table>");

            foreach (var row in worksheet.RowsUsed())
            {
                sb.Append("<tr>");
                foreach (var cell in row.CellsUsed())
                {
                    sb.Append($"<td>{cell.GetValue<string>()}</td>");
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
