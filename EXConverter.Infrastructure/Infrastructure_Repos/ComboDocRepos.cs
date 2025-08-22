using EXConverter.Application.Application_Serv;
using Microsoft.Extensions.Logging;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.PortableExecutable;
using System.Text;
using System.Threading.Tasks;

namespace EXConverter.Infrastructure.Infrastructure_Repos
{
    public class ComboDocRepos : IComboDocService
    {
        ILogger<ComboDocRepos> _logger;
        public byte[] MergeDocuments(List<byte[]> documents)
        {
            var outputDocument = new PdfDocument();

            int totalDocuments = documents?.Count ?? 0;
            Console.WriteLine($"Merging {totalDocuments} documents...");

            foreach (var docBytes in documents)
            {
                if (docBytes == null || docBytes.Length == 0)
                {
                    _logger.LogInformation("Skipped null or empty document.");
                    continue;
                }

                try
                {
                    using var stream = new MemoryStream(docBytes);
                    var inputDocument = PdfReader.Open(stream, PdfDocumentOpenMode.Import);

                    if (inputDocument.PageCount == 0)
                    {
                        _logger.LogInformation("Skipped document with 0 pages.");
                        continue;
                    }

                    for (int i = 0; i < inputDocument.PageCount; i++)
                    {
                        outputDocument.AddPage(inputDocument.Pages[i]);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError($"Skipped invalid PDF ", ex.Message);
                }
            }

            if (outputDocument.PageCount == 0)
            {
                throw new InvalidOperationException("Cannot merge documents: no valid pages found.");
            }

            using var outputStream = new MemoryStream();
            outputDocument.Save(outputStream);
            return outputStream.ToArray();
        }



        public List<byte[]> SplitDocument(byte[] documentBytes, List<int> pageIndices)
        {
            var result = new List<byte[]>();

            using var stream = new MemoryStream(documentBytes);
            var inputDocument = PdfReader.Open(stream, PdfDocumentOpenMode.Import);
            var pageCount = inputDocument.PageCount;

            Console.WriteLine($"PDF has {pageCount} pages");
            Console.WriteLine($"Requested split at: {string.Join(",", pageIndices)}");

            // Sanity check
            var splitPoints = pageIndices
                .Where(i => i > 0 && i < pageCount) 
                .OrderBy(i => i)
                .Distinct()
                .ToList();

            splitPoints.Add(pageCount); 

            int start = 0;

            foreach (var end in splitPoints)
            {
                if (start >= end) continue;

                var outputDocument = new PdfDocument();

                for (int i = start; i < end; i++)
                {
                    outputDocument.AddPage(inputDocument.Pages[i]);
                }

                using var outputStream = new MemoryStream();
                outputDocument.Save(outputStream);
                result.Add(outputStream.ToArray());

                _logger.LogInformation($"Created part: pages {start} to {end - 1}");
                start = end;
            }

            return result;
        }
    }
}
