using System;
using System.IO;

namespace EPPlus.ExcelParser.ExcelFileAggregate
{
    public class ExcelFile : IExcelFile
    {
        public Stream ExcelFileStream { get; private set; }

        public ExcelFile(Stream excelFileStream)
        {
            if (excelFileStream == null || excelFileStream.Length == 0)
            {
                throw new ArgumentException("must not be null or empty", nameof(excelFileStream));
            }

            ExcelFileStream = excelFileStream;
        }
    }
}