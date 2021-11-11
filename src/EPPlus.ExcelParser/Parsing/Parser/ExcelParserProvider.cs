using EPPlus.ExcelParser.Parsing.Mapper;
using OfficeOpenXml;
using System;
using System.IO;

namespace EPPlus.ExcelParser.Parsing.Parser
{
    public static class ExcelParserProvider
    {
        public static ExcelParser<TObject> CreateWithInlineMapper<TObject>(
            ExcelPackage excelPackage,
            Func<ExcelInlineRowMapper, TObject> mapper,
            bool hasHeaders = true)
            where TObject : class
        {
            return new ExcelParser<TObject>(excelPackage, hasHeaders).SetMapping(mapper);
        }
        public static ExcelParser<TObject> CreateNew<TObject>(
            ExcelPackage excelPackage,
            bool hasHeaders = true)
            where TObject : class
        {
            return new ExcelParser<TObject>(excelPackage, hasHeaders);
        }

        public static ExcelPackage GetExcelPackageFromStream(Stream stream)
            => new ExcelPackage(stream);
    }
}
