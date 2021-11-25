using OfficeOpenXml;
using System.Collections.Generic;

namespace EPPlus.ExcelParser.Parsing
{
    public class ExcelParserResult<TObject>
    {
        public IEnumerable<TObject> MappedObjects { get; private set; }
        public ExcelPackage ExcelResult { get; set; }
        public bool IsValid { get; set; }

        private ExcelParserResult(IEnumerable<TObject> mappedObjects, ExcelPackage excelPackage, bool isValid)
        {
            MappedObjects = mappedObjects;
            ExcelResult = excelPackage;
            IsValid = isValid;
        }

        public static ExcelParserResult<TObject> CreateNew(IEnumerable<TObject> mappedObjects, ExcelPackage excelPackage,
            bool isValid)
        {
            return new ExcelParserResult<TObject>(mappedObjects, excelPackage, isValid);
        }
    }
}