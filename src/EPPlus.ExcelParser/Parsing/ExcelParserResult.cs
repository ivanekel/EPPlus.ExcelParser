using OfficeOpenXml;
using System.Collections.Generic;

namespace EPPlus.ExcelParser.Parsing
{
    public class ExcelParserResult<TObject>
    {
        public List<TObject> MappedObjects { get; private set; }
        public ExcelPackage ExcelResult { get; set; }
        public bool IsValid { get; set; }

        private ExcelParserResult(List<TObject> mappedObjects, ExcelPackage excelPackage, bool isValid)
        {
            MappedObjects = mappedObjects;
            ExcelResult = excelPackage;
            IsValid = isValid;
        }

        public static ExcelParserResult<TObject> CreateNew(List<TObject> mappedObjects, ExcelPackage excelPackage,
            bool isValid)
        {
            return new ExcelParserResult<TObject>(mappedObjects, excelPackage, isValid);
        }
    }
}