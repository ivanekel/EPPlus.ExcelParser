using System.Collections.Generic;

namespace EPPlus.ExcelParser
{
    public class ExcelParserResult<TObject>
    {
        public List<TObject> MappedObjects { get; private set; }
        public byte[] ExcelResult { get; set; }
        public bool IsValid { get; set; }

        private ExcelParserResult(List<TObject> mappedObjects, byte[] excelFileByteResult, bool isValid)
        {
            MappedObjects = mappedObjects;
            ExcelResult = excelFileByteResult;
            IsValid = isValid;
        }

        public static ExcelParserResult<TObject> CreateNew(List<TObject> mappedObjects, byte[] excelFileByteResult,
            bool isValid)
        {
            return new ExcelParserResult<TObject>(mappedObjects, excelFileByteResult, isValid);
        }
    }
}