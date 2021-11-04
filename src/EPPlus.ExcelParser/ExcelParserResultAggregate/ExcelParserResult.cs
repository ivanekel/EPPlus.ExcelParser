using System.Collections.Generic;

namespace EPPlus.ExcelParser.ExcelParserResultAggregate
{
    public class ExcelParserResult<T> : IExcelParserResult<T>
    {
        public List<T> ResultList { get; private set; }
        public bool ContainsValidationErrors { get; private set; }
        public byte[] ValidatedExcelFileArray { get; private set; }

        private ExcelParserResult(List<T> resultList, bool containsErrors, byte[] resultFileArray)
        {
            ResultList = resultList;
            ContainsValidationErrors = containsErrors;
            ValidatedExcelFileArray = resultFileArray;
        }

        public static ExcelParserResult<T> CreateInstance(List<T> resultList, bool containsErrors, byte[] resultFileArray)
        {
            return new ExcelParserResult<T>(resultList, containsErrors, resultFileArray);
        }
    }
}