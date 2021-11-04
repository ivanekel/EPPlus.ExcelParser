using System.Collections.Generic;

namespace EPPlus.ExcelParser.ExcelParserResultAggregate
{
    public interface IExcelParserResult<T>
    {
        List<T> ResultList { get; }
        bool ContainsValidationErrors { get; }
        byte[] ValidatedExcelFileArray { get; }
    }
}