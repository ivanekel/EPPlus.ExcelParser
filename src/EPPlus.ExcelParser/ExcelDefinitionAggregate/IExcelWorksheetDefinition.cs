using System.Collections.Generic;
using EPPlus.ExcelParser.ExcelColumnDefinitionAggregate;

namespace EPPlus.ExcelParser.ExcelDefinitionAggregate
{
    public interface IExcelWorksheetDefinition
    {
        int WorksheetIndex { get; }
        bool HasHeaders { get; }
        List<IExcelColumnDefinition> Columns { get; }
    }
}