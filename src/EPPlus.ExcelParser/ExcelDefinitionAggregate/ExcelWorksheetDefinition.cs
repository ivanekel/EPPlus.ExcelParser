using System.Collections.Generic;
using EPPlus.ExcelParser.ExcelColumnDefinitionAggregate;

namespace EPPlus.ExcelParser.ExcelDefinitionAggregate
{
    public class ExcelWorksheetDefinition : IExcelWorksheetDefinition
    {
        public int WorksheetIndex { get; }
        public bool HasHeaders { get; }
        public List<IExcelColumnDefinition> Columns { get; private set; }

        private ExcelWorksheetDefinition(int worksheetIndex, bool hasHeaders)
        {
            WorksheetIndex = worksheetIndex;
            HasHeaders = hasHeaders;
            Columns = new List<IExcelColumnDefinition>();
        }

        public static ExcelWorksheetDefinition NewExcelDefinition(int worksheetIndex = 0, bool hasHeaders = true)
        {
            return new ExcelWorksheetDefinition(worksheetIndex, hasHeaders);
        }
    }
}