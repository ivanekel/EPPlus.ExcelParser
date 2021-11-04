using System.IO;
using EPPlus.ExcelParser.ExcelDefinitionAggregate;

namespace EPPlus.ExcelParser.ExcelFileAggregate
{
    public class ExcelFileDefinition : IExcelFileDefinition
    {
        public Stream ExcelFileStream { get; private set; }
        public IExcelWorksheetDefinition ExcelWorksheetDefinition { get; private set; }

        public ExcelFileDefinition(Stream excelFileStream, IExcelWorksheetDefinition excelWorksheetDefinition)
        {
            ExcelFileStream = excelFileStream;
            ExcelWorksheetDefinition = excelWorksheetDefinition;
        }

        public void AddExcelWorksheetDefinition(IExcelWorksheetDefinition worksheetDefinition)
        {
            ExcelWorksheetDefinition = worksheetDefinition;
        }
    }
}