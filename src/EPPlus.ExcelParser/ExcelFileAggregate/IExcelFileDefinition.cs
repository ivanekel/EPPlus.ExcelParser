using EPPlus.ExcelParser.ExcelDefinitionAggregate;

namespace EPPlus.ExcelParser.ExcelFileAggregate
{
    public interface IExcelFileDefinition : IExcelFile
    {
        IExcelWorksheetDefinition ExcelWorksheetDefinition { get; }
        void AddExcelWorksheetDefinition(IExcelWorksheetDefinition worksheetDefinition);
    }
}