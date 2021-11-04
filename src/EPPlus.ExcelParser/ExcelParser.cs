using System;
using System.Collections.Generic;
using System.Drawing;
using System.Dynamic;
using System.IO;
using System.Linq;
using EPPlus.ExcelParser.ExcelDefinitionAggregate;
using EPPlus.ExcelParser.ExcelFileAggregate;
using EPPlus.ExcelParser.ExcelParserResultAggregate;
using OfficeOpenXml;

namespace EPPlus.ExcelParser
{
    public static class ExcelParser
    {
        private static IExcelParserResult<dynamic> ValidateAndConvertToDynamic(IExcelFileDefinition fileDefinition)
        {
            var excelWorksheetDefinition = fileDefinition.ExcelWorksheetDefinition;
            var hasErrors = false;

            using (var excelPackage = new ExcelPackage(fileDefinition.ExcelFileStream))
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[excelWorksheetDefinition.WorksheetIndex];

                foreach (var validator in excelWorksheetDefinition.Columns)
                {
                    if (string.IsNullOrEmpty(validator.ColumnPropertyName))
                        throw new ArgumentException($"Property name for column {validator.Column} not set");
                }

                var rowStart = excelWorksheetDefinition.HasHeaders ? 2 : 1;

                var uniqueColumns = excelWorksheetDefinition.Columns.Where(o => o.IsUnique).Select(o => o.Column).ToList();
                var uniqueValues = new HashSet<(int column, string value)>();

                // dynamic object instance
                var validObjectList = new List<dynamic>();

                object ConvertValue(string value, Type type) => string.IsNullOrEmpty(value) ? string.Empty :
                    Convert.ChangeType(value, type);

                bool ValidateExcelCell(int row, int column,
                    ICollection<(Func<string, bool> validationPredicate, Color failColor)> validationCases,
                    bool isUnique, Color uniqueFailColor)
                {
                    try
                    {
                        for (int j = 0; j < validationCases.Count; j++)
                        {
                            var standardCheck = validationCases.ElementAt(j);
                            if (!standardCheck.validationPredicate(worksheet.Cells[row, column].Value?.ToString()))
                            {
                                worksheet.Row(row).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(standardCheck.failColor);
                                return true;
                            }
                        }

                        if (isUnique &&
                            uniqueValues.Contains((column, worksheet.Cells[row, column]?.Value?.ToString()))) //complexity O(1)
                        {
                            worksheet.Row(row).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(uniqueFailColor);
                            return true;
                        }

                        worksheet.Row(row).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.Green);
                    }
                    catch
                    {
                        worksheet.Row(row).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.Red);
                        return true;
                    }
                    return false;
                }

                for (var row = rowStart; row <= worksheet.Dimension.Rows; row++)
                {
                    var excelRowValid = true;
                    // iterate over columns
                    for (int i = 0; i < excelWorksheetDefinition.Columns.Count; i++)
                    {
                        //validate single column

                        var columnValidator = excelWorksheetDefinition.Columns[i];
                        if (ValidateExcelCell(row, columnValidator.Column,
                            columnValidator.Validators, columnValidator.IsUnique, columnValidator.UniqueFailColor))
                        {
                            hasErrors = true;
                            excelRowValid = false;
                            break;
                        }
                        if (uniqueColumns.IndexOf(columnValidator.Column) != -1)
                        {
                            uniqueValues.Add((columnValidator.Column,worksheet.Cells[row, columnValidator.Column].Value?.ToString()));
                        }
                    }

                    // setup object
                    if (excelRowValid)
                    {
                        var instance = new ExpandoObject();
                        var expandoDict = instance as IDictionary<string, object>;

                        for (int i = 0; i < excelWorksheetDefinition.Columns.Count; i++) // iterate over columns and create object properties
                        {
                            var cellValue = worksheet.Cells[row, excelWorksheetDefinition.Columns[i].Column]?.Value?.ToString() ?? string.Empty;
                            var propertyName = excelWorksheetDefinition.Columns[i].ColumnPropertyName;
                            var columnType = excelWorksheetDefinition.Columns[i].TypeOfColumn;
                            if (expandoDict.ContainsKey(propertyName))
                            {
                                expandoDict[propertyName] = ConvertValue(cellValue, columnType);
                            }
                            else
                            {
                                expandoDict.Add(propertyName, ConvertValue(cellValue, columnType));
                            }
                        }
                        validObjectList.Add(instance);
                    }
                }


                return ExcelParserResult<dynamic>.CreateInstance(validObjectList, hasErrors, excelPackage.GetAsByteArray());
            }
        }

        public static IExcelFile Load(Stream excelFileStream)
        {
            return new ExcelFile(excelFileStream);
        }
        public static IExcelFileDefinition Validate(this IExcelFile excelFile, IExcelWorksheetDefinition excelWorksheetDefinition)
        {
            return new ExcelFileDefinition(excelFile.ExcelFileStream, excelWorksheetDefinition);
        }
        public static IExcelParserResult<dynamic> ToDynamicResult(this IExcelFileDefinition fileDefinition)
        {
            return ValidateAndConvertToDynamic(fileDefinition);
        }
    }
}