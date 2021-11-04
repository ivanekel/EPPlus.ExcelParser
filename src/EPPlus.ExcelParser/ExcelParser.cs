using System;
using System.Collections.Generic;
using System.Drawing;
using System.Dynamic;
using System.IO;
using System.Linq;
using EPPlus.ExcelParser.ExcelColumnDefinitionAggregate;
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
            var columns = excelWorksheetDefinition.Columns;
            ValidateColumnPropertyNames(ref columns);
            var hasErrors = false;
            var rowStart = excelWorksheetDefinition.HasHeaders ? 2 : 1;
            var uniqueColumns = excelWorksheetDefinition.Columns.Where(o => o.IsUnique).Select(o => o.Column).ToList();
            var uniqueValues = new HashSet<(int column, string value)>();
            var validObjectList = new List<dynamic>();
            
            using (var excelPackage = new ExcelPackage(fileDefinition.ExcelFileStream))
            {
                var worksheet = excelPackage.Workbook.Worksheets[excelWorksheetDefinition.WorksheetIndex];

                for (var row = rowStart; row <= worksheet.Dimension.Rows; row++)
                {
                    var excelRowValid = true;
                    // iterate over columns
                    for (int i = 0; i < columns.Count; i++)
                    {
                        var columnValidator = columns[i];
                        var validators = columnValidator.Validators;

                        if (ValidateExcelCell(row, columnValidator.Column, ref validators,
                            columnValidator.IsUnique, columnValidator.UniqueFailColor, ref worksheet, ref uniqueValues))
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
                        validObjectList.Add(GetObjectInstance(row, ref worksheet, ref excelWorksheetDefinition));
                    }
                }

                return ExcelParserResult<dynamic>.CreateInstance(validObjectList, hasErrors, excelPackage.GetAsByteArray());
            }
        }

        private static void ValidateColumnPropertyNames(ref List<IExcelColumnDefinition> columns)
        {
            foreach (var validator in columns)
            {
                if (string.IsNullOrEmpty(validator.ColumnPropertyName))
                    throw new ArgumentException($"Property name for column {validator.Column} not set");
            }
        }

        private static object ConvertValue(string value, Type type) => string.IsNullOrEmpty(value) ? string.Empty :
            Convert.ChangeType(value, type);

        private static bool ValidateExcelCell(int row, int column,
            ref List<(Func<string, bool> validationPredicate, Color failColor)> validationCases,
            bool isUnique, Color uniqueFailColor, ref ExcelWorksheet worksheet,
            ref HashSet<(int column, string value)> uniqueValues)
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

                if (isUnique && uniqueValues.Contains((column, worksheet.Cells[row, column]?.Value?.ToString()))) //complexity O(1)
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

        public static dynamic GetObjectInstance(int row, ref ExcelWorksheet worksheet, ref IExcelWorksheetDefinition excelWorksheetDefinition)
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
            return expandoDict;
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