using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using FluentValidation;
using FluentValidation.Validators;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

namespace EPPlus.ExcelParser
{
    public static class ExcelParser
    {
        public static ExcelParser<T> CreateNew<T>(
            ExcelPackage excelFile,
            bool hasHeaders,
            Func<ExcelRowMapper, T> mapper)
        {
            return new ExcelParser<T>(excelFile, hasHeaders, mapper);
        }
    }

    public class ExcelParser<T>
    {
        private readonly ExcelPackage _excelPackage;
        private readonly Func<ExcelRowMapper, T> _mapper;
        private readonly bool _hasHeaders;
        private ExcelInlineValidator<T> _validation;
        
        internal ExcelParser(ExcelPackage excelFile, bool hasHeaders, Func<ExcelRowMapper, T> mapper)
        {
            _excelPackage = excelFile;
            _hasHeaders = hasHeaders;
            _mapper = mapper;
        }

        public ExcelParser<T> SetValidation(Action<ExcelInlineValidator<T>> validatorBuilder)
        {
            _validation = new ExcelInlineValidator<T>();
            validatorBuilder(_validation);
            return this;
        }

        public object GetResult()
        {
            var worksheet = _excelPackage.Workbook.Worksheets.First();
            var rowStart = _hasHeaders ? 2 : 1;

            for (var row = rowStart; row <= worksheet.Dimension.Rows; row++)
            {
                var excelRowMapper = new ExcelRowMapper(worksheet, row);
                var mappedObject = _mapper(excelRowMapper);


                var validationResult = _validation?.Validate(mappedObject);


                if (validationResult.IsValid)
                {
                    //check unique
                    continue;
                }

                var colorValidation = validationResult
                    .Errors
                    .FirstOrDefault(o => o.ErrorMessage == "InvalidColorDefined");

                if (colorValidation != null)
                {
                    var invalidColor = Enum.Parse<KnownColor>(colorValidation.ErrorCode);
                }
            }

            return null;
        }
    }

    public class ExcelRowMapper
    {
        private readonly ExcelWorksheet _worksheet;
        private readonly int _row;

        public ExcelRowMapper(ExcelWorksheet worksheet, int row)
        {
            _worksheet = worksheet;
            _row = row;
        }

        public T GetValue<T>(int column)
        {
            return _worksheet.Cells[_row, column].GetValue<T>();
        }
    }

    public static class ExcelValidatorOptions
    {
        public static IRuleBuilderOptions<T, TProperty> WithRowColor<T, TProperty>(
            this IRuleBuilderOptions<T, TProperty> rule, KnownColor invalidColor = KnownColor.Red)
        {
            return rule.WithMessage("InvalidColorDefined").WithErrorCode(invalidColor.ToString());
        }
    }


    public class ExcelInlineValidator<T> : InlineValidator<T>
    {
        private readonly Dictionary<string, KnownColor> _uniqueProperties;
        public Dictionary<string, KnownColor> UniqueProperties => _uniqueProperties;

        public ExcelInlineValidator()
        {
            _uniqueProperties = new Dictionary<string, KnownColor>();
        }

        public IRuleBuilderInitial<T, TProperty> RuleFor<TProperty>(Expression<Func<T, TProperty>> expression, bool isUnique = false,
            KnownColor uniqueFailColor = KnownColor.Yellow)
        {
            var propertyInfo = (expression.Body as MemberExpression).Member as PropertyInfo;
            if (propertyInfo == null)
            {
                throw new ArgumentException("Invalid property");
            }

            if (_uniqueProperties.ContainsKey(propertyInfo.Name))
            {
                throw new ArgumentException($"unique rule already set for property {propertyInfo.Name}");
            }

            _uniqueProperties.Add(propertyInfo.Name, uniqueFailColor);


            return RuleFor(expression);
        }
    }


    // public static class ExcelParser
    // {
    //     private static IExcelParserResult<dynamic> ValidateAndConvertToDynamic(IExcelFileDefinition fileDefinition)
    //     {
    //         var excelWorksheetDefinition = fileDefinition.ExcelWorksheetDefinition;
    //         var columns = excelWorksheetDefinition.Columns;
    //         ValidateColumnPropertyNames(columns);
    //         var hasErrors = false;
    //         var rowStart = excelWorksheetDefinition.HasHeaders ? 2 : 1;
    //         var uniqueColumns = excelWorksheetDefinition.Columns.Where(o => o.IsUnique).Select(o => o.Column).ToList();
    //         var uniqueValues = new HashSet<(int column, string value)>();
    //         var validObjectList = new List<dynamic>();
    //
    //         using (var excelPackage = new ExcelPackage(fileDefinition.ExcelFileStream))
    //         {
    //             var worksheet = excelPackage.Workbook.Worksheets[excelWorksheetDefinition.WorksheetIndex];
    //
    //             for (var row = rowStart; row <= worksheet.Dimension.Rows; row++)
    //             {
    //                 var excelRowValid = true;
    //                 // iterate over columns
    //                 for (int i = 0; i < columns.Count; i++)
    //                 {
    //                     var columnValidator = columns[i];
    //                     var validators = columnValidator.Validators;
    //
    //                     if (ValidateExcelCell(row, columnValidator.Column, validators,
    //                         columnValidator.IsUnique, columnValidator.UniqueFailColor, worksheet, uniqueValues))
    //                     {
    //                         hasErrors = true;
    //                         excelRowValid = false;
    //                         break;
    //                     }
    //
    //                     if (uniqueColumns.IndexOf(columnValidator.Column) != -1)
    //                     {
    //                         uniqueValues.Add((columnValidator.Column,worksheet.Cells[row, columnValidator.Column].Value?.ToString()));
    //                     }
    //                 }
    //                 // setup object
    //                 if (excelRowValid)
    //                 {
    //                     validObjectList.Add(GetObjectInstance(row, worksheet, excelWorksheetDefinition));
    //                 }
    //             }
    //
    //             return ExcelParserResult<dynamic>.CreateInstance(validObjectList, hasErrors, excelPackage.GetAsByteArray());
    //         }
    //     }
    //
    //     private static void ValidateColumnPropertyNames(List<IExcelColumnDefinition> columns)
    //     {
    //         foreach (var validator in columns)
    //         {
    //             if (string.IsNullOrEmpty(validator.ColumnPropertyName))
    //                 throw new ArgumentException($"Property name for column {validator.Column} not set");
    //         }
    //     }
    //
    //     private static object ConvertValue(string value, Type type) => string.IsNullOrEmpty(value) ? string.Empty :
    //         Convert.ChangeType(value, type);
    //
    //     private static bool ValidateExcelCell(int row, int column,
    //         List<(Func<string, bool> validationPredicate, Color failColor)> validationCases,
    //         bool isUnique, Color uniqueFailColor, ExcelWorksheet worksheet,
    //         HashSet<(int column, string value)> uniqueValues)
    //     {
    //         try
    //         {
    //             for (int j = 0; j < validationCases.Count; j++)
    //             {
    //                 var standardCheck = validationCases.ElementAt(j);
    //                 if (!standardCheck.validationPredicate(worksheet.Cells[row, column].Value?.ToString()))
    //                 {
    //                     worksheet.Row(row).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
    //                     worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(standardCheck.failColor);
    //                     return true;
    //                 }
    //             }
    //
    //             if (isUnique && uniqueValues.Contains((column, worksheet.Cells[row, column]?.Value?.ToString()))) //complexity O(1)
    //             {
    //                 worksheet.Row(row).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
    //                 worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(uniqueFailColor);
    //                 return true;
    //             }
    //
    //             worksheet.Row(row).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
    //             worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.Green);
    //         }
    //         catch
    //         {
    //             worksheet.Row(row).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
    //             worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.Red);
    //             return true;
    //         }
    //         return false;
    //     }
    //
    //     public static dynamic GetObjectInstance(int row, ExcelWorksheet worksheet, IExcelWorksheetDefinition excelWorksheetDefinition)
    //     {
    //         var instance = new ExpandoObject();
    //         var expandoDict = instance as IDictionary<string, object>;
    //
    //         for (int i = 0; i < excelWorksheetDefinition.Columns.Count; i++) // iterate over columns and create object properties
    //         {
    //             var cellValue = worksheet.Cells[row, excelWorksheetDefinition.Columns[i].Column]?.Value?.ToString() ?? string.Empty;
    //             var propertyName = excelWorksheetDefinition.Columns[i].ColumnPropertyName;
    //             var columnType = excelWorksheetDefinition.Columns[i].TypeOfColumn;
    //             if (expandoDict.ContainsKey(propertyName))
    //             {
    //                 expandoDict[propertyName] = ConvertValue(cellValue, columnType);
    //             }
    //             else
    //             {
    //                 expandoDict.Add(propertyName, ConvertValue(cellValue, columnType));
    //             }
    //         }
    //         return expandoDict;
    //     }
    //
    //     public static IExcelFile Load(Stream excelFileStream)
    //     {
    //         return new ExcelFile(excelFileStream);
    //     }
    //
    //     public static IExcelFileDefinition Validate(this IExcelFile excelFile, IExcelWorksheetDefinition excelWorksheetDefinition)
    //     {
    //         return new ExcelFileDefinition(excelFile.ExcelFileStream, excelWorksheetDefinition);
    //     }
    //
    //     public static IExcelParserResult<dynamic> ToDynamicResult(this IExcelFileDefinition fileDefinition)
    //     {
    //         return ValidateAndConvertToDynamic(fileDefinition);
    //     }
    // }
}