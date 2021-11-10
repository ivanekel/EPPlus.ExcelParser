using System;
using System.Collections.Generic;
using System.Drawing;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using EPPlus.ExcelParser.ExcelColumnDefinitionAggregate;
using EPPlus.ExcelParser.ExcelDefinitionAggregate;
using EPPlus.ExcelParser.ExcelFileAggregate;
using EPPlus.ExcelParser.ExcelParserResultAggregate;
using FluentValidation;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

namespace EPPlus.ExcelParser
{
    public static class ExcelParser
    {
        public static ExcelParser<TObject> CreateNew<TObject>(
            ExcelPackage excelFile,
            bool hasHeaders,
            Func<ExcelRowMapper, TObject> mapper)
        {
            return new ExcelParser<TObject>(excelFile, hasHeaders, mapper);
        }
    }

    public class ExcelParser<TObject>
    {
        private readonly ExcelPackage _excelPackage;
        private readonly Func<ExcelRowMapper, TObject> _mapper;
        private readonly bool _hasHeaders;
        private ExcelInlineValidator<TObject> _validation;

        internal ExcelParser(ExcelPackage excelFile, bool hasHeaders, Func<ExcelRowMapper, TObject> mapper)
        {
            _excelPackage = excelFile;
            _hasHeaders = hasHeaders;
            _mapper = mapper;
        }

        public ExcelParser<TObject> SetValidation(Action<ExcelInlineValidator<TObject>> validatorBuilder)
        {
            _validation = new ExcelInlineValidator<TObject>();
            validatorBuilder(_validation);
            return this;
        }

        public object GetResult()
        {
            var worksheet = _excelPackage.Workbook.Worksheets.First();
            var rowStart = _hasHeaders ? 2 : 1;
            var uniqueValues = new HashSet<(string, string)>();
            var mappedObjectList = new List<TObject>();
            var containsErrors = false;

            for (var row = rowStart; row <= worksheet.Dimension.Rows; row++)
            {
                var excelRowMapper = new ExcelRowMapper(worksheet, row);
                var mappedObject = _mapper(excelRowMapper);

                var validationResult = _validation?.Validate(mappedObject);
                containsErrors = !validationResult.IsValid;

                if (validationResult.IsValid)
                {
                    var rowValid = true;
                    // unique properties exist
                    if (_validation.UniqueProperties.Count != 0)
                    {
                        for (int i = 0; i < _validation.UniqueProperties.Count; i++)
                        {
                            var uniqueDefinition = _validation.UniqueProperties.ElementAt(i);
                            var property = (uniqueDefinition.Key,
                                typeof(T).GetProperty(uniqueDefinition.Key).GetValue(mappedObject).ToString());

                            if (uniqueValues.Contains(property)) // value already exists mark as duplicate
                            {
                                worksheet.Row(row).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(uniqueDefinition.Value);
                                rowValid = false;
                                containsErrors = true;
                                break;
                            }

                            uniqueValues.Add(property);
                        }
                    }

                    if (rowValid)
                    {
                        mappedObjectList.Add(mappedObject);
                    }

                    continue;
                }

                var colorValidation =
                    validationResult.Errors.FirstOrDefault(o => o.ErrorMessage == "InvalidColorDefined");

                if (colorValidation != null)
                {
                    var invalidColor = Enum.Parse<KnownColor>(colorValidation.ErrorCode);
                    worksheet.Row(row).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.FromKnownColor(invalidColor));
                }
                else
                {
                    worksheet.Row(row).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.Red);
                }
            }

            return ExcelParserResult<TObject>.CreateNew(mappedObjectList, _excelPackage.GetAsByteArray(),
                !containsErrors);
        }
    }

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

    public class ExcelRowMapper
    {
        private readonly ExcelWorksheet _worksheet;
        private readonly int _row;


        public ExcelRowMapper(ExcelWorksheet worksheet, int row)
        {
            _worksheet = worksheet;
            _row = row;
        }

        public TCustomProperty GetValue<TCustomProperty>(int column)
        {
            return _worksheet.Cells[_row, column].GetValue<TCustomProperty>();
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


    public class ExcelInlineValidator<TCustomObject> : InlineValidator<TCustomObject>
    {
        private readonly Dictionary<string, Color> _uniqueProperties;
        public Dictionary<string, Color> UniqueProperties => _uniqueProperties;

        public ExcelInlineValidator()
        {
            _uniqueProperties = new Dictionary<string, Color>();
        }

        public IRuleBuilderInitial<TCustomObject, TProperty> RuleFor<TProperty>(
            Expression<Func<TCustomObject, TProperty>> expression,
            bool isUnique = false,
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

            _uniqueProperties.Add(propertyInfo.Name, Color.FromKnownColor(uniqueFailColor));


            return RuleFor(expression);
        }
    }
}