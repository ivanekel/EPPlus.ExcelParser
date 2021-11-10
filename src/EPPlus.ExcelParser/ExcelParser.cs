using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using FluentValidation.Results;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

namespace EPPlus.ExcelParser
{
    public static class ExcelParser
    {
        public static ExcelParser<TObject> CreateNew<TObject>(
            Stream excelFileStream,
            Func<ExcelRowMapper, TObject> mapper,
            bool hasHeaders = true)
        {
            return new ExcelParser<TObject>(excelFileStream, hasHeaders, mapper);
        }
    }

    public class ExcelParser<TObject> : IDisposable
    {
        private readonly Stream _excelFileStream;
        private readonly Func<ExcelRowMapper, TObject> _mapper;
        private readonly bool _hasHeaders;
        private ExcelInlineValidator<TObject> _validation;

        internal ExcelParser(Stream excelFileStream, bool hasHeaders, Func<ExcelRowMapper, TObject> mapper)
        {
            _excelFileStream = excelFileStream;
            _hasHeaders = hasHeaders;
            _mapper = mapper;
        }

        public ExcelParser<TObject> SetValidation(Action<ExcelInlineValidator<TObject>> validatorBuilder)
        {
            _validation = new ExcelInlineValidator<TObject>();
            validatorBuilder(_validation);
            return this;
        }

        public ExcelParserResult<TObject> GetResult(bool raiseCastExceptions = false)
        {
            using (var excelPackage = new ExcelPackage(_excelFileStream))
            {
                var worksheet = excelPackage.Workbook.Worksheets.First();
                var rowStart = _hasHeaders ? 2 : 1;
                var uniqueValues = new HashSet<(string, string)>();
                var mappedObjectList = new List<TObject>();
                var containsErrors = false;

                for (var row = rowStart; row <= worksheet.Dimension.Rows; row++)
                {
                    var excelRowMapper = new ExcelRowMapper(worksheet, row);
                    TObject mappedObject;
                    try
                    {
                        mappedObject = _mapper(excelRowMapper);
                    }
                    catch
                    {
                        if (raiseCastExceptions) throw;

                        SetWorksheetRowInvalid(worksheet, row);
                        continue;
                    }

                    var validationResult = _validation.Validate(mappedObject);
                    containsErrors = !validationResult.IsValid;

                    if (validationResult.IsValid == false)
                    {
                        SetExcelRowValidationColor(worksheet, row, validationResult);
                        continue;
                    }

                    if (UniqueExcelRowValid(worksheet, row, _validation.UniqueProperties, uniqueValues,
                        mappedObject))
                    {
                        SetWorksheetRowValid(worksheet, row);
                        mappedObjectList.Add(mappedObject);
                    }
                    else
                    {
                        containsErrors = true;
                    }
                }

                return ExcelParserResult<TObject>.CreateNew(mappedObjectList, excelPackage.GetAsByteArray(),
                    !containsErrors);
            }
        }

        private void SetWorksheetRowValid(ExcelWorksheet worksheet, int row)
        {
            worksheet.Row(row).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.Green);
        }

        private void SetWorksheetRowInvalid(ExcelWorksheet worksheet, int row)
        {
            worksheet.Row(row).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.Red);
        }

        private bool UniqueExcelRowValid(ExcelWorksheet worksheet,
            int row,
            Dictionary<string, Color> uniqueProperties,
            HashSet<(string, string)> uniqueValues,
            TObject objectToValidate)
        {
            if (uniqueProperties.Count == 0)
            {
                return true;
            }

            for (int i = 0; i < uniqueProperties.Count; i++)
            {
                var (propertyName, failColor) = uniqueProperties.ElementAt(i);
                var propertyValue = typeof(TObject).GetProperty(propertyName).GetValue(objectToValidate).ToString();

                var property = (propertyName, propertyValue);

                if (uniqueValues.Contains(property)) // value already exists mark as duplicate
                {
                    worksheet.Row(row).Style.Fill.PatternType =
                        OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    worksheet.Row(row).Style.Fill.BackgroundColor.SetColor(failColor);
                    return false;
                }

                uniqueValues.Add(property);
            }

            return true;
        }

        private void SetExcelRowValidationColor(ExcelWorksheet worksheet, int row, ValidationResult validationResult)
        {
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

        public void Dispose()
        {
            _excelFileStream?.Dispose();
        }
    }
}