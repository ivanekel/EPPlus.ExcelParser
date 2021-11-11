using EPPlus.ExcelParser.Parsing.Validator;
using FluentValidation.Results;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace EPPlus.ExcelParser.Parsing.Parser
{
    internal class Marking
    {
        internal void MarkRowAsValid(ExcelWorksheet worksheet, int rowNumber)
               => SetRowColor(worksheet, rowNumber, Color.Green);

        internal void MarkRowAsInvalid(ExcelWorksheet worksheet, int rowNumber)
            => SetRowColor(worksheet, rowNumber, Color.Red);

        internal void MarkRowAsInvalidFromValidationResult(ExcelWorksheet worksheet, int rowNumber, ValidationResult validationResult)
            => SetRowColor(worksheet, rowNumber, GetColorFromValidationErrors(validationResult.Errors) ?? Color.Red);

        internal void SetRowColor(ExcelWorksheet worksheet, int rowNumber, Color color)
        {
            var row = worksheet.Row(rowNumber);
            row.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            row.Style.Fill.BackgroundColor.SetColor(color);
        }

        private Color? GetColorFromValidationErrors(IEnumerable<ValidationFailure> validationFailures)
        {
            var colorValidation = validationFailures
                .FirstOrDefault(o => o.ErrorCode.StartsWith(ValidationExtensions.ColorCodePrefix));

            return colorValidation == null
                ? (Color?)null
                : Color.FromKnownColor(
                    Enum.Parse<KnownColor>(
                        colorValidation
                            .ErrorCode
                            .Substring(ValidationExtensions.ColorCodePrefix.Length)));
        }

    }
}
