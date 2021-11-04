using System;
using System.Drawing;
using EPPlus.ExcelParser.ExcelColumnDefinitionAggregate;
using EPPlus.ExcelParser.ExcelDefinitionAggregate;

namespace EPPlus.ExcelParser
{
    public static class ExcelParserExtensions
    {
        public static IExcelColumnDefinition ExcelColumn<T>(int column, string columnPropertyName)
        {
            return new ExcelColumnDefinition(column, columnPropertyName, typeof(T));
        }

        public static ExcelWorksheetDefinition ExcelRuleFor(this ExcelWorksheetDefinition worksheetDefinition, IExcelColumnDefinition columnColumnDefinition)
        {
            worksheetDefinition.Columns.Add(columnColumnDefinition);
            return worksheetDefinition;
        }

        public static IExcelColumnDefinition NotNull(this IExcelColumnDefinition builder, KnownColor failColor = KnownColor.Red)
        {
            builder.AddValidator(o => o != null, Color.FromKnownColor(failColor));
            return builder;
        }

        public static IExcelColumnDefinition NotEmpty(this IExcelColumnDefinition builder, KnownColor failColor = KnownColor.Red)
        {
            builder.AddValidator(o => !string.IsNullOrEmpty(o), Color.FromKnownColor(failColor));
            return builder;
        }

        public static IExcelColumnDefinition Unique(this IExcelColumnDefinition builder, KnownColor failColor = KnownColor.Yellow)
        {
            builder.MarkAsUnique(Color.FromKnownColor(failColor));
            return builder;
        }

        public static IExcelColumnDefinition Ignore(this IExcelColumnDefinition builder)
        {
            builder.AddValidator(o => true, Color.Empty);
            return builder;
        }
        
        public static IExcelColumnDefinition GreaterThan<T>(this IExcelColumnDefinition builder, T valueToCompare, KnownColor failColor = KnownColor.Red)
            where T : IComparable<T>, IComparable
        {
            builder.AddValidator(o =>
            {
                if (o == null)
                {
                    return false;
                }
                if (builder.TypeOfColumn != typeof(T))
                {
                    throw new ArgumentException("The column and value to compare types must match", nameof(valueToCompare));
                }

                var value = (T)Convert.ChangeType(o, typeof(T));
                if (value.CompareTo(valueToCompare) > 0)
                {
                    return true;
                }

                return false;
            }, Color.FromKnownColor(failColor));
            return builder;
        }

        public static IExcelColumnDefinition LessThan<T>(this IExcelColumnDefinition builder, T valueToCompare, KnownColor failColor = KnownColor.Red)
            where T : IComparable<T>, IComparable
        {
            builder.AddValidator(o =>
            {
                if (o == null || valueToCompare == null)
                {
                    return false;
                }
                if (builder.TypeOfColumn != typeof(T))
                {
                    throw new ArgumentException("The column and value to compare types must match", nameof(valueToCompare));
                }

                var value = (T)Convert.ChangeType(o, typeof(T));
                if (value.CompareTo(valueToCompare) < 0)
                {
                    return true;
                }

                return false;
            }, Color.FromKnownColor(failColor));
            return builder;
        }

        public static IExcelColumnDefinition Must(this IExcelColumnDefinition builder, Func<string, bool> validationPredicate, KnownColor failColor = KnownColor.Red)
        {
            builder.AddValidator(validationPredicate, Color.FromKnownColor(failColor));
            return builder;
        }
    }
}