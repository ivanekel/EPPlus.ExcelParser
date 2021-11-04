using System;
using System.Collections.Generic;
using System.Drawing;

namespace EPPlus.ExcelParser.ExcelColumnDefinitionAggregate
{
    public interface IExcelColumnDefinition
    {
        int Column { get; }
        string ColumnPropertyName { get; }
        List<(Func<string, bool> validationPredicate, Color failColor)> Validators { get; }
        bool IsUnique { get; }
        Color UniqueFailColor { get; }

        Type TypeOfColumn { get; }
        void AddValidator(Func<string, bool> validationPredicate, Color failColor);
        void MarkAsUnique(Color failColor);
    }
}