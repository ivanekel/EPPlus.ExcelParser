using System;
using System.Collections.Generic;
using System.Drawing;

namespace EPPlus.ExcelParser.ExcelColumnDefinitionAggregate
{
    public class ExcelColumnDefinition : IExcelColumnDefinition
    {
        public Type TypeOfColumn { get; private set; }
        public bool IsUnique { get; private set; }
        public Color UniqueFailColor { get; private set; }

        public int Column { get; private set;}
        public string ColumnPropertyName { get;private set; }
        public List<(Func<string, bool> validationPredicate, Color failColor)> Validators { get; private set;}


        public ExcelColumnDefinition(int column, string columnPropertyName, Type typeOfColumn)
        {
            if (column <= 0)
            {
                throw new ArgumentException("Must be more than zero", nameof(column));
            }
            Column = column;

            if (string.IsNullOrEmpty(columnPropertyName))
            {
                throw new ArgumentException("Must be not null or empty", nameof(columnPropertyName));
            }
            ColumnPropertyName = columnPropertyName;
            Validators = new List<(Func<string, bool> validationPredicate, Color failColor)>();
            TypeOfColumn = typeOfColumn;
        }


        public void AddValidator(Func<string, bool> validationPredicate, Color failColor)
        {
            Validators.Add((validationPredicate,failColor));
        }

        public void MarkAsUnique(Color failColor)
        {
            IsUnique = true;
            UniqueFailColor = failColor;
        }
    }
}