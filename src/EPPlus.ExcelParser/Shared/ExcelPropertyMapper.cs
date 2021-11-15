using OfficeOpenXml;
using System;
using System.Linq.Expressions;
using System.Reflection;

namespace EPPlus.ExcelParser.Shared
{
    internal class ExcelPropertyMapper<TObject, TProperty> : IExcelPropertyMapper<TObject>
    {
        public int ColumnNumber { get;}
        private Expression<Func<TObject, TProperty>> _expression;
        private Func<TObject, TProperty> _value = null;
        private MethodInfo _valueGetter = null;


        internal ExcelPropertyMapper(int columnNumber, Expression<Func<TObject, TProperty>> expression)
        {
            ColumnNumber = columnNumber;
            _expression = expression;
        }

        public void MapFromExcel(TObject target, ExcelRange cell)
        {
            var memberSelectorExpression = _expression.Body as MemberExpression;
            if (memberSelectorExpression != null)
            {
                var property = memberSelectorExpression.Member as PropertyInfo;
                if (property != null)
                {
                    property.SetValue(target, GetValueFromCell(cell, property.PropertyType), null);
                }
            }
        }

        public void MapToExcel(TObject target, ExcelRange cell)
        {
            if (_value == null) _value = _expression.Compile();
            cell.Value = _value(target);
        }


        private object GetValueFromCell(ExcelRange cell, Type valueType)
        {
            if (_valueGetter == null)
            {
                var cellValueType = cell.GetType();
                var cellGetterMethod = cellValueType.GetMethod(nameof(cell.GetValue));
                _valueGetter = cellGetterMethod.MakeGenericMethod(valueType);
            }

            return _valueGetter.Invoke(cell, null);
        }
    }
}
