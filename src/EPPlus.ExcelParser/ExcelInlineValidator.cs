using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq.Expressions;
using System.Reflection;
using FluentValidation;

namespace EPPlus.ExcelParser
{
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
            if (isUnique)
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
            }


            return base.RuleFor(expression);
        }
    }
}