using System.Drawing;
using EPPlus.ExcelParser;
using FluentValidation;
using OfficeOpenXml;

namespace TestApp
{
    class Program
    {
        static void Main(string[] args)
        {
            var parser = ExcelParser
                .CreateNew(new ExcelPackage(),
                    true,
                    mapper => new Dto()
                    {
                        Name = mapper.GetValue<string>(1),
                        Surname = mapper.GetValue<string>(2),
                        Age = mapper.GetValue<int?>(3),
                        Height = mapper.GetValue<double?>(4)
                    })
                .SetValidation(rules =>
                {
                    rules.RuleFor(o => o.Name).NotEmpty();
                    rules.RuleFor(o => o.Surname).NotEmpty();
                    rules.RuleFor(o => o.Height).NotNull().GreaterThan(100).WithRowColor();
                    rules.RuleFor(o => o.Age).NotNull().GreaterThan(0);
                })
                .GetResult();
        }
    }


    public class Dto
    {
        public string Name { get; set; }
        public string Surname { get; set; }
        public int? Age { get; set; }
        public double? Height { get; set; }
    }
}