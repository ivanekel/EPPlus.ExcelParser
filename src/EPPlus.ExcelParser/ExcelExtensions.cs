using OfficeOpenXml;
using System.IO;
using System.Threading.Tasks;

namespace EPPlus.ExcelParser
{
    public static class ExcelExtensions
    {
        public static async Task SaveToFileAsync(this ExcelPackage excelPackage, string filePath)
        {
            using var file = File.OpenWrite(filePath);
            excelPackage.Save();
            var position = excelPackage.Stream.Position;
            excelPackage.Stream.Seek(0L, SeekOrigin.Begin);
            await excelPackage.Stream.CopyToAsync(file);
            excelPackage.Stream.Seek(position, SeekOrigin.Begin);
        }
    }
}