using ClosedXML;
using ClosedXML.Excel;
namespace HeatingReportMaker.ConsoleApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string filePath = @"C:\Users\Nipton\Desktop\HeatingReportMaker\Входные.xlsx";

            Console.Write("Введите номер квартиры: ");
            if (!int.TryParse(Console.ReadLine(), out int numberApartment))
            {
                Console.WriteLine("Введено некорректное значение.");
                Environment.Exit(0);
            }
            using var wb = new XLWorkbook(filePath);
            var ws = wb.Worksheet(1);
            Console.WriteLine(ws.Cell("A1").Value);
        }
    }
}
