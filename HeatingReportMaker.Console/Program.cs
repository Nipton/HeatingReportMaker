using HeatingReportMaker.Core.Services;

namespace HeatingReportMaker.ConsoleApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //string filePath = @"C:\Users\Nipton\source\repos\HeatingReportMaker\HeatingReportMaker.Tests\TestData\Test.xlsx";
            Console.Write("Введите расположение файла: ");
            string? filePath = Console.ReadLine();
            if (string.IsNullOrEmpty(filePath))
            {
                Console.WriteLine("Путь не указан!");
                return;
            }
            if (!File.Exists(filePath))
            {
                Console.WriteLine("Файл не найден!");
                return;
            }
            Console.Write("Введите номер квартиры: ");
            if (!int.TryParse(Console.ReadLine(), out int numberApartment))
            {
                Console.WriteLine("Введено некорректное значение.");
                Environment.Exit(0);
            }
            bool round = true;
            
            string? stringRound = "";
            do
            {
                Console.Write("Округлять Гкал? y/n: ");
                stringRound = Console.ReadLine();
                if (stringRound is null)
                    continue;
                if(stringRound.ToLower() == "y")
                    round = true;
                if (stringRound.ToLower() == "n")
                    round = false;

            } while(stringRound == null || (stringRound.ToLower() != "y" && stringRound.ToLower() != "n"));

            ExcelDataReader reader = new ExcelDataReader();
            var res = reader.ReadApartmentData(filePath, numberApartment, round);
            if (!res.Success)
            {
                Console.WriteLine(res.Message);
                return;
            }
            ExcelReportGenerator generator = new ExcelReportGenerator();
            generator.GenerateReport(res.ApartmentHeating!);
        }
    }
}
