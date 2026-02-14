using HeatingReportMaker.Core.Services;
using HeatingReportMaker.Core.Exceptions;

namespace HeatingReportMaker.ConsoleApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.Write("Введите расположение файла: ");
            string? filePath = Console.ReadLine();
            if (string.IsNullOrEmpty(filePath))
            {
                Console.WriteLine("Путь не указан!");
                Console.WriteLine("Нажмите любую клавишу для выхода.");
                Console.ReadKey();
                return;
            }
            if (!File.Exists(filePath))
            {
                Console.WriteLine("Файл не найден!");
                Console.WriteLine("Нажмите любую клавишу для выхода.");
                Console.ReadKey();
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
                {
                    Console.WriteLine("Неверное значение. Повторите ввод.");
                    continue;
                }
                if(stringRound.ToLower() == "y")
                    round = true;
                if (stringRound.ToLower() == "n")
                    round = false;
                else Console.WriteLine("Неверное значение. Повторите ввод.");

            } while(stringRound == null || (stringRound.ToLower() != "y" && stringRound.ToLower() != "n"));
            Console.WriteLine("Формирование отчёта.");
            try
            {
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
            catch (CellValueException ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine("Нажмите любую клавишу для выхода.");
                Console.ReadKey();
                Environment.Exit(1);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Возникла непредвиденная ошибка при выполнении программы. Нажмите любую клавишу для выхода.");
                Console.ReadKey();
                Environment.Exit(1);
            }
            Console.WriteLine("Отчёт сформирован. Нажмите любую клавишу для выхода.");
            Console.ReadKey();
        }
    }
}
