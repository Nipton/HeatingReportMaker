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
            Console.WriteLine();
        }
    }
}
