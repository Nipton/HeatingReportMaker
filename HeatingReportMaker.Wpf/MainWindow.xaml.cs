using HeatingReportMaker.Core.Exceptions;
using HeatingReportMaker.Core.Models;
using HeatingReportMaker.Core.Services;
using Microsoft.Win32;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace HeatingReportMaker.Wpf
{
    public partial class MainWindow : Window
    {
        private readonly ExcelDataReader _reader;
        private readonly ExcelReportGenerator _excelGenerator;
        private readonly WordReportGenerator _wordGenerator;
        private readonly SolidColorBrush redBrush = new SolidColorBrush(Colors.Red);
        private readonly SolidColorBrush greenBrush = new SolidColorBrush(Colors.Green);
        private readonly SolidColorBrush orangeBrush = new SolidColorBrush(Colors.Orange);
        public MainWindow()
        {
            InitializeComponent();
            _reader = new ExcelDataReader();
            _excelGenerator = new ExcelReportGenerator();
            _wordGenerator = new WordReportGenerator();
        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Title = "Выберите файл с данными",
                Filter = "Excel файлы|*.xlsx;*.xls",
                CheckFileExists = true
            };
            if (dialog.ShowDialog() == true)
            {
                filePathTextBox.Text = dialog.FileName;
            }
        }

        private async void GenerateButton_Click(object sender, RoutedEventArgs e)
        {
            statusTextBlock.Text = string.Empty;
            statusTextBlock.Foreground = SystemColors.ControlTextBrush;
            string filePath = filePathTextBox.Text;
            string stringNumberApartment = apartmentNumberTextBox.Text;
            int numberApartment = 0;
            if (string.IsNullOrEmpty(filePath))
            {
                statusTextBlock.Text = "Путь не указан!";
                statusTextBlock.Foreground = redBrush;
                return;
            }
            else if (!File.Exists(filePath))
            {
                statusTextBlock.Text = "Файл не найден!";
                statusTextBlock.Foreground = redBrush;
                return;
            }
            else if (string.IsNullOrEmpty(stringNumberApartment))
            {
                statusTextBlock.Text = "Номер квартиры не указан!";
                statusTextBlock.Foreground = redBrush;
                return;
            }
            else if (!int.TryParse(stringNumberApartment, out numberApartment))
            {
                statusTextBlock.Text = "Некорректный номер квартиры.";
                statusTextBlock.Foreground = redBrush;
                return;
            }
            statusTextBlock.Text = "Формирование отчёта.";
            statusTextBlock.Foreground = orangeBrush;
            bool roundGkal = roundGkalCheckBox.IsChecked ?? true;
            bool wordReport = createWordReportCheckBox.IsChecked ?? false;
            ApartmentReadResult? result = null;
            try
            {
                result = await Task.Run(() => _reader.ReadApartmentData(filePath, numberApartment, roundGkal));
                if (!result.Success)
                {
                    statusTextBlock.Text = result.Message;
                    statusTextBlock.Foreground = redBrush;
                    return;
                }
                await Task.Run(() => _excelGenerator.GenerateReport(result.ApartmentHeating!));
            }
            catch (IOException ex) when (ex.Message.Contains("занят") || ex.Message.Contains("used by another process"))
            {
                statusTextBlock.Text = "Ошибка!";
                statusTextBlock.Foreground = redBrush;
                MessageBox.Show("Файл открыт в другой программе. Закройте его и повторите попытку.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            catch (CellValueException ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                Application.Current.Shutdown();
            }
            catch (UnauthorizedAccessException)
            {
                MessageBox.Show("Ошибка доступа к файлу.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            catch (Exception)
            {
                MessageBox.Show("Возникла непредвиденная ошибка при выполнении программы.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                Application.Current.Shutdown();
            }
            if (wordReport)
            {
                try
                {
                    await Task.Run(() => _wordGenerator.GenerateReport(result!.ApartmentHeating!));
                }
                catch (TemplateException ex)
                {
                    statusTextBlock.Text = "Ошибка!";
                    statusTextBlock.Foreground = redBrush;
                    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                catch 
                {
                    MessageBox.Show("Отчёт в ворде не был сформирован. Возникла непредвиденная ошибка при выполнении программы.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    Application.Current.Shutdown();
                }
            }
            statusTextBlock.Text = "Готово. Файлы сохранены на рабочий стол.";
            statusTextBlock.Foreground = greenBrush;
        }
    }
}