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

namespace HeatingReportMaker.Wpf
{
    public partial class MainWindow : Window
    {
        private readonly ExcelDataReader _reader;
        private readonly ExcelReportGenerator _excelGenerator;
        private readonly WordReportGenerator _wordGenerator;
        private SolidColorBrush redBrush = new SolidColorBrush(Colors.Red);

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

        private void GenerateButton_Click(object sender, RoutedEventArgs e)
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
            }
            else if (!File.Exists(filePath))
            {
                statusTextBlock.Text = "Файл не найден!";
                statusTextBlock.Foreground = redBrush;
            }
            else if (string.IsNullOrEmpty(stringNumberApartment))
            {
                statusTextBlock.Text = "Номер квартиры не указан!";
                statusTextBlock.Foreground = redBrush;
            }
            else if (!int.TryParse(stringNumberApartment, out numberApartment))
            {
                statusTextBlock.Text = "Некорректный номер квартиры.";
                statusTextBlock.Foreground = redBrush;
            }
            bool roundGkal = roundGkalCheckBox.IsChecked ?? true;
            bool wordReport = createWordReportCheckBox.IsChecked ?? false;
            ApartmentReadResult? result = null;
            try
            {
                result = _reader.ReadApartmentData(filePath, numberApartment, roundGkal);
                if (!result.Success)
                {
                    statusTextBlock.Text = result.Message;
                    statusTextBlock.Foreground = redBrush;
                }
                ExcelReportGenerator generator = new ExcelReportGenerator();
                generator.GenerateReport(result.ApartmentHeating!);
            }
            catch (IOException ex) when (ex.Message.Contains("занят") || ex.Message.Contains("used by another process"))
            {
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
                ShowFatalErrorAndExit();
            }
            if (wordReport)
            {
                try
                {
                    _wordGenerator.GenerateReport(result!.ApartmentHeating!);

                }
                catch (TemplateException ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                catch 
                {
                    ShowFatalErrorAndExit();
                }
            }
        }
        private void ShowFatalErrorAndExit()
        {
            MessageBox.Show("Возникла непредвиденная ошибка при выполнении программы.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            Application.Current.Shutdown();
        }
    }
}