using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using HeatingReportMaker.Core.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HeatingReportMaker.Core.Services
{
    public class ExcelReportGenerator
    {
        public void GenerateReport(ApartmentHeating data)
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add();
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string reportDate = data.ReportPeriod.Split(' ').Take(2).Aggregate((a, b) => $"{a} {b}");
            string filePath = Path.Combine(desktopPath, $"Расчет стоимости отопления {data.Address} кв. {data.ApartmentNumber} за {reportDate.ToLower()} г.xlsx");

            ws.Range("A1:B1").Merge();
            ws.Range("A2:B2").Merge();
            ws.Range("A3:B3").Merge();
            ws.Range("A4:B4").Merge();
            ws.Cell("A1").Value = "Расчёт стоимости отопления";
            ws.Cell("A1").Style.Font.SetBold().Font.SetFontSize(14).Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center).Alignment.SetVertical(XLAlignmentVerticalValues.Center).Fill.SetBackgroundColor(XLColor.LightGray);
            ws.Cell("A2").SetValue($"Отчетный период: {reportDate}").Style.Font.SetBold();
            ws.Cell("A3").SetValue($"Адрес дома: {data.Address}").Style.Font.SetBold();
            ws.Cell("A4").SetValue($"Квартира № {data.ApartmentNumber}").Style.Font.SetBold();  

            ws.Cell("A6").SetValue("Площадь жилых помещений, кв.м").Style.Font.SetBold();
            ws.Cell("B6").SetValue(data.LivingArea).Style.Font.SetBold();
            ws.Cell("A7").SetValue("Площадь квартиры, кв.м").Style.Font.SetBold();
            ws.Cell("B7").SetValue(data.ApartmentArea).Style.Font.SetBold();
            ws.Cell("A8").SetValue("Расход ТЭ на отопление по показаниям ОДПУ, Гкал").Style.Font.SetBold();
            ws.Cell("B8").SetValue(data.BuildingHeatConsumption).Style.Font.SetBold();
            ws.Cell("A9").SetValue("Тариф на отопление");
            ws.Cell("B9").SetValue(data.Tariff);
            ws.Cell("A10").SetValue("Расчет по распределителям, кв.м");
            ws.Cell("B10").SetValue(data.CalculationByDistributors);
            ws.Cell("A11").SetValue("Расчет по площади, кв.м");
            ws.Cell("B11").SetValue(data.CalculationByArea);
            ws.Cell("A12").SetValue($"Расход тепловой энергии на отопление на МОП {(data.MopHeatPercentage * 100):F2}%, Гкал");
            ws.Cell("B12").SetValue(Math.Round(data.MopHeatConsumption, 3));
            ws.Cell("A13").SetValue("Доля тепловой энергии на отопление МОП, приходящаяся на квартиру, Гкал");
            ws.Cell("B13").SetValue(Math.Round(data.ApartmentMopHeatShare, 4));
            ws.Cell("A14").SetValue("Плата за общедомовое тепло, руб.");
            ws.Cell("B14").SetValue(Math.Round(data.MopCharge, 2));
            ws.Cell("A15").SetValue($"Индивидуальное потребление {(data.IndividualPercentage * 100):F2}%, Гкал");
            ws.Cell("B15").SetValue(Math.Round(data.IndividualHeatConsumption, 3));
            ws.Cell("A16").SetValue("Расход тепловой энергии на 1 м² инд. отоление, Гкал/м²");
            ws.Cell("B16").SetValue(Math.Round(data.HeatPerSquareMeter, 6));
            ws.Cell("A17").SetValue("Сумма всех показаний (потребленных условных единиц) со всех теплораспределителей в доме, у.е.").Style.Font.SetBold();
            ws.Cell("B17").SetValue(data.TotalHouseConsumption).Style.Font.SetBold();
            ws.Cell("A18").SetValue("Объем тепла по распределителям, Гкал");
            ws.Cell("B18").SetValue(Math.Round(data.HeatVolumeByDistributors, 3));
            ws.Cell("A19").SetValue("Объем тепла на  1 условную единицу, Гкал/у.е.");
            ws.Cell("B19").SetValue(Math.Round(data.HeatVolumePerUnit, 8));
            ws.Cell("A20").SetValue("Коэффициент квартирный");
            ws.Cell("B20").SetValue(data.ApartmentCoefficient);
            var currentRow = ws.Row(21);
            foreach(var distributor in data.Distributors)
            {
                currentRow.Cell(1).SetValue($"Показания распределителя №{distributor.DistributorNumber}").Style.Font.SetBold();
                currentRow.Cell(2).SetValue(distributor.DistributorReadings).Style.Font.SetBold();
                currentRow.Cells(1, 2).Style.Fill.BackgroundColor = XLColor.MistyRose;
                currentRow = currentRow.RowBelow();
                currentRow.Cell(1).SetValue("Коэффициент радиатора");
                currentRow.Cell(2).SetValue(distributor.RadiatorCoefficient);
                currentRow = currentRow.RowBelow();
                currentRow.Cell(1).SetValue("Пересчитанные показания");
                currentRow.Cell(2).SetValue(distributor.RecalculatedReading);
                currentRow = currentRow.RowBelow();
            }
            currentRow.Cell(1).SetValue("Сумма пересчит. показаний кв., ед");
            currentRow.Cell(2).SetValue(Math.Round(data.RecalculatedReading, 4));
            currentRow = currentRow.RowBelow();
            currentRow.Cell(1).SetValue("Индивидуальное потребление, Гкал");
            currentRow.Cell(2).SetValue(Math.Round(data.TotalIndividualConsumption, 4));
            currentRow.Cells(1, 2).Style.Fill.BackgroundColor = XLColor.FromArgb(226, 239, 217);
            currentRow = currentRow.RowBelow();
            currentRow.Cell(1).SetValue("Плата за индивидуальное потребление, руб.");
            currentRow.Cell(2).SetValue(Math.Round(data.TotalIndividualCharge, 2));
            currentRow.Cells(1,2).Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 153);
            currentRow = currentRow.RowBelow();
            currentRow.Cell(1).SetValue("Итоговая сумма, руб.");
            currentRow.Cell(2).SetValue(Math.Round(data.TotalCharge, 2));
            currentRow.Cells(1, 2).Style.Fill.BackgroundColor = XLColor.FromArgb(169, 209, 141);
            currentRow = currentRow.RowBelow();
            currentRow.Cell(1).SetValue("Всего Гкал");
            currentRow.Cell(2).SetValue(Math.Round(data.TotalGcal, 4));

            ws.Range("A6", $"B{currentRow.RowNumber()}").Style.Border.SetInsideBorder(XLBorderStyleValues.Thin);
            ws.Range("A6", $"B{currentRow.RowNumber()}").Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
            ws.Range("A8:B8").Style.Fill.BackgroundColor = XLColor.FromArgb(197, 224, 178);
            ws.Range("A13:B13").Style.Fill.BackgroundColor = XLColor.FromArgb(226, 239, 217);
            ws.Range("A14:B14").Style.Fill.BackgroundColor = XLColor.FromArgb(255, 255, 153);
            ws.Column("A").Width = 60;
            ws.Column("A").Style.Alignment.WrapText = true;            
            ws.Column("B").Width = 15;
            ws.Column("B").Style.Alignment.WrapText = true;
            ws.Row(1).Height = 30;
            wb.SaveAs(filePath);
        }
    }
}
