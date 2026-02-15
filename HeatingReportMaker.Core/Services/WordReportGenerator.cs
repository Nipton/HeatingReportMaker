using HeatingReportMaker.Core.Exceptions;
using HeatingReportMaker.Core.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace HeatingReportMaker.Core.Services
{
    public class WordReportGenerator
    {
        public void GenerateReport(ApartmentHeating data)
        {
            string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "ReportTemplate.docx");            
            string reportDate = data.ReportPeriod.Split(' ').Take(2).Aggregate((a, b) => $"{a} {b}");
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string outputPath = Path.Combine(desktopPath, $"Расчет стоимости отопления за {reportDate.ToLower()} г.docx");
            if (!File.Exists(templatePath))
            {
                throw new TemplateNotFoundException("Шаблон для Word отчёта не был найден. Поместите шаблон ReportTemplate.docx в папку Templates в корне программы.");
            }
            using var doc = DocX.Load(templatePath);
            doc.ReplaceTextSimple("{ReportPeriod}", reportDate.ToLower());
            doc.ReplaceTextSimple("{Address}", data.Address);
            doc.ReplaceTextSimple("{ApartmentNumber}", data.ApartmentNumber.ToString());
            doc.ReplaceTextSimple("{IndividualPercentage}", (data.IndividualPercentage * 100).ToString("F2"));
            doc.ReplaceTextSimple("{MopHeatPercentage}", (data.MopHeatPercentage * 100).ToString("F2"));
            doc.ReplaceTextSimple("{BuildingHeatConsumption}", Math.Round(data.BuildingHeatConsumption, 4).ToString());
            doc.ReplaceTextSimple("{LivingArea}", data.LivingArea.ToString());
            doc.ReplaceTextSimple("{CalculationByArea}", data.CalculationByArea.ToString());
            doc.ReplaceTextSimple("{CalculationByDistributors}", data.CalculationByDistributors.ToString());
            doc.ReplaceTextSimple("{ApartmentArea}", data.ApartmentArea.ToString());
            doc.ReplaceTextSimple("{TotalHouseConsumption}", data.TotalHouseConsumption.ToString());
            doc.ReplaceTextSimple("{RecalculatedReading}", data.RecalculatedReading.ToString());
            doc.ReplaceTextSimple("{Distributors}", GenerateDistributorsText(data.Distributors, data.ApartmentCoefficient));
            doc.ReplaceTextSimple("{MopHeatConsumption}", Math.Round(data.MopHeatConsumption, 3).ToString());
            doc.ReplaceTextSimple("{ApartmentMopHeatShare}", Math.Round(data.ApartmentMopHeatShare, 4).ToString());
            doc.ReplaceTextSimple("{Tariff}", data.Tariff.ToString());
            doc.ReplaceTextSimple("{MopCharge}", Math.Round(data.MopCharge, 2).ToString());
            doc.ReplaceTextSimple("{IndividualHeatConsumption}", Math.Round(data.IndividualHeatConsumption, 3).ToString());
            doc.ReplaceTextSimple("{HeatPerSquareMeter}", Math.Round(data.HeatPerSquareMeter, 6).ToString());
            doc.ReplaceTextSimple("{HeatVolumeByDistributors}", Math.Round(data.HeatVolumeByDistributors, 3).ToString());
            doc.ReplaceTextSimple("{HeatVolumePerUnit}", Math.Round(data.HeatVolumePerUnit, 8).ToString());
            doc.ReplaceTextSimple("{TotalIndividualConsumption}", Math.Round(data.TotalIndividualConsumption, 4).ToString());
            doc.ReplaceTextSimple("{TotalIndividualCharge}", Math.Round(data.TotalIndividualCharge, 2).ToString());
            doc.ReplaceTextSimple("{TotalCharge}", Math.Round(data.TotalCharge, 2).ToString());
            doc.SaveAs(outputPath);
        }
        private string GenerateDistributorsText(List<HeatDistributor> distributors, decimal apartmentCoefficient)
        {
            if (!distributors.Any())
                return "Нет данных по распределителям.";
            StringBuilder sb = new StringBuilder();
            foreach(var distributor in distributors)
            {
                sb.AppendLine($"{distributor.DistributorReadings} (показания распределителя) × {apartmentCoefficient} (коэфф-т квартирный) × {distributor.RadiatorCoefficient} (коэфф-т радиатора) = {distributor.RecalculatedReading} у.е.");
            }
            return sb.ToString();
        }
    }
    internal static class WordExtensions
    {
        internal static void ReplaceTextSimple(this DocX doc, string placeholder, string replace)
        {
            var option = new StringReplaceTextOptions
            {
                SearchValue = placeholder,
                NewValue = replace
            };
            doc.ReplaceText(option);
        }
    }
}
