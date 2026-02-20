using ClosedXML;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using HeatingReportMaker.Core.Exceptions;
using HeatingReportMaker.Core.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HeatingReportMaker.Core.Services
{
    public class ExcelDataReader
    {
        public ApartmentReadResult ReadApartmentData(string filePath, int apartmentNumber, bool round)
        {
            ApartmentReadResult result = new ApartmentReadResult();
            using var wb = new XLWorkbook(filePath);
            var ws = wb.Worksheet(1);
            string address = ws.Cell("A1").GetString();
            var rows = ws.RowsUsed(r =>
            {
                if (r.Cell("A").TryGetValue(out int value))
                    return value == apartmentNumber;
                return false;
            }).ToList();
            if (!rows.Any())
            {
                result.Message = "Указанная квартира не найдена.";
                return result;
            }
            var row = rows.First();
            if(string.Equals(row.Cell("E").GetString(), "площади"))
            {
                result.Message = "Отчет не будет сгенерирован, так как для этой квартиры расчет производится по площади, а не по показаниям.";
                return result;
            }
            ApartmentHeating apartment = new ApartmentHeating()
            {
                ApartmentNumber = apartmentNumber,
                Address = address,
                ReportPeriod = row.Cell("AR").GetString(),
                ApartmentArea = GetCellValueDecimal(row, "B"),
                LivingArea = GetCellValueDecimal(row, "AA"),
                BuildingHeatConsumption = GetCellValueDecimal(row, "AT"),
                Tariff = GetCellValueDecimal(row, "AM"),
                CalculationByDistributors = GetCellValueDecimal(row, "AW"),
                CalculationByArea = GetCellValueDecimal(row, "AV"),
                TotalHouseConsumption = GetCellValueDecimal(row, "AU"),
                HeatVolumeByDistributors = GetCellValueDecimal(row, "BE"),
                ApartmentCoefficient = GetCellValueDecimal(row, "D"),
                TotalGcal = GetCellValueDecimal(row, "M"),
                RecalculatedReading = GetCellValueDecimal(row, "F")
            };
            try
            {
                string[] allPercentages = row.Cell("AO").GetString().Split("/");
                apartment.MopHeatPercentage = decimal.Parse(allPercentages[0], CultureInfo.InvariantCulture) / 100;
                apartment.IndividualPercentage = decimal.Parse(allPercentages[1], CultureInfo.InvariantCulture) / 100;
            }
            catch (Exception ex)
            {
                throw new CellValueException($"Ошибка при считывании значения из ячейки AO{row.RowNumber()}");
            }
            foreach (var rowD in rows)
            {
                HeatDistributor hd = new HeatDistributor();
                hd.DistributorNumber = rowD.Cell("Q").GetString();
                hd.DistributorReadings = GetCellValueDecimal(rowD, "S");
                hd.RadiatorCoefficient = GetCellValueDecimal(rowD, "O");
                hd.RecalculatedReading = GetCellValueDecimal(rowD, "U");
                apartment.Distributors.Add(hd);
            }
            CalculateTotals(apartment, round);
            result.Success = true;
            result.ApartmentHeating = apartment;           
            return result;
        }
        private void CalculateTotals(ApartmentHeating apartment, bool round)
        {
            apartment.MopHeatConsumption = apartment.BuildingHeatConsumption * apartment.MopHeatPercentage;
            apartment.IndividualHeatConsumption = apartment.BuildingHeatConsumption * apartment.IndividualPercentage;
            apartment.HeatPerSquareMeter = apartment.IndividualHeatConsumption / apartment.LivingArea;
            apartment.HeatVolumePerUnit = apartment.HeatVolumeByDistributors / apartment.TotalHouseConsumption;
            decimal tempTotalIndividualConsumption = apartment.RecalculatedReading* apartment.HeatVolumePerUnit;
            apartment.TotalIndividualConsumption = round ? Math.Round(tempTotalIndividualConsumption, 4) : tempTotalIndividualConsumption;
            apartment.TotalIndividualCharge = apartment.TotalIndividualConsumption * apartment.Tariff;
            decimal tempMopHeatShare = apartment.MopHeatConsumption / apartment.LivingArea * apartment.ApartmentArea;
            apartment.ApartmentMopHeatShare = round ? Math.Round(tempMopHeatShare, 4) : tempMopHeatShare;
            apartment.MopCharge = apartment.ApartmentMopHeatShare * apartment.Tariff;
            apartment.TotalCharge = apartment.TotalIndividualCharge + apartment.MopCharge;
        }
        private decimal GetCellValueDecimal(IXLRow row, string cellAddress)
        {
            try
            {
                var value = row.Cell(cellAddress);
                if (string.Equals(value.GetString(), "-"))
                    return 0m;
                return value.GetValue<decimal>();
            }
            catch (Exception ex)
            {
                throw new CellValueException($"Ошибка при считывании значения из ячейки {cellAddress}{row.RowNumber()}");
            }
        }
    }
}
