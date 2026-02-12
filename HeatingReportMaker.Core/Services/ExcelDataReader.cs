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
        public ApartmentHeating? ReadApartmentData(string filePath, int apartmentNumber, bool round)
        {
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
                return null;
            var row = rows.First();
            ApartmentHeating apartment = new ApartmentHeating();
            apartment.ApartmentNumber = apartmentNumber;
            apartment.Address = address;
            apartment.ReportPeriod = row.Cell("AR").GetString();
            apartment.ApartmentArea = GetCellValueDecimal(row, "B");
            apartment.LivingArea = GetCellValueDecimal(row, "AA");
            apartment.BuildingHeatConsumption = GetCellValueDecimal(row, "AT");
            apartment.Tariff = GetCellValueDecimal(row, "AM");
            apartment.CalculationByDistributors = GetCellValueDecimal(row, "AW");
            apartment.CalculationByArea = GetCellValueDecimal(row, "AV");
            apartment.TotalHouseConsumption = GetCellValueDecimal(row, "AU");
            apartment.HeatVolumeByDistributors = GetCellValueDecimal(row, "BE");
            apartment.ApartmentCoefficient = GetCellValueDecimal(row, "D");
            apartment.TotalGcal = GetCellValueDecimal(row, "M");
            apartment.RecalculatedReading = GetCellValueDecimal(row, "F");
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
            return apartment;
        }
        private void CalculateTotals(ApartmentHeating apartment, bool round)
        {
            apartment.MopHeatConsumption = apartment.BuildingHeatConsumption * apartment.MopHeatPercentage;
            apartment.IndividualHeatConsumption = apartment.BuildingHeatConsumption * apartment.IndividualPercentage;
            apartment.HeatPerSquareMeter = apartment.IndividualHeatConsumption / apartment.LivingArea;
            apartment.HeatVolumePerUnit = apartment.HeatVolumeByDistributors / apartment.TotalHouseConsumption;
            apartment.TotalIndividualConsumption = apartment.RecalculatedReading * apartment.HeatVolumePerUnit;
            apartment.TotalIndividualCharge = apartment.TotalIndividualConsumption * apartment.Tariff;
            decimal tempValue = apartment.MopHeatConsumption / apartment.LivingArea * apartment.ApartmentArea;
            apartment.ApartmentMopHeatShare = round ? Math.Round(tempValue, 4) : tempValue;
            apartment.MopCharge = apartment.ApartmentMopHeatShare * apartment.Tariff;
            apartment.TotalCharge = apartment.TotalIndividualCharge + apartment.MopCharge;

        }
        private decimal GetCellValueDecimal(IXLRow row, string cellAddress)
        {
            try
            {
                return row.Cell(cellAddress).GetValue<decimal>();
            }
            catch (Exception ex)
            {
                throw new CellValueException($"Ошибка при считывании значения из ячейки {cellAddress}{row.RowNumber()}");
            }
        }
    }
}
