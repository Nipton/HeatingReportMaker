using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
using HeatingReportMaker.Core.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HeatingReportMaker.Tests.IntegrationTests
{
    public class ExcelDataReaderIntegrationTests
    {
        [Fact]
        public void ReadApartmentData_ValidFile_ReturnsCorrectData()
        {
            var testFile = Path.Combine("TestData", "Test.xlsx");
            ExcelDataReader reader = new ExcelDataReader();

            var result = reader.ReadApartmentData(testFile, 4, false).ApartmentHeating;
            var resultWithRound = reader.ReadApartmentData(testFile, 4, true).ApartmentHeating;

            Assert.NotNull(result); 
            Assert.NotNull(resultWithRound);
            Assert.Equal(4, result.ApartmentNumber);
            Assert.Equal(56.3m, result.ApartmentArea);
            Assert.Equal(0.2142m, result.MopHeatPercentage);
            Assert.Equal(0.7858m, result.IndividualPercentage);
            Assert.Equal(114.14188926m, result.MopHeatConsumption);
            Assert.Equal(418.73341074m, result.IndividualHeatConsumption, 8);
            Assert.Equal(2208.1977m, result.RecalculatedReading);
            Assert.Equal(0.000393875941582334m, result.HeatVolumePerUnit, 8);
            Assert.Equal(0.0188403940887187m, result.HeatPerSquareMeter, 15);
            Assert.Equal(0.869755948287444m, result.TotalIndividualConsumption, 4);

            Assert.Equal(0.289138430767549m, result.ApartmentMopHeatShare, 15);
            Assert.Equal(846.212771174462m, result.MopCharge, 12);
            Assert.Equal(2545.48864117441m, result.TotalIndividualCharge, 2);
            Assert.Equal(3391.70141234887m, result.TotalCharge, 2);

            Assert.Equal(0.2891m, resultWithRound.ApartmentMopHeatShare);
            Assert.Equal(846.100297m, resultWithRound.MopCharge, 6);
            Assert.Equal(2545.617566m, resultWithRound.TotalIndividualCharge, 2);
            Assert.Equal(3391.717863m, resultWithRound.TotalCharge, 2);
        }
    }
}
