using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HeatingReportMaker.Core.Models
{
    public class ApartmentHeating
    {
        public int ApartmentNumber { get; set; } // Номер квартиры
        public string Address { get; set; } = string.Empty; // Адресс дома
        public string ReportPeriod { get; set; } = string.Empty; // Отчетный период
        public decimal ApartmentArea { get; set; } // Площадь квартиры
        public decimal LivingArea { get; set; } // Площадь жилых помещений
        public decimal BuildingHeatConsumption { get; set; } // Расход ТЭ на отопление по показаниям ОДПУ
        public decimal Tariff {  get; set; } // Тариф на отопление
        public decimal CalculationByDistributors { get; set; } // Расчет по распределителям
        public decimal CalculationByArea { get; set; } // Расчет по площади
        public decimal MopHeatPercentage { get; set; } // Процент отопления на МОП
        public decimal MopHeatConsumption { get; set; } // Расход тепловой энергии на отопление на МОП 
        public decimal ApartmentMopHeatShare { get; set; } // Доля тепловой энергии на отопление МОП, приходящаяся на квартиру
        public decimal MopCharge { get; set; } // Плата за общедомовое тепло
        public decimal IndividualPercentage { get; set; } // Процент на индивидуальное потребление
        public decimal IndividualHeatConsumption { get; set; } // Индивидуальное потребление
        public decimal HeatPerSquareMeter { get; set; } // Расход тепловой энергии на 1 м² инд. отопление
        public decimal TotalHouseConsumption { get; set; } // Сумма всех показаний (потребленных условных единиц) со всех теплораспределителей в доме
        public decimal HeatVolumeByDistributors { get; set; } // Объем тепла по распределителям
        public decimal HeatVolumePerUnit { get; set; } // Объем тепла на  1 условную единицу
        public decimal ApartmentCoefficient { get; set; } // Коэффициент квартирный
        public decimal RecalculatedReading { get; set; } // Пересчитанные показания
        public decimal TotalIndividualConsumption { get; set; } // Индивидуальное потребление
        public decimal TotalIndividualCharge { get; set; } // Плата за индивидуальное потребление
        public decimal TotalCharge { get; set; } // Итоговая сумма
        public decimal TotalGcal { get; set; } // Всего Гкал
        public List<HeatDistributor> Distributors { get; set; } = new();
    }
}
