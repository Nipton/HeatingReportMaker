using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HeatingReportMaker.Core.Models
{
    public class HeatDistributor
    {
        public string DistributorNumber { get; set; } = "";
        public decimal RadiatorCoefficient { get; set; } // Коэффициент радиатора
        public decimal RecalculatedReading { get; private set; } // Пересчитанные показания
    }
}
