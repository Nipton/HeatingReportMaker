using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HeatingReportMaker.Core.Models
{
    public class ApartmentReadResult
    {
        public bool Success { get; set; }
        public ApartmentHeating? ApartmentHeating { get; set; }
        public string? Message { get; set; }
    }
}
