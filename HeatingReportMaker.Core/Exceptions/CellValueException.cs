using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HeatingReportMaker.Core.Exceptions
{
    internal class CellValueException : Exception
    {
        public CellValueException(string message) : base(message) { }
    }
}
