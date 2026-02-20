using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HeatingReportMaker.Core.Exceptions
{
    public class CellValueException : Exception
    {
        public CellValueException(string message) : base(message) { }
    }
    public class TemplateException : Exception
    {
        public TemplateException(string message) : base(message) { }
    }
}
