using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXmlCopy
{
    class Program
    {
        static void Main(string[] args)
        {
            Report.Copy(@"C:\Users\dli\Desktop\Working Files\OpenXML\EH-eCQM Report.xlsx",
                      @"C:\Users\dli\Desktop\Working Files\OpenXML\copy.xlsx");
        }
    }
}
