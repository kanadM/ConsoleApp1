using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApp1
{
    public class ReconWorksheet
    {
        public List<ReconnRec> COD { get; set; }
        public List<ReconnRec> PRE_PAID { get; set; }
        public ReconWorksheet()
        {
            COD = new List<ReconnRec>();
            PRE_PAID= new List<ReconnRec>();
        }
    }
}
