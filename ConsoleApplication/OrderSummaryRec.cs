using System;
using System.Collections.Generic;
using System.Text;

namespace ConsoleApplication
{
    public class OrderSummaryRec
    {
        //select [Order Status],count(*) TotalRowsConsidered,[Trapigo Remarks] from SUPERVISOR$28July$ where [order id] is not null  group by [Order Status], [Trapigo Remarks]
        public string OrderStatus { get; set; }
        public int conut { get; set; }
        public string TrapigoRemarks { get; set; }
    }
}
