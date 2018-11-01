using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace ConsoleApp1
{
    class DeliveryChargeRec
    {
        public DeliveryChargeRec()
        {

        }

        [DisplayName("IRCTC ID")]
        public string IRCTC_ID { get; set; }

        [DisplayName("Outlet Name")]
        public string Outlet_Name { get; set; }

        [DisplayName("Order Status")]
        public string Order_Status { get; set; }

        [DisplayName("Vendor Name")]
        public string Vendor_Name { get; set; }

        [DisplayName("Delivery Date")]
        public string Delivery_Date { get; set; }

        [DisplayName("Delivery Charges")]
        public string Delivery_Charges { get; set; }


    }
}
