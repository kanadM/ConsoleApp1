using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace ConsoleApplication
{
    public class SupervisorRec
    {
        public SupervisorRec()
        {

        }

        [DisplayName("Order Id")]
        public string Order_Id { get; set; }

        [DisplayName("Vendor Name")]
        public string Vendor_Name { get; set; }

        [DisplayName("Delivery Date")]
        public string Delivery_Date { get; set; }

        [DisplayName("Transaction Type")]
        public string Transaction_Type { get; set; }

        [DisplayName("Order Status")]
        public string Order_Status { get; set; }

        [DisplayName("Pickup Boy")]
        public string Pickup_Boy { get; set; }

        [DisplayName("Delivery Boy")]
        public string Delivery_Boy { get; set; }

        [DisplayName("IRCTC Dashboard Amount")]
        public string IRCTC_Dashboard_Amount { get; set; }

        [DisplayName("Amount received from customer")]
        public string Amount_received_from_customer { get; set; }

        [DisplayName("Delivery Charges")]
        public string Delivery_Charges { get; set; }

        [DisplayName("Bulk order charges")]
        public string Bulk_Order_Charges { get; set; }

        [DisplayName("Trapigo Payment To Vendor")]
        public string Trapigo_Payment_To_Vendor { get; set; }

        [DisplayName("Trapigo Remarks")]
        public string Trapigo_Remarks { get; set; }
        public bool _IsCancelled
        {
            get
            {
                return Order_Status.ToUpper().Trim() != "DELIVERED";
            } 
        }
    }
}
