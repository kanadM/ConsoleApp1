using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace ConsoleApplication
{

    public class ReconnRec
    { 
        public bool NotFoundInSupervisorReport { get; set; }
        public bool NotFoundInMasterReport { get; set; }
        internal bool IsCancelled
        {
            get
            {
                var res = Actual_order_Status.ToUpper().Trim() != "DELIVERED";
                if (IsUndelivered)
                    res = false;
                return res;
            }
        }

        internal bool IsUndelivered
        {
            get
            {
                return Actual_order_Status.ToUpper().Trim().StartsWith("PICKED UP") && Actual_order_Status.ToUpper().Trim().Contains("NOT");
            }
        }

        public ReconnRec()
        {

        }

        [DisplayName("Sr.No.")]
        public string Sr_No { get; set; }

        [DisplayName("Order Id")]
        public string OrderId { get; set; }

        [DisplayName("Vendor Name")]
        public string Vendor_Name { get; set; }

        [DisplayName("Outlet Name")]
        public string Outlet_Name { get; set; }

        [DisplayName("Delivery Date")]
        public string Delivery_Date { get; set; }

        [DisplayName("Transaction Type")]
        public string Transaction_Type { get; set; }

        [DisplayName("Order Status	")]
        public string Order_Status { get; set; }

        [DisplayName("Actual order Status")]
        public string Actual_order_Status { get; set; }

        [DisplayName("Order Amount")]
        public string Order_Amount { get; set; }

        [DisplayName("Delivery charges")]
        public string Delivery_charges { get; set; }

        [DisplayName("Actual Amount paid to vendor")]
        public string Actual_Amount_paid_to_vendor { get; set; }

        [DisplayName("Canclled Order/Disc	Remarks -Supervisor")]
        public string Canclled_Order_Disc_Remarks_Supervisor { get; set; }

        [DisplayName("Remarks -Supervisor")]
        public string Remarks_Supervisor { get; set; }

        [DisplayName("Remarks - Reconcilor")]
        public string Remarks_Reconcilor { get; set; }

        [DisplayName("Final Remarks")]
        public string Final_Remarks { get; set; }
    }
}
