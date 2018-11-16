using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace ConsoleApplication
{
    public class MasterRec
    {
        public MasterRec()
        {

        }

        [DisplayName("Serial No.")]
        public string Serial_No_ { get; set; }

        [DisplayName("Order Number")]
        public string Order_Number { get; set; }

        [DisplayName("Order Id")]
        public string Order_Id { get; set; }

        [DisplayName("Vendor Name")]
        public string Vendor_Name { get; set; }

        [DisplayName("Vendor Type")]
        public string Vendor_Type { get; set; }

        [DisplayName("Outlet Name")]
        public string Outlet_Name { get; set; }

        [DisplayName("Outlet Phone")]
        public string Outlet_Phone { get; set; }

        [DisplayName("Outlet Email")]
        public string Outlet_Email { get; set; }

        [DisplayName("Date of Booking")]
        public string Date_of_Booking { get; set; }

        [DisplayName("Delivery Date")]
        public string Delivery_Date { get; set; }

        [DisplayName("Delivery Station")]
        public string Delivery_Station { get; set; }

        [DisplayName("Transaction Type")]
        public string Transaction_Type { get; set; }

        [DisplayName("Order Status")]
        public string Order_Status { get; set; }

        [DisplayName("Amount")]
        public string Amount { get; set; }

        [DisplayName("PNR No.")]
        public string PNR_No_ { get; set; }

        [DisplayName("Coach")]
        public string Coach { get; set; }

        [DisplayName("Berth")]
        public string Berth { get; set; }

        [DisplayName("Train No.")]
        public string Train_No_ { get; set; }

        [DisplayName("Customer Phone")]
        public string Customer_Phone { get; set; }

        [DisplayName("Booked By")]
        public string Booked_By { get; set; }


    }
}
