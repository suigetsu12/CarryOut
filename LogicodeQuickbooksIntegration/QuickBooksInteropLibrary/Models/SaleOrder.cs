using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuickBooksInteropLibrary.Models
{
    public class SaleOrder
    {
        public string ItemID { get; set; }
        //public string NameItem { get; set; }
        public List<SalesOrderItem> ListItem { get; set; }

        public string CustomerName { get; set; } // Custommer:JOB
        public string CustomerID { get; set; }

        public DateTime TxnDate { get; set; }  //Date

        public string BillAddress { get; set; }

        public int RefNumber { get; set; }   //S.O.NO

        public string ClassRefID { get; set; }  // Class

        public string ShipAddress { get; set; }  // Ship to

        public string TermsRefName { get; set; } //Terms
        public string TermRefID { get; set; }
        public DateTime DueDate { get; set; } //ETA

        public string NameSalesRepRef { get; set; } // REP

        public string FOB { get; set; } // FOB

        public string Memo { get; set; } //Memo

        public bool IsToBePrinted { get; set; }
        public bool IsToBeEmailed { get; set; }
        public bool IsManuallyClosed { get; set; }
        public bool IsFullyInvoiced { get; set; }
        public int ExchangeRate { get; set; }
        public string EditSequence { get; set; }
        public string TemplateRefName { get; set; }// Template
        public string TemplateRefID { get; set; }
        public string PostNumber { get; set; }
        public SaleOrder()
        {
            //NameItem = string.Empty;
            ListItem = null;
            CustomerName = string.Empty;
            TxnDate = DateTime.Now;
            BillAddress = string.Empty;
            ShipAddress = string.Empty;
            RefNumber = 0;
            TermsRefName = string.Empty;
            TemplateRefID = string.Empty;
            DueDate = DateTime.Now;
            NameSalesRepRef = string.Empty;
            FOB = string.Empty;
            Memo = string.Empty;
            IsToBePrinted = false;
            IsToBeEmailed = false;
            IsManuallyClosed = false;
            IsFullyInvoiced = false;
            ExchangeRate = 1;
            TemplateRefName = string.Empty;
            TemplateRefID = "80000018-1317851269";
            EditSequence = string.Empty;
            PostNumber = string.Empty;
        }     
    }
    public class SalesOrderItem
    {
        public string ItemID;
        public int Quantity;

        public SalesOrderItem()
        {
            Quantity = 1;
        }
    }

    //public class Item
    //{
    //    public string productId;
    //    public int quantity;

    //    public Item()
    //    {
    //        productId = string.Empty;
    //        quantity = 1;
    //    }
    //}

}
