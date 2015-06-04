using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuickBooksInteropLibrary.Models
{
    public class Inventory
    {
        public int OdooId { get; set; }
        public string QBListID { get; set; }
        public string Name { get; set; }
        public string FullName { get; set; }
        public bool isActive { get; set; }
        public int SubLevel { get; set; }
        public string ManufacturerPartNumber { get; set; }
        public string SalesDesc { get; set; }
        public string PurchaseDesc { get; set; }
        public double SalesPrice { get; set; }
        public int QuantityOnHand { get; set; }
        public double PurchaseCost { get; set; }
        public double AverageCost { get; set; }
        public int QuantityOnOrder { get; set; }
        public int QuantityOnSalesOrder { get; set; }
        public Ref ParentRef { get; set; }
        public Ref IncomeAccountRef { get; set; }
        public Ref COGSAccountRef { get; set; }
        public Ref AssetAccountRef { get; set; }
        public Ref PrefVendorRef { get; set; }
        public string EditSequence { get; set; }
        public int Measure { get; set; }
        public int Category { get; set; }
        public int PucharseUnitOfMeasure { get; set; }
        public string Type { get; set; }
        public bool ApplyPurchase { get; set; }
        public int templateproduct { get; set; }

        public Inventory()
        {
            QBListID = string.Empty;
            Name = string.Empty;
            FullName = string.Empty;
            isActive = false;
            SubLevel = 0;
            ManufacturerPartNumber = string.Empty;
            SalesDesc = string.Empty;
            PurchaseDesc = string.Empty;
            SalesPrice = 0;
            QuantityOnHand = 0;
            PurchaseCost = 0;
            AverageCost = 0;
            QuantityOnOrder = 0;
            QuantityOnSalesOrder = 0;
            ParentRef = new Ref();
            ParentRef.ListID=string.Empty;
            ParentRef.FullName = string.Empty;
            ParentRef.Type = string.Empty;
            IncomeAccountRef = new Ref();
            IncomeAccountRef.ListID = string.Empty;
            IncomeAccountRef.FullName = string.Empty;
            IncomeAccountRef.Type = string.Empty;
            COGSAccountRef = new Ref();
            COGSAccountRef.ListID = string.Empty;
            COGSAccountRef.FullName = string.Empty;
            COGSAccountRef.Type = string.Empty;
            AssetAccountRef = new Ref();
            AssetAccountRef.ListID = string.Empty;
            AssetAccountRef.FullName = string.Empty;
            AssetAccountRef.Type = string.Empty;
            PrefVendorRef = new Ref();
            PrefVendorRef.ListID = string.Empty;
            PrefVendorRef.FullName = string.Empty;
            PrefVendorRef.Type = string.Empty;
            EditSequence = string.Empty;
            Measure = 1;
            templateproduct = 1;
            Category = 1;
            ApplyPurchase = true;
            Type = "consu";
            PucharseUnitOfMeasure = 1;
        }
    }

    public class Ref
    {
        public string ListID { get; set; }
        public string FullName { get; set; }
        public string Type { get; set; }
    }

    public enum ImageType
    {
        Small,
        Medium,
        Large
    }

    public class Category
    {
        public int Level { get; set; }
        public string Name { get; set; }
        public List<Category> SubItems { get; set; }
        public string LagasseId { get; set; }
        public string OdooId { get; set; }
        public Category()
        {
            SubItems = new List<Category>();
        }
    }
}
