using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuickBooksInteropLibrary.Models
{
    public class Customer
    {
        public int OdooId { get; set; }
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public string Street { get; set; }
        public string City { get; set; }
        public string Zip { get; set; }
        public string Website { get; set; }
        public string Email { get; set; }
        public bool IsSupplier { get; set; }
        public bool IsCustomer { get; set; }
        public string NotifyEmail { get; set; }
        public bool IsEmployee { get; set; }
        public string Phone { get; set; }
        public bool IsActive { get; set; }
        public bool IsCompany { get; set; }

        public string EditSequence { get; set; }

        public Customer()
        {
            Name = string.Empty;
            DisplayName = string.Empty;
            Street = string.Empty;
            City = string.Empty;
            Zip = string.Empty;
            Website = string.Empty;
            Email = string.Empty;
            IsSupplier = false;
            NotifyEmail = "always";
            IsEmployee = false;
            Phone = string.Empty;
            IsActive = false;
            IsCompany = false;
            EditSequence = string.Empty;
        }

        public string QBListId { get; set; }
    }
}
