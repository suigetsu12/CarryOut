using Interop.QBFC13;
using QuickBooksInteropLibrary.Models;
using QuickBooksInteropLibrary.SessionFramework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuickBooksInteropLibrary
{
    public class QuickbooksInterop
    {
        SessionManager sessionManager;
        private short maxVersion;

        public List<Customer> LoadCustomers()
        {
            string request = "CustomerQueryRq";
            ConnectToQB();
            int count = getCount(request);
            IMsgSetResponse responseMsgSet = processRequestFromQB(buildCustomerQueryRq(new string[] { "FullName", "FirstName", "Email", "Phone", "Website", "Salutation", "BillAddress", "ListID", "EditSequence" }, null, null));
            var customerList = parseCustomerQueryRs(responseMsgSet, count);
            disconnectFromQB();
            return customerList;
        }

        public Customer LoadCustomer(string listId, string name)
        {
            //
            string request = "CustomerQueryRq";
            ConnectToQB();
            //int count = getCount(request);
            IMsgSetResponse responseMsgSet = processRequestFromQB(buildCustomerQueryRq(new string[] { "FullName", "FirstName", "Email", "Phone", "Website", "Salutation", "BillAddress", "ListID", "EditSequence" }, listId, name));
            var customerList = parseCustomerQueryRs(responseMsgSet, 1).FirstOrDefault();
            disconnectFromQB();
            return customerList;
        }

        public List<Inventory> LoadInventories()
        {
            string request = "ItemInventoryQueryRq";
            ConnectToQB();
            int count = getCountInventory(request);
            IMsgSetResponse responseMsgSet = processRequestFromQB(buildItemInventoryQueryRq(new string[] { "AssetAccountRef", "AverageCost", "BarCodeValue", "ClassRef", "COGSAccountRef", "DataExtRetList", "EditSequence", "ExternalGUID", "FullName", "IncomeAccountRef", "IsActive", "IsTaxIncluded", "ListID", "ManufacturerPartNumber", "Max", "Name", "ParentRef", "PrefVendorRef", "PurchaseCost", "PurchaseDesc", "PurchaseTaxCodeRef", "QuantityOnHand", "QuantityOnOrder", "QuantityOnSalesOrder", "ReorderPoint", "SalesDesc", "SalesPrice", "SalesTaxCodeRef", "Sublevel", "TimeCreated", "TimeModified", "Type", "UnitOfMeasureSetRef", "EditSequence" }, null, null));
            var itemInventoryList = parseItemInventoryQueryRs(responseMsgSet, count);
            disconnectFromQB();
            return itemInventoryList;
        }

        private int getCountInventory(string request)
        {
            IMsgSetResponse responseMsgSet = processRequestFromQB(buildDataCountItemInventoryQuery(request));
            int count = parseRsForCount(responseMsgSet);
            return count;
        }

        private IMsgSetRequest buildDataCountItemInventoryQuery(string request)
        {
            IMsgSetRequest requestMsgSet = sessionManager.getMsgSetRequest();
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
            IItemInventoryQuery inventoryQuery = requestMsgSet.AppendItemInventoryQueryRq();
            inventoryQuery.metaData.SetValue(ENmetaData.mdMetaDataOnly);
            return requestMsgSet;
        }

        public Inventory LoadInventory(string listId, string name)
        {
            string request = "ItemInventoryQueryRq";
            ConnectToQB();
            IMsgSetResponse responseMsgSet = processRequestFromQB(buildItemInventoryQueryRq(new string[] { "AssetAccountRef", "AverageCost", "BarCodeValue", "ClassRef", "COGSAccountRef", "DataExtRetList", "EditSequence", "ExternalGUID", "FullName", "IncomeAccountRef", "IsActive", "IsTaxIncluded", "ListID", "ManufacturerPartNumber", "Max", "Name", "ParentRef", "PrefVendorRef", "PurchaseCost", "PurchaseDesc", "PurchaseTaxCodeRef", "QuantityOnHand", "QuantityOnOrder", "QuantityOnSalesOrder", "ReorderPoint", "SalesDesc", "SalesPrice", "SalesTaxCodeRef", "Sublevel", "TimeCreated", "TimeModified", "Type", "UnitOfMeasureSetRef", "EditSequence" }, listId, name));
            var inventList = parseItemInventoryQueryRs(responseMsgSet, 1).FirstOrDefault();
            disconnectFromQB();
            return inventList;
        }

        private List<Inventory> parseItemInventoryQueryRs(IMsgSetResponse responseMsgSet, int count)
        {
            string[] retVal = new string[count];
            var result = new List<Inventory>();
            IResponse response = responseMsgSet.ResponseList.GetAt(0);
            int statusCode = response.StatusCode;
            if (statusCode == 0)
            {
                IItemInventoryRetList itemInventRetList = response.Detail as IItemInventoryRetList;
                //ICustomerRetList custRetList = response.Detail as ICustomerRetList;

                for (int i = 0; i < count; i++)
                {
                    var itemInvent = new Inventory();

                    if (itemInventRetList.GetAt(i).ListID != null)
                    {
                        itemInvent.QBListID = itemInventRetList.GetAt(i).ListID.GetValue().ToString();
                    }

                    if (itemInventRetList.GetAt(i).Name != null)
                    {
                        itemInvent.Name = itemInventRetList.GetAt(i).Name.GetValue().ToString();
                    }

                    if (itemInventRetList.GetAt(i).FullName != null)
                    {
                        itemInvent.FullName = itemInventRetList.GetAt(i).FullName.GetValue().ToString();
                    }
                    if (itemInventRetList.GetAt(i).IsActive != null)
                    {
                        itemInvent.isActive = Convert.ToBoolean(itemInventRetList.GetAt(i).IsActive.GetValue());
                    }

                    if (itemInventRetList.GetAt(i).Sublevel != null)
                    {
                        itemInvent.SubLevel = Convert.ToInt32(itemInventRetList.GetAt(i).Sublevel.GetValue());
                    }

                    if (itemInventRetList.GetAt(i).ManufacturerPartNumber != null)
                    {
                        itemInvent.ManufacturerPartNumber = itemInventRetList.GetAt(i).ManufacturerPartNumber.GetValue().ToString();
                    }

                    if (itemInventRetList.GetAt(i).SalesDesc != null)
                    {
                        itemInvent.SalesDesc = itemInventRetList.GetAt(i).SalesDesc.GetValue().ToString();
                    }

                    if (itemInventRetList.GetAt(i).PurchaseDesc != null)
                    {
                        itemInvent.PurchaseDesc = itemInventRetList.GetAt(i).PurchaseDesc.GetValue().ToString();
                    }

                    if (itemInventRetList.GetAt(i).SalesPrice != null)
                    {
                        itemInvent.SalesPrice = Convert.ToDouble(itemInventRetList.GetAt(i).SalesPrice.GetValue());
                    }

                    if (itemInventRetList.GetAt(i).QuantityOnHand != null)
                    {
                        itemInvent.QuantityOnHand = Convert.ToInt32(itemInventRetList.GetAt(i).QuantityOnHand.GetValue());
                    }

                    if (itemInventRetList.GetAt(i).PurchaseCost != null)
                    {
                        itemInvent.PurchaseCost = Convert.ToDouble(itemInventRetList.GetAt(i).PurchaseCost.GetValue());
                    }

                    if (itemInventRetList.GetAt(i).AverageCost != null)
                    {
                        itemInvent.AverageCost = Convert.ToDouble(itemInventRetList.GetAt(i).AverageCost.GetValue());
                    }

                    if (itemInventRetList.GetAt(i).QuantityOnOrder != null)
                    {
                        itemInvent.QuantityOnOrder = Convert.ToInt32(itemInventRetList.GetAt(i).QuantityOnOrder.GetValue());
                    }

                    if (itemInventRetList.GetAt(i).QuantityOnSalesOrder != null)
                    {
                        itemInvent.QuantityOnSalesOrder = Convert.ToInt32(itemInventRetList.GetAt(i).QuantityOnSalesOrder.GetValue());
                    }

                    if (itemInventRetList.GetAt(i).ParentRef != null)
                    {
                        itemInvent.ParentRef = new Ref();
                        if (itemInventRetList.GetAt(i).ParentRef.ListID != null)
                        {
                            itemInvent.ParentRef.ListID = itemInventRetList.GetAt(i).ParentRef.ListID.GetValue().ToString();
                        }
                        if (itemInventRetList.GetAt(i).ParentRef.FullName != null)
                        {
                            itemInvent.ParentRef.FullName = itemInventRetList.GetAt(i).ParentRef.FullName.GetValue().ToString();
                        }
                        if (itemInventRetList.GetAt(i).ParentRef.Type != null)
                        {
                            itemInvent.ParentRef.Type = itemInventRetList.GetAt(i).ParentRef.Type.GetValue().ToString();
                        }
                    }

                    if (itemInventRetList.GetAt(i).IncomeAccountRef != null)
                    {
                        itemInvent.IncomeAccountRef = new Ref();
                        if (itemInventRetList.GetAt(i).IncomeAccountRef.ListID != null)
                        {
                            itemInvent.IncomeAccountRef.ListID = itemInventRetList.GetAt(i).IncomeAccountRef.ListID.GetValue().ToString();
                        }
                        if (itemInventRetList.GetAt(i).IncomeAccountRef.FullName != null)
                        {
                            itemInvent.IncomeAccountRef.FullName = itemInventRetList.GetAt(i).IncomeAccountRef.FullName.GetValue().ToString();
                        }
                        if (itemInventRetList.GetAt(i).IncomeAccountRef.Type != null)
                        {
                            itemInvent.IncomeAccountRef.Type = itemInventRetList.GetAt(i).IncomeAccountRef.Type.GetValue().ToString();
                        }
                    }

                    if (itemInventRetList.GetAt(i).COGSAccountRef != null)
                    {
                        itemInvent.COGSAccountRef = new Ref();
                        if (itemInventRetList.GetAt(i).COGSAccountRef.ListID != null)
                        {
                            itemInvent.COGSAccountRef.ListID = itemInventRetList.GetAt(i).COGSAccountRef.ListID.GetValue().ToString();
                        }
                        if (itemInventRetList.GetAt(i).COGSAccountRef.FullName != null)
                        {
                            itemInvent.COGSAccountRef.FullName = itemInventRetList.GetAt(i).COGSAccountRef.FullName.GetValue().ToString();
                        }
                        if (itemInventRetList.GetAt(i).COGSAccountRef.Type != null)
                        {
                            itemInvent.COGSAccountRef.Type = itemInventRetList.GetAt(i).COGSAccountRef.Type.GetValue().ToString();
                        }
                    }

                    if (itemInventRetList.GetAt(i).AssetAccountRef != null)
                    {
                        itemInvent.AssetAccountRef = new Ref();
                        if (itemInventRetList.GetAt(i).AssetAccountRef.ListID != null)
                        {
                            itemInvent.AssetAccountRef.ListID = itemInventRetList.GetAt(i).AssetAccountRef.ListID.GetValue().ToString();
                        }
                        if (itemInventRetList.GetAt(i).AssetAccountRef.FullName != null)
                        {
                            itemInvent.AssetAccountRef.FullName = itemInventRetList.GetAt(i).AssetAccountRef.FullName.GetValue().ToString();
                        }
                        if (itemInventRetList.GetAt(i).AssetAccountRef.Type != null)
                        {
                            itemInvent.AssetAccountRef.Type = itemInventRetList.GetAt(i).AssetAccountRef.Type.GetValue().ToString();
                        }
                    }

                    if (itemInventRetList.GetAt(i).PrefVendorRef != null)
                    {
                        itemInvent.PrefVendorRef = new Ref();
                        if (itemInventRetList.GetAt(i).PrefVendorRef.ListID != null)
                        {
                            itemInvent.PrefVendorRef.ListID = itemInventRetList.GetAt(i).PrefVendorRef.ListID.GetValue().ToString();
                        }
                        if (itemInventRetList.GetAt(i).PrefVendorRef.FullName != null)
                        {
                            itemInvent.PrefVendorRef.FullName = itemInventRetList.GetAt(i).PrefVendorRef.FullName.GetValue().ToString();
                        }
                        if (itemInventRetList.GetAt(i).PrefVendorRef.Type != null)
                        {
                            itemInvent.PrefVendorRef.Type = itemInventRetList.GetAt(i).PrefVendorRef.Type.GetValue().ToString();
                        }
                    }

                    if (itemInventRetList.GetAt(i).EditSequence != null)
                    {
                        var editSequence = itemInventRetList.GetAt(i).EditSequence.GetValue().ToString();
                        if (editSequence != null)
                        {
                            itemInvent.EditSequence = editSequence;
                        }
                    }

                    result.Add(itemInvent);
                }
            }
            return result;
        }

        private IMsgSetRequest buildItemInventoryQueryRq(string[] includeRetElement, string listId, string fullname)
        {
            IMsgSetRequest requestMsgSet = sessionManager.getMsgSetRequest();
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
            IItemInventoryQuery invenQuery = requestMsgSet.AppendItemInventoryQueryRq();
            if (listId != null)
            {
                invenQuery.ORListQueryWithOwnerIDAndClass.ListIDList.Add(listId);
            }
            if (fullname != null)
            {
                invenQuery.ORListQueryWithOwnerIDAndClass.FullNameList.Add(fullname);
            }
            for (int x = 0; x < includeRetElement.Length; x++)
            {
                invenQuery.IncludeRetElementList.Add(includeRetElement[x]);
            }
            return requestMsgSet;
        }

        private void ConnectToQB()
        {
            sessionManager = SessionManager.getInstance();
            maxVersion = sessionManager.QBsdkMajorVersion;
        }
        private int getCount(string request)
        {
            IMsgSetResponse responseMsgSet = processRequestFromQB(buildDataCountQuery(request));
            int count = parseRsForCount(responseMsgSet);
            return count;
        }
        private int parseRsForCount(IMsgSetResponse responseMsgSet)
        {
            int ret = -1;
            try
            {
                IResponse response = responseMsgSet.ResponseList.GetAt(0);
                ret = response.retCount;
            }
            catch (Exception e)
            {
                // MessageBox.Show("Error encountered: " + e.Message);
                ret = -1;
            }
            return ret;
        }
        private IMsgSetResponse processRequestFromQB(IMsgSetRequest requestSet)
        {
            try
            {
                //MessageBox.Show(requestSet.ToXMLString());
                IMsgSetResponse responseSet = sessionManager.doRequest(true, ref requestSet);
                //MessageBox.Show(responseSet.ToXMLString());
                return responseSet;
            }
            catch (Exception e)
            {
                //MessageBox.Show(e.Message);
                return null;
            }
        }
        private IMsgSetRequest buildDataCountQuery(string request)
        {
            IMsgSetRequest requestMsgSet = sessionManager.getMsgSetRequest();
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
            ICustomerQuery custQuery = requestMsgSet.AppendCustomerQueryRq();
            custQuery.metaData.SetValue(ENmetaData.mdMetaDataOnly);
            return requestMsgSet;
        }
        private IMsgSetRequest buildCustomerQueryRq(string[] includeRetElement, string listId, string fullname)
        {
            IMsgSetRequest requestMsgSet = sessionManager.getMsgSetRequest();
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
            ICustomerQuery custQuery = requestMsgSet.AppendCustomerQueryRq();
            if (listId != null)
            {

                //custQuery.ORCustomerListQuery.FullNameList.Add(fullName);
                custQuery.ORCustomerListQuery.ListIDList.Add(listId);

            }
            if (fullname != null)
            {

                //custQuery.ORCustomerListQuery.FullNameList.Add(fullName);
                custQuery.ORCustomerListQuery.FullNameList.Add(fullname);

            }
            //custQuery.ORCustomerListQuery.CustomerListFilter.FromModifiedDate.SetValue(DateTime.Now.Subtract(new TimeSpan(1, 0, 0, 0)), false);
            //custQuery.ORCustomerListQuery.CustomerListFilter.ToModifiedDate.SetValue(DateTime.Now, false);
            for (int x = 0; x < includeRetElement.Length; x++)
            {
                custQuery.IncludeRetElementList.Add(includeRetElement[x]);
            }
            return requestMsgSet;
        }

        private IMsgSetRequest buildCustomerUpdateRq(Customer customer)
        {
            IMsgSetRequest requestMsgSet = sessionManager.getMsgSetRequest();
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
            ICustomerMod custQuery = requestMsgSet.AppendCustomerModRq();
            //custQuery.Name.SetValue(customer.Name);
            custQuery.ListID.SetValue(customer.QBListId);
            //custQuery.BillAddress.City.SetValue(customer.City);
            custQuery.Email.SetValue(customer.Email);
            custQuery.Phone.SetValue(customer.Phone);
            //custQuery.BillAddress.PostalCode.SetValue(customer.Zip);
            //custQuery.BillAddress.Addr2.SetValue(customer.Street);
            custQuery.EditSequence.SetValue(customer.EditSequence);
            return requestMsgSet;
        }



        private IMsgSetRequest buildCustomerAddRq(Customer customer)
        {
            IMsgSetRequest requestMsgSet = sessionManager.getMsgSetRequest();
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
            ICustomerAdd custQuery = requestMsgSet.AppendCustomerAddRq();
            //custQuery.ExternalGUID.SetValue(customer.OdooId.ToString());
            custQuery.Name.SetValue(customer.Name);
            //custQuery.Name.SetValue(customer.Name);
            //custQuery.BillAddress.City.SetValue(customer.City);
            custQuery.Email.SetValue(customer.Email);
            //custQuery.Phone.SetValue(customer.Phone);
            //custQuery.BillAddress.PostalCode.SetValue(customer.Zip);
            //custQuery.BillAddress.Addr2.SetValue(customer.Street);
            return requestMsgSet;
        }

        private bool parseCustomerCRUDRs(IMsgSetResponse responseMsgSet)
        {
            /*
             <?xml version="1.0" ?> 
             <QBXML>
             <QBXMLMsgsRs>
             <CustomerQueryRs requestID="1" statusCode="0" statusSeverity="Info" statusMessage="Status OK">
                 <CustomerRet>
                     <FullName>Abercrombie, Kristy</FullName> 
                 </CustomerRet>
             </CustomerQueryRs>
             </QBXMLMsgsRs>
             </QBXML>    
            */

            IResponse response = responseMsgSet.ResponseList.GetAt(0);
            int statusCode = response.StatusCode;
            if (statusCode == 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private List<Customer> parseCustomerQueryRs(IMsgSetResponse responseMsgSet, int count)
        {
            /*
             <?xml version="1.0" ?> 
             <QBXML>
             <QBXMLMsgsRs>
             <CustomerQueryRs requestID="1" statusCode="0" statusSeverity="Info" statusMessage="Status OK">
                 <CustomerRet>
                     <FullName>Abercrombie, Kristy</FullName> 
                 </CustomerRet>
             </CustomerQueryRs>
             </QBXMLMsgsRs>
             </QBXML>    
            */
            string[] retVal = new string[count];
            var result = new List<Customer>();
            IResponse response = responseMsgSet.ResponseList.GetAt(0);
            int statusCode = response.StatusCode;
            if (statusCode == 0)
            {
                ICustomerRetList custRetList = response.Detail as ICustomerRetList;

                for (int i = 0; i < count; i++)
                {
                    var customer = new Customer();
                    string fullName = null;
                    if (custRetList.GetAt(i).FullName != null)
                    {
                        fullName = custRetList.GetAt(i).FullName.GetValue().ToString();
                        if (fullName != null)
                        {
                            //customer.Name = fullName;
                            customer.DisplayName = fullName;
                        }
                    }
                    if (custRetList.GetAt(i).Name != null)
                    {
                        var name = custRetList.GetAt(i).Name.GetValue().ToString();
                        if (name != null)
                        {
                            //customer.Name = fullName;
                            customer.Name = name;
                        }
                    }
                    if (custRetList.GetAt(i).ListID != null)
                    {
                        var qbListId = custRetList.GetAt(i).ListID.GetValue().ToString();
                        if (qbListId != null)
                        {
                            //customer.Name = fullName;
                            customer.QBListId = qbListId;
                        }
                    }
                    if (custRetList.GetAt(i).EditSequence != null)
                    {
                        var editSequence = custRetList.GetAt(i).EditSequence.GetValue().ToString();
                        if (editSequence != null)
                        {
                            //customer.Name = fullName;
                            customer.EditSequence = editSequence;
                        }
                    }

                    IAddress billAddress = null;
                    if (custRetList.GetAt(i).BillAddress != null)
                    {
                        billAddress = custRetList.GetAt(i).BillAddress;
                        string addr1 = "", addr2 = "", addr3 = "", addr4 = "", addr5 = "";
                        string city = "", state = "", postalcode = "";
                        if (billAddress != null)
                        {
                            var fullAddress = string.Empty;
                            // if (billAddress.Addr1 != null) fullAddress += billAddress.Addr1.GetValue().ToString();
                            if (billAddress.Addr2 != null) fullAddress = billAddress.Addr2.GetValue().ToString();
                            // if (billAddress.Addr3 != null) fullAddress += "\n" + billAddress.Addr3.GetValue().ToString();
                            //  if (billAddress.Addr4 != null) fullAddress += "\n" + billAddress.Addr4.GetValue().ToString();
                            // if (billAddress.Addr5 != null) fullAddress += "\n" + billAddress.Addr5.GetValue().ToString();
                            //  if (billAddress.City != null) customer.City = billAddress.City.GetValue().ToString();
                            //  if (billAddress.State != null) state = billAddress.State.GetValue().ToString();
                            if (billAddress.PostalCode != null) customer.Zip = billAddress.PostalCode.GetValue().ToString();
                            customer.Street = fullAddress;
                        }
                    }
                    if (custRetList.GetAt(i).Email != null)
                    {
                        var email = custRetList.GetAt(i).Email.GetValue().ToString();
                        if (email != null)
                        {
                            //customer.Name = fullName;
                            customer.Email = email;
                        }
                    }
                    if (custRetList.GetAt(i).Phone != null)
                    {
                        var phone = custRetList.GetAt(i).Phone.GetValue().ToString();
                        if (phone != null)
                        {
                            //customer.Name = fullName;
                            customer.Phone = phone;
                        }
                    }

                    customer.IsActive = true;
                    customer.IsCustomer = true;
                    customer.IsSupplier = false;
                    customer.NotifyEmail = "always";
                    customer.IsCompany = false;

                    //IAddress shipAddress = null;
                    //if (custRetList.GetAt(i).ShipAddress != null)
                    //{
                    //    shipAddress = custRetList.GetAt(i).ShipAddress;
                    //    string addr1 = "", addr2 = "", addr3 = "", addr4 = "", addr5 = "";
                    //    string city = "", state = "", postalcode = "";
                    //    if (shipAddress != null)
                    //    {
                    //        if (shipAddress.Addr1 != null) addr1 = shipAddress.Addr1.GetValue().ToString();
                    //        if (shipAddress.Addr1 != null) addr1 = shipAddress.Addr1.GetValue().ToString();
                    //        if (shipAddress.Addr2 != null) addr2 = shipAddress.Addr2.GetValue().ToString();
                    //        if (shipAddress.Addr3 != null) addr3 = shipAddress.Addr3.GetValue().ToString();
                    //        if (shipAddress.Addr4 != null) addr4 = shipAddress.Addr4.GetValue().ToString();
                    //        if (shipAddress.Addr5 != null) addr5 = shipAddress.Addr5.GetValue().ToString();
                    //        if (shipAddress.City != null) city = shipAddress.City.GetValue().ToString();
                    //        if (shipAddress.State != null) state = shipAddress.State.GetValue().ToString();
                    //        if (shipAddress.PostalCode != null) postalcode = shipAddress.PostalCode.GetValue().ToString();

                    //        // RESUME HERE
                    //        retVal[i] = addr1 + "\r\n" + addr2 + "\r\n"
                    //            + addr3 + "\r\n"
                    //            + city + "\r\n" + state + "\r\n" + postalcode;
                    //    }
                    //}
                    string currencyRef = null;
                    if (custRetList.GetAt(i).CurrencyRef != null)
                    {
                        currencyRef = custRetList.GetAt(i).CurrencyRef.FullName.GetValue().ToString();
                        if (currencyRef != null)
                        {
                            retVal[i] = currencyRef;
                        }
                    }

                    result.Add(customer);
                }
            }
            return result;
        }
        private void disconnectFromQB()
        {
            if (sessionManager != null)
            {
                try
                {
                    sessionManager.endSession();
                    sessionManager.closeConnection();
                    sessionManager = null;
                }
                catch (Exception e)
                {
                    // MessageBox.Show(e.Message);
                }
            }
        }

        public bool UpdateCustomer(Customer customer)
        {
            ConnectToQB();
            var tempCustomer = LoadCustomer(customer.QBListId, null);
            customer.EditSequence = tempCustomer.EditSequence;
            ConnectToQB();
            IMsgSetResponse responseMsgSet = processRequestFromQB(buildCustomerUpdateRq(customer));
            var result = parseCustomerCRUDRs(responseMsgSet);
            disconnectFromQB();
            return result;
        }

        public bool AddCustomer(Customer customer)
        {
            ConnectToQB();
            IMsgSetResponse responseMsgSet = processRequestFromQB(buildCustomerAddRq(customer));
            var result = parseCustomerCRUDRs(responseMsgSet);
            disconnectFromQB();
            return result;
        }

        public bool DeleteCustomer(string qblistid)
        {
            ConnectToQB();
            IMsgSetResponse responseMsgSet = processRequestFromQB(buildCustomerDeleteRq(qblistid));
            var result = parseCustomerCRUDRs(responseMsgSet);
            disconnectFromQB();
            return result;
        }

        #region Inventory
        public bool AddInventory(Inventory inventory)
        {
            ConnectToQB();
            IMsgSetResponse responseMsgSet = processRequestFromQB(buildItemInventoryAddRq(inventory));
            var result = parseCustomerCRUDRs(responseMsgSet);
            disconnectFromQB();
            return result;
        }

        public bool UpdateInventory(Inventory inventory)
        {
            ConnectToQB();
            var tempInventory = LoadInventory(inventory.QBListID, null);
            inventory.EditSequence = tempInventory.EditSequence;
            ConnectToQB();
            IMsgSetResponse responseMsgSet = processRequestFromQB(buildItemInventoryUpdateRq(inventory));
            var result = parseCustomerCRUDRs(responseMsgSet);
            disconnectFromQB();
            return result;
        }

        public bool DeleteInventory(string qblistid)
        {
            ConnectToQB();
            IMsgSetResponse responseMsgSet = processRequestFromQB(buildInventoryDeleteRq(qblistid));
            var result = parseCustomerCRUDRs(responseMsgSet);
            disconnectFromQB();
            return result;
        }

        private IMsgSetRequest buildItemInventoryAddRq(Inventory inventory)
        {
            IMsgSetRequest requestMsgSet = sessionManager.getMsgSetRequest();
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
            IItemInventoryAdd inventQuery = requestMsgSet.AppendItemInventoryAddRq();
            inventQuery.Name.SetValue(inventory.Name);
            inventQuery.SalesPrice.SetValue(inventory.SalesPrice);
            inventQuery.IsActive.SetValue(inventory.isActive);
            inventQuery.SalesDesc.SetValue(inventory.SalesDesc);
            if (inventory.ParentRef != null)
            {
                if (inventory.ParentRef.FullName != "") inventQuery.ParentRef.FullName.SetValue(inventory.ParentRef.FullName);
                if (inventory.ParentRef.ListID != "") inventQuery.ParentRef.ListID.SetValue(inventory.ParentRef.ListID);
            }
            if (inventory.ParentRef.FullName == "" && inventory.ParentRef.ListID == "")
            {
                inventQuery.IsActive.SetValue(true);
                inventQuery.PurchaseCost.SetValue(0);
                inventQuery.QuantityOnHand.SetValue(0);
            }
            inventQuery.AssetAccountRef.FullName.SetValue("Inventories");
            inventQuery.AssetAccountRef.ListID.SetValue("8000003D-1262129210");
            return requestMsgSet;
        }

        private IMsgSetRequest buildItemInventoryUpdateRq(Inventory inventory)
        {
            IMsgSetRequest requestMsgSet = sessionManager.getMsgSetRequest();
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
            IItemInventoryMod inventQuery = requestMsgSet.AppendItemInventoryModRq();
            inventQuery.ListID.SetValue(inventory.QBListID);
            inventQuery.Name.SetValue(inventory.Name);
            inventQuery.SalesPrice.SetValue(inventory.SalesPrice);
            inventQuery.IsActive.SetValue(inventory.isActive);
            inventQuery.SalesDesc.SetValue(inventory.SalesDesc);
            inventQuery.EditSequence.SetValue(inventory.EditSequence);

            return requestMsgSet;
        }

        private IMsgSetRequest buildInventoryDeleteRq(string qblistid)
        {
            IMsgSetRequest requestMsgSet = sessionManager.getMsgSetRequest();
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
            IListDel inventQuery = requestMsgSet.AppendListDelRq();
            inventQuery.ListDelType.SetValue(ENListDelType.ldtItemInventory);
            inventQuery.ListID.SetValue(qblistid);
            return requestMsgSet;
        }
        #endregion

        #region SaleOder

        public List<SaleOrder> LoadSaleOrder()
        {
            string request = "SalesOrderQueryRq";
            ConnectToQB();
            int count = getCountSaleOrder(request);
            IMsgSetResponse responseMsgSet = processRequestFromQB(buildSalesOrderQueryRq(new string[] { "CustomerRef", "SalesOrderLineRet" }, null, null));
            var saleorderList = parseSalesOrderQueryRs(responseMsgSet, count);
            //disconnectFromQB;
            return saleorderList;
        }

        public SaleOrder LoadSaleOrder(string listId, string name)
        {
            //
            string request = "SalesOrderQueryRq";
            ConnectToQB();
            //int count = getCount(request);
            IMsgSetResponse responseMsgSet = processRequestFromQB(buildSalesOrderQueryRq(new string[] { "CustomerRef", "SalesOrderLineRet" }, listId, name));
            var saleorderList = parseSalesOrderQueryRs(responseMsgSet, 1).FirstOrDefault();
            disconnectFromQB();
            return saleorderList;
        }

        private int getCountSaleOrder(string request)
        {
            IMsgSetResponse responseMsgSet = processRequestFromQB(buildDataCountSaleOrderQuery(request));
            int count = parseRsForCount(responseMsgSet);
            return count;
        }

        private IMsgSetRequest buildDataCountSaleOrderQuery(string request)
        {
            IMsgSetRequest requestMsgSet = sessionManager.getMsgSetRequest();
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
            ISalesOrderQuery saleorderQuery = requestMsgSet.AppendSalesOrderQueryRq();
            saleorderQuery.metaData.SetValue(ENmetaData.mdMetaDataOnly);
            return requestMsgSet;
        }

        private IMsgSetRequest buildSalesOrderAddRq(SaleOrder salesOrder)
        {
            IMsgSetRequest requestMsgSet = sessionManager.getMsgSetRequest();
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
            ISalesOrderAdd salesOrderQuery = requestMsgSet.AppendSalesOrderAddRq();

            salesOrderQuery.CustomerRef.ListID.SetValue(salesOrder.CustomerID);

            foreach (var list in salesOrder.ListItem)
            {
                IORSalesOrderLineAdd line = salesOrderQuery.ORSalesOrderLineAddList.Append();
                line.SalesOrderLineAdd.ItemRef.ListID.SetValue(list.ItemID);
                line.SalesOrderLineAdd.Quantity.SetValue(list.Quantity);
            }

            //salesOrderQuery.BillAddress.Addr1.SetValue(salesOrder.BillAddress);
            //salesOrderQuery.ShipAddress.Addr1.SetValue(salesOrder.ShipAddress);
            //salesOrderQuery.Memo.SetValue(salesOrder.Memo);
            //salesOrderQuery.ClassRef.FullName.SetValue(salesOrder.ClassRefID);
            salesOrderQuery.TemplateRef.ListID.SetValue(salesOrder.TemplateRefID);
            //salesOrderQuery.PONumber.SetValue(salesOrder.PostNumber);
            //salesOrderQuery.TermsRef.ListID.SetValue(salesOrder.TermRefID);
            //salesOrderQuery.FOB.SetValue(salesOrder.FOB);

            return requestMsgSet;
        }

        private List<SaleOrder> parseSalesOrderQueryRs(IMsgSetResponse responseMsgSet, int count)
        {
            string[] retVal = new string[count];
            var result = new List<SaleOrder>();
            IResponse response = responseMsgSet.ResponseList.GetAt(0);

            int statusCode = response.StatusCode;
            if (statusCode == 0)
            {
                ISalesOrderRetList saleorderRetList = response.Detail as ISalesOrderRetList;

                for (int i = 0; i < count; i++)
                {
                    var salesorder = new SaleOrder();

                    if (saleorderRetList.GetAt(i).CustomerRef.FullName != null)
                    {
                        var customername = saleorderRetList.GetAt(i).CustomerRef.FullName.GetValue().ToString();
                        if (customername != null)
                        {
                            salesorder.CustomerName = customername;
                        }
                    }

                    //if (saleorderRetList.GetAt(i).ORSalesOrderLineRetList.GetAt(0).SalesOrderLineRet.ItemRef != null)
                    //{
                    //    var nameItem = saleorderRetList.GetAt(i).ORSalesOrderLineRetList.GetAt(0).SalesOrderLineRet.ItemRef.FullName.GetValue().ToString();
                    //    if (nameItem != null)
                    //    {
                    //        salesorder.NameItem = nameItem;
                    //    }
                    //}
                    var a = saleorderRetList.GetAt(i).ORSalesOrderLineRetList.Append().SalesOrderLineRet.ItemRef;
                    if (saleorderRetList.GetAt(i).ORSalesOrderLineRetList.GetAt(0).SalesOrderLineRet.ItemRef != null)
                    {
                        var qblistId = saleorderRetList.GetAt(i).ORSalesOrderLineRetList.GetAt(0).SalesOrderLineRet.ItemRef.ListID.GetValue().ToString();
                        if (qblistId != null)
                        {
                            salesorder.ItemID = qblistId;
                        }
                    }

                    if (saleorderRetList.GetAt(i).TxnDate != null)
                    {
                        var date = saleorderRetList.GetAt(i).TxnDate.GetValue().ToString();
                        if (date != null)
                        {
                            salesorder.TxnDate = DateTime.Parse(date);
                        }
                    }

                    result.Add(salesorder);
                }
            }
            return result;
        }

        public bool AddSalesOrder(SaleOrder salesOrder)
        {
            ConnectToQB();
            IMsgSetResponse responseMsgSet = processRequestFromQB(buildSalesOrderAddRq(salesOrder));
            var result = parseSalesOrderCRUDRs(responseMsgSet);
            disconnectFromQB();
            return result;
        }

        private bool parseSalesOrderCRUDRs(IMsgSetResponse responseMsgSet)
        {
            //MessageBox.Show(responseMsgSet.ToXMLString());
            IResponse response = responseMsgSet.ResponseList.GetAt(0);
            int statusCode = response.StatusCode;
            if (statusCode == 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private IMsgSetRequest buildSalesOrderQueryRq(string[] includeRetElement, string listId, string name)
        {
            IMsgSetRequest requestMsgSet = sessionManager.getMsgSetRequest();
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
            ISalesOrderQuery SaleOrderQuery = requestMsgSet.AppendSalesOrderQueryRq();


            for (int x = 0; x < includeRetElement.Length; x++)
            {
                SaleOrderQuery.IncludeRetElementList.Add(includeRetElement[x]);
            }
            return requestMsgSet;
        }

        #endregion

        private IMsgSetRequest buildCustomerDeleteRq(string qblistid)
        {
            IMsgSetRequest requestMsgSet = sessionManager.getMsgSetRequest();
            requestMsgSet.Attributes.OnError = ENRqOnError.roeContinue;
            IListDel custQuery = requestMsgSet.AppendListDelRq();
            //custQuery.Name.SetValue(customer.Name);
            custQuery.ListDelType.SetValue(ENListDelType.ldtCustomer);
            custQuery.ListID.SetValue(qblistid);
            //custQuery.BillAddress.City.SetValue(customer.City);
            return requestMsgSet;
        }
    }
}
