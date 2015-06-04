using System;
using System.Diagnostics;
using System.Text;
using System.Windows.Forms;  // For using MessageBox.
using QBSDKEVENTLib; // In order to implement IQBEventCallback.
using System.Runtime.InteropServices;  // For use of the GuidAttribute, ProgIdAttribute and ClassInterfaceAttribute.
using System.Xml;
using System.Linq;
using System.Xml.Linq;
using QuickBooksInteropLibrary;
using Npgsql;
using System.Configuration; //XML Parsing


namespace SubscribeAndHandleQBEvent
{
    [
      Guid("80552C03-B271-4cb2-B09F-0271E1DE808C"),  // We indicate a specific CLSID for "SubscribeAndHandleQBEvent.EventHandlerObj" for convenience of searching the registry.
      ProgId("SubscribeAndHandleQBEvent.EventHandlerObj"),  // This ProgId is used by default. Not 100% necessary.
      ClassInterface(ClassInterfaceType.None)
    ]

    public class EventHandlerObj :
        ReferenceCountedObjectBase, // EventHandlerObj is derived from ReferenceCountedObjectBase so that we can track its creation and destruction.
        IQBEventCallback  // this must implement the IQBEventCallback interface.
    {
        QuickbooksInterop qbInterop = new QuickbooksInterop();
        string _OdooConnectionString = ConfigurationSettings.AppSettings["OdooConnectionString"].ToString();
        public EventHandlerObj()
        {
            // ReferenceCountedObjectBase constructor will be invoked.
            Console.WriteLine("EventHandlerObj constructor.");
        }

        ~EventHandlerObj()
        {
            // ReferenceCountedObjectBase destructor will be invoked.
            Console.WriteLine("EventHandlerObj destructor.");
        }

        //Call back function which would be invoked from the QB
        public void inform(string strMessage)
        {
            try
            {
                StringBuilder sb = new StringBuilder(strMessage);
                XmlDocument outputXMLDoc = new XmlDocument();
                outputXMLDoc.LoadXml(strMessage);
                XmlNodeList qbXMLMsgsRsNodeList = outputXMLDoc.GetElementsByTagName("QBXMLEvents");
                XmlNode childNode = qbXMLMsgsRsNodeList.Item(0).FirstChild;

                // handle the event based on type of event
                switch (childNode.Name)
                {
                    case "DataEvent":
                        //Handle Data Event Here
                        MessageBox.Show(sb.ToString(), "DATA EVENT - From QB");
                        DataEventHandling(sb.ToString());
                        break;

                    case "UIExtensionEvent":
                        //Handle UI Extension Event HERE
                        MessageBox.Show(sb.ToString(), "UI EXTENSION EVENT - From QB");
                        break;

                    default:
                        MessageBox.Show(sb.ToString(), "Response From QB");
                        break;
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("Unexpected error in processing the response from QB - " + ex.Message);
            }
        }

        public void DataEventHandling(string xml)
        {
            var doc = XDocument.Parse(xml);
            var listType = doc.Descendants("ListEventType").FirstOrDefault();
            if (listType != null)
            {
                var listName = listType.Value.ToString();
                var listIdNode = doc.Descendants("ListID").FirstOrDefault();
                var listId = string.Empty;
                if (listIdNode != null)
                {
                    listId = listIdNode.Value.ToString();
                }
                switch (listName.ToUpper())
                {
                    case "CUSTOMER":
                        var operationNode = doc.Descendants("ListEventOperation").FirstOrDefault();
                        if (operationNode != null)
                        {
                            var action = operationNode.Value.ToString();

                            switch (action.ToUpper())
                            {
                                case "ADD":
                                    AddNewCustomer(listId);
                                    break;
                                case "MODIFY":
                                    if (!string.IsNullOrEmpty(listId))
                                    {
                                        if (!SyncManager.IsInprogressSync(listId))
                                        {
                                            SyncManager.StartSync(listId);
                                            UpdateExistingCustomer(listId);
                                        }
                                        else
                                        {
                                            SyncManager.CompleteSync(listId);
                                        }
                                    }
                                    break;
                                case "DELETE":
                                    if (!string.IsNullOrEmpty(listId))
                                    {
                                        if (!SyncManager.IsInprogressSync(listId))
                                        {
                                            SyncManager.StartSync(listId);
                                            DeleteExistingCustomer(listId);
                                        }
                                        else
                                        {
                                            SyncManager.CompleteSync(listId);
                                        }
                                    }
                                    break;
                            }
                        }
                        break;
                    #region Inventory
                    case "ITEMINVENTORY":
                        var operationNodeInventory = doc.Descendants("ListEventOperation").FirstOrDefault();
                        if (operationNodeInventory != null)
                        {
                            var action = operationNodeInventory.Value.ToString();

                            switch (action.ToUpper())
                            {
                                case "ADD":
                                    AddNewItemInventory(listId);
                                    break;
                                case "MODIFY":
                                    if (!string.IsNullOrEmpty(listId))
                                    {
                                        if (!SyncManager.IsInprogressSync(listId))
                                        {
                                            SyncManager.StartSync(listId);
                                            UpdateExistingItemInventory(listId);
                                        }
                                        else
                                        {
                                            SyncManager.CompleteSync(listId);
                                        }
                                    }
                                    break;
                                case "DELETE":
                                    if (!string.IsNullOrEmpty(listId))
                                    {
                                        if (!SyncManager.IsInprogressSync(listId))
                                        {
                                            SyncManager.StartSync(listId);
                                            DeleteExistingItemInventory(listId);
                                        }
                                        else
                                        {
                                            SyncManager.CompleteSync(listId);
                                        }
                                    }
                                    break;
                            }
                        }
                        break;
                    #endregion
                }               
            }
        }

        #region Inventory
        private void DeleteExistingItemInventory(string listId)
        {
            using (NpgsqlConnection conn = new NpgsqlConnection(_OdooConnectionString))
            {
                conn.Open();
                NpgsqlCommand cmd = new NpgsqlCommand(@"delete from product_category where qblistid = :qblistid", conn);
                cmd.Parameters.Add(new NpgsqlParameter("qblistid", listId));
                cmd.ExecuteNonQuery();

                NpgsqlCommand cmdproduct = new NpgsqlCommand(@"delete from product_product where qblistid = :qblistid", conn);
                cmdproduct.Parameters.Add(new NpgsqlParameter("qblistid", listId));
                cmdproduct.ExecuteNonQuery();

                NpgsqlCommand cmdtemplate = new NpgsqlCommand(@"delete from product_template where qblistid = :qblistid", conn);
                cmdtemplate.Parameters.Add(new NpgsqlParameter("qblistid", listId));
                cmdtemplate.ExecuteNonQuery();
            }
        }

        public int GetParentIdCategory(string qblistid)
        {
            using (NpgsqlConnection conn = new NpgsqlConnection(_OdooConnectionString))
            {
                conn.Open();
                NpgsqlCommand cmd = new NpgsqlCommand(@"Select id from product_category where qblistid=:qblistid", conn);
                cmd.Parameters.Add(new NpgsqlParameter("qblistid", qblistid));
                return Convert.ToInt32(cmd.ExecuteScalar());
            }
        }

         private void AddNewItemInventory(string listId)
        {
            var inventory = qbInterop.LoadInventory(listId, null);
            int parentID = 0;
            if (!String.IsNullOrEmpty(inventory.ParentRef.ListID))
            {
                parentID = GetParentIdCategory(inventory.ParentRef.ListID);
            }
            if (inventory.SubLevel == 0)
            {
                parentID = 4;//Other Product Category
            }
            using (NpgsqlConnection conn = new NpgsqlConnection(_OdooConnectionString))
            {
                conn.Open();

                if (inventory.SalesPrice != 0 || inventory.PurchaseCost!=0)
                {
                    NpgsqlCommand cmd = new NpgsqlCommand(@"insert into product_template (list_price, color, write_uid, mes_type, uom_id, create_date, uos_coeff, create_uid, sale_ok, categ_id, uom_po_id, description_sale, write_date, active, rental, name, type, sale_delay, purchase_ok, website_sequence, website_published, website_size_x, website_size_y, qblistid, edit_sequence, is_qb_notification) values(:list_price, :color, :write_uid, :mes_type, :uom_id, :create_date, :uos_coeff, :create_uid, :sale_ok, :categ_id, :uom_po_id, :description_sale, :write_date, :active, :rental, :name, :type, :sale_delay, :purchase_ok, :website_sequence, :website_published, :website_size_x, :website_size_y, :qblistid, :edit_sequence, :is_qb_notification) returning id ", conn);
                    cmd.Parameters.Add(new NpgsqlParameter("list_price", inventory.SalesPrice));
                    cmd.Parameters.Add(new NpgsqlParameter("color", "0"));
                    cmd.Parameters.Add(new NpgsqlParameter("write_uid", 1));
                    cmd.Parameters.Add(new NpgsqlParameter("mes_type", "fixed"));
                    cmd.Parameters.Add(new NpgsqlParameter("uom_id", 1));
                    cmd.Parameters.Add(new NpgsqlParameter("create_date", DateTime.Now));
                    cmd.Parameters.Add(new NpgsqlParameter("uos_coeff", 1));
                    cmd.Parameters.Add(new NpgsqlParameter("create_uid", 1));
                    cmd.Parameters.Add(new NpgsqlParameter("sale_ok", false));
                    cmd.Parameters.Add(new NpgsqlParameter("categ_id", parentID));
                    cmd.Parameters.Add(new NpgsqlParameter("uom_po_id", 1));
                    cmd.Parameters.Add(new NpgsqlParameter("description_sale", inventory.SalesDesc));
                    cmd.Parameters.Add(new NpgsqlParameter("write_date", DateTime.Now));
                    cmd.Parameters.Add(new NpgsqlParameter("active", true));
                    cmd.Parameters.Add(new NpgsqlParameter("rental", false));
                    cmd.Parameters.Add(new NpgsqlParameter("name", inventory.FullName));
                    cmd.Parameters.Add(new NpgsqlParameter("type", "consu"));
                    cmd.Parameters.Add(new NpgsqlParameter("sale_delay", 7));
                    cmd.Parameters.Add(new NpgsqlParameter("purchase_ok", false));
                    cmd.Parameters.Add(new NpgsqlParameter("website_sequence", "0"));
                    cmd.Parameters.Add(new NpgsqlParameter("website_published", false));
                    cmd.Parameters.Add(new NpgsqlParameter("website_size_x", 1));
                    cmd.Parameters.Add(new NpgsqlParameter("website_size_y", 1));
                    cmd.Parameters.Add(new NpgsqlParameter("qblistid", inventory.QBListID));
                    cmd.Parameters.Add(new NpgsqlParameter("edit_sequence", inventory.EditSequence));
                    cmd.Parameters.Add(new NpgsqlParameter("is_qb_notification", "true"));

                    var idTemplate = cmd.ExecuteScalar().ToString();
                  
                    NpgsqlCommand cmdpd = new NpgsqlCommand(@"insert into product_product (create_date, name_template, create_uid, product_tmpl_id, write_uid, write_date, active, qblistid, edit_sequence) values(:create_date, :name_template, :create_uid, :product_tmpl_id, :write_uid, :write_date, :active, :qblistid, :edit_sequence)", conn);
                    cmdpd.Parameters.Add(new NpgsqlParameter("create_date", DateTime.Now));
                    cmdpd.Parameters.Add(new NpgsqlParameter("name_template", inventory.FullName));
                    cmdpd.Parameters.Add(new NpgsqlParameter("create_uid", 1));
                    cmdpd.Parameters.Add(new NpgsqlParameter("product_tmpl_id", idTemplate));
                    cmdpd.Parameters.Add(new NpgsqlParameter("write_uid", 1));
                    cmdpd.Parameters.Add(new NpgsqlParameter("write_date", DateTime.Now));
                    cmdpd.Parameters.Add(new NpgsqlParameter("active", true));
                    cmdpd.Parameters.Add(new NpgsqlParameter("qblistid", inventory.QBListID));
                    cmdpd.Parameters.Add(new NpgsqlParameter("edit_sequence", inventory.EditSequence));
                    cmdpd.ExecuteNonQuery();

                    DeleteRowQBListIDEmptyHandle("product_product");
                  
                    DeleteRowQBListIDEmptyHandle("product_template");
                }
                else
                {
                    NpgsqlCommand cmd = new NpgsqlCommand(@"insert into product_category (create_uid, create_date, name, write_uid, parent_id, write_date, type, qblistid, edit_sequence)                                                    values(:create_uid, :create_date, :name,:write_uid, :parent_id, :write_date, :type, :qblistid, :edit_sequence)", conn);
                    cmd.Parameters.Add(new NpgsqlParameter("create_uid", 1));
                    cmd.Parameters.Add(new NpgsqlParameter("create_date", DateTime.Now));
                    cmd.Parameters.Add(new NpgsqlParameter("name", inventory.FullName));
                    cmd.Parameters.Add(new NpgsqlParameter("write_uid", 1));
                    cmd.Parameters.Add(new NpgsqlParameter("parent_id", parentID));
                    cmd.Parameters.Add(new NpgsqlParameter("write_date", DateTime.Now));
                    cmd.Parameters.Add(new NpgsqlParameter("type", "normal"));
                    cmd.Parameters.Add(new NpgsqlParameter("qblistid", inventory.QBListID));
                    cmd.Parameters.Add(new NpgsqlParameter("edit_sequence", inventory.EditSequence));
                    cmd.ExecuteNonQuery();
                 
                    DeleteRowQBListIDEmptyHandle("product_category");
                }
            }
        }

         private void DeleteRowQBListIDEmptyHandle(string table)
         {
             using (NpgsqlConnection conn = new NpgsqlConnection(_OdooConnectionString))
             {
                 conn.Open();
                 NpgsqlCommand delete = new NpgsqlCommand(@"delete from " + table + " where coalesce(qblistid,'') = ''", conn);
                 delete.ExecuteNonQuery();
             }
         }

         private void UpdateExistingItemInventory(string listId)
        {

            var updatedItemInventory = qbInterop.LoadInventory(listId, null);
            using (NpgsqlConnection conn = new NpgsqlConnection(_OdooConnectionString))
            {
                conn.Open();
                NpgsqlCommand cmdCheck = new NpgsqlCommand(@"Select * from product_category where qblistid=:qblistid", conn);
                cmdCheck.Parameters.Add(new NpgsqlParameter("qblistid", updatedItemInventory.QBListID));
                NpgsqlDataReader dr = cmdCheck.ExecuteReader();
                conn.Close();
                conn.Open();
                if (dr.HasRows==true)
                {
                    NpgsqlCommand cmd = new NpgsqlCommand(
"update product_category set \"name\" = :name, \"edit_sequence\" = :edit_sequence where qblistid = :qblistid", conn);
                    cmd.Parameters.Add(new NpgsqlParameter("name", updatedItemInventory.FullName));
                    cmd.Parameters.Add(new NpgsqlParameter("edit_sequence", updatedItemInventory.EditSequence));
                    cmd.Parameters.Add(new NpgsqlParameter("qblistid", updatedItemInventory.QBListID));
                    cmd.ExecuteNonQuery();
                }
                else
                {
                    NpgsqlCommand cmdProd = new NpgsqlCommand(
"update product_product set \"name_template\" = :name_template, \"edit_sequence\" = :edit_sequence where qblistid = :qblistid", conn);
                    cmdProd.Parameters.Add(new NpgsqlParameter("name_template", updatedItemInventory.FullName));
                    cmdProd.Parameters.Add(new NpgsqlParameter("edit_sequence", updatedItemInventory.EditSequence));
                    cmdProd.Parameters.Add(new NpgsqlParameter("qblistid", updatedItemInventory.QBListID));
                    cmdProd.ExecuteNonQuery();
                    NpgsqlCommand cmd = new NpgsqlCommand(
    "update product_template set \"name\" = :name, \"list_price\" = :list_price, \"description_sale\" = :description_sale, \"edit_sequence\" = :edit_sequence, \"is_qb_notification\" = :is_qb_notification where qblistid = :qblistid", conn);
                    cmd.Parameters.Add(new NpgsqlParameter("name", updatedItemInventory.FullName));
                    cmd.Parameters.Add(new NpgsqlParameter("list_price", updatedItemInventory.SalesPrice));
                    cmd.Parameters.Add(new NpgsqlParameter("description_sale", updatedItemInventory.SalesDesc));
                    cmd.Parameters.Add(new NpgsqlParameter("edit_sequence", updatedItemInventory.EditSequence));
                    cmd.Parameters.Add(new NpgsqlParameter("is_qb_notification", "true"));
                    cmd.Parameters.Add(new NpgsqlParameter("qblistid", updatedItemInventory.QBListID));
                    cmd.ExecuteNonQuery();
                }
            }
        }
        #endregion

        private void DeleteExistingCustomer(string listId)
        {
            //var updatedCustomer = qbInterop.LoadCustomer(listId);
            using (NpgsqlConnection conn = new NpgsqlConnection(_OdooConnectionString))
            {
                conn.Open();
                NpgsqlCommand cmd = new NpgsqlCommand(@"delete from res_partner where qblistid = :qblistid", conn);
                cmd.Parameters.Add(new NpgsqlParameter("qblistid", listId));
                cmd.ExecuteNonQuery();
            }
        }

        private void AddNewCustomer(string listId)
        {
            var updatedCustomer = qbInterop.LoadCustomer(listId, null);
            using (NpgsqlConnection conn = new NpgsqlConnection(_OdooConnectionString))
            {
                conn.Open();

                NpgsqlCommand cmd = new NpgsqlCommand(@"insert into res_partner (name, display_name, street, city, zip, website, email, supplier, is_company, customer, notify_email, employee, phone, active,qblistid, edit_sequence, is_qb_notification) 
                                        values(:name, :display_name, :street, :city, :zip, :website, :email, :supplier, :is_company, :customer, :notify_email, :employee, :phone, :active, :qblistid, :edit_sequence, :is_qb_notification)", conn);
                cmd.Parameters.Add(new NpgsqlParameter("name", updatedCustomer.DisplayName));
                cmd.Parameters.Add(new NpgsqlParameter("display_name", updatedCustomer.DisplayName));
                cmd.Parameters.Add(new NpgsqlParameter("street", updatedCustomer.Street));
                cmd.Parameters.Add(new NpgsqlParameter("city", updatedCustomer.City));
                cmd.Parameters.Add(new NpgsqlParameter("zip", updatedCustomer.Zip));
                cmd.Parameters.Add(new NpgsqlParameter("website", updatedCustomer.Website));
                cmd.Parameters.Add(new NpgsqlParameter("email", updatedCustomer.Email));
                cmd.Parameters.Add(new NpgsqlParameter("supplier", updatedCustomer.IsSupplier));
                cmd.Parameters.Add(new NpgsqlParameter("is_company", updatedCustomer.IsCompany));
                cmd.Parameters.Add(new NpgsqlParameter("customer", updatedCustomer.IsCustomer));
                cmd.Parameters.Add(new NpgsqlParameter("notify_email", updatedCustomer.NotifyEmail));
                cmd.Parameters.Add(new NpgsqlParameter("employee", updatedCustomer.IsEmployee));
                cmd.Parameters.Add(new NpgsqlParameter("phone", updatedCustomer.Phone));
                cmd.Parameters.Add(new NpgsqlParameter("active", updatedCustomer.IsActive));
                cmd.Parameters.Add(new NpgsqlParameter("qblistid", updatedCustomer.QBListId));
                cmd.Parameters.Add(new NpgsqlParameter("edit_sequence", updatedCustomer.EditSequence));
                cmd.Parameters.Add(new NpgsqlParameter("is_qb_notification", "true"));
                cmd.ExecuteNonQuery();

                NpgsqlCommand delete = new NpgsqlCommand(@"delete from res_partner where coalesce(qblistid,'') = ''", conn);
                delete.ExecuteNonQuery();
            }
        }

        private void UpdateExistingCustomer(string listId)
        {

            var updatedCustomer = qbInterop.LoadCustomer(listId, null);
            using (NpgsqlConnection conn = new NpgsqlConnection(_OdooConnectionString))
            {
                conn.Open();
                NpgsqlCommand cmd = new NpgsqlCommand(
"update res_partner set \"name\" = :name, \"display_name\" = :display_name, \"street\" = :street, \"city\" = :city, \"zip\" = :zip,"
                + "\"website\" = :website, \"email\" = :email, \"supplier\" = :supplier, \"is_company\" = :is_company, \"customer\" = :customer,"
                + "\"notify_email\" = :notify_email, \"employee\" = :employee, \"phone\" = :phone, \"active\" = :active, \"edit_sequence\" = :edit_sequence, \"is_qb_notification\" = :is_qb_notification where qblistid = :qblistid", conn);
                cmd.Parameters.Add(new NpgsqlParameter("name", updatedCustomer.DisplayName));
                cmd.Parameters.Add(new NpgsqlParameter("display_name", updatedCustomer.DisplayName));
                cmd.Parameters.Add(new NpgsqlParameter("street", updatedCustomer.Street));
                cmd.Parameters.Add(new NpgsqlParameter("city", updatedCustomer.City));
                cmd.Parameters.Add(new NpgsqlParameter("zip", updatedCustomer.Zip));
                cmd.Parameters.Add(new NpgsqlParameter("website", updatedCustomer.Website));
                cmd.Parameters.Add(new NpgsqlParameter("email", updatedCustomer.Email));
                cmd.Parameters.Add(new NpgsqlParameter("supplier", updatedCustomer.IsSupplier));
                cmd.Parameters.Add(new NpgsqlParameter("is_company", updatedCustomer.IsCompany));
                cmd.Parameters.Add(new NpgsqlParameter("customer", updatedCustomer.IsCustomer));
                cmd.Parameters.Add(new NpgsqlParameter("notify_email", updatedCustomer.NotifyEmail));
                cmd.Parameters.Add(new NpgsqlParameter("employee", updatedCustomer.IsEmployee));
                cmd.Parameters.Add(new NpgsqlParameter("phone", updatedCustomer.Phone));
                cmd.Parameters.Add(new NpgsqlParameter("active", updatedCustomer.IsActive));
                cmd.Parameters.Add(new NpgsqlParameter("qblistid", updatedCustomer.QBListId));
                cmd.Parameters.Add(new NpgsqlParameter("edit_sequence", updatedCustomer.EditSequence));
                cmd.Parameters.Add(new NpgsqlParameter("is_qb_notification", "true"));
                cmd.ExecuteNonQuery();
            }
        }
    }

    class EventHandlerObjClassFactory : ClassFactoryBase
    {
        public override void virtual_CreateInstance(IntPtr pUnkOuter, ref Guid riid, out IntPtr ppvObject)
        {
            Console.WriteLine("EventHandlerObjClassFactory.CreateInstance().");
            Console.WriteLine("Requesting Interface : " + riid.ToString());

            if (riid == Marshal.GenerateGuidForType(typeof(IQBEventCallback)) ||
                riid == SubscribeAndHandleQBEvent.IID_IDispatch ||
                riid == SubscribeAndHandleQBEvent.IID_IUnknown)
            {
                EventHandlerObj EventHandlerObj_New = new EventHandlerObj();

                ppvObject = Marshal.GetComInterfaceForObject(EventHandlerObj_New, typeof(IQBEventCallback));
            }
            else
            {
                throw new COMException("No interface", unchecked((int)0x80004002));
            }
        }
    }
}