using Npgsql;
using QuickBooksInteropLibrary;
using QuickBooksInteropLibrary.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;

namespace SubscribeAndHandleQBEvent.PostgreSQL
{
    public class PostgreListener
    {
        QuickbooksInterop qbInterop = new QuickbooksInterop();
        string _OdooListenerConnectionString = ConfigurationSettings.AppSettings["OdooListenerConnectionString"].ToString();
        string _OdooConnectionString = ConfigurationSettings.AppSettings["OdooConnectionString"].ToString();
        public void ListenCustomChange()
        {
            var conn = new NpgsqlConnection(_OdooListenerConnectionString);

            conn.Open();

            NpgsqlCommand command = new NpgsqlCommand(ConfigurationSettings.AppSettings["CustomerOdooCommandListener"].ToString(), conn);
            command.ExecuteNonQuery();

            conn.Notification += new NotificationEventHandler(NotificationSupportHelper);
        }

        #region Inventory
        public void ListenInventoryCategoryChange()
        {
            var conn = new NpgsqlConnection(_OdooListenerConnectionString);

            conn.Open();

            NpgsqlCommand command = new NpgsqlCommand(ConfigurationSettings.AppSettings["InventoryCategoryOdooCommandListener"].ToString(), conn);
            command.ExecuteNonQuery();

            conn.Notification += new NotificationEventHandler(NotificationSupportHelperForInventoryCategory);
        }

        public void ListenItemInventoryChange()
        {
            var conn = new NpgsqlConnection(_OdooListenerConnectionString);

            conn.Open();

            NpgsqlCommand command = new NpgsqlCommand(ConfigurationSettings.AppSettings["ItemInventoryOdooCommandListener"].ToString(), conn);
            command.ExecuteNonQuery();

            conn.Notification += new NotificationEventHandler(NotificationSupportHelperForItemInventory);
        }
        #endregion


        #region SaleOrder

        public void ListenSalesOrderChange()
        {
            var conn = new NpgsqlConnection(_OdooListenerConnectionString);

            conn.Open();

            NpgsqlCommand command = new NpgsqlCommand(ConfigurationSettings.AppSettings["SaleOrderOdooCommandListener"].ToString(), conn);
            command.ExecuteNonQuery();

            conn.Notification += new NotificationEventHandler(NotificationSupportHelperSaleOrder);
        }

        private void NotificationSupportHelperSaleOrder(object sender, NpgsqlNotificationEventArgs e)
        {
            var message = e.AdditionalInformation;
            var info = message.Split(';');
            var qblistid = info[0];
            var action = info[1];
            var odooId = info[2];
            SaleOrder salesOrder = new SaleOrder();

            switch (action.ToUpper())
            {
                case "UPDATE":
                    using (NpgsqlConnection conn = new NpgsqlConnection(_OdooConnectionString))
                    {
                        conn.Open();
                        NpgsqlCommand cmd = new NpgsqlCommand(@"SELECT id FROM sale_order WHERE id = :odooId", conn);
                        cmd.Parameters.Add("odooId", odooId);
                        NpgsqlDataReader dr = cmd.ExecuteReader();
                        string salesOrderId = string.Empty;
                        string customerID = string.Empty;
                        string orderID = string.Empty;
                        while (dr.Read())
                        {
                            salesOrderId = dr[0].ToString();
                        }
                        NpgsqlCommand cmd2 = new NpgsqlCommand(@"SELECT product_uom_qty, order_partner_id, product_id, order_id FROM sale_order_line WHERE order_id = :salesOrderId", conn);
                        cmd2.Parameters.Add("salesOrderId", salesOrderId);

                        NpgsqlDataReader dr2 = cmd2.ExecuteReader();

                        List<SalesOrderItem> listsale = new List<SalesOrderItem>();
                        List<SalesOrderItem> listitem = new List<SalesOrderItem>();

                        while (dr2.Read())
                        {
                            customerID = dr2[1].ToString();
                            orderID = dr2[3].ToString();
                            var productId = dr2[2].ToString();
                            int quantity = Convert.ToInt32(dr2[0]);
                            listitem.Add(new SalesOrderItem() { ItemID = dr2[2].ToString(), Quantity = Convert.ToInt32(dr2[0]) });
                        }
                        NpgsqlCommand cmd3 = new NpgsqlCommand(@"SELECT qblistid FROM res_partner WHERE id = :customerID", conn);
                        cmd3.Parameters.Add("customerID", customerID);
                        NpgsqlDataReader dr3 = cmd3.ExecuteReader();

                        var customerListid = string.Empty;
                        while (dr3.Read())
                        {
                            customerListid = dr3[0].ToString();
                        }


                        foreach (var id in listitem)
                        {
                            NpgsqlCommand cmd4 = new NpgsqlCommand(@"SELECT qblistid FROM product_product WHERE id = :productID", conn);
                            cmd4.Parameters.Add("productID", id.ItemID);
                            NpgsqlDataReader dr4 = cmd4.ExecuteReader();

                            var productListid = string.Empty;

                            while (dr4.Read())
                            {
                                var Quantity = 1;
                                productListid = dr4[0].ToString();
                                Quantity = id.Quantity;

                                listsale.Add(new SalesOrderItem() { ItemID = productListid, Quantity = Quantity });
                            }
                        }

                        salesOrder.CustomerID = customerListid;
                        salesOrder.ListItem = listsale;
                        qbInterop.AddSalesOrder(salesOrder);
                    }
                    break;
            }
        }

        #endregion


        public void NotificationSupportHelper(object sender, NpgsqlNotificationEventArgs e)
        {
            var message = e.AdditionalInformation;
            var info = message.Split(';');
            var qblistid = info[0];
            var action = info[1];
            var odooId = info[2];
            Customer customer = new Customer();

            switch (action.ToUpper())
            {
                case "INSERT":
                    using (NpgsqlConnection conn = new NpgsqlConnection(_OdooConnectionString))
                    {
                        conn.Open();
                        NpgsqlCommand cmd = new NpgsqlCommand(@"SELECT name, display_name, street, city, zip, website, email, phone, qblistid, edit_sequence FROM res_partner where id = :odooId", conn);
                        cmd.Parameters.Add(new NpgsqlParameter("odooId", odooId));
                        NpgsqlDataReader dr = cmd.ExecuteReader();

                        // Output rows
                        while (dr.Read())
                        {
                            customer.Name = dr[0].ToString();
                            customer.DisplayName = dr[1].ToString();
                            customer.Street = dr[2].ToString();
                            customer.City = dr[3].ToString();
                            customer.Zip = dr[4].ToString();
                            customer.Website = dr[5].ToString();
                            customer.Email = dr[6].ToString();
                            customer.Phone = dr[7].ToString();
                            customer.QBListId = dr[8].ToString();
                            customer.EditSequence = dr[9].ToString();
                            customer.OdooId = int.Parse(odooId);
                        }

                        qbInterop.AddCustomer(customer);
                    }
                    break;
                case "UPDATE":
                    if (!string.IsNullOrEmpty(qblistid))
                    {
                        if (!SyncManager.IsInprogressSync(qblistid))
                        {
                            SyncManager.StartSync(qblistid);
                            using (NpgsqlConnection conn = new NpgsqlConnection(_OdooConnectionString))
                            {
                                conn.Open();
                                NpgsqlCommand cmd = new NpgsqlCommand(@"SELECT name, display_name, street, city, zip, website, email, phone, qblistid, edit_sequence FROM res_partner where qblistid = :qblistid", conn);
                                cmd.Parameters.Add(new NpgsqlParameter("qblistid", qblistid));
                                NpgsqlDataReader dr = cmd.ExecuteReader();

                                // Output rows
                                while (dr.Read())
                                {
                                    customer.Name = dr[0].ToString();
                                    customer.DisplayName = dr[1].ToString();
                                    customer.Street = dr[2].ToString();
                                    customer.City = dr[3].ToString();
                                    customer.Zip = dr[4].ToString();
                                    customer.Website = dr[5].ToString();
                                    customer.Email = dr[6].ToString();
                                    customer.Phone = dr[7].ToString();
                                    customer.QBListId = dr[8].ToString();
                                    customer.EditSequence = dr[9].ToString();
                                }
                            }
                            qbInterop.UpdateCustomer(customer);
                        }
                        else
                        {
                            SyncManager.CompleteSync(qblistid);
                        }
                    }
                    break;
                case "DELETE":
                    if (!string.IsNullOrEmpty(qblistid))
                    {
                        if (!SyncManager.IsInprogressSync(qblistid))
                        {
                            SyncManager.StartSync(qblistid);
                            qbInterop.DeleteCustomer(qblistid);
                        }
                        else
                        {
                            SyncManager.CompleteSync(qblistid);
                        }
                    }
                    break;
            }
        }

        #region Inventory
        public void NotificationSupportHelperForInventoryCategory(object sender, NpgsqlNotificationEventArgs e)
        {
            var message = e.AdditionalInformation;
            var info = message.Split(';');
            var qblistid = info[0];
            var action = info[1];
            var odooId = info[2];

            Inventory inventory = new Inventory();

            switch (action.ToUpper())
            {
                case "INSERT":
                    using (NpgsqlConnection conn = new NpgsqlConnection(_OdooConnectionString))
                    {
                        conn.Open();
                        NpgsqlCommand cmd = new NpgsqlCommand(@"select b.name, b.qblistid, b.edit_sequence, a.name as name_parent, a.qblistid as qblistid_parent, a.edit_sequence as edit_sequence_parent from product_category a left join product_category b on b.parent_id=a.id where b.id = :odooId", conn);
                        cmd.Parameters.Add(new NpgsqlParameter("odooId", odooId));
                        NpgsqlDataReader dr = cmd.ExecuteReader();
                        if (dr.HasRows == true && dr[4].ToString()!="")
                        {
                            // Output rows
                            while (dr.Read())
                            {
                                inventory.Name = dr[0].ToString();
                                inventory.QBListID = dr[1].ToString();
                                inventory.EditSequence = dr[2].ToString();
                                inventory.ParentRef.FullName = dr[3].ToString();
                                inventory.ParentRef.ListID = dr[4].ToString();
                                inventory.OdooId = int.Parse(odooId);
                            }
                        }
                        else
                        {
                            NpgsqlCommand cmdCateg = new NpgsqlCommand(@"select name, qblistid, edit_sequence from product_category where id = :odooId", conn);
                            cmdCateg.Parameters.Add(new NpgsqlParameter("odooId", odooId));
                            NpgsqlDataReader drCateg = cmdCateg.ExecuteReader();
                            while (drCateg.Read())
                            {
                                inventory.Name = drCateg[0].ToString();
                                inventory.QBListID = drCateg[1].ToString();
                                inventory.EditSequence = drCateg[2].ToString();
                                inventory.OdooId = int.Parse(odooId);
                            }
                        }
                        qbInterop.AddInventory(inventory);
                    }
                    break;
                case "UPDATE":
                    if (!string.IsNullOrEmpty(qblistid))
                    {
                        if (!SyncManager.IsInprogressSync(qblistid))
                        {
                            SyncManager.StartSync(qblistid);
                            using (NpgsqlConnection conn = new NpgsqlConnection(_OdooConnectionString))
                            {
                                conn.Open();
                                NpgsqlCommand cmd = new NpgsqlCommand(@"SELECT name, qblistid, edit_sequence FROM product_category where qblistid = :qblistid", conn);
                                cmd.Parameters.Add(new NpgsqlParameter("qblistid", qblistid));
                                NpgsqlDataReader dr = cmd.ExecuteReader();

                                // Output rows
                                while (dr.Read())
                                {
                                    if (dr[0].ToString().Contains(":"))
                                    {
                                        string[] temp = dr[0].ToString().Split(':');
                                        inventory.Name = temp[temp.Length - 1];
                                    }
                                    else
                                    {
                                        inventory.Name = dr[0].ToString();
                                    }
                                    inventory.QBListID = dr[1].ToString();
                                    inventory.EditSequence = dr[2].ToString();
                                }
                            }
                            qbInterop.UpdateInventory(inventory);
                        }
                        else
                        {
                            SyncManager.CompleteSync(qblistid);
                        }
                    }
                    break;
                case "DELETE":
                    if (!string.IsNullOrEmpty(qblistid))
                    {
                        if (!SyncManager.IsInprogressSync(qblistid))
                        {
                            SyncManager.StartSync(qblistid);
                            qbInterop.DeleteInventory(qblistid);
                        }
                        else
                        {
                            SyncManager.CompleteSync(qblistid);
                        }
                    }
                    break;
            }
        }

        public void NotificationSupportHelperForItemInventory(object sender, NpgsqlNotificationEventArgs e)
        {
            var message = e.AdditionalInformation;
            var info = message.Split(';');
            var qblistid = info[0];
            var action = info[1];
            var odooId = info[2];

            Inventory inventory = new Inventory();

            switch (action.ToUpper())
            {
                case "INSERT":
                    using (NpgsqlConnection conn = new NpgsqlConnection(_OdooConnectionString))
                    {
                        conn.Open();
                        NpgsqlCommand cmd = new NpgsqlCommand(@"SELECT t.name,t.list_price, t.description_sale, t.active, t.qblistid, t.edit_sequence, c.name as name_parent, c.qblistid as qblistid_parent FROM product_template t Left join product_category c on c.id = t.categ_id where t.id = :odooId", conn);
                        cmd.Parameters.Add(new NpgsqlParameter("odooId", odooId));
                        NpgsqlDataReader dr = cmd.ExecuteReader();

                        // Output rows
                        while (dr.Read())
                        {
                            inventory.Name = dr[0].ToString();
                            inventory.SalesPrice = Convert.ToDouble(dr[1]);
                            inventory.SalesDesc = dr[2].ToString();
                            inventory.isActive = Convert.ToBoolean(dr[3]);
                            inventory.QBListID = dr[4].ToString();
                            inventory.EditSequence = dr[5].ToString();
                            inventory.ParentRef.FullName = dr[6].ToString();
                            inventory.ParentRef.ListID = dr[7].ToString();
                            inventory.OdooId = int.Parse(odooId);
                        }
                        qbInterop.AddInventory(inventory);
                    }
                    break;
                case "UPDATE":
                    if (!string.IsNullOrEmpty(qblistid))
                    {
                        if (!SyncManager.IsInprogressSync(qblistid))
                        {
                            SyncManager.StartSync(qblistid);
                            using (NpgsqlConnection conn = new NpgsqlConnection(_OdooConnectionString))
                            {
                                conn.Open();
                                NpgsqlCommand cmd = new NpgsqlCommand(@"SELECT name,list_price, description_sale, active, qblistid, edit_sequence FROM product_template where qblistid = :qblistid", conn);
                                cmd.Parameters.Add(new NpgsqlParameter("qblistid", qblistid));
                                NpgsqlDataReader dr = cmd.ExecuteReader();

                                // Output rows
                                while (dr.Read())
                                {
                                    if (dr[0].ToString().Contains(":"))
                                    {
                                        string[] temp = dr[0].ToString().Split(':');
                                        inventory.Name = temp[temp.Length - 1];
                                    }
                                    else
                                    {
                                        inventory.Name = dr[0].ToString();
                                    }
                                    inventory.SalesPrice = Convert.ToDouble(dr[1]);
                                    inventory.SalesDesc = dr[2].ToString();
                                    inventory.isActive = Convert.ToBoolean(dr[3]);
                                    inventory.QBListID = dr[4].ToString();
                                    inventory.EditSequence = dr[5].ToString();
                                }
                            }
                            qbInterop.UpdateInventory(inventory);
                        }
                        else
                        {
                            SyncManager.CompleteSync(qblistid);
                        }
                    }
                    break;
                case "DELETE":
                    if (!string.IsNullOrEmpty(qblistid))
                    {
                        if (!SyncManager.IsInprogressSync(qblistid))
                        {
                            SyncManager.StartSync(qblistid);
                            qbInterop.DeleteInventory(qblistid);
                        }
                        else
                        {
                            SyncManager.CompleteSync(qblistid);
                        }
                    }
                    break;
            }
        }
        #endregion
    }
}
