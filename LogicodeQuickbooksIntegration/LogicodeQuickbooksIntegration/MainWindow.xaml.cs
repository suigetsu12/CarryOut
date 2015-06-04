using Interop.QBFC13;
using Npgsql;
using NpgsqlTypes;
using QuickBooksInteropLibrary;
using QuickBooksInteropLibrary.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml.Linq;

using System.Xml;
using Interop.QBXMLRP2;

namespace LogicodeQuickbooksIntegration
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        QuickbooksInterop qbInterop = new QuickbooksInterop();
        string _lagasseFolderPath = string.Empty;
        string _odooConnectionString = string.Empty;
        string _lagasseMdbFile = string.Empty;
        Dictionary<string, string> _idMapping = new Dictionary<string, string>();

        public MainWindow()
        {
            InitializeComponent();
            _lagasseFolderPath = ConfigurationSettings.AppSettings["LagasseFolderPath"];
            _odooConnectionString = ConfigurationSettings.AppSettings["OdooConnectionString"];
            _lagasseMdbFile = ConfigurationSettings.AppSettings["LagasseMdbFile"];
        }

        private void bulkImport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (ckbCustomers.IsChecked == true)
                {
                    var customerList = qbInterop.LoadCustomers();
                    ImportCustomersToOdoo(customerList);
                }
                if (ckbInventories.IsChecked == true)
                {
                    var itemInventoryList = qbInterop.LoadInventories();
                    ImportItemInventoriesToOdoo(itemInventoryList);
                    //LoadItemInventoryXML();
                }
                if (ckbInvoices.IsChecked == true)
                {
                    LoadInvoiceXML();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void LoadInvoiceXML()
        {
            XmlDocument inputXMLDoc = new XmlDocument();
            inputXMLDoc.AppendChild(inputXMLDoc.CreateXmlDeclaration("1.0", null, null));
            inputXMLDoc.AppendChild(inputXMLDoc.CreateProcessingInstruction("qbxml", "version=\"8.0\""));
            XmlElement qbXML = inputXMLDoc.CreateElement("QBXML");
            inputXMLDoc.AppendChild(qbXML);
            XmlElement qbXMLMsgsRq = inputXMLDoc.CreateElement("QBXMLMsgsRq");
            qbXML.AppendChild(qbXMLMsgsRq);
            qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
            XmlElement custAddRq = inputXMLDoc.CreateElement("SalesOrderQueryRq");
            qbXMLMsgsRq.AppendChild(custAddRq);
            custAddRq.SetAttribute("requestID", "1371");

            string input = inputXMLDoc.OuterXml;
            //step3: do the qbXMLRP request
            RequestProcessor2 rp = null;
            string ticket = null;
            string response = null;
            try
            {
                rp = new RequestProcessor2();
                rp.OpenConnection("", "IDN SalesOrderQueryRq C# sample");
                ticket = rp.BeginSession("", QBFileMode.qbFileOpenDoNotCare);
                response = rp.ProcessRequest(ticket, input);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("COM Error Description = " + ex.Message, "COM error");
                return;
            }
            finally
            {
                if (ticket != null)
                {
                    rp.EndSession(ticket);
                }
                if (rp != null)
                {
                    rp.CloseConnection();
                }
            };
            XmlDocument outputXMLDoc = new XmlDocument();
            outputXMLDoc.LoadXml(response);
            outputXMLDoc.Save(@"D:\SaleOrder.xml");
        }

        private void LoadItemInventoryXML()
        {
            XmlDocument inputXMLDoc = new XmlDocument();
            inputXMLDoc.AppendChild(inputXMLDoc.CreateXmlDeclaration("1.0", null, null));
            inputXMLDoc.AppendChild(inputXMLDoc.CreateProcessingInstruction("qbxml", "version=\"8.0\""));
            XmlElement qbXML = inputXMLDoc.CreateElement("QBXML");
            inputXMLDoc.AppendChild(qbXML);
            XmlElement qbXMLMsgsRq = inputXMLDoc.CreateElement("QBXMLMsgsRq");
            qbXML.AppendChild(qbXMLMsgsRq);
            qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
            XmlElement custAddRq = inputXMLDoc.CreateElement("ItemInventoryQueryRq");
            qbXMLMsgsRq.AppendChild(custAddRq);
            custAddRq.SetAttribute("requestID", "1234");

            string input = inputXMLDoc.OuterXml;
            //step3: do the qbXMLRP request
            RequestProcessor2 rp = null;
            string ticket = null;
            string response = null;
            try
            {
                rp = new RequestProcessor2();
                rp.OpenConnection("", "IDN ItemInventoryQueryRq C# sample");
                ticket = rp.BeginSession("", QBFileMode.qbFileOpenDoNotCare);
                response = rp.ProcessRequest(ticket, input);

            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show("COM Error Description = " + ex.Message, "COM error");
                return;
            }
            finally
            {
                if (ticket != null)
                {
                    rp.EndSession(ticket);
                }
                if (rp != null)
                {
                    rp.CloseConnection();
                }
            };

            //step4: parse the XML response and show a message
            XmlDocument outputXMLDoc = new XmlDocument();
            outputXMLDoc.LoadXml(response);
            XmlNodeList qbXMLMsgsRsNodeList = outputXMLDoc.GetElementsByTagName("CustomerQueryRq");

            if (qbXMLMsgsRsNodeList.Count == 1) //it's always true, since we added a single Customer
            {
                System.Text.StringBuilder popupMessage = new System.Text.StringBuilder();

                XmlAttributeCollection rsAttributes = qbXMLMsgsRsNodeList.Item(0).Attributes;
                //get the status Code, info and Severity
                string retStatusCode = rsAttributes.GetNamedItem("statusCode").Value;
                string retStatusSeverity = rsAttributes.GetNamedItem("statusSeverity").Value;
                string retStatusMessage = rsAttributes.GetNamedItem("statusMessage").Value;
                popupMessage.AppendFormat("statusCode = {0}, statusSeverity = {1}, statusMessage = {2}",
                    retStatusCode, retStatusSeverity, retStatusMessage);

                //get the CustomerRet node for detailed info

                //a CustomerAddRs contains max one childNode for "CustomerRet"
                XmlNodeList custAddRsNodeList = qbXMLMsgsRsNodeList.Item(0).ChildNodes;
                if (custAddRsNodeList.Count == 1 && custAddRsNodeList.Item(0).Name.Equals("CustomerRet"))
                {
                    XmlNodeList custRetNodeList = custAddRsNodeList.Item(0).ChildNodes;

                } // End of customerRet

                MessageBox.Show(popupMessage.ToString(), "QuickBooks response");
            }
        }

        /*Inventory Start*/

        private void ImportItemInventoriesToOdoo(List<Inventory> itemInventoryList)
        {

            List<string> excList = ExceptionErrorList(itemInventoryList);

            int parentID = 1;
            List<Inventory> itemInventorySub0List = itemInventoryList.Where(i => i.SubLevel == 0).ToList();
            ImportInventory(itemInventorySub0List, parentID, excList);

            List<Inventory> itemInventorySub1List = itemInventoryList.Where(i => i.SubLevel == 1).ToList();
            ImportInventory(itemInventorySub1List, parentID, excList);

            List<Inventory> itemInventorySub2List = itemInventoryList.Where(i => i.SubLevel == 2).ToList();
            ImportInventory(itemInventorySub2List, parentID, excList);

            List<Inventory> itemInventorySub3List = itemInventoryList.Where(i => i.SubLevel == 3).ToList();
            ImportInventory(itemInventorySub3List, parentID, excList);
        }

        /*get ListID excetion error when category have cost>0*/
        private List<string> ExceptionErrorList(List<Inventory> itemInventoryList)
        {
            List<string> tempList = new List<string>();
            List<Inventory> lstSub0 = itemInventoryList.Where(i => i.SubLevel == 0 && (i.SalesPrice != 0 || i.PurchaseCost != 0)).ToList();
            foreach (var item in lstSub0)
            {
                int temp = itemInventoryList.Count(w => w.SubLevel > 0 && w.ParentRef.ListID.Contains(item.QBListID));
                if (temp > 0)
                {
                    tempList.Add(item.QBListID);
                }
            }

            List<Inventory> lstSub1 = itemInventoryList.Where(i => i.SubLevel == 1 && (i.SalesPrice != 0 || i.PurchaseCost != 0)).ToList();
            foreach (var item in lstSub1)
            {
                int temp = itemInventoryList.Count(w => w.SubLevel > 1 && w.ParentRef.ListID.Contains(item.QBListID));
                if (temp > 0)
                {
                    tempList.Add(item.QBListID);
                }
            }

            List<Inventory> lstSub2 = itemInventoryList.Where(i => i.SubLevel == 2 && (i.SalesPrice != 0 || i.PurchaseCost != 0)).ToList();
            foreach (var item in lstSub2)
            {
                int temp = itemInventoryList.Count(w => w.SubLevel > 2 && w.ParentRef.ListID.Contains(item.QBListID));
                if (temp > 0)
                {
                    tempList.Add(item.QBListID);
                }
            }

            return tempList;
        }

        private void ImportInventory(List<Inventory> itemList, int parentID, List<string> excList)
        {
            foreach (var item in itemList)
            {
                if (!String.IsNullOrEmpty(item.ParentRef.ListID))
                {
                    parentID = GetParentIdCategory(item.ParentRef.ListID);
                }

                if (item.SalesPrice == 0 && item.PurchaseCost == 0)
                {
                    ImportCategoryToOdoo(item, parentID);
                }
                else
                {
                    bool find = excList.Exists(x => x.Contains(item.QBListID));
                    if (find == true)
                        ImportCategoryToOdoo(item, parentID);
                    else
                        ImportProductItemToOdoo(item, parentID);
                }
            }
        }

        private void ImportCategoryToOdoo(Inventory item, int parentID)
        {
            using (NpgsqlConnection conn = new NpgsqlConnection(_odooConnectionString))
            {
                conn.Open();
                NpgsqlCommand cmd = new NpgsqlCommand(@"insert into product_category (create_uid, create_date, name, write_uid, parent_id, write_date, type, qblistid, edit_sequence)                                                    values(:create_uid, :create_date, :name,:write_uid, :parent_id, :write_date, :type, :qblistid, :edit_sequence)", conn);
                cmd.Parameters.Add(new NpgsqlParameter("create_uid", 1));
                cmd.Parameters.Add(new NpgsqlParameter("create_date", DateTime.Now));
                cmd.Parameters.Add(new NpgsqlParameter("name", item.FullName));
                cmd.Parameters.Add(new NpgsqlParameter("write_uid", 1));
                cmd.Parameters.Add(new NpgsqlParameter("parent_id", parentID));
                cmd.Parameters.Add(new NpgsqlParameter("write_date", DateTime.Now));
                cmd.Parameters.Add(new NpgsqlParameter("type", "normal"));
                cmd.Parameters.Add(new NpgsqlParameter("qblistid", item.QBListID));
                cmd.Parameters.Add(new NpgsqlParameter("edit_sequence", item.EditSequence));
                cmd.ExecuteNonQuery();
            }
        }

        private void ImportProductItemToOdoo(Inventory item, int parentID)
        {
            if (item.SubLevel == 0)
            {
                parentID = 4;//Other Product Category
            }
            using (NpgsqlConnection conn = new NpgsqlConnection(_odooConnectionString))
            {
                conn.Open();
                NpgsqlCommand cmdpdTemp = new NpgsqlCommand(@"insert into product_template (list_price, color, write_uid, mes_type, uom_id, create_date, uos_coeff, create_uid, sale_ok, categ_id, uom_po_id, description_sale, write_date, active, rental, name, type, sale_delay, purchase_ok, website_sequence, website_published, website_size_x, website_size_y, qblistid, edit_sequence) values(:list_price, :color, :write_uid, :mes_type, :uom_id, :create_date, :uos_coeff, :create_uid, :sale_ok, :categ_id, :uom_po_id, :description_sale, :write_date, :active, :rental, :name, :type, :sale_delay, :purchase_ok, :website_sequence, :website_published, :website_size_x, :website_size_y, :qblistid, :edit_sequence) returning id ", conn);
                cmdpdTemp.Parameters.Add(new NpgsqlParameter("list_price", item.SalesPrice));
                cmdpdTemp.Parameters.Add(new NpgsqlParameter("color", "0"));
                cmdpdTemp.Parameters.Add(new NpgsqlParameter("write_uid", 1));
                cmdpdTemp.Parameters.Add(new NpgsqlParameter("mes_type", "fixed"));
                cmdpdTemp.Parameters.Add(new NpgsqlParameter("uom_id", 1));
                cmdpdTemp.Parameters.Add(new NpgsqlParameter("create_date", DateTime.Now));
                cmdpdTemp.Parameters.Add(new NpgsqlParameter("uos_coeff", 1));
                cmdpdTemp.Parameters.Add(new NpgsqlParameter("create_uid", 1));
                cmdpdTemp.Parameters.Add(new NpgsqlParameter("sale_ok", true));
                cmdpdTemp.Parameters.Add(new NpgsqlParameter("categ_id", parentID));
                cmdpdTemp.Parameters.Add(new NpgsqlParameter("uom_po_id", 1));
                cmdpdTemp.Parameters.Add(new NpgsqlParameter("description_sale", item.SalesDesc));
                cmdpdTemp.Parameters.Add(new NpgsqlParameter("write_date", DateTime.Now));
                cmdpdTemp.Parameters.Add(new NpgsqlParameter("active", true));
                cmdpdTemp.Parameters.Add(new NpgsqlParameter("rental", false));
                cmdpdTemp.Parameters.Add(new NpgsqlParameter("name", item.FullName));
                cmdpdTemp.Parameters.Add(new NpgsqlParameter("type", item.Type));
                cmdpdTemp.Parameters.Add(new NpgsqlParameter("sale_delay", 7));
                cmdpdTemp.Parameters.Add(new NpgsqlParameter("purchase_ok", item.ApplyPurchase));
                cmdpdTemp.Parameters.Add(new NpgsqlParameter("website_sequence", "0"));
                cmdpdTemp.Parameters.Add(new NpgsqlParameter("website_published", false));
                cmdpdTemp.Parameters.Add(new NpgsqlParameter("website_size_x", 1));
                cmdpdTemp.Parameters.Add(new NpgsqlParameter("website_size_y", 1));
                cmdpdTemp.Parameters.Add(new NpgsqlParameter("qblistid", item.QBListID));
                cmdpdTemp.Parameters.Add(new NpgsqlParameter("edit_sequence", item.EditSequence));
                var idTemplate = cmdpdTemp.ExecuteScalar().ToString();

                NpgsqlCommand cmdpd = new NpgsqlCommand(@"insert into product_product (create_date, name_template, create_uid, product_tmpl_id, write_uid, write_date, active, qblistid, edit_sequence) values(:create_date, :name_template, :create_uid, :product_tmpl_id, :write_uid, :write_date, :active, :qblistid, :edit_sequence)", conn);
                cmdpd.Parameters.Add(new NpgsqlParameter("create_date", DateTime.Now));
                cmdpd.Parameters.Add(new NpgsqlParameter("name_template", item.FullName));
                cmdpd.Parameters.Add(new NpgsqlParameter("create_uid", 1));
                cmdpd.Parameters.Add(new NpgsqlParameter("product_tmpl_id", idTemplate));
                cmdpd.Parameters.Add(new NpgsqlParameter("write_uid", 1));
                cmdpd.Parameters.Add(new NpgsqlParameter("write_date", DateTime.Now));
                cmdpd.Parameters.Add(new NpgsqlParameter("active", true));
                cmdpd.Parameters.Add(new NpgsqlParameter("qblistid", item.QBListID));
                cmdpd.Parameters.Add(new NpgsqlParameter("edit_sequence", item.EditSequence));
                cmdpd.ExecuteNonQuery();
            }
        }

        public int GetParentIdCategory(string QBListID)
        {
            int rs;
            using (NpgsqlConnection conn = new NpgsqlConnection(_odooConnectionString))
            {
                conn.Open();
                NpgsqlCommand cmd = new NpgsqlCommand(@"Select id from product_category where qblistid=:qblistid", conn);
                cmd.Parameters.Add(new NpgsqlParameter("qblistid", QBListID));
                rs = Convert.ToInt32(cmd.ExecuteScalar());
            }
            return rs;
        }

        /*Inventory End*/
        private void ImportCustomersToOdoo(List<Customer> customerList)
        {
            using (NpgsqlConnection conn = new NpgsqlConnection(_odooConnectionString))
            {
                conn.Open();

                foreach (var customer in customerList)
                {
                    NpgsqlCommand cmd = new NpgsqlCommand(@"insert into res_partner (name, display_name, street, city, zip, website, email, supplier, is_company, customer, notify_email, employee, phone, active, qblistid, edit_sequence) 
                                                    values(:name, :display_name, :street, :city, :zip, :website, :email, :supplier, :is_company, :customer, :notify_email, :employee, :phone, :active, :qblistid, :edit_sequence)", conn);
                    cmd.Parameters.Add(new NpgsqlParameter("name", customer.DisplayName));
                    cmd.Parameters.Add(new NpgsqlParameter("display_name", customer.DisplayName));
                    cmd.Parameters.Add(new NpgsqlParameter("street", customer.Street));
                    cmd.Parameters.Add(new NpgsqlParameter("city", customer.City));
                    cmd.Parameters.Add(new NpgsqlParameter("zip", customer.Zip));
                    cmd.Parameters.Add(new NpgsqlParameter("website", customer.Website));
                    cmd.Parameters.Add(new NpgsqlParameter("email", customer.Email));
                    cmd.Parameters.Add(new NpgsqlParameter("supplier", customer.IsSupplier));
                    cmd.Parameters.Add(new NpgsqlParameter("is_company", customer.IsCompany));
                    cmd.Parameters.Add(new NpgsqlParameter("customer", customer.IsCustomer));
                    cmd.Parameters.Add(new NpgsqlParameter("notify_email", customer.NotifyEmail));
                    cmd.Parameters.Add(new NpgsqlParameter("employee", customer.IsEmployee));
                    cmd.Parameters.Add(new NpgsqlParameter("phone", customer.Phone));
                    cmd.Parameters.Add(new NpgsqlParameter("active", customer.IsActive));
                    cmd.Parameters.Add(new NpgsqlParameter("qblistid", customer.QBListId));
                    cmd.Parameters.Add(new NpgsqlParameter("edit_sequence", customer.EditSequence));
                    cmd.ExecuteNonQuery();
                }
            }

            var xml = "<?xml version=\"1.0\" ?> <?qbxml version=\"5.0\" ?> <QBXML> <QBXMLEvents> <DataEvent> <CompanyFilePath>C:\\Users\\Public\\Documents\\Intuit\\QuickBooks\\Sample Company Files\\QuickBooks 2012\\sample_product-based business.qbw</CompanyFilePath> <HostInfo> <ProductName>QuickBooks Accountant 2012</ProductName> <MajorVersion>22</MajorVersion> <MinorVersion>0</MinorVersion> <Country>US</Country> </HostInfo> <ListEvent> <ListEventType>Customer</ListEventType> <ListEventOperation>Modify</ListEventOperation> <ListID>150000-933272658</ListID> </ListEvent> <DataEventRecoveryTime>2015-05-05T16:06:02+07:00</DataEventRecoveryTime> </DataEvent> </QBXMLEvents> </QBXML>";

            //Process.Start("SubscribeAndHandleQBEvent.exe");
        }

        private void ImportInventorysToOdoo(List<Inventory> inventoryList)
        {
            using (NpgsqlConnection conn = new NpgsqlConnection(_odooConnectionString))
            {
                conn.Open();

                foreach (var inventory in inventoryList)
                {
                    NpgsqlCommand cmd = new NpgsqlCommand(@"insert into product_template (name, active, list_price, description_sale, description_purchase,uom_id, categ_id, uom_po_id, type,purchase_ok, edit_sequence, qblistid, is_qb_notification) 
                                        values(:name, :active, :list_price, :description_sale, :description_purchase,:uom_id, :categ_id, :uom_po_id, :type, :purchase_ok, :edit_sequence, :qblistid, :is_qb_notification)", conn);
                    cmd.Parameters.Add(new NpgsqlParameter("name", inventory.Name));
                    cmd.Parameters.Add(new NpgsqlParameter("active", inventory.isActive));
                    cmd.Parameters.Add(new NpgsqlParameter("list_price", inventory.SalesPrice));
                    cmd.Parameters.Add(new NpgsqlParameter("description_sale", inventory.SalesDesc));
                    cmd.Parameters.Add(new NpgsqlParameter("description_purchase", inventory.PurchaseDesc));
                    cmd.Parameters.Add(new NpgsqlParameter("uom_id", inventory.Measure));
                    cmd.Parameters.Add(new NpgsqlParameter("categ_id", inventory.Category));
                    cmd.Parameters.Add(new NpgsqlParameter("uom_po_id", inventory.PucharseUnitOfMeasure));
                    cmd.Parameters.Add(new NpgsqlParameter("type", inventory.Type));
                    cmd.Parameters.Add(new NpgsqlParameter("purchase_ok", inventory.ApplyPurchase));
                    cmd.Parameters.Add(new NpgsqlParameter("qblistid", inventory.QBListID));
                    cmd.Parameters.Add(new NpgsqlParameter("edit_sequence", inventory.EditSequence));
                    cmd.Parameters.Add(new NpgsqlParameter("is_qb_notification", "true"));
                    cmd.ExecuteNonQuery();

                    string str = "select last_value from product_template_id_seq";
                    NpgsqlCommand command = new NpgsqlCommand(str, conn);
                    var template_id = command.ExecuteScalar();

                    NpgsqlCommand cmd2 = new NpgsqlCommand(@"insert into product_product(create_date, name_template, create_uid, product_tmpl_id, write_uid, write_date, active, qblistid, edit_sequence, is_qb_notification) values(:create_date, :name_template, :create_uid, :product_tmpl_id, :write_uid, :write_date, :active, :qblistid, :edit_sequence, :is_qb_notification)", conn);
                    cmd2.Parameters.Add(new NpgsqlParameter("create_date", DateTime.Now));
                    cmd2.Parameters.Add(new NpgsqlParameter("name_template", inventory.Name));
                    cmd2.Parameters.Add(new NpgsqlParameter("create_uid", 1));
                    cmd2.Parameters.Add(new NpgsqlParameter("product_tmpl_id", template_id));
                    cmd2.Parameters.Add(new NpgsqlParameter("write_uid", 1));
                    cmd2.Parameters.Add(new NpgsqlParameter("write_date", DateTime.Now));
                    cmd2.Parameters.Add(new NpgsqlParameter("active", true));
                    cmd2.Parameters.Add(new NpgsqlParameter("qblistid", inventory.QBListID));
                    cmd2.Parameters.Add(new NpgsqlParameter("edit_sequence", inventory.EditSequence));
                    cmd2.Parameters.Add(new NpgsqlParameter("is_qb_notification", "true"));
                    cmd2.ExecuteNonQuery();
                }
            }
        }

        private void importLagFile_Click(object sender, RoutedEventArgs e)
        {
            // Set Access connection and select strings.
            // The path to BugTypes.MDB must be changed if you build 
            // the sample from the command line:
            ImportAllProduct();
            var rootCategory = GetCategoryTree();
            ImportCategoriesIntoOdoo(rootCategory);
            MapProductWithCategories(rootCategory);
        }

        private void MapProductWithCategories(Category rootCategory)
        {
            string strAccessSelect = "SELECT * FROM LAG_ECDB2_HIERARCHY";

            // Create the dataset and add the ITEM table to it:
            DataSet myDataSet = new DataSet();
            OleDbConnection myAccessConn = null;
            try
            {
                myAccessConn = new OleDbConnection(_lagasseMdbFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: Failed to create a database connection. \n{0}", ex.Message);
                return;
            }

            try
            {
                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(myDataSet, "CATEGORIES");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: Failed to retrieve the required data from the DataBase.\n{0}", ex.Message);
                return;
            }
            finally
            {
                myAccessConn.Close();
            }
            DataRowCollection dra = myDataSet.Tables["CATEGORIES"].Rows;
            using (NpgsqlConnection conn = new NpgsqlConnection(_odooConnectionString))
            {
                conn.Open();
                foreach (DataRow dr in dra)
                {
                    var productId = dr["lag_prodid"];
                    var lvl1 = dr["lag_lvl1_id"];
                    var lvl2 = dr["lag_lvl2_id"];
                    var lvl3 = dr["lag_lvl3_id"];
                    var sequenceId = dr["seq_id"];


                    NpgsqlCommand cmd = new NpgsqlCommand(@"insert into product_public_category_product_template_rel (product_template_id, product_public_category_id) 
                                                    values(:productId, :categoryId)", conn);
                    cmd.Parameters.Add(new NpgsqlParameter("productId", _idMapping[productId.ToString().Trim()]));
                    var categoryOdooId = string.Empty;
                    var firstLevel = rootCategory.SubItems.FirstOrDefault(item => string.Equals(item.LagasseId, lvl1.ToString(), StringComparison.OrdinalIgnoreCase));
                    if (firstLevel != null)
                    {
                        var secondLevel = firstLevel.SubItems.FirstOrDefault(item => string.Equals(item.LagasseId, lvl2.ToString(), StringComparison.OrdinalIgnoreCase));
                        if (secondLevel != null)
                        {
                            var thirdLevel = secondLevel.SubItems.FirstOrDefault(item => string.Equals(item.LagasseId, lvl3.ToString(), StringComparison.OrdinalIgnoreCase));
                            if (thirdLevel != null)
                            {
                                categoryOdooId = thirdLevel.OdooId;
                            }
                        }
                    }
                    cmd.Parameters.Add(new NpgsqlParameter("categoryId",
                        categoryOdooId));
                    cmd.ExecuteNonQuery();

                }
            }
        }

        private Category GetCategoryTree()
        {
            string strAccessSelect = "SELECT * FROM LAG_ECDB2_HIERARCHY";
            var rootCategory = new Category();
            // Create the dataset and add the ITEM table to it:
            DataSet myDataSet = new DataSet();
            OleDbConnection myAccessConn = null;
            try
            {
                myAccessConn = new OleDbConnection(_lagasseMdbFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: Failed to create a database connection. \n{0}", ex.Message);
                return rootCategory;
            }

            try
            {

                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(myDataSet, "CATEGORIES");

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: Failed to retrieve the required data from the DataBase.\n{0}", ex.Message);
                return rootCategory;
            }
            finally
            {
                myAccessConn.Close();
            }
            DataRowCollection dra = myDataSet.Tables["CATEGORIES"].Rows;

            foreach (DataRow dr in dra)
            {
                // Print the CategoryID as a subscript, then the CategoryName:
                var lvl1 = dr["lag_lvl1"];
                var lvl2 = dr["lag_lvl2"];
                var lvl3 = dr["lag_lvl3"];
                var lvl1id = dr["lag_lvl1_id"];
                var lvl2id = dr["lag_lvl2_id"];
                var lvl3id = dr["lag_lvl3_id"];
                var sequenceId = dr["seq_id"];
                var lvl1item = new Category()
                {
                    Name = lvl1.ToString(),
                    LagasseId = lvl1id.ToString(),
                    Level = 1
                };
                var lvl2item = new Category()
                {
                    Name = lvl2.ToString(),
                    LagasseId = lvl2id.ToString(),
                    Level = 2
                };
                var lvl3item = new Category()
                {
                    Name = lvl3.ToString(),
                    LagasseId = lvl3id.ToString(),
                    Level = 3
                };

                var existLvl1 = rootCategory.SubItems.FirstOrDefault(item => string.Equals(item.Name, lvl1.ToString(), StringComparison.OrdinalIgnoreCase));
                if (existLvl1 == null)
                {
                    rootCategory.SubItems.Add(lvl1item);
                    existLvl1 = lvl1item;
                }

                var existLvl2 = existLvl1.SubItems.FirstOrDefault(item => string.Equals(item.Name, lvl2.ToString(), StringComparison.OrdinalIgnoreCase));
                if (existLvl2 == null)
                {
                    existLvl1.SubItems.Add(lvl2item);
                    existLvl2 = lvl2item;
                }

                var existLvl3 = existLvl2.SubItems.FirstOrDefault(item => string.Equals(item.Name, lvl3.ToString(), StringComparison.OrdinalIgnoreCase));
                if (existLvl3 == null)
                {
                    existLvl2.SubItems.Add(lvl3item);
                    existLvl3 = lvl3item;
                }
            }

            return rootCategory;
        }

        private void ImportAllProduct()
        {

            string strAccessSelect = "SELECT sg.sku_gp_nm as Name, it.item_wgt as Weight, it.item_depth as Depth, it.item_wdt as Width, ";
            strAccessSelect += "it.item_hgt as Height, it.list_amt as Price, sg.sling_pnt_1 as Des1, sg.sling_pnt_2 as Des2, sg.sling_pnt_3 as Des3, ";
            strAccessSelect += "sg.sling_pnt_4 as Des4, sg.sling_pnt_5 as Des5, sg.sling_pnt_6 as Des6 , it.lag_prodid as ProductID, it.sgl_img as Img, it.prod_dsc as ProdDsc ";
            strAccessSelect += "FROM ITEM it ";
            strAccessSelect += "LEFT JOIN SKU_GROUP sg ON it.sku_gp_id=sg.sku_gp_id";
            // Create the dataset and add the Categories table to it:
            DataSet myDataSet = new DataSet();
            OleDbConnection myAccessConn = null;
            try
            {
                myAccessConn = new OleDbConnection(_lagasseMdbFile);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: Failed to create a database connection. \n{0}", ex.Message);
                return;
            }

            try
            {

                OleDbCommand myAccessCommand = new OleDbCommand(strAccessSelect, myAccessConn);
                OleDbDataAdapter myDataAdapter = new OleDbDataAdapter(myAccessCommand);

                myAccessConn.Open();
                myDataAdapter.Fill(myDataSet, "Products");

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: Failed to retrieve the required data from the DataBase.\n{0}", ex.Message);
                return;
            }
            finally
            {
                myAccessConn.Close();
            }

            // A dataset can contain multiple tables, so let's get them
            // all into an array:
            DataTableCollection dta = myDataSet.Tables;

            DataColumnCollection drc = myDataSet.Tables["Products"].Columns;

            DataRowCollection dra = myDataSet.Tables["Products"].Rows;

            using (NpgsqlConnection conn = new NpgsqlConnection(_odooConnectionString))
            {
                conn.Open();

                foreach (DataRow dr in dra)
                {

                    string desc_sales = "";
                    desc_sales += dr["ProdDsc"].ToString();
                    string web_desc = "<section class=\"mt16 mb16\">";
                    web_desc += "<div class=\"container\">";
                    if (dr[1].ToString() != "") web_desc += "<div class=\"row\">" + "Weight: " + float.Parse(dr[1].ToString()) + "</div>";
                    if (dr[2].ToString() != "") web_desc += "<div class=\"row\">" + "Weight: " + float.Parse(dr[2].ToString()) + "</div>";
                    if (dr[3].ToString() != "") web_desc += "<div class=\"row\">" + "Weight: " + float.Parse(dr[3].ToString()) + "</div>";
                    if (dr[4].ToString() != "") web_desc += "<div class=\"row\">" + "Weight: " + float.Parse(dr[4].ToString()) + "</div>";

                    if (dr[6].ToString() != "") web_desc += "<div class=\"row\">" + dr[6] + "</div>";
                    if (dr[7].ToString() != "") web_desc += "<div class=\"row\">" + dr[7] + "</div>";
                    if (dr[8].ToString() != "") web_desc += "<div class=\"row\">" + dr[8] + "</div>";
                    if (dr[9].ToString() != "") web_desc += "<div class=\"row\">" + dr[9] + "</div>";
                    if (dr[10].ToString() != "") web_desc += "<div class=\"row\">" + dr[10] + "</div>";
                    if (dr[11].ToString() != "") web_desc += "<div class=\"row\">" + dr[11] + "</div>";
                    var imageFileName = dr["Img"].ToString();
                    web_desc += "</div>";
                    web_desc += "</section>";
                    NpgsqlCommand cmd = new NpgsqlCommand(@"insert into product_template (id, list_price, weight, color, image, image_medium, image_small,
write_uid, mes_type, uom_id, create_date, uos_coeff, create_uid, sale_ok, categ_id, company_id, uom_po_id, description_sale,
write_date, active, rental, name, type, sale_delay, purchase_ok, website_sequence, website_published, website_description,
website_size_x, website_size_y, lag_prodid) values 
(DEFAULT, :list_price, :weight, :color, :image, :image_medium, :image_small, :write_uid, :mes_type, :uom_id, :create_date, :uos_coeff, :create_uid, :sale_ok, :categ_id,
:company_id, :uom_po_id, :description_sale, :write_date, :active, :rental, :name, :type, :sale_delay, :purchase_ok,
:website_sequence, :website_published, :website_description, :website_size_x, :website_size_y, :lag_prodid) returning id", conn);
                    cmd.Parameters.Add(new NpgsqlParameter("list_price", dr[5]));
                    cmd.Parameters.Add(new NpgsqlParameter("weight", dr[3]));
                    cmd.Parameters.Add(new NpgsqlParameter("color", "0"));
                    cmd.Parameters.Add(new NpgsqlParameter("image", GetImageBase64EncodedString(ImageType.Large, _lagasseFolderPath, imageFileName)));
                    cmd.Parameters.Add(new NpgsqlParameter("image_medium", GetImageBase64EncodedString(ImageType.Medium, _lagasseFolderPath, imageFileName)));
                    cmd.Parameters.Add(new NpgsqlParameter("image_small", GetImageBase64EncodedString(ImageType.Small, _lagasseFolderPath, imageFileName)));
                    cmd.Parameters.Add(new NpgsqlParameter("write_uid", 1));
                    cmd.Parameters.Add(new NpgsqlParameter("mes_type", "fixed"));
                    cmd.Parameters.Add(new NpgsqlParameter("uom_id", 1));
                    cmd.Parameters.Add(new NpgsqlParameter("create_date", DateTime.Now));
                    cmd.Parameters.Add(new NpgsqlParameter("uos_coeff", 1.00));
                    cmd.Parameters.Add(new NpgsqlParameter("create_uid", 1));
                    cmd.Parameters.Add(new NpgsqlParameter("sale_ok", true));
                    cmd.Parameters.Add(new NpgsqlParameter("categ_id", 1));
                    cmd.Parameters.Add(new NpgsqlParameter("company_id", 1));
                    cmd.Parameters.Add(new NpgsqlParameter("uom_po_id", 1));
                    cmd.Parameters.Add(new NpgsqlParameter("description_sale", desc_sales));
                    cmd.Parameters.Add(new NpgsqlParameter("write_date", DateTime.Now));
                    cmd.Parameters.Add(new NpgsqlParameter("active", true));
                    cmd.Parameters.Add(new NpgsqlParameter("rental", false));
                    cmd.Parameters.Add(new NpgsqlParameter("name", dr[0]));
                    cmd.Parameters.Add(new NpgsqlParameter("type", "product"));
                    cmd.Parameters.Add(new NpgsqlParameter("sale_delay", 7));
                    cmd.Parameters.Add(new NpgsqlParameter("purchase_ok", true));
                    cmd.Parameters.Add(new NpgsqlParameter("website_sequence", "0"));
                    cmd.Parameters.Add(new NpgsqlParameter("website_published", true));
                    cmd.Parameters.Add(new NpgsqlParameter("website_description", web_desc));
                    cmd.Parameters.Add(new NpgsqlParameter("website_size_x", 1));
                    cmd.Parameters.Add(new NpgsqlParameter("website_size_y", 1));
                    cmd.Parameters.Add(new NpgsqlParameter("lag_prodid", dr[12]));
                    var odooId = cmd.ExecuteScalar().ToString();
                    _idMapping.Add(dr[12].ToString(), odooId);

                    NpgsqlCommand cmd2 = new NpgsqlCommand(@"insert into product_product(create_date, name_template, create_uid, product_tmpl_id, write_uid, write_date, active) values(:create_date, :name_template, :create_uid, :product_tmpl_id, :write_uid, :write_date, :active)", conn);
                    cmd2.Parameters.Add(new NpgsqlParameter("create_date", DateTime.Now));
                    cmd2.Parameters.Add(new NpgsqlParameter("name_template", dr[0]));
                    cmd2.Parameters.Add(new NpgsqlParameter("create_uid", 1));
                    cmd2.Parameters.Add(new NpgsqlParameter("product_tmpl_id", odooId));
                    cmd2.Parameters.Add(new NpgsqlParameter("write_uid", 1));
                    cmd2.Parameters.Add(new NpgsqlParameter("write_date", DateTime.Now));
                    cmd2.Parameters.Add(new NpgsqlParameter("active", true));
                    cmd2.ExecuteNonQuery();
                }
            }
        }

        private string GetImageBase64EncodedString(ImageType imageType, string rootFolderPath, string fileName)
        {
            var urlTemplate = rootFolderPath + "/{0}/" + fileName;
            var fullFileUrl = string.Empty;
            switch (imageType)
            {
                case ImageType.Small:
                    fullFileUrl = string.Format(urlTemplate, "lgsi100");
                    break;
                case ImageType.Medium:
                    fullFileUrl = string.Format(urlTemplate, "lgsi240");
                    break;
                case ImageType.Large:
                    fullFileUrl = string.Format(urlTemplate, "lgsi400");
                    break;
            }
            byte[] imageArray = System.IO.File.ReadAllBytes(fullFileUrl);
            string base64ImageRepresentation = Convert.ToBase64String(imageArray);
            return base64ImageRepresentation;
        }
        private void ImportCategoriesIntoOdoo(Category rootCategory)
        {

            var firstLevel = rootCategory.SubItems.Where(item => item.Level == 1).ToList();
            using (NpgsqlConnection conn = new NpgsqlConnection(_odooConnectionString))
            {
                conn.Open();

                foreach (var level in firstLevel)
                {
                    NpgsqlCommand cmd = new NpgsqlCommand(@"insert into product_public_category (id, name) 
                                                    values(DEFAULT, :name) returning id", conn);
                    cmd.Parameters.Add(new NpgsqlParameter("name", level.Name));

                    var currentId = cmd.ExecuteScalar().ToString();
                    var subItems = level.SubItems;
                    if (subItems.Count > 0)
                    {
                        InsertChildIntoParentCategory(subItems, currentId);
                    }
                }
            }
        }

        private void InsertChildIntoParentCategory(List<Category> listItems, string parentId)
        {
            using (NpgsqlConnection conn = new NpgsqlConnection(_odooConnectionString))
            {
                conn.Open();

                foreach (var item in listItems)
                {
                    NpgsqlCommand cmd = new NpgsqlCommand(@"insert into product_public_category (id, name, parent_id) 
                                                    values(DEFAULT, :name, :parent_id) returning id", conn);
                    cmd.Parameters.Add(new NpgsqlParameter("name", item.Name));
                    cmd.Parameters.Add(new NpgsqlParameter("parent_id", parentId));
                    var currentId = cmd.ExecuteScalar().ToString();
                    item.OdooId = currentId;

                    var subItems = item.SubItems;
                    if (subItems.Count > 0)
                    {
                        InsertChildIntoParentCategory(subItems, currentId);
                    }
                }
            }
        }
    }
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
