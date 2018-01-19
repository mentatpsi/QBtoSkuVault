using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using QBToSkuVault.Model;
using QBFC13Lib;
using QBXMLRP2Lib;
using System.Xml;
using System.Net;
using System.Collections.Specialized;
using Newtonsoft.Json;
using System.IO;
using System.Runtime.InteropServices;

namespace QBToSkuVault
{
    public class Program
    {
        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;

        static void Main(string[] args)
        {
            var tokens = getTokens();
            var handle = GetConsoleWindow();
            ShowWindow(handle, SW_HIDE);

            Items it = new Items();
            var cur = it.itemInventories;
            CreateProductsRequest cpreq = new CreateProductsRequest();
            cpreq.UserToken = tokens["UserToken"].ToString();
            cpreq.TenantToken = tokens["TenantToken"].ToString();

            SetItemQuantitiesRequest siqreq = new SetItemQuantitiesRequest();
            siqreq.UserToken = tokens["UserToken"].ToString();
            siqreq.TenantToken = tokens["TenantToken"].ToString();
            
            List<CreateProductRequest> critems = new List<CreateProductRequest>();
            List<UpdateQuantityRequest> siitems = new List<UpdateQuantityRequest>();

            int count = 0;
            foreach (var item in cur)
            {
                if (count % 100 == 0)
                {
                    System.Console.WriteLine("Updating 100 items");
                    cpreq.Items = critems;
                    siqreq.Items = siitems;
                    sendObject(cpreq, "https://app.skuvault.com/api/products/createProducts?json");
                    sendObject(siqreq, "https://app.skuvault.com/api/inventory/setItemQuantities?json");
                    critems = new List<CreateProductRequest>();
                    siitems = new List<UpdateQuantityRequest>();
                    System.Threading.Thread.Sleep(5000);
                }
                //
                if (Int32.Parse(item.QuantityOnHand) >= 0)
                {
                    var cprequest = createProducts(item.FullName, tokens["TenantToken"].ToString(), tokens["UserToken"].ToString());
                    var uqrequest = updateQuantity(item.FullName, item.QuantityOnHand, tokens["TenantToken"].ToString(), tokens["UserToken"].ToString());
                    critems.Add(cprequest);
                    siitems.Add(uqrequest);
                    count = count + 1;

                }

                else
                {
                    var cprequest = createProducts(item.FullName, tokens["TenantToken"].ToString(), tokens["UserToken"].ToString());
                    var uqrequest = updateQuantity(item.FullName, "0", tokens["TenantToken"].ToString(), tokens["UserToken"].ToString());
                    critems.Add(cprequest);
                    siitems.Add(uqrequest);
                    count = count + 1;
                }

                
            }

            if (critems.Count > 0)
            {
                System.Console.WriteLine(String.Format("Updating {0} items", critems.Count.ToString()));
                cpreq.Items = critems;
                siqreq.Items = siitems;
                sendObject(cpreq, "https://app.skuvault.com/api/products/createProducts?json");
                sendObject(siqreq, "https://app.skuvault.com/api/inventory/setItemQuantities?json");
            }
            System.Console.WriteLine(cur.Count);

        }
        string workCurrency = "$";
        string homeCurrency = "$";
        decimal exchangeRate = 1.0m;

        private static string ticket;
        private static QBXMLRP2Lib.RequestProcessor2 rp;
        private static string maxVersion;
        //private static string companyFile = "C:\\Users\\Public\\Documents\\Intuit\\QuickBooks\\Sample Company Files\\QuickBooks Enterprise Solutions 15.0\\sample_product-based business.qbw";
        private static string companyFile = "G:\\QuickBooks\\TIM.QBW";
        private static QBFileMode mode = QBFileMode.qbFileOpenDoNotCare;

        private static string appID = "IDN123";
        private static string appName = "QBToSkuVault";

        public static class Http
        {
            public static byte[] Post(string uri, NameValueCollection pairs)
            {
                byte[] response = null;
                using (WebClient client = new WebClient())
                {
                    response = client.UploadValues(uri, pairs);
                }
                return response;
            }
            public static byte[] Post(string uri, byte[] pairs)
            {
                byte[] response = null;
                using (WebClient client = new WebClient())
                {
                    response = client.UploadData(uri, pairs);
                }
                return response;
            }
        }

        public static void sendObject(Object request, string uri)
        {
            var json = JsonConvert.SerializeObject(request);
            HttpPost(uri, json);
        }

        public static void HttpPost(string uri, string json)  {
            var httpWebRequest = (HttpWebRequest)WebRequest.Create(uri);
            httpWebRequest.ContentType = "application/json";
            httpWebRequest.Method = "POST";

            using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
            {
                    streamWriter.Write(json);
                    streamWriter.Flush();
                    streamWriter.Close();
            }
            try
            {
                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                System.Console.WriteLine("Posting Successful");
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    var result = streamReader.ReadToEnd();
                }
            }
            catch (Exception e) {
                var x = e;
                System.Console.WriteLine("Bad Request");
            }
        }
        public static Dictionary<string,object> getTokens()
        {
            var tokens = new Dictionary<string, string>();
            
            var data = new NameValueCollection() {
                { "Email", "tomer@timcorpofnj.com" },
                { "Password", "yrvellayVarona3" }
            };
            var response = Http.Post("https://app.skuvault.com/api/gettokens?format=json", data);
            System.Console.WriteLine(response);
            string result = System.Text.Encoding.UTF8.GetString(response);
            Dictionary<string, object> dictionaryLevelOne = JsonConvert.DeserializeObject<Dictionary<string, object>>(result);
            return dictionaryLevelOne;
        }


        public static CreateProductRequest createProducts(string sku, string tenant, string user)
        {
            var tokens = new Dictionary<string, string>();

            var supplierinfo = new NameValueCollection()
            {
                { "SupplierName", "Tim Corp of NJ" },
                { "isPrimary", "true" },
            };
            var request = new CreateProductRequest();
            var supplierInfo = new SupplierInfo();
            supplierInfo.isPrimary = "true";
            supplierInfo.SupplierName = "Tim Corp of NJ";
            request.SupplierInfo = supplierInfo;
            request.Brand = "Generic";
            request.Classification = "General";
            request.Sku = sku;

            return request;
        }

        public static UpdateQuantityRequest updateQuantity(string sku, string quantity, string tenant, string user)
        {
            var request = new UpdateQuantityRequest();
            request.Sku = sku;
            request.WarehouseId = "18694";
            request.LocationCode = "LOC-01";
            request.Quantity = quantity;

            return request;
        }


        /*public static Dictionary<string, object> getProducts(Dictionary<string,object> tokens)
        {
            var data = new NameValueCollection() {
                { "PageSize", "10000" },
                { "TenantToken", tokens["TenantToken"].ToString()},
                { "UserToken", tokens["UserToken"].ToString()}
            };
            var response = Http.Post("https://app.skuvault.com/api/products/getProducts?format=json", data);
            string result = System.Text.Encoding.UTF8.GetString(response);
            System.Console.WriteLine(response);
            Dictionary<string, object> dictionaryLevelOne = JsonConvert.DeserializeObject<Dictionary<string, object>>(result);
            Dictionary<string, object> dictionaryLevelTwo = JsonConvert.DeserializeObject<Dictionary<string, List>>(dictionaryLevelOne["Products"]);
            return dictionaryLevelOne;
        }*/




        private List<Dictionary<string, string>> getItemInventory()
        {
            string request = "ItemInventoryQueryRq";
            connectToQB();
            int count = getCount(request);
            string response = processRequestFromQB(buildItemInventoryQueryRqXML(new string[] { "ListID", "FullName", "QuantityOnHand", "ManufacturerPartNumber" }, null));
            return parseItemInvenQueryRs(response, count);
        }

        public virtual int parseRsForCount(string xml, string request)
        {
            int ret = -1;
            try
            {
                XmlNodeList RsNodeList = null;
                XmlDocument Doc = new XmlDocument();
                Doc.LoadXml(xml);
                string tagname = request.Replace("Rq", "Rs");
                RsNodeList = Doc.GetElementsByTagName(tagname);
                System.Text.StringBuilder popupMessage = new System.Text.StringBuilder();
                XmlAttributeCollection rsAttributes = RsNodeList.Item(0).Attributes;
                XmlNode retCount = rsAttributes.GetNamedItem("retCount");
                ret = Convert.ToInt32(retCount.Value);
            }
            catch (Exception e)
            {
                System.Console.WriteLine("Error encountered: " + e.Message);
                ret = -1;
            }
            return ret;
        }

        private int getCount(string request)
        {
            string response = processRequestFromQB(buildDataCountQuery(request));
            int count = parseRsForCount(response, request);
            return count;
        }
        public virtual string buildDataCountQuery(string request)
        {
            string input = "";
            XmlDocument inputXMLDoc = new XmlDocument();
            XmlElement qbXMLMsgsRq = buildRqEnvelope(inputXMLDoc, maxVersion);
            XmlElement queryRq = inputXMLDoc.CreateElement(request);
            queryRq.SetAttribute("metaData", "MetaDataOnly");
            qbXMLMsgsRq.AppendChild(queryRq);
            input = inputXMLDoc.OuterXml;
            return input;
        }
        private string processRequestFromQB(string request)
        {
            try
            {
                return rp.ProcessRequest(ticket, request);
            }
            catch (Exception e)
            {
                System.Console.WriteLine(e.Message);
                return null;
            }
        }

        private XmlElement buildRqEnvelope(XmlDocument doc, string maxVer)
        {
            doc.AppendChild(doc.CreateXmlDeclaration("1.0", null, null));
            doc.AppendChild(doc.CreateProcessingInstruction("qbxml", "version=\"" + maxVer + "\""));
            XmlElement qbXML = doc.CreateElement("QBXML");
            doc.AppendChild(qbXML);
            XmlElement qbXMLMsgsRq = doc.CreateElement("QBXMLMsgsRq");
            qbXML.AppendChild(qbXMLMsgsRq);
            qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
            return qbXMLMsgsRq;
        }
        private string buildItemInventoryQueryRqXML(string[] includeRetElement, string fullName)
        {
            string xml = "";
            XmlDocument xmlDoc = new XmlDocument();
            XmlElement qbXMLMsgsRq = buildRqEnvelope(xmlDoc, maxVersion);
            qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
            XmlElement ItemQueryRq = xmlDoc.CreateElement("ItemInventoryQueryRq");
            qbXMLMsgsRq.AppendChild(ItemQueryRq);
            if (fullName != null)
            {
                XmlElement fullNameElement = xmlDoc.CreateElement("FullName");
                ItemQueryRq.AppendChild(fullNameElement).InnerText = fullName;
            }
            for (int x = 0; x < includeRetElement.Length; x++)
            {
                XmlElement includeRet = xmlDoc.CreateElement("IncludeRetElement");
                ItemQueryRq.AppendChild(includeRet).InnerText = includeRetElement[x];
            }
            ItemQueryRq.SetAttribute("requestID", "2");
            xml = xmlDoc.OuterXml;
            return xml;
        }
        private static void connectToQB()
        {
            rp = new QBXMLRP2Lib.RequestProcessor2();
            rp.OpenConnection(appID, appName);
            ticket = rp.BeginSession(companyFile, mode);
            System.Array versions = rp.QBXMLVersionsForSession[ticket];
            int len = versions.Length - 1;
            maxVersion = versions.GetValue(len).ToString();
        }

        public class Items
        {
            public string[] items;
            public List<Dictionary<string, string>> itemInventory;
            public List<ItemInventory> itemInventories;
            public Items()
            {
                itemInventories = new List<ItemInventory>();
                Program p1 = new Program();
                itemInventory = p1.getItemInventory();
                string resultCSV;
                foreach (Dictionary<string, string> item in itemInventory)
                {

                    if (item.Count != 0)
                    {
                        ItemInventory tempRes = new ItemInventory();

                        tempRes.ListID = item["ListID"];
                        tempRes.FullName = item["FullName"];
                        tempRes.QuantityOnHand = item["QuantityOnHand"];
                        if (item.ContainsKey("ManufacturerPartNumber"))
                        {
                            tempRes.ManufacturerPartNumber = item["ManufacturerPartNumber"];
                        }
                        itemInventories.Add(tempRes);
                    }
                }
                //resultCSV = ServiceStack.StringExtensions.ToCsv<List<ItemInventory>>(itemInventories);
                //System.IO.File.WriteAllText("C:\\temp\\resultcsv.csv", resultCSV);

            }


        }

        private static bool IsDecimal(string theValue)
        {
            bool returnVal = false;
            try
            {
                Convert.ToDouble(theValue, System.Globalization.CultureInfo.CurrentCulture);
                returnVal = true;
            }
            catch
            {
                returnVal = false;
            }
            finally
            {
            }

            return returnVal;
        }

        private List<Dictionary<string, string>> parseItemInvenQueryRs(string xml, int count)
        {
            /*
              <?xml version="1.0" ?> 
            - <QBXML>
            - <QBXMLMsgsRs>
            - <ItemQueryRs requestID="2" statusCode="0" statusSeverity="Info" statusMessage="Status OK">
            - <ItemServiceRet>
  	            <ListID>20000-933272655</ListID> 
  	            <TimeCreated>1999-07-29T11:24:15-08:00</TimeCreated> 
  	            <TimeModified>2007-12-15T11:32:53-08:00</TimeModified> 
  	            <EditSequence>1197747173</EditSequence> 
  	            <Name>Installation</Name> 
  	            <FullName>Installation</FullName> 
  	            <IsActive>true</IsActive> 
  	            <Sublevel>0</Sublevel> 
            - 	<SalesTaxCodeRef>
  		            <ListID>20000-999022286</ListID> 
  		            <FullName>Non</FullName> 
  	            </SalesTaxCodeRef>
            - 	<SalesOrPurchase>
  		            <Desc>Installation labor</Desc> 
  		            <Price>35.00</Price> 
            - 		<AccountRef>
  			            <ListID>190000-933270541</ListID> 
  			            <FullName>Construction Income:Labor Income</FullName> 
  		            </AccountRef>
  	            </SalesOrPurchase>
              </ItemServiceRet>
              </ItemQueryRs>
              </QBXMLMsgsRs>
              </QBXML>
            */
            string[] retVal = new string[count];
            System.IO.StringReader rdr = new System.IO.StringReader(xml);
            System.Xml.XPath.XPathDocument doc = new System.Xml.XPath.XPathDocument(rdr);
            System.Xml.XPath.XPathNavigator nav = doc.CreateNavigator();
            List<Dictionary<string, string>> result = new List<Dictionary<string, string>>();
            result.Add(new Dictionary<string, string>());

            if (nav != null)
            {
                nav.MoveToFirstChild();
            }
            bool more = true;
            int x = 0;
            while (more)
            {
                switch (nav.LocalName)
                {
                    case "ListID":
                        if (String.IsNullOrEmpty(retVal[x]))
                        {
                            retVal[x] = nav.Value.Trim();

                            //Initialize Dictionary
                            result[x]["ListID"] = nav.Value.Trim();
                            result[x]["FullName"] = null;
                            result[x]["ManufacturerPartNumber"] = null;
                            result[x]["QuantityOnHand"] = null;
                            result[x]["Price"] = null;
                        }
                        more = nav.MoveToNext();
                        continue;
                    case "QBXML":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "QBXMLMsgsRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "ItemInventoryQueryRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "ItemServiceRet":
                    case "ItemNonInventoryRet":
                    case "ItemOtherChargeRet":
                    case "ItemInventoryRet":
                    case "ItemInventoryAssemblyRet":
                    case "ItemFixedAssetRet":
                    case "ItemSubtotalRet":
                    case "ItemDiscountRet":
                    case "ItemPaymentRet":
                    case "ItemSalesTaxRet":
                    case "ItemSalesTaxGroupRet":
                    case "ItemGroupRet":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "FullName":
                        retVal[x] = retVal[x] + "<" + nav.Value.Trim();
                        result[x]["FullName"] = nav.Value.Trim();
                        more = nav.MoveToNext();

                        continue;
                    case "QuantityOnHand":
                        retVal[x] = retVal[x] + "<" + nav.Value.Trim();
                        result[x]["QuantityOnHand"] = nav.Value.Trim();


                        x++;

                        result.Add(new Dictionary<string, string>());

                        more = nav.MoveToParent();
                        more = nav.MoveToNext();
                        continue;
                    case "ManufacturerPartNumber":
                        retVal[x] = retVal[x] + "<" + nav.Value.Trim();
                        result[x]["ManufacturerPartNumber"] = nav.Value.Trim();
                        more = nav.MoveToNext();
                        continue;
                    case "SalesOrPurchase":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "Desc":
                    case "Price":
                        string val = nav.Value.Trim();
                        result[x]["Price"] = nav.Value.Trim();
                        decimal price = 0.0m;
                        if (IsDecimal(val))
                        {
                            price = Convert.ToDecimal(val);
                            if (exchangeRate != 1.0m) { price = price / exchangeRate; }
                            retVal[1] = price.ToString("N2");
                        }
                        else
                        {
                            retVal[0] = val;
                        }

                        more = nav.MoveToNext();
                        continue;
                    default:
                        more = nav.MoveToNext();
                        continue;
                }
            }
            return result;
        }

    }
}
