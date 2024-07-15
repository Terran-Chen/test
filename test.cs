using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Management;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Collections;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System.Drawing;
using System.Net.Mail;
using System.Data;
using System.Data.SqlClient;
using ViewSonic.WebLogIn;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium;
using System.Threading;
using System.DirectoryServices;
using System.Xml;
using System.Xml.XPath;
using System.Reflection;

namespace ConsoleApplication1
{
    class Program
    {
        private static bool CheckValidationResult(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors errors)
        {
            return true; //总是接受 
        }

        private static WebProxy m_proxy;
        public static WebProxy Proxy
        {
            get
            {
                return m_proxy;
            }
            set
            {
                m_proxy = value;
            }
        }

        private static string m_requestUrl;

        //public NetworkAdapter[] GetNetworkAdapters()
        //{
        //    ArrayList macs = new ArrayList();

        //    ManagementClass mc = new ManagementClass("Win32_NetworkAdapterConfiguration");
        //    ManagementObjectCollection moc = mc.GetInstances();
        //    foreach (ManagementObject mo in moc)
        //    {
        //        if (Convert.ToBoolean(mo["ipEnabled"]) == true)
        //        {
        //            string mac = mo["MacAddress"].ToString();
        //            string[] ips = (string[])mo["IPAddress"];

        //            NetworkAdapter networkAdapter = new NetworkAdapter();
        //            networkAdapter.MAC = mac;
        //            networkAdapter.IP = ips;

        //            macs.Add(networkAdapter);
        //        }
        //    }
        //    return (NetworkAdapter[])macs.ToArray(typeof(NetworkAdapter));
        //}

        static int GetSeed()
        {
            byte[] bytes = new byte[4];
            System.Security.Cryptography.RNGCryptoServiceProvider rng = new System.Security.Cryptography.RNGCryptoServiceProvider();
            rng.GetBytes(bytes);
            return BitConverter.ToInt32(bytes, 0);
        }


        static void login()
        {
            //ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;




















            //EdgeOptions options = new EdgeOptions();
            //options.BinaryLocation = @"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe";

            //EdgeDriver driver = new EdgeDriver(options);
            //driver.Navigate().GoToUrl("https://subscribers.servright.com/custLogin.php");

            //driver.FindElement(By.Name("username")).Clear();
            //driver.FindElement(By.Id("username")).SendKeys("dings");
            //driver.FindElement(By.Name("password")).Clear();
            //driver.FindElement(By.Name("password")).SendKeys("dings@123");
            //Thread.Sleep(1000);
            //driver.FindElement(By.Name("submit")).Click();

            //CookieContainer cookies = new CookieContainer();

            //foreach (var item in driver.Manage().Cookies.AllCookies)
            //{
            //    cookies.Add(new System.Net.Cookie { Name = item.Name, Value = item.Value, Path = item.Path, Domain = item.Domain, Expires = item.Expiry ?? DateTime.MaxValue, Secure = item.Secure });
            //}



























            HttpWebRequest request;
            HttpWebResponse response;
            Stream stream;
            StreamReader reader;

            CookieContainer cookieContainer = new CookieContainer();

            request = WebRequest.Create("https://subscribers.servright.com/custLogin.php") as HttpWebRequest;
            request.Method = "GET";
            request.AllowAutoRedirect = true;
            request.CookieContainer = cookieContainer;
            response = request.GetResponse() as HttpWebResponse;
            reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8);
            string html = reader.ReadToEnd();



            byte[] data = Encoding.UTF8.GetBytes(string.Format("dest=&ePassword=76a2574cfe82e6a3f4c812b73b2bffa1f02b4235&username=dings&password=dings%40123&submit=Login"));

            request = WebRequest.Create("https://subscribers.servright.com/custLogin.php") as HttpWebRequest;
            request.Expect = null;
            request.AllowAutoRedirect = true;
            request.Method = "POST";
            request.ContentType = "application/x-www-form-urlencoded";
            request.CookieContainer = cookieContainer;
            request.ContentLength = data.Length;
            stream = request.GetRequestStream();
            stream.Write(data, 0, data.Length);
            stream.Close();
            response = request.GetResponse() as HttpWebResponse;


            request = WebRequest.Create("https://subscribers.servright.com/oneCall.php?call_number=17190497") as HttpWebRequest;
            request.Method = "GET";
            request.CookieContainer = cookieContainer;

            response = request.GetResponse() as HttpWebResponse;
            reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8);
            html = reader.ReadToEnd();


            string asdf = Regex.Match(html, @"<b>Call Num:</b> \w+").Value;

        }


        static void outlook()
        {
            //Microsoft.Office.Interop.Outlook.Application app = new Microsoft.Office.Interop.Outlook.Application();
            //Microsoft.Office.Interop.Outlook.NameSpace mynamespace;
            //Microsoft.Office.Interop.Outlook.MAPIFolder myFolder;
            //OLFolder = OLNameS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);



            Microsoft.Office.Interop.Outlook.Application OL = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.AddressLists addrlist = OL.Session.AddressLists;
            Microsoft.Office.Interop.Outlook.AddressList alist = OL.Session.GetGlobalAddressList();

            foreach (Microsoft.Office.Interop.Outlook.AddressEntry item in alist.AddressEntries)
            {

                //if (string.Compare(item.Name, "gsc-is", true) == 0 || string.Compare(item.Name, "terran chen", true) == 0)
                if (((dynamic)item).Members != null && ((dynamic)item).Members.count > 0)
                {
                    Console.WriteLine(item.Name);
                    var bbb = ((dynamic)item).Members[1];
                }
            }



            Microsoft.Office.Interop.Outlook.AddressEntry entry = alist.AddressEntries["terran chen"];
            Microsoft.Office.Interop.Outlook.ContactItem contact = entry.GetContact();













            Microsoft.Office.Interop.Outlook.NameSpace OLNameS;
            Microsoft.Office.Interop.Outlook.MAPIFolder OLFolder;
            Microsoft.Office.Interop.Outlook.DistListItem OLDistListItem;//用于接收组信息
            OLNameS = OL.GetNamespace("MAPI");
            OLFolder = OLNameS.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderContacts);
            //var aa = OLFolder.Items.GetFirst();
            int contactItemCout = OLFolder.Items.Count;

            var aa = OLFolder.Name;

            foreach (Microsoft.Office.Interop.Outlook.ContactItem item in OLFolder.Items)
            {
                Console.WriteLine(string.Format("{0} {1} {2}", item.FirstName, item.LastName, item.Email1Address));
            }


        }


        public static object EvaluateConstString(string strExpress)
        {
            if (string.IsNullOrEmpty(strExpress))
                return null;
            strExpress = strExpress.Replace("/", " div ");
            XmlDocument doc = new XmlDocument();
            XPathNavigator navigator = doc.CreateNavigator();

            return navigator.Evaluate(strExpress);
        }
        static void Main(string[] args)
        {


            Console.WriteLine(Convert.ToBoolean(EvaluateConstString(" contains(',110,', ',110,')")));





            DateTime date111 = new DateTime(2024, 6, 1);
            DateTime date211 = new DateTime(2024, 12, 31);



            Console.Write((int)3.9);


            if (date111.AddYears(3) < new DateTime(2024, 6, 1))
            {

            }
            else if (date111.AddYears(3) <= new DateTime(2024, 12, 31))
            {

            }





            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

            HttpWebRequest request = WebRequest.CreateHttp("https://viewsonic-fs-sandbox.freshservice.com/api/v2/tickets/50");
            request.Method = "GET";
            request.ContentType = "application/json";
            request.Headers["Authorization"] = "Basic " + Convert.ToBase64String(Encoding.Default.GetBytes("ybeePbt84yqZD2Ai6b83:X"));

            HttpWebResponse response = request.GetResponse() as HttpWebResponse;
            StreamReader contentReader = new StreamReader(response.GetResponseStream());
            string aaaa = contentReader.ReadToEnd();

            Console.WriteLine(aaaa);




















            Regex reg = new Regex("(https://vsi-ap01.viewsonic.com.tw/ECRECN/)[f|w]/");

            reg.Match("https://vsi-ap01.viewsonic.com.tw/ECRECN/w/WebUI/ECN/ECNView.aspx?a2V5PWJjZTZlN2Y0LTZiMjYtNGIxMi1iNDdmLTg3Y2M5M2ZiYWI1Mw%3d%3d");























            Console.WriteLine(System.Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile("1", "MD5"));





            ViewSonic.Helper.MailHelper.tbl_MailQueue mail = new ViewSonic.Helper.MailHelper.tbl_MailQueue
            {
                mailTO = "terran.chen@viewsonic.com",
                Subject = "test",
                Contect = "test11test11test11",
            };

            ViewSonic.Helper.MailHelper.MailQueue_Add(mail);







            List<decimal?> listddd = new List<decimal?>();
            listddd.Add(10);
            listddd.Add(new Nullable<decimal>());
            listddd.Add(20);

            Console.WriteLine(listddd.Sum(o => o));

            //outlook();



            //FileInfo file = new FileInfo("D:\\Projects\\Study\\ConsoleApplication1\\ConsoleApplication1\\bin\\Debug\\DLL20230704.txt");

            //FtpWebRequest ftp_request = (FtpWebRequest)FtpWebRequest.Create(string.Format("ftp://121.40.141.12/{0}", "DLL20230704.txt"));

            //ftp_request.Credentials = new NetworkCredential("wmsyp", "680GH9835Ry");
            //ftp_request.KeepAlive = false;
            //ftp_request.UsePassive = "0" == "1";
            //ftp_request.Method = WebRequestMethods.Ftp.UploadFile;
            //ftp_request.UseBinary = true;
            //ftp_request.ContentLength = file.Length;
            //int buffLength = 2048;
            //byte[] buff = new byte[buffLength];
            //int contentLen;
            //FileStream fs = file.OpenRead();

            //string err_msg = "Success";
            //try
            //{
            //    using (Stream strm = ftp_request.GetRequestStream())
            //    {
            //        contentLen = fs.Read(buff, 0, buffLength);
            //        while (contentLen != 0)
            //        {
            //            strm.Write(buff, 0, contentLen);
            //            contentLen = fs.Read(buff, 0, buffLength);
            //        }
            //        strm.Close();

            //    }
            //}
            //catch { }
































            //connection_string_service.Connection_StringSoapClient abc = new connection_string_service.Connection_StringSoapClient();
            //string str111 = abc.GetConnection_String_Cache("Oracle_ReadOnly");
            //return;

            //test();
            //return;



            //login();


            StringBuilder sb = new StringBuilder();
            string chars = "0123456789ABCDEFGHJKLMNOPQRSTUVWXYZabcdefghjkmnopqrstuvwxyz";

            for (int i = 0; i < 21; i++)
            {
                sb = new StringBuilder();

                for (int j = 0; j < 8; j++)
                {
                    Random rd = new Random(GetSeed());
                    sb.Append(chars[rd.Next(0, 62)].ToString());
                }

                Console.Write(sb.ToString());

                Console.Write("||");

                Console.Write(System.Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile(sb.ToString(), "sha1"));

                Console.WriteLine(System.Environment.NewLine);
            }





















            ICellStyle style;









            IWorkbook book111;
            ISheet sheet111;
            using (FileStream fs3 = new FileStream(@"d:\test\888.xlsx", FileMode.Open, FileAccess.Read))
            {
                book111 = new XSSFWorkbook(fs3);
            }

            sheet111 = book111.GetSheetAt(0);
            Console.WriteLine(sheet111.LastRowNum);




            DateTime date = Convert.ToDateTime("2022/Aug/01");

            //7E5761391CF1DDDCAA149A179061963EAFE0062F

            Console.WriteLine(HashEncrypt.EncryptSha1("Xux33393"));

            Console.WriteLine(Path.GetExtension("adfasdfasdf.aaa"));



            DateTime ddd = new DateTime(2022, 8, 29);

            int dddddd = (ddd.Subtract(new DateTime(2022, 1, 3)).Days + 1) / 7 + 1;





            string sql = @"declare @formID varchar(20),@caseno int 
set @caseno= 13423

select @formID=formid from tblPAF where caseno=@caseno 

select a.*,region,producttype,name from ad_user a 
join ad_user_group b on a.userid=b.userid 
left outer join tblUser_ProductType c on b.userid=c.userid and b.groupid=c.groupid 
left outer join tblUser_Region d on b.userid=d.userid and b.groupid=d.groupid left outer join ad_group e on b.groupid=e.groupid 
where name in ('Region_PM','Region_PM_Head') and 
producttype in (select product_type from tblPAF where formid=@formID) and 
d.region in ('VSE') order by name ";

            DataTable dt = new DataTable();
            SqlDataAdapter aa = new SqlDataAdapter(sql, @"Data Source=vstw-ch-db02.vsi.viewsonic.com;Initial Catalog=ePAF;Persist Security Info=True;User ID=ePAF_Writer;Password=ePAF_Writer;max pool size=200;");

            int iiii = aa.Fill(dt);
























            string a1 = "VA2406-H#ABMPPJS0100083_ACMPBXXXX00039V10000#V2.rar";
            bool bbb = Regex.IsMatch(a1, "(?i)(VA2406-H|sdfasdf|cbsertrewqt)(#?\\w*)(#V[0-9.]+(.zip|.rar))$");























            Image im = Image.FromFile("20220209181515.pdf");
            int towidth = 100;
            int toheight = 100;
            int x = 0, y = 0;
            int ow = im.Width;
            int oh = im.Height;























































            Match mm = Regex.Match("回复:  - Request for Approval.  Application #:OBO-ADF-2206-0002-GLOBAL", "(?i)(program|claim)");


            string aaaaa = Regex.Replace("Program - Request for Approval. Application #: OBO-ADF-2205-0002", ".*#: ", "");





            IWorkbook book11, book_new;
            ISheet sheet_new;
            using (FileStream fs2 = new FileStream("other charges  -US.xlsx", FileMode.Open, FileAccess.Read))
            {
                book11 = new XSSFWorkbook(fs2);
            }
            book_new = new XSSFWorkbook();
            sheet_new = book_new.CreateSheet("aa");

            ISheet sheet = book11.GetSheetAt(2);
            IEnumerator rows = sheet.GetRowEnumerator();

            int index = 0;

            rows.MoveNext();
            rows.MoveNext();
            while (rows.MoveNext())
            {
                IRow row = rows.Current as IRow;
                if (row.GetCell(13).CellComment != null)
                {
                    string comment = row.GetCell(13).CellComment.String.String;

                    var list = Regex.Matches(comment, @"(?i)(Lift Gate|Inside Delivery)[ \w]* \$*[\d.]+");

                    if (list.Count > 0)
                    {
                        IRow row_new = sheet_new.CreateRow(index++);

                        row_new.CreateCell(0).SetCellValue(row.GetCell(7).ToString());
                        for (int i = 0; i < list.Count; i++)
                        {
                            row_new.CreateCell(1 + i * 2).SetCellValue(list[i].ToString());

                            row_new.CreateCell(2 + i * 2).SetCellValue(Regex.Matches(list[i].ToString(), @"[\d.]+$")[0].ToString());



                        }
                    }


                }
            }

            using (FileStream fs1 = new FileStream("other charges  -US_new.xlsx", FileMode.Create, FileAccess.Write))
            {
                book_new.Write(fs1);
            }





            //[[a-z|A-Z| |\(|\)]+\s+([\d.]+)]

            //var list = Regex.Matches(str1, @"(?i)(Lift Gate|Inside Delivery)[ \w]* \$*[\d.]+");
















            List<string> llll = new List<string> { "", "1", "2", null };
            List<string> lll1 = llll.Where(o => o != "").ToList();















































































            //ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;



            Console.WriteLine(System.Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile("", "sha1"));


















            //HttpWebRequest request = WebRequest.Create("https://api.bamboohr.com/api/gateway.php/viewsonic/v1/reports/90?format=csv") as HttpWebRequest;
            //request.Method = "GET";
            //request.ContentType = ("application/json");
            //request.Headers.Add("Authorization", "Basic ODBmOTk2M2YwMjZlMjlkMGJkNzJlMWI0NzE5NTk3MGQ0MzM4ZDI0Yzp4");
            //string postData = "{\"filters\":{\"status\":{\"value\":\"active\"}}\"fields\":[\"firstName\",\"lastName\",\"workEmail\",\"supervisor\"],\"title\":\"IS_Users\"}";


            ////using (var requestStream = request.GetRequestStream())
            ////{
            ////    var bytes = Encoding.UTF8.GetBytes(postData);
            ////    requestStream.Write(bytes, 0, bytes.Length);
            ////}
            //HttpWebResponse response = request.GetResponse() as HttpWebResponse;

            //StreamReader contentReader = new StreamReader(response.GetResponseStream());
            //string ret = contentReader.ReadToEnd();

            //File.WriteAllText("90.txt", ret);




            //HttpWebRequest request = WebRequest.Create("https://api.bamboohr.com/api/gateway.php/viewsonic/v1/reports/custom?onlyCurrent=true&format=csv") as HttpWebRequest;
            //request.Method = "POST";
            //request.ContentType = ("application/json");
            //request.Headers.Add("Authorization", "Basic ODBmOTk2M2YwMjZlMjlkMGJkNzJlMWI0NzE5NTk3MGQ0MzM4ZDI0Yzp4");
            //string postData = "{\"filters\":{\"status\":{\"value\":\"Active\"}},\"fields\":[\"firstName\",\"lastName\",\"workEmail\",\"supervisor\",\"status\"],\"title\":\"IS_Users\"}";


            //using (var requestStream = request.GetRequestStream())
            //{
            //    var bytes = Encoding.UTF8.GetBytes(postData);
            //    requestStream.Write(bytes, 0, bytes.Length);
            //}
            //HttpWebResponse response = request.GetResponse() as HttpWebResponse;

            //StreamReader contentReader = new StreamReader(response.GetResponseStream());
            //string ret = contentReader.ReadToEnd();




            //for (int i = 2000; i < 5000; i++)
            //{
            //    try
            //    {
            //        HttpWebRequest request = WebRequest.Create("https://api.bamboohr.com/api/gateway.php/viewsonic/v1/reports/" + i + "?format=csv") as HttpWebRequest;
            //        request.Method = "GET";
            //        request.ContentType = ("application/json");
            //        request.Headers.Add("Authorization", "Basic ODBmOTk2M2YwMjZlMjlkMGJkNzJlMWI0NzE5NTk3MGQ0MzM4ZDI0Yzp4");
            //        string postData = "{\"filters\":{\"status\":{\"value\":\"active\"}}\"fields\":[\"firstName\",\"lastName\",\"workEmail\",\"supervisor\"],\"title\":\"IS_Users\"}";


            //        //using (var requestStream = request.GetRequestStream())
            //        //{
            //        //    var bytes = Encoding.UTF8.GetBytes(postData);
            //        //    requestStream.Write(bytes, 0, bytes.Length);
            //        //}
            //        HttpWebResponse response = request.GetResponse() as HttpWebResponse;

            //        StreamReader contentReader = new StreamReader(response.GetResponseStream());
            //        string ret = contentReader.ReadToEnd();

            //        File.WriteAllText(i + ".txt", ret);
            //    }
            //    catch
            //    { }
            //}

            //Console.Read();


























            #region 11


            string companySubdomain = "viewsonic";    // Your company's subdomain here
            string userSecretKey = "80f9963f026e29d0bd72e1b47195970d4338d24c";        // Your user's api key here

            BambooAPIClient cl = new BambooAPIClient(companySubdomain);
            cl.setSecretKey(userSecretKey);

            string[] fields = {
                "id",
                "employeeNumber",
                "lastName",
                "firstName",
                "jobTitle",
                "hireDate",
                "department",
                "division",
                "supervisor",
                "supervisorId",
                "workemail",
                "location",
                "city",
                "status",
                "terminationDate",
                "gender",

            };

            //string[] fields = { 
            //    "employeeNumber",
            //    "lastName",
            //    "middleName",
            //    "jobTitle",
            //                  };

            int employeeId = 27;

            BambooHTTPResponse resp;

            Console.WriteLine("--- Custom Report ---");
            //resp = cl.getEmployeesReport("csv", "Custom report", fields);

            resp = cl.getReportByID("json", "90");

            //File.WriteAllText("aa.txt", resp.getContentString());

            Model obj = Newtonsoft.Json.JsonConvert.DeserializeObject<Model>(resp.getContentString());


            //var obj = Newtonsoft.Json.JsonConvert.DeserializeObject<Employee>(resp.getContentString());

            Console.WriteLine(resp.getContentString());


            #endregion

            //Console.WriteLine("--- Get time off requests ---");
            //Hashtable filters = new Hashtable();
            //filters.Add("start", "2012-01-01");
            //filters.Add("end", "2012-01-31");
            //filters.Add("status", "approved");

            //resp = cl.getTimeOffRequests(filters);
            //Console.WriteLine(resp.getContentString());

            //Console.WriteLine("--- Update employee ---");
            //Hashtable values = new Hashtable();
            //values.Add("firstName", "VB");
            //values.Add("lastName", "User 2");
            //values.Add("status", "active");

            //resp = cl.updateEmployee(41290, values);
            //Console.WriteLine(resp.responseCode);
            //Console.WriteLine(resp.getContentString());


            //return;






















            string bb1 = ViewSonic.WebLogIn.DESEncrypt.EncryptByConfigKey("VSct98877");
            string bb2 = ViewSonic.WebLogIn.DESEncrypt.EncryptByConfigKey("vscn@2004@vsi");
            string bb3 = ViewSonic.WebLogIn.DESEncrypt.EncryptByConfigKey("R0llerC04ster");
            string bb4 = ViewSonic.WebLogIn.DESEncrypt.EncryptByConfigKey("VSct98877");
            string bb5 = ViewSonic.WebLogIn.DESEncrypt.EncryptByConfigKey("vsiAPP");





            IWorkbook book;
            using (FileStream fs1 = new FileStream(@"d:\ESQUIRE SYSTEM TECHNOLOGY (PTY) LTD11111.xls", FileMode.Open, FileAccess.Read))
            {
                book = new HSSFWorkbook(fs1);
            }




            Convert.ToDateTime("20210101");


            DateTime date1;
            if (!DateTime.TryParse("2020/8/7", out date1))
            {
                date1 = DateTime.Now;
            }




            if (true)
            {
                Console.WriteLine("aaaaaaaaaaaa");
            }
            else
            {


































                Task.GetAllTasks();


                //NetworkAdapter[] networkAdapters = (new GetNetworkAdapter()).GetNetworkAdapters();




                string url = "https://adf.viewsonic.com/MIS/ADFSystemUI/Program/MISProgramViewEmail.aspx?a2V5PTU3MjkmSXNUb0FwcGxpY2FudD0wJkVtYWlsPTEmaXNTaG93Q29zdD0w";

                Login();
                string rtn = GetUrlResponseContent(url);
                Logoff();
                System.GC.Collect();
                return;

















                //string user = ConfigurationManager.AppSettings["DomainUserId"];
                //HttpWebRequest request;

                //if (url.StartsWith("https", StringComparison.OrdinalIgnoreCase))
                //{
                //    ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(CheckValidationResult);
                //    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Ssl3;
                //    request = WebRequest.Create(url) as HttpWebRequest;
                //    request.ProtocolVersion = HttpVersion.Version10;
                //}
                //else
                //{
                //    request = WebRequest.Create(url) as HttpWebRequest;
                //}

                ////AD认证
                //    string passWord = "0EDSFFYHFyxDjlGnuoV+ww==";
                //    passWord = ViewSonic.WebLogIn.DESEncrypt.DecryptByConfigKey(passWord);
                //    string[] ary = "vsa\\vsap".Split('\\');
                //    NetworkCredential credential = new NetworkCredential(ary[1], passWord, ary[0]);
                //    CredentialCache cache = new CredentialCache();
                //    cache.Add(request.RequestUri, "NTLM", credential);
                //    request.Credentials = cache;

                //request.Proxy = m_proxy;
                //request.Method = "GET";

                //request.AllowAutoRedirect = false;
                ////request_login.Timeout = 300000;
                //WebResponse response_login = request.GetResponse();





                DateTime time = DateTime.Now;

                int temp1, temp2;

                string str;

                time = DateTime.Now;
                for (int i = 0; i < 10000; i++)
                {
                    temp1 = i;
                }

                Console.WriteLine((DateTime.Now - time).TotalSeconds);




                time = DateTime.Now;
                for (int i = 0; i < 10000; i++)
                {
                    str = i.ToString();
                }
                Console.WriteLine((DateTime.Now - time).TotalSeconds);







                Console.ReadKey();
            }
        }

        private static void test()
        {
            DirectoryEntry userEntry2 = new DirectoryEntry(ConfigurationManager.AppSettings["ldap"]);
            DirectorySearcher rootSearcher = new DirectorySearcher(userEntry2);
            rootSearcher.PropertiesToLoad.Add("mailnickname");
            rootSearcher.PropertiesToLoad.Add("samAccountName");
            rootSearcher.PropertiesToLoad.Add("name");                      // Full name…
            rootSearcher.PropertiesToLoad.Add("mail");                        // Primary email addy…
            rootSearcher.PropertiesToLoad.Add("telephoneNumber");  // Phone #...
            rootSearcher.PropertiesToLoad.Add("displayName");
            rootSearcher.PropertiesToLoad.Add("title");
            rootSearcher.PropertiesToLoad.Add("TP");
            rootSearcher.PropertiesToLoad.Add("description");
            rootSearcher.PropertiesToLoad.Add("company");
            rootSearcher.PropertiesToLoad.Add("sn");
            rootSearcher.PropertiesToLoad.Add("givenName");
            rootSearcher.PropertiesToLoad.Add("mobile");
            rootSearcher.PropertiesToLoad.Add("pager");
            rootSearcher.PropertiesToLoad.Add("streetaddress");
            rootSearcher.PropertiesToLoad.Add("physicalDeliveryOfficeName");
            rootSearcher.PropertiesToLoad.Add("facsimileTelephoneNumber");
            rootSearcher.PropertiesToLoad.Add("co");
            rootSearcher.PropertiesToLoad.Add("department");
            rootSearcher.PropertiesToLoad.Add("postalCode");
            rootSearcher.PropertiesToLoad.Add("l");
            rootSearcher.PropertiesToLoad.Add("c");
            rootSearcher.PropertiesToLoad.Add("st");
            rootSearcher.PropertiesToLoad.Add("pwdLastSet");
            rootSearcher.PropertiesToLoad.Add("userAccountControl");
            rootSearcher.Filter = "(&(objectClass=user))";
            rootSearcher.PageSize = 10000;
            SearchResultCollection searchResults = null;
            searchResults = rootSearcher.FindAll();
            if (searchResults != null)
            {
                Console.WriteLine(searchResults.Count);
            }
        }

        private static List<string> m_loginAuthCookie = new List<string>();
        private static string ServerHttpHead = "https://adf.viewsonic.com/MIS";

        public static void Login()
        {
            System.Net.ServicePointManager.DefaultConnectionLimit = 50;
            m_loginAuthCookie.Clear();
            string head = ServerHttpHead.EndsWith("/") ? ServerHttpHead : ServerHttpHead + "/";
            string indexUrl = head + "index.aspx";
            HttpWebRequest request_login = BuildRequest(indexUrl);
            request_login.Method = "GET";
            request_login.AllowAutoRedirect = false;
            //request_login.Timeout = 300000;
            WebResponse response_login = request_login.GetResponse();

            if (response_login != null)
                response_login.Close();
            if (request_login != null)
                request_login.Abort();
            return;



        }

        protected static HttpWebRequest BuildRequest(string url)
        {
            string user = "vsa\\vsap";
            HttpWebRequest request;

            if (url.StartsWith("https", StringComparison.OrdinalIgnoreCase))
            {
                ServicePointManager.ServerCertificateValidationCallback = new RemoteCertificateValidationCallback(CheckValidationResult);
                //ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls | SecurityProtocolType.Ssl3;
                request = WebRequest.Create(url) as HttpWebRequest;
                request.ProtocolVersion = HttpVersion.Version10;
            }
            else
            {
                request = WebRequest.Create(url) as HttpWebRequest;
            }

            //AD??
            string passWord = "0EDSFFYHFyxDjlGnuoV+ww==";
            passWord = ViewSonic.WebLogIn.DESEncrypt.DecryptByConfigKey(passWord);
            string[] ary = user.Split('\\');
            NetworkCredential credential = new NetworkCredential(ary[1], passWord, ary[0]);
            CredentialCache cache = new CredentialCache();
            cache.Add(request.RequestUri, "NTLM", credential);
            request.Credentials = cache;

            request.Proxy = m_proxy;
            request.Method = "GET";
            return request;
        }

        public static string GetUrlResponseContent(string url)
        {
            m_requestUrl = url;
            HttpWebRequest request_Real = BuildRequest(url);
            foreach (string setCookie in m_loginAuthCookie)
            {
                string[] array = setCookie.Split(';')[0].Split('=');
                if (request_Real.CookieContainer == null)
                {
                    request_Real.CookieContainer = new CookieContainer();
                }
                //request_Real.CookieContainer.Add(new Cookie(array[0], array[1], "/", request_Real.RequestUri.Host));
            }
            try
            {
                WebResponse response_real = request_Real.GetResponse();
                using (Stream stream = response_real.GetResponseStream())
                {
                    Encoding encode = Encoding.UTF8;
                    if (encode == null)
                    {
                        encode = Encoding.Default;
                    }
                    string content;
                    using (StreamReader sr = new StreamReader(stream, encode))
                    {
                        content = sr.ReadToEnd();
                    }
                    //??????
                    Regex hrefReg1 = new Regex("<[^<]*?href=\"(.+?)\"[^>]*?>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                    MatchEvaluator evaluator1 = new MatchEvaluator(ReplaceHrefMatch);
                    if (response_real != null)
                        response_real.Close();
                    if (request_Real != null)
                        request_Real.Abort();
                    return hrefReg1.Replace(content, evaluator1);
                }
            }
            catch (Exception ee)
            {
                throw new Exception("Url:" + url, ee);
            }
        }

        private static string ReplaceHrefMatch(Match m)
        {
            string releateUrl = m.Groups[1].Value;
            string newUrl = CombinUrlPath(m_requestUrl, releateUrl);
            return m.Value.Replace(releateUrl, newUrl);
        }

        private static string CombinUrlPath(string fullUrlPath, string relatePath)
        {
            if (relatePath.IndexOf("://") > 0)
                return relatePath;
            Uri uri = new Uri(fullUrlPath);
            string dirPath;
            if (uri.Query.Length > 0)
                dirPath = uri.AbsoluteUri.Substring(0, uri.AbsoluteUri.Length - uri.Query.Length);
            else
                dirPath = uri.AbsoluteUri;
            dirPath = dirPath.Substring(0, dirPath.LastIndexOf('/') + 1);
            return new Uri(dirPath + relatePath).AbsoluteUri;
        }

        public static void Logoff()
        {
            m_loginAuthCookie.Clear();
            return;
        }
    }


    class Model
    {
        public string title { get; set; }
        public List<field> fields { get; set; }
        public List<employee> employees { get; set; }
    }

    class field
    {
        public string id { get; set; }
        public string type { get; set; }
        public string name { get; set; }
    }

    class employee
    {
        public string employeeNumber { get; set; }
        public string firstName { get; set; }
        public string lastName { get; set; }
        public string employmentHistoryStatus { get; set; }
        public string customFunctionCode { get; set; }
        public string jobTitle { get; set; }
        public string department { get; set; }
        public string division { get; set; }
        [JsonProperty("customReportsToEmp#")]
        public string customReportsToEmp { get; set; }
        public string workEmail { get; set; }
        public string workPhone { get; set; }
        public string workPhoneExtension { get; set; }
        public string hireDate { get; set; }
        public string gender { get; set; }
        public string customCountryDescription { get; set; }
        public string customCostCenterCode { get; set; }
        public string customCostCenterDescription { get; set; }
        public string customLegalEntityCode { get; set; }
    }







    //public class TiffImage
    //{
    //    private string myPath;
    //    private Guid myGuid;
    //    private FrameDimension myDimension;
    //    public ArrayList myImages = new ArrayList();
    //    private int myPageCount;
    //    private Bitmap myBMP;

    //    public TiffImage(string path)
    //    {
    //        MemoryStream ms;
    //        Image myImage;

    //        myPath = path;
    //        FileStream fs = new FileStream(myPath, FileMode.Open);
    //        myImage = Image.FromStream(fs);
    //        myGuid = myImage.FrameDimensionsList[0];
    //        myDimension = new FrameDimension(myGuid);
    //        myPageCount = myImage.GetFrameCount(myDimension);
    //        for (int i = 0; i < myPageCount; i++)
    //        {
    //            ms = new MemoryStream();
    //            myImage.SelectActiveFrame(myDimension, i);
    //            myImage.Save(ms, ImageFormat.Bmp);
    //            myBMP = new Bitmap(ms);
    //            myImages.Add(myBMP);
    //            ms.Close();
    //        }
    //        fs.Close();
    //    }
    //}

}
