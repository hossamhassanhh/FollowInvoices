using ClosedXML.Excel;
using Microsoft.Extensions.Options;
using FollowInvoices.Models;
using FollowInvoices.Utilities;
using System.Data.SqlClient;
using System.Diagnostics;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Hosting.Internal;
using System.IO;
using FollowInvoices.Data;
using DocumentFormat.OpenXml.InkML;
using Azure.Core;
using DocumentFormat.OpenXml.VariantTypes;
using DocumentFormat.OpenXml.EMMA;
using Microsoft.AspNetCore.Authorization;
using DocumentFormat.OpenXml.Office2010.Excel;
using System.Numerics;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using System.Reflection.Metadata;
using static System.Runtime.InteropServices.JavaScript.JSType;
using Microsoft.IdentityModel.Tokens;

namespace FollowInvoices.Controllers
{
    [Authorize]
    public class HomeController : Controller
    {
        //private readonly string _connectionString = "Data Source=(local);Initial Catalog=FollowInvoices;Integrated Security=True;";
        //private readonly string _connectionString = "Data Source=localhost\\SQLEXPRESS;Initial Catalog=FollowInvoices;Integrated Security=True;";
        private readonly ILogger<HomeController> _logger;
        private readonly IOptions<AppSettings> _options;
        private readonly IWebHostEnvironment _hostingEnvironment;
        private readonly IConfiguration _configuration;
        private readonly string _connectionString;
        public HomeController(IConfiguration configuration, ILogger<HomeController> logger, IOptions<AppSettings> options, IWebHostEnvironment hostingEnvironment)
        {
            _configuration = configuration;
            _connectionString = _configuration.GetConnectionString("DefaultConnection");
            _logger = logger;
            _options = options;
            _hostingEnvironment = hostingEnvironment;
        }

        //[ResponseCache(Location = ResponseCacheLocation.None, NoStore = true)]
        //[Authorize]
        public IActionResult Index()
        {
            return View();
        }
        public IActionResult AddVendor()
        {
            var vendors = new List<Vendor>();
            using (SqlConnection con = new SqlConnection(_connectionString))
            {
                con.Open();
                using (SqlCommand cmd = new SqlCommand("SELECT Id, VendorName FROM Vendors ORDER BY Id DESC", con))
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        vendors.Add(new Vendor
                        {
                            Id = reader.GetInt32(0),
                            VendorName = reader.GetString(1)
                        });
                    }
                }
            }

            return View(vendors);
        }
    

    public class Vendor
    {
        public int Id { get; set; }
        public string VendorName { get; set; }
}
        [HttpGet]
        public IActionResult ExecuteAddVendor(string vendor_name)
        {
            
            int newId;
            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                connection.Open(); // Open the connection here

                // Insert into InvoiceDetails
                string insertQueryIntoInvoiceDetails =
                    @"INSERT INTO Vendors (VendorName)
                    OUTPUT INSERTED.Id -- Adjust 'Id' to your actual primary key column
                    VALUES (@VendorName);";
                //SELECT SCOPE_IDENTITY()
                using (SqlCommand command = new SqlCommand(insertQueryIntoInvoiceDetails, connection))
                {
                    command.Parameters.AddWithValue("@VendorName", vendor_name);


                    try
                    {
                        newId = (int)command.ExecuteScalar(); // This will return the new Id.
                    }
                    catch (Exception /*ex*/)
                    {
                        //ViewBag.ErrorMessage = ex.Message;
                        ViewBag.ErrorMessage = "اسم المورد موجود بالفعل";
                        return View("Failed");
                    }
                }
            }

            return View("AddVendorDoneNotificationView");
        }
        public IActionResult Create()
        {
            DateTime currentDateOnly = DateTime.Now.Date;
            ViewBag.IsoDate = currentDateOnly.ToString("yyyy-MM-dd");
            
            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                string getIso = @"SELECT MAX(IsoNumber) FROM InvoiceDetails";

                // Execute your SQL query using the connection
                using (SqlCommand command = new SqlCommand(getIso, connection))
                {
                    object result = command.ExecuteScalar();

                    // Check if a valid number was returned
                    if (result != DBNull.Value && result != null)
                    {
                        int maxIsoNumber = Convert.ToInt32(result);
                        int nextIsoNumber = maxIsoNumber + 1;

                        // Pass the next IsoNumber to the view using ViewBag
                        ViewBag.NextIsoNumber = nextIsoNumber;
                    }
                    else
                    {
                        ViewBag.NextIsoNumber = 1;
                    }
                }
            }

            return View();
        }
        public IActionResult CreateOnSameIso()
        {
            DateTime currentDateOnly = DateTime.Now.Date;
            ViewBag.IsoDate = currentDateOnly.ToString("yyyy-MM-dd");
            
            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                string getIso = @"SELECT MAX(IsoNumber) FROM InvoiceDetails";

                // Execute your SQL query using the connection
                using (SqlCommand command = new SqlCommand(getIso, connection))
                {
                    object result = command.ExecuteScalar();

                    // Check if a valid number was returned
                    if (result != DBNull.Value && result != null)
                    {
                        int maxIsoNumber = Convert.ToInt32(result);
                        int nextIsoNumber = maxIsoNumber;

                        // Pass the next IsoNumber to the view using ViewBag
                        ViewBag.NextIsoNumber = nextIsoNumber;
                    }
                    else
                    {
                        ViewBag.NextIsoNumber = 1;
                    }
                }
            }

            return View("Create");
        }

        [HttpGet]
        public IActionResult ExecuteCreate(string userName, string isoNumber, string isoDate, string invoiceNumber, 
            string invoiceValue1, string invoiceValue2, string invoiceValue3, string currency1, string currency2,
            string currency3, string invoiceDate, string invoiceReceiptDate, string requester,
            string vendorName, string sendToRequesterDate, string area)
        {
            //string username = User.Identity.Name;
            
            IsoDetails isoDetails = new IsoDetails();
            //DateTime isoDate = DateTime.Today;
            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                connection.Open(); // Open the connection here

                //// Check if the ID already exists
                //string getId = @"SELECT IsoNumber FROM InvoiceDetails WHERE IsoNumber = @Id";
                //using (SqlCommand command = new SqlCommand(getId, connection))
                //{
                //    command.Parameters.AddWithValue("@Id", isoNumber);

                //    using (SqlDataReader reader = command.ExecuteReader())
                //    {
                //        if (reader.HasRows)
                //        {
                //            ViewBag.ErrorMessage = "!رقم المتابعة المدخل مكرر";
                //            return View("Failed");
                //        }
                //    }
                //}

                // Insert into InvoiceDetails
                string insertQueryIntoInvoiceDetails = @"INSERT INTO InvoiceDetails (IsoNumber, IsoDate, InvoiceNumber, 
                    InvoiceValue1, InvoiceValue2, InvoiceValue3, Currency1, Currency2, Currency3, InvoiceDate,
                    InvoiceReceiptDate, Requester, VendorName, UserName, SendToRequesterDate, Area, InvoiceEntree)
                    VALUES (@IsoNumber, @IsoDate, @InvoiceNumber, @InvoiceValue1, @InvoiceValue2, @InvoiceValue3, @Currency1, 
                    @Currency2, @Currency3, @InvoiceDate, @InvoiceReceiptDate, @Requester, 
                    @VendorName, @UserName, @SendToRequesterDate, @Area, @InvoiceEntree);";
                //SELECT SCOPE_IDENTITY()
                using (SqlCommand command = new SqlCommand(insertQueryIntoInvoiceDetails, connection))
                {
                    command.Parameters.AddWithValue("@UserName", userName);
                    command.Parameters.AddWithValue("@IsoNumber", Convert.ToInt64(isoNumber));
                    command.Parameters.AddWithValue("@IsoDate", isoDate);
                    command.Parameters.AddWithValue("@InvoiceNumber", invoiceNumber);
                    command.Parameters.AddWithValue("@InvoiceValue1", Convert.ToDecimal(invoiceValue1));
                    command.Parameters.AddWithValue("@InvoiceValue2",
                        string.IsNullOrWhiteSpace(invoiceValue2) || invoiceValue2 == "null"
                            ? (object)DBNull.Value
                            : Convert.ToDecimal(invoiceValue2));

                    command.Parameters.AddWithValue("@InvoiceValue3",
                        string.IsNullOrWhiteSpace(invoiceValue3) || invoiceValue3 == "null"
                            ? (object)DBNull.Value
                            : Convert.ToDecimal(invoiceValue3));
                    command.Parameters.AddWithValue("@Currency1", currency1);
                    command.Parameters.AddWithValue("@Currency2", currency2 ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@Currency3", currency3 ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@InvoiceDate", invoiceDate);
                    command.Parameters.AddWithValue("@InvoiceReceiptDate", invoiceReceiptDate);
                    command.Parameters.AddWithValue("@Requester", requester ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@VendorName", vendorName);
                    command.Parameters.AddWithValue("@SendToRequesterDate", sendToRequesterDate ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@Area", area ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@InvoiceEntree", userName ?? (object)DBNull.Value);

                    try
                    {

                        command.ExecuteNonQuery();
                        //isoNumber = command.ExecuteScalar().ToString();
                        //object result = command.ExecuteScalar();
                        //if (result != null)
                        //{
                        //    isoNumber = result.ToString();
                        //    isoDetails.ID = isoNumber;
                        //}

                    }
                    catch (Exception ex)
                    {
                        //ViewBag.ErrorMessage = ex.Message;
                        ViewBag.ErrorMessage = "يرجي التأكد من ادخال كل البيانات بطريقة صحيحة";
                        return View("Failed");
                    }
                }
                //string getIso = @"SELECT IsoNumber FROM InvoiceDetails WHERE Id = @Id";
                //using (SqlCommand command = new SqlCommand(getId, connection))
                //{
                //    command.Parameters.AddWithValue("@Id", id);

                //    using (SqlDataReader reader = command.ExecuteReader())
                //    {
                //        isoDetails.ID = reader["IsoNumber"]?.ToString();
                //        if (reader.HasRows)
                //        {
                //            ViewBag.ErrorMessage = "يرجي التأكد من ادخال كل البيانات بطريقة صحيحة";
                //            return View("Failed");
                //        }
                //    }
                //}
                
                
        //        // Insert into History
        //        string insertQueryIntoHistory = @"
        //    INSERT INTO History (Id, InvoiceStatus, Dat)
        //    VALUES (@InvoiceNumber, @InvoiceStatus, @ReceiptDate)
        //";

        //        using (SqlCommand command = new SqlCommand(insertQueryIntoHistory, connection))
        //        {
        //            command.Parameters.AddWithValue("@InvoiceNumber", invoiceNumber);
        //            command.Parameters.AddWithValue("@InvoiceStatus", invoiceStatus);
        //            command.Parameters.AddWithValue("@ReceiptDate", invoiceReceiptDate);

        //            try
        //            {
        //                command.ExecuteNonQuery();
        //            }
        //            catch (Exception ex)
        //            {
        //                ViewBag.ErrorMessage = "يرجي التأكد من ادخال كل البيانات بطريقة صحيحة";
        //                //ViewBag.ErrorMessage = ex.Message;
        //                return View("Failed");
        //            }
        //        }                
            }

            // Redirect to another action after the insert
            return View("DoneCreate", isoDetails);
        }
        [HttpGet]
        public IActionResult ExecuteEdit(string userName, string isoNumber, string isoDate, string invoiceNumber,
            string invoiceValue1, string invoiceValue2, string invoiceValue3, string currency1, string currency2,
            string currency3, string invoiceDate, string invoiceReceiptDate, string requester,
            string vendorName, string backToRequester, string toRequesterDate, string area)
        {
            //string username = User.Identity.Name;
            
            IsoDetails isoDetails = new IsoDetails();
            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                string insertQueryIntoInvoiceDetails = @"Update InvoiceDetails set
                InvoiceValue1=@InvoiceValue1, InvoiceValue2=@InvoiceValue2, InvoiceValue3=@InvoiceValue3, 
                Currency1=@Currency1, Currency2=@Currency2, Currency3=@Currency3, InvoiceDate=@InvoiceDate, 
                InvoiceReceiptDate=@InvoiceReceiptDate, Requester=@Requester,
                VendorName=@VendorName, BackToRequester = @BackToRequester, ToRequesterDate = @ToRequesterDate, Area = @Area , InvoiceEntree = @InvoiceEntree 
                    where IsoNumber=@IsoNumber and InvoiceNumber = @InvoiceNumber;";
                using (SqlCommand command = new SqlCommand(insertQueryIntoInvoiceDetails, connection))
                {
                    command.Parameters.AddWithValue("@UserName", userName);
                    command.Parameters.AddWithValue("@IsoNumber", isoNumber);
                    command.Parameters.AddWithValue("@IsoDate", isoDate);
                    command.Parameters.AddWithValue("@InvoiceNumber", invoiceNumber);
                    command.Parameters.AddWithValue("@InvoiceValue1", Convert.ToDecimal(invoiceValue1));
                    command.Parameters.AddWithValue("@InvoiceValue2",
                        string.IsNullOrWhiteSpace(invoiceValue2) || invoiceValue2 == "null"
                            ? (object)DBNull.Value
                            : Convert.ToDecimal(invoiceValue2));

                    command.Parameters.AddWithValue("@InvoiceValue3",
                        string.IsNullOrWhiteSpace(invoiceValue3) || invoiceValue3 == "null"
                            ? (object)DBNull.Value
                            : Convert.ToDecimal(invoiceValue3));
                    command.Parameters.AddWithValue("@Currency1", currency1);
                    command.Parameters.AddWithValue("@Currency2", currency2 ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@Currency3", currency3 ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@InvoiceDate", invoiceDate);
                    command.Parameters.AddWithValue("@InvoiceReceiptDate", invoiceReceiptDate);
                    command.Parameters.AddWithValue("@Requester", requester ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@VendorName", vendorName);
                    command.Parameters.AddWithValue("@BackToRequester", backToRequester);
                    command.Parameters.AddWithValue("@ToRequesterDate", toRequesterDate ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@Area", area ?? (object)DBNull.Value);
                    command.Parameters.AddWithValue("@InvoiceEntree", userName ?? (object)DBNull.Value);

                    try
                    {

                        command.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                        ViewBag.ErrorMessage = "يرجي التأكد من ادخال كل البيانات بطريقة صحيحة";
                        return View("Failed");
                    }
                }                               
            }
            return View("DoneEdit", isoDetails);
        }
        public IActionResult Exchange()
        {
            DateTime currentDateOnly = DateTime.Now.Date;
            ViewBag.IsoDate = currentDateOnly.ToString("yyyy-MM-dd");
            
            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                string getIso = @"SELECT MAX(ExchangeNumber) FROM InvoiceDetails";
                int NextExchangeNumber = 100;
                // Execute your SQL query using the connection
                using (SqlCommand command = new SqlCommand(getIso, connection))
                {
                    object result = command.ExecuteScalar();

                    // Check if a valid number was returned
                    if (result != DBNull.Value && result != null)
                    {
                        int maxIsoNumber = Convert.ToInt32(result);
                        int nextIsoNumber = maxIsoNumber + 1;

                        // Pass the next IsoNumber to the view using ViewBag
                        NextExchangeNumber = nextIsoNumber;
                    }

                    ViewBag.NextExchangeNumber = NextExchangeNumber;
                }
            }
            return View();
        }        
        [HttpGet]
        public IActionResult ExecuteExchange(string vendorName, string invoiceNumbers, string exchangeNumber, string exchangeCase, string employeeName, 
             string invoiceValue, string currency, string backToRequester, string toRequesterDate)
        {

            try
            {
                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();
                    var lst_InvoiceNumber = invoiceNumbers.Split(',');
                    #region InvoiceDetails                
                    foreach (var _invoiceNumber in lst_InvoiceNumber)
                    {
                        string getId = @"SELECT * FROM InvoiceDetails WHERE VendorName = @VendorName and InvoiceNumber = @InvoiceNumber";

                        // Check if the InvoiceDetails includes the invoice number 
                        using (SqlCommand command = new SqlCommand(getId, connection))
                        {
                            // Add parameters to the command
                            command.Parameters.AddWithValue("@VendorName", vendorName);
                            command.Parameters.AddWithValue("@InvoiceNumber", _invoiceNumber.Trim());

                            // Execute the command and retrieve data
                            using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                            {
                                DataTable dt = new DataTable();
                                adapter.Fill(dt);
                                // Check if the DataTable has data
                                if (!(dt.Rows.Count > 0))
                                {
                                    ViewBag.ErrorMessage = "يجب اضافة الفاتورة أولا";
                                    return View("Failed");
                                }
                            }
                        }
                        // Create the update query with parameters to prevent SQL injection
                        string insertQuery = @"UPDATE InvoiceDetails SET EmployeeName = @EmployeeName , ExchangeNumber = @ExchangeNumber , 
                                     ExchangeCase = @ExchangeCase, InvoiceExchangeValue = @InvoiceExchangeValue, ExchangeCurrency = @ExchangeCurrency
                                    , BackToRequester = @BackToRequester, ToRequesterDate = @ToRequesterDate 
                                    WHERE VendorName = @VendorName and InvoiceNumber = @InvoiceNumber";

                        // Execute the update query using a parameterized SqlCommand
                        using (SqlCommand command = new SqlCommand(insertQuery, connection))
                        {
                            // Add parameters to the command
                            command.Parameters.AddWithValue("@VendorName", vendorName);
                            command.Parameters.AddWithValue("@InvoiceNumber", _invoiceNumber.Trim());
                            command.Parameters.AddWithValue("@EmployeeName", employeeName);
                            command.Parameters.AddWithValue("@ExchangeNumber", exchangeNumber);
                            command.Parameters.AddWithValue("@ExchangeCase", exchangeCase);
                            command.Parameters.AddWithValue("@InvoiceExchangeValue", invoiceValue);
                            command.Parameters.AddWithValue("@ExchangeCurrency", currency);
                            command.Parameters.AddWithValue("@BackToRequester", backToRequester);
                            command.Parameters.AddWithValue("@ToRequesterDate", toRequesterDate ?? (object)DBNull.Value);

                            command.ExecuteNonQuery();
                        }
                    }
                    #endregion InvoiceDetails
                    return View("ExchangeDoneNotificationView");
                }
            }
            catch (Exception)
            {
                ViewBag.ErrorMessage = "حدث خطأ فى ادخال البيانات";
                return View("Failed");
            }
            
        }
        public IActionResult FinalizeExchange()
        {
            return View();
        }
        [HttpGet]
        public IActionResult ExecuteFinalizeExchange(string vendorName, string invoiceNumbers, string exchangeCase, string finalizeExchangeDate)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();
                    var lst_InvoiceNumber = invoiceNumbers.Split(',');
                    #region InvoiceDetails                
                    foreach (var _invoiceNumber in lst_InvoiceNumber)
                    {
                        string getId = @"SELECT * FROM InvoiceDetails WHERE VendorName = @VendorName and InvoiceNumber = @InvoiceNumber";

                        // Check if the InvoiceDetails includes the invoice number 
                        using (SqlCommand command = new SqlCommand(getId, connection))
                        {
                            // Add parameters to the command
                            command.Parameters.AddWithValue("@VendorName", vendorName);
                            command.Parameters.AddWithValue("@InvoiceNumber", _invoiceNumber.Trim());

                            // Execute the command and retrieve data
                            using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                            {
                                DataTable dt = new DataTable();
                                adapter.Fill(dt);
                                // Check if the DataTable has data
                                if (!(dt.Rows.Count > 0))
                                {
                                    ViewBag.ErrorMessage = "يجب اضافة الفاتورة أولا";
                                    return View("Failed");
                                }
                            }
                        }
                        // Create the update query with parameters to prevent SQL injection
                        string insertQuery = @"UPDATE InvoiceDetails SET FinalizeExchangeCase = @ExchangeCase, FinalizeExchangeDate = @FinalizeExchangeDate WHERE VendorName = @VendorName and InvoiceNumber = @InvoiceNumber";

                        // Execute the update query using a parameterized SqlCommand
                        using (SqlCommand command = new SqlCommand(insertQuery, connection))
                        {
                            // Add parameters to the command
                            command.Parameters.AddWithValue("@VendorName", vendorName);
                            command.Parameters.AddWithValue("@InvoiceNumber", _invoiceNumber.Trim());
                            command.Parameters.AddWithValue("@ExchangeCase", exchangeCase);
                            command.Parameters.AddWithValue("@FinalizeExchangeDate", finalizeExchangeDate);

                            command.ExecuteNonQuery();
                        }
                    }
                    #endregion InvoiceDetails
                    return View("FinalizeExchangeDoneNotificationView");
                }
            }
            catch (Exception)
            {
                ViewBag.ErrorMessage = "حدث خطأ فى ادخال البيانات";
                return View("Failed");
            }            
        }
        public IActionResult Edit()
        {
            return View();
        }
        //[HttpGet]
        //public IActionResult ExecuteEdit(string id, string invoiceStatus, string invoiceSapStatus)
        //{

        //    

        //    using (SqlConnection connection = new SqlConnection(_connectionString))
        //    {
        //        connection.Open();
        //        string getId = @"SELECT Id FROM InvoiceDetails WHERE Id = @Id";

        //        // Execute your SQL query using the connection
        //        using (SqlCommand command = new SqlCommand(getId, connection))
        //        {
        //            // Add parameters to the command
        //            command.Parameters.AddWithValue("@Id", id);

        //            // Execute the command and retrieve data
        //            using (SqlDataAdapter adapter = new SqlDataAdapter(command))
        //            {
        //                DataTable dt = new DataTable();
        //                adapter.Fill(dt);
        //                // Check if the DataTable has data
        //                if (!(dt.Rows.Count > 0))
        //                {
        //                    ViewBag.ErrorMessage = "يجب اضافة الفاتورة أولا";
        //                    return View("Failed");
        //                }
        //            }
        //        }
        //        #region InvoiceDetails
        //        // Create the insert query with parameters to prevent SQL injection
        //        string insertQuery = @"UPDATE InvoiceDetails SET Statu = @Statu , SapStatus = @SapStatus WHERE Id = @Id";

        //        // Execute the insert query using a parameterized SqlCommand
        //        using (SqlCommand command = new SqlCommand(insertQuery, connection))
        //        {
        //            // Add parameters to the command
        //            command.Parameters.AddWithValue("@Id", id);
        //            command.Parameters.AddWithValue("@Statu", invoiceStatus);
        //            command.Parameters.AddWithValue("@SapStatus", invoiceSapStatus);

        //            try
        //            {
        //                command.ExecuteNonQuery();
        //                //throw new Exception("An error occurred.");
        //            }
        //            catch (Exception ex)
        //            {
        //                //ViewBag.ErrorMessage = ex.Message;
        //                ViewBag.ErrorMessage = "يرجي التأكد من ادخال كل البيانات بطريقة صحيحة";
        //                return View("Failed");
        //            }
        //        }
        //        #endregion InvoiceDetails
        //        #region History
        //        //Insert into InvoiceDetails
        //        // Create the insert query with parameters to prevent SQL injection
        //        string insertQueryIntoHistory = @"INSERT INTO History (Id, InvoiceStatus, Dat)
        //    VALUES (@Id, @InvoiceStatus, @Dat)";

        //        // Execute the insert query using a parameterized SqlCommand
        //        using (SqlCommand command = new SqlCommand(insertQueryIntoHistory, connection))
        //        {
        //            // Add parameters to the command
        //            command.Parameters.AddWithValue("@Id", id);
        //            command.Parameters.AddWithValue("@InvoiceStatus", invoiceStatus);
        //            command.Parameters.AddWithValue("@Dat", DateTime.Now.Date);

        //            try
        //            {
        //                // Execute the insert command
        //                command.ExecuteNonQuery();
        //                //throw new Exception("An error occurred.");
        //            }
        //            catch (Exception ex)
        //            {
        //                // Pass the exception to the view
        //                ViewBag.ErrorMessage = "حدث خطأ فى ادخال البيانات";
        //                return View("Failed");
        //            }

        //        }
        //        #endregion
        //        // Redirect to another action after the insert
        //        return View("DoneEdit");
        //    }

        //}
        public IActionResult PassToExchange()
        {
            return View();
        }
        [HttpGet]
        public IActionResult ExecutePassToExchange(string id, string employeeName)
        {            
            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                connection.Open();
                string getId = @"SELECT Id FROM InvoiceDetails WHERE Id = @Id";

                // Execute your SQL query using the connection
                using (SqlCommand command = new SqlCommand(getId, connection))
                {
                    // Add parameters to the command
                    command.Parameters.AddWithValue("@Id", id);

                    // Execute the command and retrieve data
                    using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                    {
                        DataTable dt = new DataTable();
                        adapter.Fill(dt);
                        // Check if the DataTable has data
                        if (!(dt.Rows.Count > 0))
                        {
                            ViewBag.ErrorMessage = "يجب اضافة الفاتورة أولا";
                            return View("Failed");
                        }
                    }
                }
                #region InvoiceDetails
                // Create the insert query with parameters to prevent SQL injection
                string insertQuery = @"UPDATE InvoiceDetails SET EmployeeName = @EmployeeName WHERE Id = @Id";

                // Execute the insert query using a parameterized SqlCommand
                using (SqlCommand command = new SqlCommand(insertQuery, connection))
                {
                    // Add parameters to the command
                    command.Parameters.AddWithValue("@Id", id);
                    command.Parameters.AddWithValue("@EmployeeName", employeeName);

                    try
                    {
                        command.ExecuteNonQuery();
                        //throw new Exception("An error occurred.");
                    }
                    catch (Exception ex)
                    {
                        ViewBag.ErrorMessage = ex.Message;
                        ViewBag.ErrorMessage = "يرجي التأكد من ادخال كل البيانات بطريقة صحيحة";
                        return View("Failed");
                    }
                }
                #endregion InvoiceDetails
                
                return View("PassToExchangeDoneNotificationView");
            }

        }
        public IActionResult Query()
        {
            return View();
        }
        public IActionResult LastStep()
        {
            return View();
        }
        public IActionResult FilterByTasweyatReportView()
        {
            return View();
        }
        [HttpGet]
        public IActionResult FilterByTasweyatReport(string date)
        {
            try
            {
                string id = "تسويات";
                ViewBag.Code = id;
                // Connection string

                // SQL query with parameterized query
                string sqlQuery = "SELECT \r\n  UserName As 'اسم المستخدم', \r\n  EmployeeName As 'اسم الموظف', ExchangeNumber As 'رقم طلب الصرف',\r\n FORMAT(InvoiceReceiptDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستلام',\r\n FORMAT(InvoiceDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الفاتورة',\r\n    FORMAT(ToRequesterDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الارسال للجهة الطالبة',\r\n    Requester AS 'الجهه الطالبة', \r\n    Currency2 AS 'العملة  ٢', \r\n    InvoiceValue2 AS 'قيمة الفاتورة  ٢', \r\n    Currency1 AS 'العملة  ١', \r\n    InvoiceValue1 AS 'قيمة الفاتورة  ١', \r\n    VendorName AS 'اسم المورد', \r\n    InvoiceNumber AS 'رقم الفاتورة', \r\n    IsoDate AS 'تاريخ الايزو', \r\n    IsoNumber AS 'رقم الايزو' \r\nFROM \r\n    InvoiceDetails where FinalizeExchangeCase = @Code and InvoiceDate = @InvoiceDate";
                //string sqlQuery = "SELECT \r\n    one.EmployeeName As 'اسم الموظف',\r\n FORMAT(two.Dat, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستلام',\r\n    two.InvoiceStatus AS 'حالة القاتورة',\r\n    one.Requester AS 'الجهه الطالبة',\r\n    FORMAT(one.Dat, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستحقاق', \r\n    one.SapStatus AS 'حالة الفاتورة على ساب', \r\n    one.Currency AS 'العملة', \r\n    one.Valu AS 'قيمة الفاتورة', \r\n    one.VendorName AS 'اسم المورد', \r\n    one.VendorCode AS 'كود المورد',\r\n    one.Contrac AS 'وصف العقد', \r\n    one.Descriptio AS 'وصف الفاتورة', \r\n    one.InvoiceNumber AS 'رقم الفاتورة', \r\n    one.IsoDate AS 'تاريخ الايزو', \r\n    one.IsoNumber AS 'رقم الايزو', \r\n    one.Id AS 'رقم استعلام ساب'\r\nFROM \r\n    InvoiceDetails AS one \r\nJOIN \r\n    [dbo].[History] AS two ON one.InvoiceNumber = two.Id \r\nWHERE \r\n    one.Id = @Code";

                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();
                    // Execute your SQL query using the connection
                    using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                    {
                        // Add parameters to the command
                        command.Parameters.AddWithValue("@Code", id);
                        command.Parameters.AddWithValue("@InvoiceDate", date);

                        // Execute the command and retrieve data
                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            DataTable dt = new DataTable();
                            adapter.Fill(dt);

                            // Create a new DataTable to hold distinct values
                            DataTable distinctDt = new DataTable();

                            // Add columns to the new DataTable based on the original DataTable's schema
                            foreach (DataColumn column in dt.Columns)
                            {
                                distinctDt.Columns.Add(column.ColumnName, column.DataType);
                            }

                            // Get distinct rows from the original DataTable
                            var distinctRows = dt.AsEnumerable().Distinct(DataRowComparer.Default);

                            // Add distinct rows to the new DataTable
                            foreach (DataRow row in distinctRows)
                            {
                                distinctDt.Rows.Add(row.ItemArray);
                            }

                            // Pass the distinct DataTable to the ViewBag
                            ViewBag.DataTable = distinctDt;

                            return View("ViewResult");
                        }

                    }
                }
            }
            catch (Exception)
            {

                ViewBag.ErrorMessage = "يرجي ادخال البيانات المطلوبة بشكل صحيح";
                return View("Failed");
            }            
        }
        public IActionResult FilterBy7sabatReportView()
        {
            return View();
        }
        [HttpGet]
        public IActionResult FilterBy7sabatReport(string date)
        {
            try
            {
                string id = "الحسابات";
                ViewBag.Code = id;
                // Connection string

                // SQL query with parameterized query
                string sqlQuery = "SELECT \r\n  UserName As 'اسم المستخدم', \r\n  EmployeeName As 'اسم الموظف', ExchangeNumber As 'رقم طلب الصرف',\r\n FORMAT(InvoiceReceiptDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستلام',\r\n FORMAT(InvoiceDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الفاتورة',\r\n    FORMAT(ToRequesterDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الارسال للجهة الطالبة',\r\n    Requester AS 'الجهه الطالبة', \r\n    Currency2 AS 'العملة  ٢', \r\n    InvoiceValue2 AS 'قيمة الفاتورة  ٢', \r\n    Currency1 AS 'العملة  ١', \r\n    InvoiceValue1 AS 'قيمة الفاتورة  ١', \r\n    VendorName AS 'اسم المورد', \r\n    InvoiceNumber AS 'رقم الفاتورة', \r\n    IsoDate AS 'تاريخ الايزو', \r\n    IsoNumber AS 'رقم الايزو' \r\nFROM \r\n    InvoiceDetails where FinalizeExchangeCase = @Code and InvoiceDate = @InvoiceDate";
                //string sqlQuery = "SELECT \r\n    one.EmployeeName As 'اسم الموظف',\r\n FORMAT(two.Dat, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستلام',\r\n    two.InvoiceStatus AS 'حالة القاتورة',\r\n    one.Requester AS 'الجهه الطالبة',\r\n    FORMAT(one.Dat, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستحقاق', \r\n    one.SapStatus AS 'حالة الفاتورة على ساب', \r\n    one.Currency AS 'العملة', \r\n    one.Valu AS 'قيمة الفاتورة', \r\n    one.VendorName AS 'اسم المورد', \r\n    one.VendorCode AS 'كود المورد',\r\n    one.Contrac AS 'وصف العقد', \r\n    one.Descriptio AS 'وصف الفاتورة', \r\n    one.InvoiceNumber AS 'رقم الفاتورة', \r\n    one.IsoDate AS 'تاريخ الايزو', \r\n    one.IsoNumber AS 'رقم الايزو', \r\n    one.Id AS 'رقم استعلام ساب'\r\nFROM \r\n    InvoiceDetails AS one \r\nJOIN \r\n    [dbo].[History] AS two ON one.InvoiceNumber = two.Id \r\nWHERE \r\n    one.Id = @Code";

                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();
                    // Execute your SQL query using the connection
                    using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                    {
                        // Add parameters to the command
                        command.Parameters.AddWithValue("@Code", id);
                        command.Parameters.AddWithValue("@InvoiceDate", date);

                        // Execute the command and retrieve data
                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            DataTable dt = new DataTable();
                            adapter.Fill(dt);

                            // Create a new DataTable to hold distinct values
                            DataTable distinctDt = new DataTable();

                            // Add columns to the new DataTable based on the original DataTable's schema
                            foreach (DataColumn column in dt.Columns)
                            {
                                distinctDt.Columns.Add(column.ColumnName, column.DataType);
                            }

                            // Get distinct rows from the original DataTable
                            var distinctRows = dt.AsEnumerable().Distinct(DataRowComparer.Default);

                            // Add distinct rows to the new DataTable
                            foreach (DataRow row in distinctRows)
                            {
                                distinctDt.Rows.Add(row.ItemArray);
                            }

                            // Pass the distinct DataTable to the ViewBag
                            ViewBag.DataTable = distinctDt;

                            return View("ViewResult");
                        }

                    }
                }
            }
            catch (Exception)
            {
                ViewBag.ErrorMessage = "يرجي ادخال البيانات المطلوبة بشكل صحيح";
                return View("Failed");
            }            
        }
        public IActionResult FilterByMasrafyaReportView()
        {
            return View();
        }
        [HttpGet]
        public IActionResult FilterByMasrafyaReport(string date)
        {
            try
            {
                string id = "المصرفية";
                ViewBag.Code = id;
                // Connection string

                // SQL query with parameterized query
                string sqlQuery = "SELECT \r\n  UserName As 'اسم المستخدم', \r\n  EmployeeName As 'اسم الموظف', ExchangeNumber As 'رقم طلب الصرف',\r\n FORMAT(InvoiceReceiptDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستلام',\r\n FORMAT(InvoiceDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الفاتورة',\r\n    FORMAT(ToRequesterDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الارسال للجهة الطالبة',\r\n    Requester AS 'الجهه الطالبة', \r\n    Currency2 AS 'العملة  ٢', \r\n    InvoiceValue2 AS 'قيمة الفاتورة  ٢', \r\n    Currency1 AS 'العملة  ١', \r\n    InvoiceValue1 AS 'قيمة الفاتورة  ١', \r\n    VendorName AS 'اسم المورد', \r\n    InvoiceNumber AS 'رقم الفاتورة', \r\n    IsoDate AS 'تاريخ الايزو', \r\n    IsoNumber AS 'رقم الايزو' \r\nFROM \r\n    InvoiceDetails where FinalizeExchangeCase = @Code and InvoiceDate = @InvoiceDate";
                //string sqlQuery = "SELECT \r\n    one.EmployeeName As 'اسم الموظف',\r\n FORMAT(two.Dat, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستلام',\r\n    two.InvoiceStatus AS 'حالة القاتورة',\r\n    one.Requester AS 'الجهه الطالبة',\r\n    FORMAT(one.Dat, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستحقاق', \r\n    one.SapStatus AS 'حالة الفاتورة على ساب', \r\n    one.Currency AS 'العملة', \r\n    one.Valu AS 'قيمة الفاتورة', \r\n    one.VendorName AS 'اسم المورد', \r\n    one.VendorCode AS 'كود المورد',\r\n    one.Contrac AS 'وصف العقد', \r\n    one.Descriptio AS 'وصف الفاتورة', \r\n    one.InvoiceNumber AS 'رقم الفاتورة', \r\n    one.IsoDate AS 'تاريخ الايزو', \r\n    one.IsoNumber AS 'رقم الايزو', \r\n    one.Id AS 'رقم استعلام ساب'\r\nFROM \r\n    InvoiceDetails AS one \r\nJOIN \r\n    [dbo].[History] AS two ON one.InvoiceNumber = two.Id \r\nWHERE \r\n    one.Id = @Code";

                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();
                    // Execute your SQL query using the connection
                    using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                    {
                        // Add parameters to the command
                        command.Parameters.AddWithValue("@Code", id);
                        command.Parameters.AddWithValue("@InvoiceDate", date);

                        // Execute the command and retrieve data
                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            DataTable dt = new DataTable();
                            adapter.Fill(dt);

                            // Create a new DataTable to hold distinct values
                            DataTable distinctDt = new DataTable();

                            // Add columns to the new DataTable based on the original DataTable's schema
                            foreach (DataColumn column in dt.Columns)
                            {
                                distinctDt.Columns.Add(column.ColumnName, column.DataType);
                            }

                            // Get distinct rows from the original DataTable
                            var distinctRows = dt.AsEnumerable().Distinct(DataRowComparer.Default);

                            // Add distinct rows to the new DataTable
                            foreach (DataRow row in distinctRows)
                            {
                                distinctDt.Rows.Add(row.ItemArray);
                            }

                            // Pass the distinct DataTable to the ViewBag
                            ViewBag.DataTable = distinctDt;

                            return View("ViewResult");
                        }

                    }
                }
            }
            catch (Exception)
            {
                ViewBag.ErrorMessage = "يرجي ادخال البيانات المطلوبة بشكل صحيح";
                return View("Failed");
            }            
        }
        public IActionResult AllVoicesReportView()
        {
            return View();
        }
        [HttpGet]
        public IActionResult AllVoicesReport()
        {
            try
            {
                // SQL query with parameterized query
                //string sqlQuery = "SELECT \r\n    one.EmployeeName As 'اسم الموظف',\r\n FORMAT(two.Dat, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستلام',\r\n    two.InvoiceStatus AS 'حالة القاتورة',\r\n    one.Requester AS 'الجهه الطالبة',\r\n    FORMAT(one.Dat, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستحقاق', \r\n    one.SapStatus AS 'حالة الفاتورة على ساب', \r\n    one.Currency AS 'العملة', \r\n    one.Valu AS 'قيمة الفاتورة', \r\n    one.VendorName AS 'اسم المورد', \r\n    one.VendorCode AS 'كود المورد',\r\n    one.Contrac AS 'وصف العقد', \r\n    one.Descriptio AS 'وصف الفاتورة', \r\n    one.InvoiceNumber AS 'رقم الفاتورة', \r\n    one.IsoDate AS 'تاريخ الايزو', \r\n    one.IsoNumber AS 'رقم الايزو', \r\n    one.Id AS 'رقم استعلام ساب'\r\nFROM \r\n    InvoiceDetails AS one \r\nJOIN \r\n    [dbo].[History] AS two ON one.InvoiceNumber = two.Id";
                string sqlQuery = "SELECT \r\n  UserName As 'اسم المستخدم', \r\n  EmployeeName As 'اسم الموظف', ExchangeNumber As 'رقم طلب الصرف',\r\n FORMAT(InvoiceReceiptDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستلام',\r\n FORMAT(InvoiceDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الفاتورة',\r\n    FORMAT(ToRequesterDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الارسال للجهة الطالبة',\r\n    Requester AS 'الجهه الطالبة', \r\n    Currency3 AS 'العملة - 3', \r\n    InvoiceValue3 AS 'قيمة الفاتورة - 3', \r\n    Currency2 AS 'العملة  ٢', \r\n    InvoiceValue2 AS 'قيمة الفاتورة  ٢', \r\n    Currency1 AS 'العملة  ١', \r\n    InvoiceValue1 AS 'قيمة الفاتورة  ١', \r\n    VendorName AS 'اسم المورد', \r\n    InvoiceNumber AS 'رقم الفاتورة', \r\n    VendorName AS 'اسم المورد', \r\n    IsoDate AS 'تاريخ الايزو', \r\n    IsoNumber AS 'رقم الايزو'\r\nFROM \r\n    InvoiceDetails";

                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();

                    // Execute your SQL query using the connection
                    using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                    {
                        // Execute the command and retrieve data
                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            DataTable dt = new DataTable();
                            adapter.Fill(dt);

                            // Create a new DataTable to hold distinct values
                            DataTable distinctDt = new DataTable();

                            // Add columns to the new DataTable based on the original DataTable's schema
                            foreach (DataColumn column in dt.Columns)
                            {
                                distinctDt.Columns.Add(column.ColumnName, column.DataType);
                            }

                            // Get distinct rows from the original DataTable
                            var distinctRows = dt.AsEnumerable().Distinct(DataRowComparer.Default);

                            // Add distinct rows to the new DataTable
                            foreach (DataRow row in distinctRows)
                            {
                                distinctDt.Rows.Add(row.ItemArray);
                            }

                            // Pass the distinct DataTable to the ViewBag
                            ViewBag.DataTable = distinctDt;

                            return View("ViewResult");
                        }
                    }
                }
            }
            catch (Exception)
            {
                ViewBag.ErrorMessage = "يرجي ادخال البيانات المطلوبة بشكل صحيح";
                return View("Failed");
            }            
        }
        public IActionResult FilterByExchangeCaseWithDateReportView()
        {
            return View();
        }
        [HttpGet]
        public IActionResult FilterByExchangeCaseWithDateReport(string id, string date)
        {
            try
            {
                ViewBag.Code = id;
                // Connection string

                // SQL query with parameterized query
                string sqlQuery = "SELECT \r\n  UserName As 'اسم المستخدم', \r\n  EmployeeName As 'اسم الموظف', ExchangeNumber As 'رقم طلب الصرف',\r\n FORMAT(InvoiceReceiptDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستلام',\r\n FORMAT(InvoiceDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الفاتورة',\r\n    FORMAT(ToRequesterDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الارسال للجهة الطالبة',\r\n    Requester AS 'الجهه الطالبة', \r\n    Currency2 AS 'العملة  ٢', \r\n    InvoiceValue2 AS 'قيمة الفاتورة  ٢', \r\n    Currency1 AS 'العملة  ١', \r\n    InvoiceValue1 AS 'قيمة الفاتورة  ١', \r\n    VendorName AS 'اسم المورد', \r\n    InvoiceNumber AS 'رقم الفاتورة', \r\n    IsoDate AS 'تاريخ الايزو'\r\nFROM \r\n    InvoiceDetails where ExchangeCase = @Code and InvoiceDate = @InvoiceDate";
                //string sqlQuery = "SELECT \r\n    one.EmployeeName As 'اسم الموظف',\r\n FORMAT(two.Dat, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستلام',\r\n    two.InvoiceStatus AS 'حالة القاتورة',\r\n    one.Requester AS 'الجهه الطالبة',\r\n    FORMAT(one.Dat, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستحقاق', \r\n    one.SapStatus AS 'حالة الفاتورة على ساب', \r\n    one.Currency AS 'العملة', \r\n    one.Valu AS 'قيمة الفاتورة', \r\n    one.VendorName AS 'اسم المورد', \r\n    one.VendorCode AS 'كود المورد',\r\n    one.Contrac AS 'وصف العقد', \r\n    one.Descriptio AS 'وصف الفاتورة', \r\n    one.InvoiceNumber AS 'رقم الفاتورة', \r\n    one.IsoDate AS 'تاريخ الايزو', \r\n    one.IsoNumber AS 'رقم الايزو', \r\n    one.Id AS 'رقم استعلام ساب'\r\nFROM \r\n    InvoiceDetails AS one \r\nJOIN \r\n    [dbo].[History] AS two ON one.InvoiceNumber = two.Id \r\nWHERE \r\n    one.Id = @Code";

                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();
                    // Execute your SQL query using the connection
                    using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                    {
                        // Add parameters to the command
                        command.Parameters.AddWithValue("@Code", id);
                        command.Parameters.AddWithValue("@InvoiceDate", date);

                        // Execute the command and retrieve data
                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            DataTable dt = new DataTable();
                            adapter.Fill(dt);

                            // Create a new DataTable to hold distinct values
                            DataTable distinctDt = new DataTable();

                            // Add columns to the new DataTable based on the original DataTable's schema
                            foreach (DataColumn column in dt.Columns)
                            {
                                distinctDt.Columns.Add(column.ColumnName, column.DataType);
                            }

                            // Get distinct rows from the original DataTable
                            var distinctRows = dt.AsEnumerable().Distinct(DataRowComparer.Default);

                            // Add distinct rows to the new DataTable
                            foreach (DataRow row in distinctRows)
                            {
                                distinctDt.Rows.Add(row.ItemArray);
                            }

                            // Pass the distinct DataTable to the ViewBag
                            ViewBag.DataTable = distinctDt;

                            return View("ViewResult");
                        }

                    }
                }
            }
            catch (Exception)
            {
                ViewBag.ErrorMessage = "يرجي ادخال البيانات المطلوبة بشكل صحيح";
                return View("Failed");
            }            
        }
        public IActionResult FilterBySapTempCodeReportView()
        {
            return View();
        }
        [HttpGet]
        public IActionResult FilterBySapTempCodeReport(string id)
        {
            try
            {
                ViewBag.Code = id;
                // Connection string

                string checkInvoiceValue2 = "SELECT InvoiceValue2 \r\nFROM \r\n    InvoiceDetails where IsoNumber = @Code";
                // SQL query with parameterized query
                string sqlQuery = "SELECT \r\n  UserName As 'اسم المستخدم',\r\n FORMAT(InvoiceReceiptDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستلام',\r\n FORMAT(InvoiceDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الفاتورة',\r\n    FORMAT(ToRequesterDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الارسال للجهة الطالبة',\r\n    Requester AS 'الجهه الطالبة', \r\n    Currency2 AS 'العملة  ٢', \r\n    InvoiceValue2 AS 'قيمة الفاتورة  ٢', \r\n    Currency1 AS 'العملة  ١', \r\n    InvoiceValue1 AS 'قيمة الفاتورة  ١', \r\n    VendorName AS 'اسم المورد', \r\n    InvoiceNumber AS 'رقم الفاتورة', \r\n    IsoDate AS 'تاريخ الايزو', \r\nIsoNumber AS 'رقم الايزو' \r\nFROM \r\n    InvoiceDetails where IsoNumber = @Code";
                //string sqlQuery = "SELECT \r\n    one.EmployeeName As 'اسم الموظف',\r\n FORMAT(two.Dat, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستلام',\r\n    two.InvoiceStatus AS 'حالة القاتورة',\r\n    one.Requester AS 'الجهه الطالبة',\r\n    FORMAT(one.Dat, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستحقاق', \r\n    one.SapStatus AS 'حالة الفاتورة على ساب', \r\n    one.Currency AS 'العملة', \r\n    one.Valu AS 'قيمة الفاتورة', \r\n    one.VendorName AS 'اسم المورد', \r\n    one.VendorCode AS 'كود المورد',\r\n    one.Contrac AS 'وصف العقد', \r\n    one.Descriptio AS 'وصف الفاتورة', \r\n    one.InvoiceNumber AS 'رقم الفاتورة', \r\n    one.IsoDate AS 'تاريخ الايزو', \r\n    one.IsoNumber AS 'رقم الايزو', \r\n    one.Id AS 'رقم استعلام ساب'\r\nFROM \r\n    InvoiceDetails AS one \r\nJOIN \r\n    [dbo].[History] AS two ON one.InvoiceNumber = two.Id \r\nWHERE \r\n    one.Id = @Code";

                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand(checkInvoiceValue2, connection))
                    {
                        command.Parameters.AddWithValue("@Code", id);
                        object result = command.ExecuteScalar();

                        // Check if a valid number was returned
                        if (result == DBNull.Value || result == null)
                        {
                            sqlQuery = "SELECT \r\n  UserName As 'اسم المستخدم',\r\n FORMAT(InvoiceReceiptDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستلام',\r\n FORMAT(InvoiceDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الفاتورة',\r\n    FORMAT(ToRequesterDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الارسال للجهة الطالبة',\r\n    Requester AS 'الجهه الطالبة', \r\n    Currency1 AS 'العملة  ١', \r\n    InvoiceValue1 AS 'قيمة الفاتورة  ١', \r\n    VendorName AS 'اسم المورد', \r\n    InvoiceNumber AS 'رقم الفاتورة', \r\n    IsoDate AS 'تاريخ الايزو'\r\n,\r\n IsoNumber AS 'رقم الايزو' FROM \r\n    InvoiceDetails where IsoNumber = @Code";
                        }
                    }
                    // Execute your SQL query using the connection
                    using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                    {
                        // Add parameters to the command
                        command.Parameters.AddWithValue("@Code", id);

                        // Execute the command and retrieve data
                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            DataTable dt = new DataTable();
                            adapter.Fill(dt);

                            // Create a new DataTable to hold distinct values
                            DataTable distinctDt = new DataTable();

                            // Add columns to the new DataTable based on the original DataTable's schema
                            foreach (DataColumn column in dt.Columns)
                            {
                                distinctDt.Columns.Add(column.ColumnName, column.DataType);
                            }

                            // Get distinct rows from the original DataTable
                            var distinctRows = dt.AsEnumerable().Distinct(DataRowComparer.Default);

                            // Add distinct rows to the new DataTable
                            foreach (DataRow row in distinctRows)
                            {
                                distinctDt.Rows.Add(row.ItemArray);
                            }

                            // Pass the distinct DataTable to the ViewBag
                            ViewBag.DataTable = distinctDt;

                            return View("ViewResult");
                        }

                    }
                }
            }
            catch (Exception)
            {
                ViewBag.ErrorMessage = "يرجي ادخال البيانات المطلوبة بشكل صحيح";
                return View("Failed");
            }            
        }
        public IActionResult FilterByInvoiceNumberReportView()
        {
            return View();
        }
        [HttpGet]
        public IActionResult FilterByInvoiceNumberReport(string id)
        {
            try
            {
                ViewBag.Code = id;
                // Connection string

                string queryCheckInvoiceValue2 = "SELECT InvoiceValue2 AS 'قيمة الفاتورة  ٢' FROM InvoiceDetails  where InvoiceNumber = @Code";

                // SQL query with parameterized query
                string sqlQuery = "SELECT \r\n  UserName As 'اسم المستخدم',\r\n FORMAT(InvoiceReceiptDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستلام',\r\n FORMAT(InvoiceDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الفاتورة',\r\n    FORMAT(ToRequesterDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الارسال للجهة الطالبة',\r\n    Requester AS 'الجهه الطالبة', \r\n    Currency2 AS 'العملة ٢', \r\n    InvoiceValue2 AS 'قيمة الفاتورة ٢', \r\n    Currency1 AS 'العملة ١', \r\n    InvoiceValue1 AS 'قيمة الفاتورة ١', \r\n InvoiceNumber AS 'رقم الفاتورة', \r\n    VendorName AS 'اسم المورد', \r\n    IsoDate AS 'تاريخ الايزو', \r\n    IsoNumber AS 'رقم الايزو'\r\nFROM \r\n    InvoiceDetails where InvoiceNumber = @Code";
                //string sqlQuery = "SELECT \r\n    one.EmployeeName As 'اسم الموظف',\r\n FORMAT(two.Dat, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستلام',\r\n    two.InvoiceStatus AS 'حالة القاتورة',\r\n    one.Requester AS 'الجهه الطالبة',\r\n    FORMAT(one.Dat, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستحقاق', \r\n    one.SapStatus AS 'حالة الفاتورة على ساب', \r\n    one.Currency AS 'العملة', \r\n    one.Valu AS 'قيمة الفاتورة', \r\n    one.VendorName AS 'اسم المورد', \r\n    one.VendorCode AS 'كود المورد',\r\n    one.Contrac AS 'وصف العقد', \r\n    one.Descriptio AS 'وصف الفاتورة', \r\n    one.InvoiceNumber AS 'رقم الفاتورة', \r\n    one.IsoDate AS 'تاريخ الايزو', \r\n    one.IsoNumber AS 'رقم الايزو', \r\n    one.Id AS 'رقم استعلام ساب'\r\nFROM \r\n    InvoiceDetails AS one \r\nJOIN \r\n    [dbo].[History] AS two ON one.InvoiceNumber = two.Id \r\nWHERE \r\n    one.Id = @Code";

                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();
                    using (SqlCommand command = new SqlCommand(queryCheckInvoiceValue2, connection))
                    {
                        command.Parameters.AddWithValue("@Code", id);
                        object result = command.ExecuteScalar();

                        // Check if a valid number was returned
                        if (result == DBNull.Value || result == null)
                        {
                            sqlQuery = "SELECT \r\n  UserName As 'اسم المستخدم',\r\n FORMAT(InvoiceReceiptDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستلام',\r\n FORMAT(InvoiceDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الفاتورة',\r\n    FORMAT(ToRequesterDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الارسال للجهة الطالبة',\r\n    Requester AS 'الجهه الطالبة',\r\n    Currency1 AS 'العملة ١', \r\n    InvoiceValue1 AS 'قيمة الفاتورة ١', \r\n InvoiceNumber AS 'رقم الفاتورة', \r\n    VendorName AS 'اسم المورد', \r\n    IsoDate AS 'تاريخ الايزو', \r\n    IsoNumber AS 'رقم الايزو'\r\nFROM \r\n    InvoiceDetails where InvoiceNumber = @Code";
                        }
                    }

                    // Execute your SQL query using the connection
                    using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                    {
                        // Add parameters to the command
                        command.Parameters.AddWithValue("@Code", id);

                        // Execute the command and retrieve data
                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            DataTable dt = new DataTable();
                            adapter.Fill(dt);

                            // Create a new DataTable to hold distinct values
                            DataTable distinctDt = new DataTable();

                            // Add columns to the new DataTable based on the original DataTable's schema
                            foreach (DataColumn column in dt.Columns)
                            {
                                distinctDt.Columns.Add(column.ColumnName, column.DataType);
                            }

                            // Get distinct rows from the original DataTable
                            var distinctRows = dt.AsEnumerable().Distinct(DataRowComparer.Default);

                            // Add distinct rows to the new DataTable
                            foreach (DataRow row in distinctRows)
                            {
                                distinctDt.Rows.Add(row.ItemArray);
                            }

                            // Pass the distinct DataTable to the ViewBag
                            ViewBag.DataTable = distinctDt;

                            return View("ViewResult");
                        }

                    }
                }
            }
            catch (Exception)
            {
                ViewBag.ErrorMessage = "يرجي ادخال البيانات المطلوبة بشكل صحيح";
                return View("Failed");
            }            
        }
        public IActionResult FilterByExchangeNumberReportView()
        {
            return View();
        }
        [HttpGet]
        public IActionResult FilterByExchangeNumberReport(string id)
        {
            try
            {
                ViewBag.Code = id;

                // SQL query with parameterized query
                string sqlQuery = "SELECT \r\n  UserName As 'اسم المستخدم', \r\n  EmployeeName As 'اسم الموظف',\r\n FORMAT(InvoiceReceiptDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستلام',\r\n FORMAT(InvoiceDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الفاتورة',\r\n    FORMAT(ToRequesterDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الارسال للجهة الطالبة',\r\n    Requester AS 'الجهه الطالبة', \r\n    Currency3 AS 'العملة - 3', \r\n    InvoiceValue3 AS 'قيمة الفاتورة - 3', \r\n    Currency2 AS 'العملة  ٢', \r\n    InvoiceValue2 AS 'قيمة الفاتورة  ٢', \r\n    Currency1 AS 'العملة  ١', \r\n    InvoiceValue1 AS 'قيمة الفاتورة  ١', \r\n    VendorName AS 'اسم المورد', \r\n    InvoiceNumber AS 'رقم الفاتورة', \r\n    IsoDate AS 'تاريخ الايزو', \r\n    IsoNumber AS 'رقم الايزو'\r\nFROM \r\n    InvoiceDetails where ExchangeNumber = @Code";

                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();

                    // Execute your SQL query using the connection
                    using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                    {
                        // Add parameters to the command
                        command.Parameters.AddWithValue("@Code", id);

                        // Execute the command and retrieve data
                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            DataTable dt = new DataTable();
                            adapter.Fill(dt);

                            // Create a new DataTable to hold distinct values
                            DataTable distinctDt = new DataTable();

                            // Add columns to the new DataTable based on the original DataTable's schema
                            foreach (DataColumn column in dt.Columns)
                            {
                                distinctDt.Columns.Add(column.ColumnName, column.DataType);
                            }

                            // Get distinct rows from the original DataTable
                            var distinctRows = dt.AsEnumerable().Distinct(DataRowComparer.Default);

                            // Add distinct rows to the new DataTable
                            foreach (DataRow row in distinctRows)
                            {
                                distinctDt.Rows.Add(row.ItemArray);
                            }

                            // Pass the distinct DataTable to the ViewBag
                            ViewBag.DataTable = distinctDt;

                            return View("ViewResult");
                        }

                    }
                }
            }
            catch (Exception)
            {
                ViewBag.ErrorMessage = "يرجي ادخال البيانات المطلوبة بشكل صحيح";
                return View("Failed");
            }            
        }
        public IActionResult FilterByInvoiceValueReportView()
        {
            return View();
        }
        [HttpGet]
        public IActionResult FilterByInvoiceValueReport(string id)
        {
            try
            {
                ViewBag.Code = id;

                // SQL query with parameterized query
                string sqlQuery = "SELECT \r\n  UserName As 'اسم المستخدم', \r\n  EmployeeName As 'اسم الموظف',\r\n FORMAT(InvoiceReceiptDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستلام',\r\n FORMAT(InvoiceDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الفاتورة',\r\n    FORMAT(ToRequesterDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الارسال للجهة الطالبة',\r\n    Requester AS 'الجهه الطالبة', \r\n    Currency3 AS 'العملة - 3', \r\n    InvoiceValue3 AS 'قيمة الفاتورة - 3', \r\n    Currency2 AS 'العملة  ٢', \r\n    InvoiceValue2 AS 'قيمة الفاتورة  ٢', \r\n    Currency1 AS 'العملة  ١', \r\n    InvoiceValue1 AS 'قيمة الفاتورة  ١', \r\n    VendorName AS 'اسم المورد', \r\n    InvoiceNumber AS 'رقم الفاتورة', \r\n    IsoDate AS 'تاريخ الايزو', \r\n    IsoNumber AS 'رقم الايزو'\r\nFROM \r\n    InvoiceDetails where InvoiceExchangeValue = @Code";

                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();

                    // Execute your SQL query using the connection
                    using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                    {
                        // Add parameters to the command
                        command.Parameters.AddWithValue("@Code", id);

                        // Execute the command and retrieve data
                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            DataTable dt = new DataTable();
                            adapter.Fill(dt);

                            // Create a new DataTable to hold distinct values
                            DataTable distinctDt = new DataTable();

                            // Add columns to the new DataTable based on the original DataTable's schema
                            foreach (DataColumn column in dt.Columns)
                            {
                                distinctDt.Columns.Add(column.ColumnName, column.DataType);
                            }

                            // Get distinct rows from the original DataTable
                            var distinctRows = dt.AsEnumerable().Distinct(DataRowComparer.Default);

                            // Add distinct rows to the new DataTable
                            foreach (DataRow row in distinctRows)
                            {
                                distinctDt.Rows.Add(row.ItemArray);
                            }

                            // Pass the distinct DataTable to the ViewBag
                            ViewBag.DataTable = distinctDt;

                            return View("ViewResult");
                        }

                    }
                }
            }
            catch (Exception)
            {
                ViewBag.ErrorMessage = "يرجي ادخال البيانات المطلوبة بشكل صحيح";
                return View("Failed");
            }            
        }
        public IActionResult FilterByRequesterReportView()
        {
            return View();
        }
        [HttpGet]
        public IActionResult FilterByRequesterReport(string id)
        {
            try
            {
                ViewBag.Code = id;

                // SQL query with parameterized query
                string sqlQuery = "SELECT \r\n  UserName As 'اسم المستخدم', \r\n  EmployeeName As 'اسم الموظف', ExchangeNumber As 'رقم طلب الصرف',\r\n FORMAT(InvoiceReceiptDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستلام',\r\n FORMAT(InvoiceDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الفاتورة',\r\n    FORMAT(ToRequesterDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الارسال للجهة الطالبة',\r\n    Requester AS 'الجهه الطالبة', \r\n    Currency3 AS 'العملة - 3', \r\n    InvoiceValue3 AS 'قيمة الفاتورة - 3', \r\n    Currency2 AS 'العملة  ٢', \r\n    InvoiceValue2 AS 'قيمة الفاتورة  ٢', \r\n    Currency1 AS 'العملة  ١', \r\n    InvoiceValue1 AS 'قيمة الفاتورة  ١', \r\n    VendorName AS 'اسم المورد', \r\n    InvoiceNumber AS 'رقم الفاتورة', \r\n    IsoDate AS 'تاريخ الايزو', \r\n    IsoNumber AS 'رقم الايزو'\r\nFROM \r\n    InvoiceDetails where Requester = @Code";

                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();

                    // Execute your SQL query using the connection
                    using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                    {
                        // Add parameters to the command
                        command.Parameters.AddWithValue("@Code", id);

                        // Execute the command and retrieve data
                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            DataTable dt = new DataTable();
                            adapter.Fill(dt);

                            // Create a new DataTable to hold distinct values
                            DataTable distinctDt = new DataTable();

                            // Add columns to the new DataTable based on the original DataTable's schema
                            foreach (DataColumn column in dt.Columns)
                            {
                                distinctDt.Columns.Add(column.ColumnName, column.DataType);
                            }

                            // Get distinct rows from the original DataTable
                            var distinctRows = dt.AsEnumerable().Distinct(DataRowComparer.Default);

                            // Add distinct rows to the new DataTable
                            foreach (DataRow row in distinctRows)
                            {
                                distinctDt.Rows.Add(row.ItemArray);
                            }

                            // Pass the distinct DataTable to the ViewBag
                            ViewBag.DataTable = distinctDt;

                            return View("ViewResult");
                        }

                    }
                }
            }
            catch (Exception)
            {
                ViewBag.ErrorMessage = "يرجي ادخال البيانات المطلوبة بشكل صحيح";
                return View("Failed");
            }            
        }
        public IActionResult FilterByVendorReportView()
        {
            return View();
        }
        [HttpGet]
        public IActionResult FilterByVendorReport(string id)
        {
            try
            {
                ViewBag.Code = id;

                // SQL query with parameterized query
                string sqlQuery = "SELECT \r\n  UserName As 'اسم المستخدم', \r\n  EmployeeName As 'اسم الموظف', ExchangeNumber As 'رقم طلب الصرف',\r\n FORMAT(InvoiceReceiptDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستلام',\r\n FORMAT(InvoiceDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الفاتورة',\r\n    FORMAT(ToRequesterDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الارسال للجهة الطالبة',\r\n    Requester AS 'الجهه الطالبة', \r\n    Currency3 AS 'العملة - 3', \r\n    InvoiceValue3 AS 'قيمة الفاتورة - 3', \r\n    Currency2 AS 'العملة  ٢', \r\n    InvoiceValue2 AS 'قيمة الفاتورة  ٢', \r\n    Currency1 AS 'العملة  ١', \r\n    InvoiceValue1 AS 'قيمة الفاتورة  ١', \r\n    VendorName AS 'اسم المورد', \r\n    InvoiceNumber AS 'رقم الفاتورة', \r\n    IsoDate AS 'تاريخ الايزو', \r\n    IsoNumber AS 'رقم الايزو'\r\nFROM \r\n    InvoiceDetails where VendorName = @Code";

                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();

                    // Execute your SQL query using the connection
                    using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                    {
                        // Add parameters to the command
                        command.Parameters.AddWithValue("@Code", id);

                        // Execute the command and retrieve data
                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            DataTable dt = new DataTable();
                            adapter.Fill(dt);

                            // Create a new DataTable to hold distinct values
                            DataTable distinctDt = new DataTable();

                            // Add columns to the new DataTable based on the original DataTable's schema
                            foreach (DataColumn column in dt.Columns)
                            {
                                distinctDt.Columns.Add(column.ColumnName, column.DataType);
                            }

                            // Get distinct rows from the original DataTable
                            var distinctRows = dt.AsEnumerable().Distinct(DataRowComparer.Default);

                            // Add distinct rows to the new DataTable
                            foreach (DataRow row in distinctRows)
                            {
                                distinctDt.Rows.Add(row.ItemArray);
                            }

                            // Pass the distinct DataTable to the ViewBag
                            ViewBag.DataTable = distinctDt;

                            return View("ViewResult");
                        }

                    }
                }
            }
            catch (Exception)
            {
                ViewBag.ErrorMessage = "يرجي ادخال البيانات المطلوبة بشكل صحيح";
                return View("Failed");
            }            
        }
        public IActionResult FilterByEmployeeReportView()
        {
            return View();
        }
        [HttpGet]
        public IActionResult FilterByEmployeeReport(string id)
        {
            try
            {
                ViewBag.Code = id;

                // SQL query with parameterized query
                string sqlQuery = "SELECT \r\n  UserName As 'اسم المستخدم', \r\n  EmployeeName As 'اسم الموظف', ExchangeNumber As 'رقم طلب الصرف',\r\n FORMAT(InvoiceReceiptDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستلام',\r\n FORMAT(InvoiceDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الفاتورة',\r\n    FORMAT(ToRequesterDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الارسال للجهة الطالبة',\r\n    Requester AS 'الجهه الطالبة', \r\n    Currency3 AS 'العملة - 3', \r\n    InvoiceValue3 AS 'قيمة الفاتورة - 3', \r\n    Currency2 AS 'العملة  ٢', \r\n    InvoiceValue2 AS 'قيمة الفاتورة  ٢', \r\n    Currency1 AS 'العملة  ١', \r\n    InvoiceValue1 AS 'قيمة الفاتورة  ١', \r\n    VendorName AS 'اسم المورد', \r\n    InvoiceNumber AS 'رقم الفاتورة', \r\n    IsoDate AS 'تاريخ الايزو', \r\n    IsoNumber AS 'رقم الايزو'\r\nFROM \r\n    InvoiceDetails where EmployeeName = @Code";

                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();

                    // Execute your SQL query using the connection
                    using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                    {
                        // Add parameters to the command
                        command.Parameters.AddWithValue("@Code", id);

                        // Execute the command and retrieve data
                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            DataTable dt = new DataTable();
                            adapter.Fill(dt);

                            // Create a new DataTable to hold distinct values
                            DataTable distinctDt = new DataTable();

                            // Add columns to the new DataTable based on the original DataTable's schema
                            foreach (DataColumn column in dt.Columns)
                            {
                                distinctDt.Columns.Add(column.ColumnName, column.DataType);
                            }

                            // Get distinct rows from the original DataTable
                            var distinctRows = dt.AsEnumerable().Distinct(DataRowComparer.Default);

                            // Add distinct rows to the new DataTable
                            foreach (DataRow row in distinctRows)
                            {
                                distinctDt.Rows.Add(row.ItemArray);
                            }

                            // Pass the distinct DataTable to the ViewBag
                            ViewBag.DataTable = distinctDt;

                            return View("ViewResult");
                        }

                    }
                }
            }
            catch (Exception)
            {
                ViewBag.ErrorMessage = "يرجي ادخال البيانات المطلوبة بشكل صحيح";
                return View("Failed");
            }            
        }
        public IActionResult FilterByInvoiceDateReportView()
        {
            return View();
        }
        [HttpGet]
        public IActionResult FilterByInvoiceDateReport(string id)
        {
            try
            {
                ViewBag.Code = id;

                // SQL query with parameterized query
                string sqlQuery = "SELECT \r\n  UserName As 'اسم المستخدم', \r\n  EmployeeName As 'اسم الموظف', ExchangeNumber As 'رقم طلب الصرف',\r\n FORMAT(InvoiceReceiptDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستلام',\r\n FORMAT(InvoiceDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الفاتورة',\r\n    FORMAT(ToRequesterDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الارسال للجهة الطالبة',\r\n    Requester AS 'الجهه الطالبة', \r\n    Currency3 AS 'العملة - 3', \r\n    InvoiceValue3 AS 'قيمة الفاتورة - 3', \r\n    Currency2 AS 'العملة  ٢', \r\n    InvoiceValue2 AS 'قيمة الفاتورة  ٢', \r\n    Currency1 AS 'العملة  ١', \r\n    InvoiceValue1 AS 'قيمة الفاتورة  ١', \r\n    VendorName AS 'اسم المورد', \r\n    InvoiceNumber AS 'رقم الفاتورة', \r\n    IsoDate AS 'تاريخ الايزو', \r\n    IsoNumber AS 'رقم الايزو'\r\nFROM \r\n    InvoiceDetails where InvoiceDate = @Code";

                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();

                    // Execute your SQL query using the connection
                    using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                    {
                        // Add parameters to the command
                        command.Parameters.AddWithValue("@Code", id);

                        // Execute the command and retrieve data
                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            DataTable dt = new DataTable();
                            adapter.Fill(dt);

                            // Create a new DataTable to hold distinct values
                            DataTable distinctDt = new DataTable();

                            // Add columns to the new DataTable based on the original DataTable's schema
                            foreach (DataColumn column in dt.Columns)
                            {
                                distinctDt.Columns.Add(column.ColumnName, column.DataType);
                            }

                            // Get distinct rows from the original DataTable
                            var distinctRows = dt.AsEnumerable().Distinct(DataRowComparer.Default);

                            // Add distinct rows to the new DataTable
                            foreach (DataRow row in distinctRows)
                            {
                                distinctDt.Rows.Add(row.ItemArray);
                            }

                            // Pass the distinct DataTable to the ViewBag
                            ViewBag.DataTable = distinctDt;

                            return View("ViewResult");
                        }

                    }
                }
            }
            catch (Exception)
            {
                ViewBag.ErrorMessage = "يرجي ادخال البيانات المطلوبة بشكل صحيح";
                return View("Failed");
            }            
        }
        public IActionResult FilterByInvoiceReceiptDateReportView()
        {
            return View();
        }
        [HttpGet]
        public IActionResult FilterByInvoiceReceiptDateReport(string id)
        {
            try
            {
                ViewBag.Code = id;

                // SQL query with parameterized query
                string sqlQuery = "SELECT \r\n  UserName As 'اسم المستخدم', \r\n  EmployeeName As 'اسم الموظف', ExchangeNumber As 'رقم طلب الصرف',\r\n FORMAT(InvoiceReceiptDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستلام',\r\n FORMAT(InvoiceDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الفاتورة',\r\n    FORMAT(ToRequesterDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الارسال للجهة الطالبة',\r\n    Requester AS 'الجهه الطالبة', \r\n    Currency3 AS 'العملة - 3', \r\n    InvoiceValue3 AS 'قيمة الفاتورة - 3', \r\n    Currency2 AS 'العملة  ٢', \r\n    InvoiceValue2 AS 'قيمة الفاتورة  ٢', \r\n    Currency1 AS 'العملة  ١', \r\n    InvoiceValue1 AS 'قيمة الفاتورة  ١', \r\n    VendorName AS 'اسم المورد', \r\n    InvoiceNumber AS 'رقم الفاتورة', \r\n    IsoDate AS 'تاريخ الايزو', \r\n    IsoNumber AS 'رقم الايزو'\r\nFROM \r\n    InvoiceDetails where InvoiceReceiptDate = @Code";

                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();

                    // Execute your SQL query using the connection
                    using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                    {
                        // Add parameters to the command
                        command.Parameters.AddWithValue("@Code", id);

                        // Execute the command and retrieve data
                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            DataTable dt = new DataTable();
                            adapter.Fill(dt);

                            // Create a new DataTable to hold distinct values
                            DataTable distinctDt = new DataTable();

                            // Add columns to the new DataTable based on the original DataTable's schema
                            foreach (DataColumn column in dt.Columns)
                            {
                                distinctDt.Columns.Add(column.ColumnName, column.DataType);
                            }

                            // Get distinct rows from the original DataTable
                            var distinctRows = dt.AsEnumerable().Distinct(DataRowComparer.Default);

                            // Add distinct rows to the new DataTable
                            foreach (DataRow row in distinctRows)
                            {
                                distinctDt.Rows.Add(row.ItemArray);
                            }

                            // Pass the distinct DataTable to the ViewBag
                            ViewBag.DataTable = distinctDt;

                            return View("ViewResult");
                        }

                    }
                }
            }
            catch (Exception)
            {
                ViewBag.ErrorMessage = "يرجي ادخال البيانات المطلوبة بشكل صحيح";
                return View("Failed");
            }            
        }
        public IActionResult FilterByAreaReportView()
        {
            return View();
        }
        [HttpGet]
        public IActionResult FilterByAreaReport(string id)
        {
            try
            {
                ViewBag.Code = id;

                // SQL query with parameterized query
                string sqlQuery = "SELECT \r\n  UserName As 'اسم المستخدم', \r\n  EmployeeName As 'اسم الموظف', ExchangeNumber As 'رقم طلب الصرف',\r\n FORMAT(InvoiceReceiptDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الاستلام',\r\n FORMAT(InvoiceDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الفاتورة',\r\n    FORMAT(ToRequesterDate, 'dd MMMM yyyy', 'ar-EG') AS 'تاريخ الارسال للجهة الطالبة',\r\n    Requester AS 'الجهه الطالبة', \r\n    Currency3 AS 'العملة - 3', \r\n    InvoiceValue3 AS 'قيمة الفاتورة - 3', \r\n    Currency2 AS 'العملة  ٢', \r\n    InvoiceValue2 AS 'قيمة الفاتورة  ٢', \r\n    Currency1 AS 'العملة  ١', \r\n    InvoiceValue1 AS 'قيمة الفاتورة  ١', \r\n    VendorName AS 'اسم المورد', \r\n    InvoiceNumber AS 'رقم الفاتورة', \r\n    IsoDate AS 'تاريخ الايزو', \r\n    IsoNumber AS 'رقم الايزو'\r\nFROM \r\n    InvoiceDetails where EmployeeName = @Code";

                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();

                    // Execute your SQL query using the connection
                    using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                    {
                        // Add parameters to the command
                        command.Parameters.AddWithValue("@Code", id);

                        // Execute the command and retrieve data
                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            DataTable dt = new DataTable();
                            adapter.Fill(dt);

                            // Create a new DataTable to hold distinct values
                            DataTable distinctDt = new DataTable();

                            // Add columns to the new DataTable based on the original DataTable's schema
                            foreach (DataColumn column in dt.Columns)
                            {
                                distinctDt.Columns.Add(column.ColumnName, column.DataType);
                            }

                            // Get distinct rows from the original DataTable
                            var distinctRows = dt.AsEnumerable().Distinct(DataRowComparer.Default);

                            // Add distinct rows to the new DataTable
                            foreach (DataRow row in distinctRows)
                            {
                                distinctDt.Rows.Add(row.ItemArray);
                            }

                            // Pass the distinct DataTable to the ViewBag
                            ViewBag.DataTable = distinctDt;

                            return View("ViewResult");
                        }

                    }
                }
            }
            catch (Exception)
            {
                ViewBag.ErrorMessage = "يرجي ادخال البيانات المطلوبة بشكل صحيح";
                return View("Failed");
            }            
        }
        private Cell CreateCell(string value)
        {
            Cell cell = new Cell(new CellValue(value));
            cell.DataType = new EnumValue<CellValues>(CellValues.String);
            return cell;
        }
        public IActionResult Delete()
        {
            return View();
        }
        [HttpGet]
        public IActionResult ExecuteDelete(string id, string invoiceNumber)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();
                    string getId = @"SELECT IsoNumber,InvoiceNumber FROM InvoiceDetails WHERE IsoNumber = @Id and InvoiceNumber = @InvoiceNumber";

                    // Execute your SQL query using the connection
                    using (SqlCommand command = new SqlCommand(getId, connection))
                    {
                        // Add parameters to the command
                        command.Parameters.AddWithValue("@Id", id);
                        command.Parameters.AddWithValue("@InvoiceNumber", invoiceNumber);

                        // Execute the command and retrieve data
                        using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                        {
                            DataTable dt = new DataTable();
                            adapter.Fill(dt);
                            // Check if the DataTable has data
                            if (!(dt.Rows.Count > 0))
                            {
                                ViewBag.ErrorMessage = "الفاتورة غير موجودة";
                                return View("Failed");
                            }
                        }
                    }
                    // Create the insert query with parameters to prevent SQL injection
                    string insertQuery = @"DELETE FROM InvoiceDetails WHERE IsoNumber = @Id and InvoiceNumber = @InvoiceNumber";

                    // Execute the insert query using a parameterized SqlCommand
                    using (SqlCommand command = new SqlCommand(insertQuery, connection))
                    {
                        // Add parameters to the command
                        command.Parameters.AddWithValue("@Id", id);
                        command.Parameters.AddWithValue("@InvoiceNumber", invoiceNumber);

                        command.ExecuteNonQuery();
                    }

                    // Redirect to another action after the insert
                    return View("DoneDelete");
                }
            }
            catch (Exception ex)
            {
                ViewBag.ErrorMessage = ex.Message;
                return View("Failed");
            }            
        }
        public IActionResult Send()
        {
            return View("ThankYou");
        }
        public IActionResult ForgotPassword()
        {
            return View();
        }
        [HttpGet]
        public IActionResult GetAllVendors()
        {
            try
            {
                string sqlQuery = "SELECT VendorName FROM \r\n    Vendors";

                var invoiceNumbers = new List<string>();

                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                    {
                        connection.Open();
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                invoiceNumbers.Add(reader["VendorName"].ToString());
                            }
                        }
                    }
                }

                return Json(invoiceNumbers);
            }
            catch (Exception ex)
            {
                return StatusCode(500, "Error fetching data: " + ex.Message);
            }            
        }        

        [HttpGet]
        public IActionResult GetVendorName(string vendorCode)
        {
            try
            {
                ViewBag.Code = vendorCode;

                // SQL query with parameterized query
                string sqlQuery = "SELECT VendorName FROM \r\n    Vendors \r\nWHERE \r\n    Id = @Code";
                string vendorName = string.Empty;
                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();

                    using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                    {
                        // Add the parameter to the command
                        command.Parameters.AddWithValue("@Code", vendorCode);

                        // Execute the query and retrieve the result
                        object result = command.ExecuteScalar();
                        if (result != null)
                        {
                            vendorName = result.ToString();
                        }
                    }
                }

                return Json(new { vendorName });
            }
            catch (Exception ex)
            {
                // Log the exception (optional)
                Console.WriteLine("Error fetching vendor name: " + ex.Message);

                // Return empty or handle the error gracefully
                return Json(new { vendorName = "test" });
            }            
        }
        [HttpGet]
        public IActionResult GetInvoiceNumbers(string isoNumber)
        {
            try
            {
                if (string.IsNullOrEmpty(isoNumber))
                    return BadRequest("ISO number is required");
                string sqlQuery = "SELECT InvoiceNumber FROM InvoiceDetails WHERE IsoNumber = @Code";

                var invoiceNumbers = new List<string>();

                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                    {
                        command.Parameters.AddWithValue("@Code", isoNumber);

                        connection.Open();
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                invoiceNumbers.Add(reader["InvoiceNumber"].ToString());
                            }
                        }
                    }
                }

                return Json(invoiceNumbers);
            }
            catch (Exception ex)
            {
                // Log or handle the exception as necessary
                return StatusCode(500, "Error fetching data: " + ex.Message);
            }            

        }
        [HttpGet]
        public IActionResult GetInvoiceNumbersByVendor(string VendorName)
        {
            if (string.IsNullOrEmpty(VendorName))
                return BadRequest("Vendor name is required");

            
            string sqlQuery = "SELECT InvoiceNumber FROM InvoiceDetails WHERE VendorName = @Code";

            var invoiceNumbers = new List<string>();

            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                {
                    command.Parameters.AddWithValue("@Code", VendorName);

                    try
                    {
                        connection.Open();
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                invoiceNumbers.Add(reader["InvoiceNumber"].ToString());
                            }
                        }
                    }
                    catch (SqlException ex)
                    {
                        // Log or handle the exception as necessary
                        return StatusCode(500, "Error fetching data: " + ex.Message);
                    }
                }
            }

            return Json(invoiceNumbers);
        }
        [HttpGet]
        public JsonResult GetInvoiceData(string isoNumber, string invoiceNumber)
        {
            if (string.IsNullOrEmpty(invoiceNumber))
            {
                return Json(new { success = false, message = "Invoice number is required." });
            }

            try
            {
                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();

                    string query = "SELECT InvoiceValue1, Currency1, InvoiceValue2, Currency2, InvoiceValue3, Currency3" +
                        ", VendorName , InvoiceDate, InvoiceReceiptDate, Requester, ToRequesterDate" +
                        ", EmployeeName, ExchangeNumber, InvoiceExchangeValue, ExchangeCurrency, IsoDate " +
                        "FROM InvoiceDetails WHERE IsoNumber = @IsoNumber and InvoiceNumber = @InvoiceNumber";

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        // Use parameterized query to prevent SQL injection
                        command.Parameters.AddWithValue("@IsoNumber", isoNumber);
                        command.Parameters.AddWithValue("@InvoiceNumber", invoiceNumber);

                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.HasRows) // Ensure there is data to read
                            {
                                // Move to the first record
                                reader.Read();
                                var invoiceData = new
                                {
                                    invoiceValue1 = reader["InvoiceValue1"]?.ToString(),
                                    currency1 = reader["Currency1"]?.ToString(),
                                    invoiceValue2 = reader["InvoiceValue2"]?.ToString(),
                                    currency2 = reader["Currency2"]?.ToString(),
                                    invoiceValue3 = reader["InvoiceValue3"]?.ToString(),
                                    currency3 = reader["Currency3"]?.ToString(),
                                    vendorName = reader["VendorName"]?.ToString(),
                                    invoiceDate = reader["InvoiceDate"] is DBNull ? null
                  : Convert.ToDateTime(reader["InvoiceDate"]).ToString("yyyy-MM-dd"),
                                    invoiceReceiptDate = reader["InvoiceReceiptDate"] is DBNull ? null
                         : Convert.ToDateTime(reader["InvoiceReceiptDate"]).ToString("yyyy-MM-dd"),
                                    requester = reader["Requester"]?.ToString(),
                                    toRequesterDate = reader["ToRequesterDate"] is DBNull ? null
                      : Convert.ToDateTime(reader["ToRequesterDate"]).ToString("yyyy-MM-dd"),
                                    employeeName = reader["EmployeeName"]?.ToString(),
                                    exchangeNumber = reader["ExchangeNumber"]?.ToString(),
                                    invoiceExchangeValue = reader["InvoiceExchangeValue"]?.ToString(),
                                    exchangeCurrency = reader["ExchangeCurrency"]?.ToString(),                                   
                                    isoDate = reader["IsoDate"] is DBNull ? null
              : Convert.ToDateTime(reader["IsoDate"]).ToString("yyyy-MM-dd")
                                };


                                return Json(new { success = true, details = invoiceData });
                            }
                            else
                            {
                                return Json(new { success = false, message = "Invoice not found." });
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "An error occurred while fetching invoice data.", error = ex.Message });
            }
        }
        [HttpGet]
        public JsonResult GetInvoiceDataByVendorAndInvoice(string vendorName, string invoiceNumber)
        {
            if (string.IsNullOrEmpty(invoiceNumber))
            {
                return Json(new { success = false, message = "Invoice number is required." });
            }
            string _invoiceNumber = invoiceNumber.Split(',')[0];
            try
            {
                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();
                    var invoiceData = new Invoicedata();
                    string query = "SELECT EmployeeName, ExchangeCurrency, BackToRequester, ToRequesterDate, Area, InvoiceEntree FROM InvoiceDetails WHERE VendorName = @VendorName and InvoiceNumber = @InvoiceNumber";
                    
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        // Use parameterized query to prevent SQL injection
                        command.Parameters.AddWithValue("@VendorName", vendorName);
                        command.Parameters.AddWithValue("@InvoiceNumber", _invoiceNumber);

                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.HasRows) // Ensure there is data to read
                            {
                                // Move to the first record
                                reader.Read();

                                invoiceData.employeeName = reader["EmployeeName"]?.ToString();
                                invoiceData.exchangeCurrency = reader["ExchangeCurrency"]?.ToString();
                                invoiceData.backToRequester = reader["BackToRequester"]?.ToString();
                                invoiceData.toRequesterDate = reader["ToRequesterDate"] is DBNull ? null
                    : Convert.ToDateTime(reader["ToRequesterDate"]).ToString("yyyy-MM-dd");
                                invoiceData.area = reader["Area"]?.ToString();
                                invoiceData.invoiceEntree = reader["InvoiceEntree"]?.ToString();



                                //return Json(new { success = true, details = invoiceData });
                            }
                            else
                            {
                                return Json(new { success = false, message = "Invoice not found." });
                            }
                        }
                    }
                    var invoiceNumbers = invoiceNumber.Split(',');
                    invoiceData.invoiceExchangeValue = 0;
                    foreach (var __invoiceNumber in invoiceNumbers)
                    {
                        string QueryToGetSumOfInvoicesExchangeValues = "SELECT InvoiceValue1, Currency1, InvoiceExchangeValue FROM InvoiceDetails WHERE VendorName = @VendorName and InvoiceNumber = @InvoiceNumber";
                        
                        using (SqlCommand command = new SqlCommand(QueryToGetSumOfInvoicesExchangeValues, connection))
                        {
                            // Use parameterized query to prevent SQL injection
                            command.Parameters.AddWithValue("@VendorName", vendorName);
                            command.Parameters.AddWithValue("@InvoiceNumber", __invoiceNumber);

                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                if (reader.HasRows) // Ensure there is data to read
                                {
                                    // Move to the first record
                                    reader.Read();
                                    if (!string.IsNullOrEmpty(reader["InvoiceExchangeValue"]?.ToString()))
                                    {
                                        invoiceData.invoiceExchangeValue = invoiceData.invoiceExchangeValue + decimal.Parse(reader["InvoiceExchangeValue"]?.ToString());
                                        invoiceData.exchangeCurrency = reader["Currency1"]?.ToString();
                                    }
                                    else
                                    {
                                        invoiceData.invoiceExchangeValue = decimal.Parse(reader["InvoiceValue1"]?.ToString());
                                        invoiceData.exchangeCurrency = reader["Currency1"]?.ToString();
                                    }

                                }
                                else
                                {
                                    return Json(new { success = false, message = "Invoice not found." });
                                }
                            }
                        }
                    }
                    return Json(new { success = true, details = invoiceData });
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "An error occurred while fetching invoice data.", error = ex.Message });
            }
        }
        [HttpGet]
        public JsonResult GetExchangeDataByVendorAndInvoice(string vendorName, string invoiceNumber)
        {
            if (string.IsNullOrEmpty(invoiceNumber))
            {
                return Json(new { success = false, message = "Invoice number is required." });
            }
            string _invoiceNumber = invoiceNumber.Split(',')[0];
            try
            {
                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();
                    var invoiceData = new Invoicedata();
                    string query = "SELECT EmployeeName, ExchangeNumber, ExchangeCurrency, BackToRequester, ToRequesterDate, Area, InvoiceEntree FROM InvoiceDetails WHERE VendorName = @VendorName and InvoiceNumber = @InvoiceNumber";

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        // Use parameterized query to prevent SQL injection
                        command.Parameters.AddWithValue("@VendorName", vendorName);
                        command.Parameters.AddWithValue("@InvoiceNumber", _invoiceNumber);

                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.HasRows) // Ensure there is data to read
                            {
                                // Move to the first record
                                reader.Read();

                                invoiceData.employeeName = reader["EmployeeName"]?.ToString();
                                invoiceData.exchangeNumber = reader["ExchangeNumber"]?.ToString();
                                invoiceData.exchangeCurrency = reader["ExchangeCurrency"]?.ToString();
                                invoiceData.backToRequester = reader["BackToRequester"]?.ToString();
                                invoiceData.toRequesterDate = reader["ToRequesterDate"] is DBNull ? null
                    : Convert.ToDateTime(reader["ToRequesterDate"]).ToString("yyyy-MM-dd");
                                invoiceData.area = reader["Area"]?.ToString();
                                invoiceData.invoiceEntree = reader["InvoiceEntree"]?.ToString();



                                //return Json(new { success = true, details = invoiceData });
                            }
                            else
                            {
                                return Json(new { success = false, message = "Invoice not found." });
                            }
                        }
                    }
                    var invoiceNumbers = invoiceNumber.Split(',');
                    invoiceData.invoiceExchangeValue = 0;
                    foreach (var __invoiceNumber in invoiceNumbers)
                    {
                        string QueryToGetSumOfInvoicesExchangeValues = "SELECT InvoiceValue1, Currency1, InvoiceExchangeValue FROM InvoiceDetails WHERE VendorName = @VendorName and InvoiceNumber = @InvoiceNumber";

                        using (SqlCommand command = new SqlCommand(QueryToGetSumOfInvoicesExchangeValues, connection))
                        {
                            // Use parameterized query to prevent SQL injection
                            command.Parameters.AddWithValue("@VendorName", vendorName);
                            command.Parameters.AddWithValue("@InvoiceNumber", __invoiceNumber);

                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                if (reader.HasRows) // Ensure there is data to read
                                {
                                    // Move to the first record
                                    reader.Read();
                                    if (!string.IsNullOrEmpty(reader["InvoiceExchangeValue"]?.ToString()))
                                    {
                                        invoiceData.invoiceExchangeValue = invoiceData.invoiceExchangeValue + decimal.Parse(reader["InvoiceExchangeValue"]?.ToString());
                                        invoiceData.exchangeCurrency = reader["Currency1"]?.ToString();
                                    }
                                    else
                                    {
                                        invoiceData.invoiceExchangeValue = decimal.Parse(reader["InvoiceValue1"]?.ToString());
                                        invoiceData.exchangeCurrency = reader["Currency1"]?.ToString();
                                    }

                                }
                                else
                                {
                                    return Json(new { success = false, message = "Invoice not found." });
                                }
                            }
                        }
                    }
                    return Json(new { success = true, details = invoiceData });
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "An error occurred while fetching invoice data.", error = ex.Message });
            }
        }
        [HttpGet]
        public JsonResult GetInvoiceDataByVendor(string vendorName, string invoiceNumber)
        {
            if (string.IsNullOrEmpty(invoiceNumber))
            {
                return Json(new { success = false, message = "Invoice number is required." });
            }

            try
            {
                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();

                    string query = "SELECT InvoiceValue1, Currency1, InvoiceValue2, Currency2, InvoiceValue3, Currency3" +
                        ", InvoiceDate, InvoiceReceiptDate, Requester, ToRequesterDate" +
                        ", EmployeeName, ExchangeNumber, InvoiceExchangeValue1, ExchangeCurrency1" +
                        ", InvoiceExchangeValue2, ExchangeCurrency2, InvoiceExchangeValue3, ExchangeCurrency3, IsoDate " +
                        "FROM InvoiceDetails WHERE VendorName = @VendorName and InvoiceNumber = @InvoiceNumber";

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        // Use parameterized query to prevent SQL injection
                        command.Parameters.AddWithValue("@VendorName", vendorName);
                        command.Parameters.AddWithValue("@InvoiceNumber", invoiceNumber);

                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.HasRows) // Ensure there is data to read
                            {
                                // Move to the first record
                                reader.Read();
                                var invoiceData = new
                                {
                                    invoiceValue1 = reader["InvoiceValue1"]?.ToString(),
                                    currency1 = reader["Currency1"]?.ToString(),
                                    invoiceValue2 = reader["InvoiceValue2"]?.ToString(),
                                    currency2 = reader["Currency2"]?.ToString(),
                                    invoiceValue3 = reader["InvoiceValue3"]?.ToString(),
                                    currency3 = reader["Currency3"]?.ToString(),
                                    invoiceDate = reader["InvoiceDate"] is DBNull ? null
                  : Convert.ToDateTime(reader["InvoiceDate"]).ToString("yyyy-MM-dd"),
                                    invoiceReceiptDate = reader["InvoiceReceiptDate"] is DBNull ? null
                         : Convert.ToDateTime(reader["InvoiceReceiptDate"]).ToString("yyyy-MM-dd"),
                                    requester = reader["Requester"]?.ToString(),
                                    toRequesterDate = reader["ToRequesterDate"] is DBNull ? null
                      : Convert.ToDateTime(reader["ToRequesterDate"]).ToString("yyyy-MM-dd"),
                                    employeeName = reader["EmployeeName"]?.ToString(),
                                    exchangeNumber = reader["ExchangeNumber"]?.ToString(),
                                    invoiceExchangeValue1 = reader["InvoiceExchangeValue1"]?.ToString(),
                                    exchangeCurrency1 = reader["ExchangeCurrency1"]?.ToString(),
                                    invoiceExchangeValue2 = reader["InvoiceExchangeValue2"]?.ToString(),
                                    exchangeCurrency2 = reader["ExchangeCurrency2"]?.ToString(),
                                    invoiceExchangeValue3 = reader["InvoiceExchangeValue3"]?.ToString(),
                                    exchangeCurrency3 = reader["ExchangeCurrency3"]?.ToString(),
                                    isoDate = reader["IsoDate"] is DBNull ? null
              : Convert.ToDateTime(reader["IsoDate"]).ToString("yyyy-MM-dd")
                                };


                                return Json(new { success = true, details = invoiceData });
                            }
                            else
                            {
                                return Json(new { success = false, message = "Invoice not found." });
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "An error occurred while fetching invoice data.", error = ex.Message });
            }
        }
        [HttpGet]
        public JsonResult GetVendorNameByInvoiceNumber(string invoiceNumber)
        {
            if (string.IsNullOrEmpty(invoiceNumber))
            {
                return Json(new { success = false, message = "Invoice number is required." });
            }

            try
            {
                using (SqlConnection connection = new SqlConnection(_connectionString))
                {
                    connection.Open();

                    string query = "SELECT VendorName FROM InvoiceDetails WHERE InvoiceNumber = @InvoiceNumber";

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        // Use parameterized query to prevent SQL injection
                        command.Parameters.AddWithValue("@InvoiceNumber", invoiceNumber);

                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            if (reader.HasRows) // Ensure there is data to read
                            {
                                // Move to the first record
                                reader.Read();
                                var invoiceData = new
                                {
                                    vendorName = reader["VendorName"]?.ToString(),                                    
                                };


                                return Json(new { success = true, details = invoiceData });
                            }
                            else
                            {
                                return Json(new { success = false, message = "Invoice not found." });
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "An error occurred while fetching invoice data.", error = ex.Message });
            }
        }
    }
}