using FollowInvoices.Utilities;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Options;
using System.Data.SqlClient;
using System.Security.Claims;

namespace FollowInvoices.Controllers
{
    public class AccountController : Controller
    {
        // Mock user store for demonstration purposes.
        private readonly Dictionary<string, string> _users = new Dictionary<string, string>
        {
            { "حسام حسن", "P@$$w0rd" },
            { "user", "password" },
            { "Sally", "Sally@123" },
            { "Walid", "Walid@123" }
        };

        //public IActionResult Login()
        //{
        //    return View();
        //}

        //[HttpPost]
        //public async Task<JsonResult> Login(string username, string password)
        //{
        //    // Validate input
        //    if (string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password))
        //    {
        //        return Json(new { success = false, message = ".اسم المستخدم و الرقم السري مطلوبان" });
        //    }

        //    // Check user credentials
        //    if (_users.TryGetValue(username, out var storedPassword) && storedPassword == password)
        //    {
        //        // Create claims for the authenticated user
        //        var claims = new List<Claim>
        //        {
        //            new Claim(ClaimTypes.Name, username),
        //            new Claim(ClaimTypes.Role, "User") // Example role claim
        //        };

        //        // Create the identity and principal for cookie authentication
        //        var identity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);
        //        var principal = new ClaimsPrincipal(identity);

        //        // Sign in the user
        //        await HttpContext.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, principal);

        //        return Json(new { success = true });
        //    }

        //    return Json(new { success = false, message = "بيانات الدخول غير صحيحة." });
        //}

        //public async Task<IActionResult> Logout()
        //{
        //    // Clear any session data if necessary
        //    HttpContext.Session.Clear();

        //    // Sign the user out from cookie-based authentication
        //    await HttpContext.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);

        //    // Redirect the user to the Login page
        //    return RedirectToAction("Login", "Account");
        //}

        //private readonly string _connectionString = "Data Source=localhost\\SQLEXPRESS;Initial Catalog=FollowInvoices;Integrated Security=True;";
        private readonly IConfiguration _configuration;
        private readonly string _connectionString;
        public AccountController(IConfiguration configuration)
        {
            _configuration = configuration;
            _connectionString = _configuration.GetConnectionString("DefaultConnection");
        }
        public IActionResult Login()
		{
			return View();
		}

		[HttpPost]
		public async Task<JsonResult> Login(string username, string password)
		{
			Console.WriteLine("username is: " + username);
			Console.WriteLine("password is: " + password);
			username = username.ToLower().Trim();
            if (string.IsNullOrWhiteSpace(username) || string.IsNullOrWhiteSpace(password))
			{
				return Json(new { success = false, message = "اسم المستخدم والرقم السري مطلوبان." });
			}

			using (var conn = new SqlConnection(_connectionString))
			{
				conn.Open();
				Console.WriteLine("connected to " + _connectionString);
				var cmd = new SqlCommand("SELECT * FROM Users WHERE Username = @username AND PasswordHash = @password", conn);
				cmd.Parameters.AddWithValue("@username", username);
				cmd.Parameters.AddWithValue("@password", password); // ملاحظة: يفضل تستخدم التشفير لاحقًا

				using (var reader = cmd.ExecuteReader())
				{
					if (reader.Read())
					{
						var fullName = reader["FullName"].ToString();
						var department = reader["DepartmentName"].ToString();
						var displayName = reader["DisplayName"].ToString();
						Console.WriteLine(displayName);
						var claims = new List<Claim>
				{
					new Claim(ClaimTypes.Name, displayName),
					new Claim("FullName", fullName),
					new Claim("Department", department),
					new Claim(ClaimTypes.Role, "User")
				};

						var identity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme);
						var principal = new ClaimsPrincipal(identity);

						await HttpContext.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, principal);

						return Json(new { success = true });
					}
				}
			}

			return Json(new { success = false, message = "بيانات الدخول غير صحيحة." });
		}
		public IActionResult ChangePasswordView()
		{
			return View();
		}

		[HttpPost]
		public async Task<JsonResult> ChangePassword(string currentPassword, string newPassword)
		{
			var username = User.Identity?.Name;
			if (string.IsNullOrEmpty(username))
			{
				return Json(new { success = false, message = "يجب تسجيل الدخول أولاً." });
			}

			using (SqlConnection conn = new SqlConnection(_connectionString))
			{
				await conn.OpenAsync();
				var command = new SqlCommand("SELECT PasswordHash FROM Users WHERE DisplayName = @username", conn);
				command.Parameters.AddWithValue("@username", username);
				var storedPassword = (string?)await command.ExecuteScalarAsync();

				if (storedPassword != currentPassword)
				{
					return Json(new { success = false, message = "الرقم السري الحالي غير صحيح." });
				}

				var updateCommand = new SqlCommand("UPDATE Users SET PasswordHash = @newPassword WHERE DisplayName = @username", conn);
				updateCommand.Parameters.AddWithValue("@newPassword", newPassword);
				updateCommand.Parameters.AddWithValue("@username", username);
				await updateCommand.ExecuteNonQueryAsync();

				return Json(new { success = true });
			}
		}

		public async Task<IActionResult> Logout()
		{
			// Clear any session data if necessary
			HttpContext.Session.Clear();

			// Sign the user out from cookie-based authentication
			await HttpContext.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme);

			// Redirect the user to the Login page
			return RedirectToAction("Login", "Account");
		} 
	}
}