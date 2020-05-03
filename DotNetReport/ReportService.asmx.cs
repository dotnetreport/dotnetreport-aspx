using ReportBuilder.Web.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Web.Script.Serialization;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Security;
using Newtonsoft.Json;
using System.Web.Script.Services;
using System.Web.Services;

namespace ReportBuilder.Demo.WebForms.DotNetReport
{
    /// <summary>
    /// Summary description for ReportService
    /// </summary>
    [WebService(Namespace = "http://tempuri.org/")]
    [WebServiceBinding(ConformsTo = WsiProfiles.BasicProfile1_1)]
    [System.ComponentModel.ToolboxItem(false)]
    [System.Web.Script.Services.ScriptService]
    public class ReportService : System.Web.Services.WebService    
    {
        private DotNetReportSettings GetSettings()
        {
            var settings = new DotNetReportSettings
            {
                ApiUrl = ConfigurationManager.AppSettings["dotNetReport.apiUrl"],
                AccountApiToken = ConfigurationManager.AppSettings["dotNetReport.accountApiToken"], // Your Account Api Token from your http://dotnetreport.com Account
                DataConnectApiToken = ConfigurationManager.AppSettings["dotNetReport.dataconnectApiToken"] // Your Data Connect Api Token from your http://dotnetreport.com Account
            };

            // Populate the values below using your Application Roles/Claims if applicable

            settings.ClientId = "";  // You can pass your multi-tenant client id here to track their reports and folders
            settings.UserId = ""; // You can pass your current authenticated user id here to track their reports and folders            
            settings.UserName = "";
            settings.CurrentUserRole = new List<string>(); // Populate your current authenticated user's roles

            settings.Users = new List<string>(); // Populate all your application's user, ex  { "Jane", "John" }
            settings.UserRoles = new List<string>(); // Populate all your application's user roles, ex  { "Admin", "Normal" }       
            settings.CanUseAdminMode = true; // Set to true only if current user can use Admin mode to setup reports and dashboard

            // An example of populating Roles using MVC web security if available
            if (Roles.Enabled && User.Identity.IsAuthenticated)
            {
                settings.UserId = User.Identity.Name;
                settings.CurrentUserRole = Roles.GetRolesForUser(User.Identity.Name).ToList();

                settings.Users = Roles.GetAllRoles().SelectMany(x => Roles.GetUsersInRole(x)).ToList();
                settings.UserRoles = Roles.GetAllRoles().ToList();
            }

            return settings;
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public object GetLookupList(string lookupSql, string connectKey)
        {
            var sql = DotNetReportHelper.Decrypt(lookupSql);

            // Uncomment if you want to restrict max records returned
            sql = sql.Replace("SELECT ", "SELECT TOP 500 ");

            var json = new StringBuilder();
            var dt = new DataTable();
            using (var conn = new SqlConnection(ConfigurationManager.ConnectionStrings[connectKey].ConnectionString))
            {
                conn.Open();
                var command = new SqlCommand(sql, conn);
                var adapter = new SqlDataAdapter(command);

                adapter.Fill(dt);
            }

            int i = 0;
            foreach (DataRow dr in dt.Rows)
            {
                json.AppendFormat("{{\"id\": \"{0}\", \"text\": \"{1}\"}}{2}", dr[0], dr[1], i != dt.Rows.Count - 1 ? "," : "");
                i += 1;
            }

            return ((new JavaScriptSerializer()).DeserializeObject("[" + json.ToString() + "]"));
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public object RunReportApi(string method, bool SaveReport, string ReportJson, bool adminMode)
        {
            return CallReportApi(method, (new JavaScriptSerializer()).Serialize(new DotNetReportApiCall
            {
                Method = method,
                ReportJson = ReportJson,
                SaveReport = SaveReport,
                adminMode = adminMode
            }));
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public object CallReportApi(string method, string model)
        {
            using (var client = new HttpClient())
            {
                var settings = GetSettings();
                var keyvalues = new List<KeyValuePair<string, string>>
                {
                    new KeyValuePair<string, string>("account", settings.AccountApiToken),
                    new KeyValuePair<string, string>("dataConnect", settings.DataConnectApiToken),
                    new KeyValuePair<string, string>("clientId", settings.ClientId),
                    new KeyValuePair<string, string>("userId", settings.UserId),
                    new KeyValuePair<string, string>("userRole", String.Join(",", settings.CurrentUserRole))
                };

                var data = JsonConvert.DeserializeObject<Dictionary<string, dynamic>>(model);
                foreach (var key in data.Keys)
                {
                    if (key != "adminMode" || (key == "adminMode" && settings.CanUseAdminMode))
                    {
                        keyvalues.Add(new KeyValuePair<string, string>(key, data[key].ToString()));
                    }
                }

                var content = new FormUrlEncodedContent(keyvalues);
                var response = client.PostAsync(new Uri(settings.ApiUrl + method), content).Result;
                var stringContent = response.Content.ReadAsStringAsync().Result;

                Context.Response.StatusCode = (int)response.StatusCode;
                return (new JavaScriptSerializer()).Deserialize<dynamic>(stringContent);
            }

        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public DotNetReportResultModel RunReport(string reportSql, string connectKey, string reportType, int pageNumber = 1, int pageSize = 50, string sortBy = null, bool desc = false)
        {
            var sql = DotNetReportHelper.Decrypt(reportSql);

            try
            {
                if (!String.IsNullOrEmpty(sortBy))
                {
                    if (sortBy.StartsWith("DATENAME(MONTH, "))
                    {
                        sortBy = sortBy.Replace("DATENAME(MONTH, ", "MONTH(");
                    }
                    if (sortBy.StartsWith("MONTH(") && sortBy.Contains(")) +") && sql.Contains("Group By"))
                    {
                        sortBy = sortBy.Replace("MONTH(", "CONVERT(VARCHAR(3), DATENAME(MONTH, ");
                    }
                    sql = sql.Substring(0, sql.IndexOf("ORDER BY")) + "ORDER BY " + sortBy + (desc ? " DESC" : "");
                }

                // Execute sql
                var dt = new DataTable();
                var dtPaged = new DataTable();
                using (var conn = new SqlConnection(ConfigurationManager.ConnectionStrings[connectKey].ConnectionString))
                {
                    conn.Open();
                    var command = new SqlCommand(sql, conn);
                    var adapter = new SqlDataAdapter(command);

                    adapter.Fill(dt);
                }

                dtPaged = (dt.Rows.Count > 0) ? dtPaged = dt.AsEnumerable().Skip((pageNumber - 1) * pageSize).Take(pageSize).CopyToDataTable() : dt;

                var model = new DotNetReportResultModel
                {
                    ReportData = DataTableToDotNetReportDataModel(dtPaged, sql),
                    Warnings = GetWarnings(sql),
                    ReportSql = sql,
                    ReportDebug = Context.Request.Url.Host.Contains("localhost"),
                    Pager = new DotNetReportPagerModel
                    {
                        CurrentPage = pageNumber,
                        PageSize = pageSize,
                        TotalRecords = dt.Rows.Count,
                        TotalPages = (int)((dt.Rows.Count / pageSize) + 1)
                    }
                };

                return model;

            }

            catch (Exception ex)
            {
                return new DotNetReportResultModel
                {
                    ReportSql = sql,
                    HasError = true,
                    Exception = ex.Message
                };
            }
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public object GetDashboards(bool adminMode = false)
        {
            var settings = GetSettings();

            using (var client = new HttpClient())
            {
                var content = new FormUrlEncodedContent(new[]
                {
                    new KeyValuePair<string, string>("account", settings.AccountApiToken),
                    new KeyValuePair<string, string>("dataConnect", settings.DataConnectApiToken),
                    new KeyValuePair<string, string>("clientId", settings.ClientId),
                    new KeyValuePair<string, string>("userId", settings.UserId),
                    new KeyValuePair<string, string>("userRole", String.Join(",", settings.CurrentUserRole)),
                    new KeyValuePair<string, string>("adminMode", adminMode.ToString()),
                });

                var response = client.PostAsync(new Uri(settings.ApiUrl + $"/ReportApi/GetDashboards"), content).Result;
                var stringContent = response.Content.ReadAsStringAsync().Result;

                Context.Response.StatusCode = (int)response.StatusCode;
                return (new JavaScriptSerializer()).Deserialize<dynamic>(stringContent);
            }
        }


        [WebMethod]
        public void DownloadExcel(string reportSql, string connectKey, string reportName)
        {

            var excel = DotNetReportHelper.GetExcelFile(reportSql, connectKey, reportName);
            Context.Response.ClearContent();

            Context.Response.AddHeader("content-disposition", "attachment; filename=" + reportName + ".xlsx");
            Context.Response.ContentType = "application/vnd.ms-excel";
            Context.Response.BinaryWrite(excel);
            Context.Response.End();
            
        }

        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public void DownloadXml(string reportSql, string connectKey, string reportName)
        {

            var xml = DotNetReportHelper.GetXmlFile(reportSql, connectKey, reportName);
            Context.Response.ClearContent();

            Context.Response.AddHeader("content-disposition", "attachment; filename=" + reportName + ".xml");
            Context.Response.ContentType = "application/xml";
            Context.Response.Write(xml);
            Context.Response.End();
            
        }


        [WebMethod]
        [ScriptMethod(ResponseFormat = ResponseFormat.Json)]
        public object GetUsersAndRoles()
        {
            var settings = GetSettings();
            return new
            {
                users = settings.CanUseAdminMode ? settings.Users : new List<string>(),
                userRoles = settings.CanUseAdminMode ? settings.UserRoles : new List<string>(),
                currentUserId = settings.UserId,
                currentUserRoles = settings.UserRoles,
                currentUserName = settings.UserName,
                allowAdminMode = settings.CanUseAdminMode
            };
        }

        private string GetWarnings(string sql)
        {
            var warning = "";
            if (sql.ToLower().Contains("cross join"))
            {
                warning += "Some data used in this report have relations that are not setup properly, so data might duplicate incorrectly.<br/>";
            }

            return warning;
        }

        public static bool IsNumericType(Type type)
        {

            switch (Type.GetTypeCode(type))
            {
                case TypeCode.Int16:
                case TypeCode.UInt16:
                case TypeCode.Int32:
                case TypeCode.UInt32:
                case TypeCode.Int64:
                case TypeCode.UInt64:
                case TypeCode.Single:
                case TypeCode.Double:
                case TypeCode.Decimal:
                    return true;

                case TypeCode.Boolean:
                case TypeCode.DateTime:
                case TypeCode.String:
                default:
                    return false;
            }
        }

        public static string GetLabelValue(DataColumn col, DataRow row)
        {
            switch (Type.GetTypeCode(col.DataType))
            {
                case TypeCode.Int16:
                case TypeCode.UInt16:
                case TypeCode.Int32:
                case TypeCode.UInt32:
                case TypeCode.Int64:
                case TypeCode.UInt64:
                case TypeCode.Single:
                    return row[col].ToString();

                case TypeCode.Double:
                case TypeCode.Decimal:
                    return @row[col].ToString();// "'" + (Convert.ToDouble(@row[col].ToString()).ToString("C")) + "'";

                case TypeCode.Boolean:
                    return (Convert.ToBoolean(@row[col]) ? "Yes" : "No");

                case TypeCode.DateTime:
                    try
                    {
                        return "'" + @Convert.ToDateTime(@row[col]).ToShortDateString() + "'";
                    }
                    catch
                    {
                        return "'" + @row[col] + "'";
                    }

                case TypeCode.String:
                default:
                    return "'" + @row[col].ToString().Replace("'", "") + "'";
            }
        }

        public static string GetFormattedValue(DataColumn col, DataRow row)
        {
            if (@row[col] != null)
            {
                switch (Type.GetTypeCode(col.DataType))
                {
                    case TypeCode.Int16:
                    case TypeCode.UInt16:
                    case TypeCode.Int32:
                    case TypeCode.UInt32:
                    case TypeCode.Int64:
                    case TypeCode.UInt64:
                    case TypeCode.Single:
                        return row[col].ToString();


                    case TypeCode.Double:
                    case TypeCode.Decimal:
                        return col.ColumnName.Contains("%")
                            ? (Convert.ToDouble(row[col].ToString()) / 100).ToString("P2")
                            : Convert.ToDouble(row[col].ToString()).ToString("C");


                    case TypeCode.Boolean:
                        return (Convert.ToBoolean(row[col]) ? "Yes" : "No");


                    case TypeCode.DateTime:
                        try
                        {
                            return Convert.ToDateTime(row[col]).ToShortDateString();
                        }
                        catch
                        {
                            return row[col] != null ? row[col].ToString() : null;
                        }

                    case TypeCode.String:
                    default:
                        if (row[col].ToString() == "System.Byte[]")
                        {

                            return "<img src=\"data:image/png;base64," + Convert.ToBase64String((byte[])row[col], 0, ((byte[])row[col]).Length) + "\" style=\"max-width: 200px;\" />";
                        }
                        else
                        {
                            return row[col].ToString();
                        }

                }
            }
            return "";
        }

        private DotNetReportDataModel DataTableToDotNetReportDataModel(DataTable dt, string sql)
        {
            var model = new DotNetReportDataModel
            {
                Columns = new List<DotNetReportDataColumnModel>(),
                Rows = new List<DotNetReportDataRowModel>()
            };

            sql = sql.Substring(0, sql.IndexOf("FROM")).Replace("SELECT", "").Trim();
            var sqlFields = Regex.Split(sql, "], (?![^\\(]*?\\))").Where(x => x != "CONVERT(VARCHAR(3)").ToArray();

            int i = 0;
            foreach (DataColumn col in dt.Columns)
            {
                var sqlField = sqlFields[i++];
                model.Columns.Add(new DotNetReportDataColumnModel
                {
                    SqlField = sqlField.Substring(0, sqlField.IndexOf("AS")).Trim(),
                    ColumnName = col.ColumnName,
                    DataType = col.DataType.ToString(),
                    IsNumeric = IsNumericType(col.DataType)
                });

            }

            foreach (DataRow row in dt.Rows)
            {
                i = 0;
                var items = new List<DotNetReportDataRowItemModel>();

                foreach (DataColumn col in dt.Columns)
                {
                    items.Add(new DotNetReportDataRowItemModel
                    {
                        Column = model.Columns[i],
                        Value = row[col] != null ? row[col].ToString() : null,
                        FormattedValue = GetFormattedValue(col, row),
                        LabelValue = GetLabelValue(col, row)
                    });
                    i += 1;
                }

                model.Rows.Add(new DotNetReportDataRowModel
                {
                    Items = items.ToArray()
                });
            }

            return model;
        }


    }

}
