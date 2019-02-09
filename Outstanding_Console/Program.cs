using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Threading;
using System.Configuration;
using System.Net.Mail;
using System.Net;

namespace Outstanding_Console
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "Google Sheets API .NET Quickstart";
        DataTable dt = new DataTable("MyTable1");
        SqlDataReader oReader2;
        
        public static object ConfigurationSettings { get; private set; }

        static void Main(string[] args)
        {
           
            UserCredential credential;
            string path = AppDomain.CurrentDomain.BaseDirectory;

            using (var stream = new FileStream(path+"\\"+"credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define request parameters.
            //String spreadsheetId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
            String spreadsheetId = "1YxQJtzUuhpN6PDA10YQtQ90ItGH0ppe7ZCRJGQp5M9s";
            //String range = "Class Data!A2:D";
            String range = "Report!A2:AE";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);

            // Prints the names and majors of students in a sample spreadsheet:
            // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            DataTable workTable = new DataTable("Outstanding");
            DataRow workRow;
            workTable.Columns.Add(new DataColumn("Date"));
            workTable.Columns.Add(new DataColumn("RefNo"));
            workTable.Columns.Add(new DataColumn("ClientName"));
            workTable.Columns.Add(new DataColumn("PendingAmt"));
            workTable.Columns.Add(new DataColumn("Comments"));
            workTable.Columns.Add(new DataColumn("Status"));
            workTable.Columns.Add(new DataColumn("DueDate"));
            workTable.Columns.Add(new DataColumn("Owner"));
            

            if (values != null && values.Count > 0)
            {
                //Console.WriteLine("Name, Major");
                foreach (var row in values)
                {
                    // Print columns A and E, which correspond to indices 0 and 4.
                    Console.WriteLine("{0}, {1}, {2}, {3} , {4}, {5}, {6}, {7}", row[1], row[2], row[3], row[4], row[5], row[7], row[8], row[10]);
                    if (row[7].ToString() == "Overdue" || row[7].ToString() == "Due Soon")
                    {
                        workRow = workTable.NewRow();
                        workRow[0] = row[1];
                        workRow[1] = row[2];
                        workRow[2] = row[3];
                        workRow[3] = row[4];
                        workRow[4] = row[5];
                        workRow[5] = row[7];
                        workRow[6] = row[8];
                        workRow[7] = row[10];
                        workTable.Rows.Add(workRow);
                    }
                    
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }

            var _connection = new SqlConnection();
            _connection.ConnectionString = "Data Source=192.168.1.225;Initial Catalog=Outstanding;Persist Security Info=True;User ID=sa;Password=Enviro1@!@#$";
            string query = "DELETE FROM OverDue";
            _connection.Open();
            SqlCommand DateCheck = new SqlCommand(query, _connection);
            DateCheck.ExecuteNonQuery();

            using (var bulkCopy = new SqlBulkCopy(_connection.ConnectionString, SqlBulkCopyOptions.KeepIdentity))
            {
                // my DataTable column names match my SQL Column names, so I simply made this loop. However if your column names don't match, just pass in which datatable name matches the SQL column name in Column Mappings
                foreach (DataColumn col in workTable.Columns)
                {
                    bulkCopy.ColumnMappings.Add(col.ColumnName, col.ColumnName);
                }

                bulkCopy.BulkCopyTimeout = 600;
                bulkCopy.DestinationTableName = "OverDue";
                bulkCopy.WriteToServer(workTable);
            }

            string EmpName = "SELECT DISTINCT Owner FROM OverDue ORDER BY Owner ASC";
            string emailID = "";
            string msgbody = "";
            List<string> EmpNameList = new List<string>();
            SqlCommand DateCheck2 = new SqlCommand(EmpName, _connection);
            using (SqlDataReader oReader = DateCheck2.ExecuteReader())
            {
                while (oReader.Read())
                {
                    EmpNameList.Add(oReader[0].ToString());
                }
                oReader.Close();
            }

            

            foreach (string Emp_Name in EmpNameList)
            {
                //string query3 = "SELECT Date,RefNo,ClientName,PendingAmt,Comments,Status,DueDate FROM OverDue where Owner = '" + Emp_Name + "'";
                string query3 = "SELECT OverDue.Date, OverDue.RefNo, OverDue.ClientName, OverDue.PendingAmt, OverDue.Comments, OverDue.Status, OverDue.DueDate, Email.Email FROM Email INNER JOIN OverDue ON Email.Name = OverDue.Owner where OverDue.Owner = '" + Emp_Name + "'";
                SqlCommand DateCheck3 = new SqlCommand(query3, _connection);
                using (SqlDataReader oReader2 = DateCheck3.ExecuteReader())
                {
                    string messageBody = "";
                    messageBody = "<font>The following are the Outstanding Invoices: </font><br><br>";
                    string htmlTableStart = "<table style=\"border-collapse:collapse; text-align:center;\" >";
                    string htmlTableEnd = "</table>";
                    string htmlHeaderRowStart = "<tr style =\"background-color:#6FA1D2; color:#ffffff;\">";
                    string htmlHeaderRowEnd = "</tr>";
                    string htmlTrStart = "<tr style =\"color:#555555;\">";
                    string htmlTrEnd = "</tr>";
                    string htmlTdStart = "<th style=\" border-color:#5c87b2; border-style:solid; border-width:thin; padding: 5px;\">";
                    string htmlTdEnd = "</th>";

                    messageBody += htmlTableStart;
                    messageBody += htmlHeaderRowStart;
                    messageBody += htmlTdStart + "Date " + htmlTdEnd;
                    messageBody += htmlTdStart + "Ref. No. " + htmlTdEnd;
                    messageBody += htmlTdStart + "Client Name " + htmlTdEnd;
                    messageBody += htmlTdStart + "Pending Amt " + htmlTdEnd;
                    messageBody += htmlTdStart + "Comments " + htmlTdEnd;
                    messageBody += htmlTdStart + "Status " + htmlTdEnd;
                    messageBody += htmlTdStart + "Due Date " + htmlTdEnd;

                    messageBody += htmlHeaderRowEnd;

                    while (oReader2.Read())
                    {
                        messageBody = messageBody + htmlTrStart;
                        messageBody = messageBody + htmlTdStart + oReader2[0].ToString() + htmlTdEnd;
                        messageBody = messageBody + htmlTdStart + oReader2[1].ToString() + htmlTdEnd;
                        messageBody = messageBody + htmlTdStart + oReader2[2].ToString() + htmlTdEnd;
                        messageBody = messageBody + htmlTdStart + oReader2[3].ToString() + htmlTdEnd;
                        messageBody = messageBody + htmlTdStart + oReader2[4].ToString() + htmlTdEnd;
                        messageBody = messageBody + htmlTdStart + oReader2[5].ToString() + htmlTdEnd;
                        messageBody = messageBody + htmlTdStart + oReader2[6].ToString() + htmlTdEnd;
                        emailID = oReader2[7].ToString();
                        messageBody = messageBody + htmlTrEnd;
                    }
                    messageBody = messageBody + htmlTableEnd;
                    msgbody = messageBody;
                }
                string HostAdd = "smtp.gmail.com";
                string FromEmailid = "notifications@envirosafetysolutions.in";
                string password = "Enviro@987";
                string port = "587";

                MailMessage mail = new MailMessage();
                mail.To.Add(emailID);
                //mail.To.Add("prasannapatnaikrcert@gmail.com");
                mail.From = new MailAddress(FromEmailid);
                mail.CC.Add("anup@envirosafetysolutions.in");
                mail.CC.Add("vipin@envirosafetysolutions.in");
                //mail.CC.Add("prasanna@envirosafetysolutions.in");
                mail.Subject = "Outstanding Report of " + Emp_Name;
                mail.Body = msgbody;
                mail.IsBodyHtml = true;

                SmtpClient cnt = new SmtpClient();
                cnt.UseDefaultCredentials = false;
                cnt.Host = HostAdd;
                cnt.Port = Convert.ToInt32(port);
                cnt.Credentials = new NetworkCredential(FromEmailid, password);
                cnt.EnableSsl = true;
                cnt.Send(mail);

            }

            //Console.Read();
            Environment.Exit(0);
        }
    }
}
