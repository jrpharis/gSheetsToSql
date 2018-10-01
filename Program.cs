using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Threading;
using System.Text;

namespace SheetsQuickstart
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "Google Sheets API .NET Quickstart";

        static void Main(string[] args)
        {
            UserCredential credential;

            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
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
            String spreadsheetId = "1xeRwJHbs4T10akqx5LNRQHobx7CQe1_C4PTv_mb3I_E";
            String range = "Sheet1!A2:G";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                    service.Spreadsheets.Values.Get(spreadsheetId, range);

            // Prints the names and majors of students in a sample spreadsheet:
            // https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
            ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;
            if (values != null && values.Count > 0)
            {
                List<DriverModel> drivers = DriverModel.Convert(values);

                using (var conn = new SqlConnection("Data Source=localhost;Initial Catalog=playground;Integrated Security=True"))
                {
                    foreach (var driver in drivers)
                    {
                        SqlCommand insert = new SqlCommand("insert into Drivers(DriverID, PayPeriod, Qualify, PaperWork, Expiration, Complaint, UpdateTime) " +
                                                           "values(@driverId, @payPeriod, @qualify, @paperWork, @expiration, @complaint, @updateTime)", conn);
                        insert.Parameters.AddWithValue("@driverId", driver.DriverID);
                        insert.Parameters.AddWithValue("@payPeriod", driver.PayPeriod);
                        insert.Parameters.AddWithValue("@qualify", driver.Qualify);
                        insert.Parameters.AddWithValue("@paperWork", driver.PaperWork);
                        insert.Parameters.AddWithValue("@expiration", driver.Expiration);
                        insert.Parameters.AddWithValue("@complaint", driver.Complaint);
                        insert.Parameters.AddWithValue("@updateTime", driver.UpdateTime);
                        try
                        {
                            conn.Open();
                            insert.ExecuteNonQuery();
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error inserting record for {driver.DriverID} - {ex.Message}");
                        }
                        finally
                        {
                            conn.Close();
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine("No data found.");
            }
            Console.Read();
        }
    }
}
