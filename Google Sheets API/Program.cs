using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace GoogleSheetsAPI4_v1console
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.Spreadsheets }; // static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "<MYSpreadsheet>";
        static string[,] Data = NewData();
        static void Main(string[] args)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");

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
                ApplicationName = "Jopa",
            });



            String spreadsheetId2 = "14C4nf4MBafsc52Fb1toQ_5xqCjItB73G7it7-SNXgLw";
            String range2 = "A1:C";  // update cell F5 
            ValueRange valueRange = new ValueRange();
            valueRange.MajorDimension = "COLUMNS";//"ROWS";//COLUMNS

            FillSpreadSheet(service, spreadsheetId2);

            Console.WriteLine("done!");
        }

        private static void FillSpreadSheet(SheetsService service, string spreadsheetId)
        {
            List<Request> requests = new List<Request>();
            for (int i = 0; i < Data.GetLength(0); i++)
            {
                List<CellData> values = new List<CellData>();

                for (int j = 0; j < Data.GetLength(1); j++)
                {
                    values.Add(new CellData
                    {
                        UserEnteredValue = new ExtendedValue
                        {
                            StringValue = Data[i, j]
                        }
                    });
                }

                requests.Add(
                    new Request
                    {
                        UpdateCells = new UpdateCellsRequest
                        {
                            Start = new GridCoordinate
                            {
                                SheetId = 0,
                                RowIndex = i,
                                ColumnIndex = 0
                            },
                            Rows = new List<RowData> { new RowData { Values = values } },
                            Fields = "userEnteredValue"
                        }
                    });
            }

            BatchUpdateSpreadsheetRequest busr = new BatchUpdateSpreadsheetRequest
            {
                Requests = requests
            };

            service.Spreadsheets.BatchUpdate(busr, spreadsheetId).Execute();
        }

        public static string[,] NewData()
        {
            string[,] Data = new string[,]
            {
            {"11", "12", "13" },
            {"21", "22", "23" },
            {"31", "32", "33" }
            };

            return Data;
        }
    }
}