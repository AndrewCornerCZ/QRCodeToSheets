using System;
using System.Collections.Generic;
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
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.IO;
using System.DirectoryServices;
using System.Threading;
using System.Timers;
using System.Windows.Threading;
using static Google.Apis.Sheets.v4.SheetsService;
using System.Security.Cryptography.X509Certificates;
using Color = Google.Apis.Sheets.v4.Data.Color;
using System.Diagnostics;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace Sheets
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public string InputValue { get; private set; }

        private DispatcherTimer timer;
        public MainWindow()
        {
            InitializeComponent();
            timer = new DispatcherTimer();
            timer.Interval = TimeSpan.FromMilliseconds(1);
            timer.Tick += Timer_Tick;
        }


        private void DebounceTimerElapsed(object sender, ElapsedEventArgs e)
        {
            // Musíme pracovat v hlavním vlákně UI
            Dispatcher.Invoke(() =>
            {
                string searchCode = SearchNumberInput.Text;

                // Vyhledávání provedeme pouze, pokud uživatel něco zadal
                if (!string.IsNullOrEmpty(searchCode))
                {

                }
                else
                {
                }
            });
        }

        private void SearchNumberInput_TextChanged(object sender, TextChangedEventArgs e)
        {
            timer.Stop();
            timer.Start();
        }

        private void AuthenticateGoogleSheets()
        {
            string[] Scopes = { SheetsService.Scope.Spreadsheets };
            var service = new SheetsService(new BaseClientService.Initializer() { HttpClientInitializer = GoogleWebAuthorizationBroker.AuthorizeAsync(new ClientSecrets { ClientId = "", ClientSecret = "" }, Scopes, "user", CancellationToken.None, new FileDataStore("MyAppsToken")).Result, ApplicationName = "Google Sheets API .NET Quickstart", });

        }
        public void Timer_Tick(object sender, EventArgs e)
        {
            timer.Stop();
            PerformSearch(SearchNumberInput.Text);
        }

        private void PerformSearch(string searchText)
        {

            string spreadsheetId = "";
            string spreadsheetRangeStart = "";
            string spreadsheetRangeEnd = "";
            string spreadsheetName = "";
            string spreadsheetPalett = SpreadsheetPalett.Text;
            int spreadsheetIdsheet = 2088831864;
            int first = 0;
            try
            {
                string[] Scopes = { SheetsService.Scope.Spreadsheets };
                var service = new SheetsService(new BaseClientService.Initializer() { HttpClientInitializer = GoogleWebAuthorizationBroker.AuthorizeAsync(new ClientSecrets { ClientId = "", ClientSecret = " }, Scopes, "user", CancellationToken.None, new FileDataStore("MyAppsToken")).Result, ApplicationName = "Google Sheets API .NET Quickstart", });
                    //Smazal jsém osobní id pro soukromí
                SpreadsheetsResource.ValuesResource.GetRequest getReguest = service.Spreadsheets.Values.Get(spreadsheetId, spreadsheetName + "!" + spreadsheetRangeStart + ":" + spreadsheetRangeEnd);
                System.Net.ServicePointManager.ServerCertificateValidationCallback = delegate (object sender2, X509Certificate certificate, X509Chain chain, System.Net.Security.SslPolicyErrors sslPolicyErrors) { return true; };
                ValueRange getResponse = getReguest.Execute();
                IList<IList<Object>> values = getResponse.Values;
                var Range = $"{spreadsheetName}!" + spreadsheetRangeStart + (values.Count + 1) + ":" + spreadsheetRangeEnd + (values.Count + 1);

                var ValueRange = new ValueRange();
                int i = 0;
                int ano = 0;
                foreach (var Column in values)
                {

                    if (values != null)
                    {
                        if (Column[1].Equals(searchText) && Column[2].Equals(spreadsheetPalett) && !Column[6].Equals(Column[7]))
                        {
                            var range = new GridRange
                            {
                                SheetId = spreadsheetIdsheet,  // Změň na správné Sheet ID
                                StartRowIndex = i,
                                EndRowIndex = i + 1,
                                StartColumnIndex = 7,
                                EndColumnIndex = 8  // Změň na správný sloupec, kde chceš změnit barvu
                            };

                            ano = 1;

                            ScrollToCell($"A{i + 1}");
                            break;
                        }

                    }
                    i++;                   
                }
                if (ano == 0)
                {
                 Result.Text = "Produkt nenalezen";

                    
                }
                Range = $"{spreadsheetName}!" + spreadsheetRangeStart + i + ":" + spreadsheetRangeEnd + i;


            }
            catch (Exception ex)
            {
            }
        }

        private void OnSearchClick(object sender, RoutedEventArgs e)
        {

        }
        static int GetLetterDifference(char letter1, char letter2)
        {
            letter1 = char.ToUpper(letter1);
            letter2 = char.ToUpper(letter2);

            return Math.Abs(letter2 - letter1);
        }

        public char LowerLetterByOne(char input)
        {
            // Zkontroluje, jestli je to písmeno a jestli není A nebo a
            if (char.IsLetter(input) && input != 'A' && input != 'a')
            {
                return (char)(input - 1);  // Sníží hodnotu písmena o 1 podle ASCII tabulky
            }
            else
            {
                // Pokud je to A nebo a, nemůže být sníženo
                return input;
            }
        }
        public char UpperLetterByOne(char input)
        {
            // Zkontroluje, jestli je to písmeno a jestli není A nebo a
            if (char.IsLetter(input) && input != 'Z' && input != 'Z')
            {
                return (char)(input + 1);  // Sníží hodnotu písmena o 1 podle ASCII tabulky
            }
            else
            {
                // Pokud je to A nebo a, nemůže být sníženo
                return input;
            }
        }
        private void ScrollToCell(string cellAddress)
        {
            string url = $"{cellAddress}";

            // Nastav WebView Source na nové URL
            webView.Source = new Uri(url);
        }

        private void ListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var selectedItem = comboBox.SelectedItem as ComboBoxItem;
            string selectedValue = selectedItem.Content.ToString();
            string spreadsheetId = "";  //Odpojil jsem tabulku, kvůli soukromí firmy rodičů
            string spreadsheetRangeStart = "A";
            string spreadsheetRangeEnd = "J";
            string spreadsheetName = "Bauhaus 24.9";
            string spreadsheetPalett = SpreadsheetPalett.Text;
            string searchText = SearchNumberInput.Text;
            string result = Result.Text;
            int spreadsheetIdsheet = 2088831864;
            var orange = new Color()
            {
                Red = 1f,
                Green = 0.647f,
                Blue = 0f,
                Alpha = 1.0f
            };
            var blue = new Color()
            {
                Red = 0.678f,
                Green = 0.847f,
                Blue = 0.902f,
                Alpha = 1.0f
            };
            var green = new Color()
            {
                Red = 0f,
                Green = 1f,
                Blue = 0f,
                Alpha = 1.0f
            };
            try
            {
                string[] Scopes = { SheetsService.Scope.Spreadsheets };
                var service = new SheetsService(new BaseClientService.Initializer() { HttpClientInitializer = GoogleWebAuthorizationBroker.AuthorizeAsync(new ClientSecrets { ClientId = "", ClientSecret = "" }, Scopes, "user", CancellationToken.None, new FileDataStore("MyAppsToken")).Result, ApplicationName = "Google Sheets API .NET Quickstart", });

                SpreadsheetsResource.ValuesResource.GetRequest getReguest = service.Spreadsheets.Values.Get(spreadsheetId, spreadsheetName + "!" + spreadsheetRangeStart + ":" + spreadsheetRangeEnd);
                System.Net.ServicePointManager.ServerCertificateValidationCallback = delegate (object sender2, X509Certificate certificate, X509Chain chain, System.Net.Security.SslPolicyErrors sslPolicyErrors) { return true; };
                ValueRange getResponse = getReguest.Execute();
                IList<IList<Object>> values = getResponse.Values;
                var Range = $"{spreadsheetName}!" + spreadsheetRangeStart + (values.Count + 1) + ":" + spreadsheetRangeEnd + (values.Count + 1);
                var color = new Color();
                var ValueRange = new ValueRange();
                int i = 0;
                foreach (var Column in values)
                {

                    if (values != null)
                    {
                        if (Column[2].Equals(spreadsheetPalett) && Column[1].Equals(searchText) && !Column[6].Equals(Column[7]))
                        {
                            Result.Text = "";
                            var request = new Request();
                            var batchUpdateRequest = new BatchUpdateSpreadsheetRequest();
                            var batchUpdate = service.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadsheetId);
                            var range = new GridRange
                            {
                                SheetId = spreadsheetIdsheet,  // Změň na správné Sheet ID
                                StartRowIndex = i,
                                EndRowIndex = i + 1,
                                StartColumnIndex = 7,
                                EndColumnIndex = 8  // Změň na správný sloupec, kde chceš změnit barvu
                            };
                            if (int.TryParse(Column[6].ToString(), out int value6) && int.TryParse(Column[7].ToString(), out int value7))
                            {
                                if (value6 == value7 + 1)
                                {
                                    if (selectedValue == "Bazar")
                                    {
                                        color = blue;
                                    }
                                    if (selectedValue == "Servis")
                                    {
                                        color = green;
                                    }
                                    if (selectedValue == "A" || selectedValue == "N" || selectedValue == "B" ||selectedValue == "C")
                                    {
                                        color = orange;
                                    }
                                    request = new Request
                                    {
                                        RepeatCell = new RepeatCellRequest
                                        {
                                            Cell = new CellData
                                            {
                                                UserEnteredFormat = new CellFormat
                                                {
                                                    BackgroundColor = color
                                                }
                                            },
                                            Range = range,
                                            Fields = "userEnteredFormat.backgroundColor"
                                        }

                                    };
                                    batchUpdateRequest = new BatchUpdateSpreadsheetRequest
                                    {
                                        Requests = new List<Request> { request }
                                    };


                                    batchUpdate = service.Spreadsheets.BatchUpdate(batchUpdateRequest, spreadsheetId);
                                    batchUpdate.Execute();

                                }   
                            }


                            var updateRange = $"{spreadsheetName}!H{i + 1}";  // Odkaz na buňku posledního sloupce
                            var lastColumnValueRequest = service.Spreadsheets.Values.Get(spreadsheetId, updateRange);
                            var categoryidetificator = lastColumnValueRequest;// Získání hodnoty z posledního sloupce
                            ValueRange lastColumnValueResponse = lastColumnValueRequest.Execute();
                            IList<IList<Object>> lastColumnValues = lastColumnValueResponse.Values;
                            int currentValue = 0;  // Výchozí hodnota

                            if (lastColumnValues != null && lastColumnValues.Count > 0 && lastColumnValues[0].Count > 0)
                            {
                                int.TryParse(lastColumnValues[0][0].ToString(), out currentValue);  // Převedení hodnoty na int
                            }
                            currentValue++;  // Zvýšení hodnoty o 1

                            var updateValue = new ValueRange
                            {
                                Values = new List<IList<object>> { new List<object> { currentValue } }  // Nová hodnota
                            };

                            var updateRequest = service.Spreadsheets.Values.Update(updateValue, spreadsheetId, updateRange);
                            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                            updateRequest.Execute();
                            updateRange = $"{spreadsheetName}!I{i + 1}";
                            lastColumnValueRequest = service.Spreadsheets.Values.Get(spreadsheetId, updateRange);
                            ValueRange lastColumnResponse = lastColumnValueRequest.Execute(); // Získání hodnoty

                            // Získání hodnoty jako string z odpovědi
                            string categoryvalue = lastColumnResponse.Values != null && lastColumnResponse.Values.Count > 0
                                ? lastColumnResponse.Values[0][0].ToString()
                                : "";
                            if (currentValue == 1)
                            {
                                categoryvalue = selectedValue;
                            }
                            else
                            {
                                categoryvalue += "," + selectedValue;
                            }
                            updateValue = new ValueRange
                            {
                                Values = new List<IList<object>> { new List<object> { categoryvalue } }  // Nová hodnota
                            };
                            updateRequest = service.Spreadsheets.Values.Update(updateValue, spreadsheetId, updateRange);
                            updateRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                            updateRequest.Execute();
                            ScrollToCell($"A{i + 1}");
                            SearchNumberInput.Text = "";
                            SearchNumberInput.Focus();
                            break;
                        }
                        i++;

                    }



                    Result.Text = "Produkt nenalezen";
                }

            }
            catch
            {

            }

        }

    }   
}
