namespace FirstExcelAddIn
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using Excel = Microsoft.Office.Interop.Excel;
    using www.mnb.hu.webservices;
    using System.ServiceModel;
    using System.Xml;
    using System.Windows.Forms;
    using Microsoft.Office.Interop.Excel;

    using FirstExcelAddIn.Models;

    public partial class ThisAddIn
    {
        private MNBArfolyamServiceSoapClient client;
        private IList<CurrencyWrapper> CurrenciesRetrieved { get; set; }
        private IList<FxDateWrapper> CurrencySnapshotDate { get; set; }

        private const int ROW_OFFSET_FOR_DATES = 3;
        private const int COLUMN_OFFSET_FOR_CURRENCIES = 2;
        private const string MNB_DATE_ATTRIBUTE_NAME = "date";
        private const string MNB_CURRENCY_ATTRIBUTE_NAME = "curr";
        private const string MNB_UNIT_ATTRIBUTE_NAME = "unit";
        private const string MNB_EXCHANGE_RATES_PER_DAY = "MNBExchangeRates/Day";
        private const string MNB_CURRENCIES_PER_CURRENCIES_PER_CURR = "MNBCurrencies/Currencies/Curr";
        private const string DATE_REQUEST_FORMAT = "yyyy-MM-dd";
        private const string DATE_OUTPUT_FORMAT = "yyyy.MM.dd.";
        private const string MNB_ENDPOINT_ADDRESS = "http://www.mnb.hu/arfolyamok.asmx";

        private DateTime startDate = new DateTime(2015, 1, 1);
        private DateTime endDate = new DateTime(2020, 4, 1);

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Upon running Excel - Open a blank new Workbook
            this.Application.Workbooks.Add();

            // Initialize the client
            var binding = new BasicHttpBinding();
            binding.MaxReceivedMessageSize = Int32.MaxValue;
            binding.MaxBufferSize = Int32.MaxValue;
            var endpointAddress = new System.ServiceModel.EndpointAddress(MNB_ENDPOINT_ADDRESS);
            client = new MNBArfolyamServiceSoapClient(binding, endpointAddress);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            client.Close();
        }

        #region MNB Excel Methods

        /// <summary>
        /// Method being called when user clicks the button on the ribbon to get MNB Exchange rates
        /// </summary>
        public void RetrieveDataFromMNB()
        {
            CurrenciesRetrieved = new List<CurrencyWrapper>();
            CurrencySnapshotDate = new List<FxDateWrapper>();

            // Rename the active worksheet to the current date
            Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);
            activeWorksheet.Name = DateTime.Now.ToString(DATE_REQUEST_FORMAT);

            // Retrieve available currencies from MNB
            GetAndProcessAvailableCurrencies();

            // Retrieve Exchange rates for given days for available currencies
            GetAndProcessExchangeRates();

            // Print header for currencies
            PrintCurrenciesHeader(activeWorksheet);

            // PrintDateTime Header
            PrintFxRateDateRangeHeader(activeWorksheet);

            // Print the fxRates
            PrintFxRates(activeWorksheet);

            // Print the base titles
            Excel.Range firstCells = activeWorksheet.Range[activeWorksheet.Cells[1, 1], activeWorksheet.Cells[2,1]];
            string[,] titles = new string[,] { { "Dátum/ISO" }, { "Egység" } };
            firstCells.Value = titles;

            firstCells.EntireColumn.AutoFit();
        }

        /// <summary>
        /// Use MNB Client to retrieve Available Currencies
        /// </summary>
        private void GetAndProcessAvailableCurrencies()
        {
            try
            {
                var currencyResponse = client.GetCurrencies(new GetCurrenciesRequestBody()).GetCurrenciesResult;

                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(currencyResponse);

                XmlNodeList nodes = xmlDoc.SelectNodes(MNB_CURRENCIES_PER_CURRENCIES_PER_CURR);

                foreach (XmlNode node in nodes)
                {
                    // Skip if node is null or empty
                    if (string.IsNullOrWhiteSpace(node.InnerText)) continue;

                    CurrencyWrapper currencyWrapper = new CurrencyWrapper()
                    {
                        CurrencyName = node.InnerText
                    };

                    bool currencyAlreadyPresent = CurrenciesRetrieved.Any(curr => curr.CurrencyName == node.InnerText);
                    if (!currencyAlreadyPresent)
                    {
                        CurrenciesRetrieved.Add(currencyWrapper);
                    }
                }
            }
            catch (XmlException)
            {
                ShowErrorMessage("Could not parse the response from MNB while querying available currencies");
            }
            catch (System.Xml.XPath.XPathException e)
            {
                ShowErrorMessage($"XML structure is inconsistent with assumption. Could not retrieve structure of {MNB_CURRENCIES_PER_CURRENCIES_PER_CURR}");
            }
            catch (Exception e)
            {
                ShowErrorMessage(e.Message);
            }
        }

        /// <summary>
        /// Query the MNB Client to retrieve Exchange rates for given currencies
        /// </summary>
        private void GetAndProcessExchangeRates()
        {
            GetAndProcessExchangeRates(startDate, endDate, CurrenciesRetrieved.Select(name => name.CurrencyName).ToList());
        }

        /// <summary>
        /// Query the MNB Client to retrieve Exchange rates for given currencies
        /// </summary>
        /// <param name="start">DateTime: The start date to be used in the query</param>
        /// <param name="end">DateTime: The end date to be used in the query</param>
        /// <param name="currencyNames">IList<string>: The names of the currencies to be used in the query</param>
        private void GetAndProcessExchangeRates(DateTime start, DateTime end, IList<string> currencyNames)
        {
            try
            {
                string startDate = start.ToString(DATE_REQUEST_FORMAT);
                string endDate = end.ToString(DATE_REQUEST_FORMAT);

                var currencies = string.Join(",", currencyNames);

                GetExchangeRatesRequestBody requestBody = new GetExchangeRatesRequestBody()
                {
                    startDate = startDate,
                    endDate = endDate,
                    currencyNames = currencies
                };

                GetExchangeRatesResponseBody exchangeRatesResponseBody   = client.GetExchangeRates(requestBody);

                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(exchangeRatesResponseBody.GetExchangeRatesResult);

                XmlNodeList rawNodes = xmlDoc.SelectNodes(MNB_EXCHANGE_RATES_PER_DAY);
                foreach (XmlNode rawNode in rawNodes)
                {
                    // Retrieve the day of the FX rate and store it in variable: date
                    XmlNode rawDate = rawNode.Attributes.GetNamedItem(MNB_DATE_ATTRIBUTE_NAME);
                    DateTime.TryParse(rawDate.Value, out DateTime date);

                    // check if date is not included
                    var dateIsAlreadyAdded = CurrencySnapshotDate.Any(item => item.Date == date);
                    if (!dateIsAlreadyAdded)
                    {
                        var newFxDate = new FxDateWrapper()
                        {
                            Date = date,
                        };
                        CurrencySnapshotDate.Add(newFxDate);
                    }

                    // Iterate through child nodes and retrieve the unit and curr from attributes - and fx rate from innerText
                    foreach (XmlNode currencyRate in rawNode.ChildNodes)
                    {
                        // Find currency by name
                        XmlNode currAttrib = currencyRate.Attributes.GetNamedItem(MNB_CURRENCY_ATTRIBUTE_NAME);
                        var currencyWrapper = CurrenciesRetrieved.FirstOrDefault(currName => currName.CurrencyName.Equals(currAttrib.Value, StringComparison.InvariantCultureIgnoreCase));

                        if (currencyWrapper is null) continue;

                        // feed currency unit to currencyWrapper
                        XmlNode unitAttrib = currencyRate.Attributes.GetNamedItem(MNB_UNIT_ATTRIBUTE_NAME);
                        if (Int32.TryParse(unitAttrib.Value, out int result))
                        {
                            currencyWrapper.RateUnit = result;
                        }

                        // Feed rate into currencyWrapper
                        currencyWrapper.CurrencyRates.Add(new Currency()
                        {
                            Date = date,
                            RawRate = currencyRate.InnerText
                        });
                    }
                }
            }
            catch (Exception e)
            {
                string error = e.Message;
                ShowErrorMessage(error);
            }
        }

        /// <summary>
        /// Prints the currencies which were requested during the query to the first row offset by default value
        /// </summary>
        /// <param name="activeWorksheet">Excel:Worksheet: worksheet where the data gets copied</param>
        private void PrintCurrenciesHeader(Excel.Worksheet activeWorksheet)
        {
            Excel.Range currencyHeader = activeWorksheet.Range[activeWorksheet.Cells[1, COLUMN_OFFSET_FOR_CURRENCIES], activeWorksheet.Cells[2, CurrenciesRetrieved.Count + COLUMN_OFFSET_FOR_CURRENCIES - 1]];

            string[,] values = new string[2, CurrenciesRetrieved.Count];
            for (int i = 0; i < CurrenciesRetrieved.Count; i++)
            {
                values[0,i] = CurrenciesRetrieved[i].CurrencyName;
                values[1, i] = CurrenciesRetrieved[i].RateUnit.ToString();
                CurrenciesRetrieved[i].ColumnReference = i + COLUMN_OFFSET_FOR_CURRENCIES;
            }

            currencyHeader.Value = values;
        }

        /// <summary>
        /// Prints the dates which were requested during the query to the first column offset by default value
        /// </summary>
        /// <param name="activeWorksheet">Excel:Worksheet: worksheet where the data gets copied</param>
        private void PrintFxRateDateRangeHeader(Excel.Worksheet activeWorksheet)
        {
            
            CurrencySnapshotDate.OrderBy(srt => srt.Date);
            Excel.Range dateRange = activeWorksheet.Range[activeWorksheet.Cells[ROW_OFFSET_FOR_DATES, 1], activeWorksheet.Cells[CurrencySnapshotDate.Count - 1 + ROW_OFFSET_FOR_DATES, 1]];

            string[,] values = new string[CurrencySnapshotDate.Count, 1];

            for (int i = 0; i < CurrencySnapshotDate.Count; i++)
            {
                values[i, 0] = CurrencySnapshotDate[i].Date.ToString(DATE_OUTPUT_FORMAT);
                CurrencySnapshotDate[i].RowReferece = i + ROW_OFFSET_FOR_DATES;
            }
        
            dateRange.Value = values;
        }

        /// <summary>
        /// Prints the fxRates in the corresponding rows : columns
        /// Rows represent dates
        /// Columns represent currencies
        /// </summary>
        /// <param name="activeWorksheet">Excel:Worksheet: worksheet where the data gets copied</param>
        private void PrintFxRates(Excel.Worksheet activeWorksheet)
        {
            foreach (CurrencyWrapper currencyWrapper in CurrenciesRetrieved)
            {
                foreach (Currency currency in currencyWrapper.CurrencyRates)
                {
                    FxDateWrapper currencySnapshotRef = CurrencySnapshotDate.FirstOrDefault(x => x.Date == currency.Date);

                    Excel.Range fxCellReference = activeWorksheet.Cells[currencySnapshotRef.RowReferece, currencyWrapper.ColumnReference];
                    var normRawRate = currency.RawRate.Replace(',', '.');
                    fxCellReference.Value = normRawRate;
                    ((dynamic)fxCellReference).NumberFormat = "0.00";
                }
            }
        }

        #endregion

        #region Access Methods

        /// <summary>
        /// Method being called when user clicks the button on the ribbon to give reasoning on the query
        /// </summary>
        public void LogReasonForQuery()
        {

        }

        #endregion

        #region Helper methods

        /// <summary>
        /// Pops up a MessageBox with given error message.
        /// </summary>
        /// <param name="error">string: the message to be shown</param>
        private void ShowErrorMessage(string error)
        {
            MessageBox.Show(error);
        }

        #endregion

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
