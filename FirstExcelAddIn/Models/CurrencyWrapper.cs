using System.Collections.Generic;

namespace FirstExcelAddIn.Models
{
    public class CurrencyWrapper
    {
        public string CurrencyName { get; set; }
        public int RateUnit { get; set; }
        public IList<Currency> CurrencyRates { get; set; }
        public int ColumnReference { get; set; }

        public CurrencyWrapper()
        {
            CurrencyRates = new List<Currency>();
        }
    }
}
