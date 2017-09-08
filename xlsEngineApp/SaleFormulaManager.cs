using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace xlsEngineApp
{
    public sealed class SaleFormulaManager
    {
        private SalesFormula _formula;

        public SaleFormulaManager()
        {
            this._formula = new SalesFormula();
        }

        public void Calulate(IEnumerable<Sales> sales)
        {
            this._formula.Date = DateTime.Now;
            this._formula.DateFrom = sales.Min(x=>x.OrderDate);
            this._formula.DateTo = sales.Max(x=>x.OrderDate);

            this._formula.ReportNo = "SALE-00-0001";
            this._formula.SaleTotalDiscount = sales.Sum(x=>x.Discount);
            this._formula.SaleTotalAmount = sales.Sum(x=>x.Amount);
        }

        public SalesFormula Forumla
        {
            get
            {
                return this._formula;
            }
        }

    }
}
