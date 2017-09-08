using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xlsEngine.v2;

namespace xlsEngineApp
{
    /// <summary>實作銷售報表匯出</summary>
    public sealed class SalesXlsEngine : xlsEngineProvider
    {
        const string repeatColumName = "account";

        protected override void SetOption(xlsEngineOption option)
        {
            option.Author = "報表管理員";
            option.TemplatePath = @"..\..\file\sale.template.xls";
            option.StartIndex = 3;

            option.Header = new RowTemplateConfig() { SheetIndex = 1, RowIndexStart = 0, RowIndexEnd = 3 };
            option.Detail = new RowTemplateConfig() { SheetIndex = 2, RowIndexStart = 0, RowIndexEnd = 0 };
            option.Footer = new RowTemplateConfig() { SheetIndex = 3, RowIndexStart = 0, RowIndexEnd = 2 };
            option.DocumentFooter = new RowTemplateConfig() { SheetIndex = 4, RowIndexStart = 0, RowIndexEnd = 1 };
        }

        protected override bool InsertHeaderRow(DataRow current, DataRow before)
        {
            if(before == null)
                return true;

            if(string.Equals(current[repeatColumName]
                            , before[repeatColumName]) == true)
            {
                return false;
            }

            return true;
        }

        protected override bool InsertFooterRow(DataRow current, DataRow after)
        {
            if (after == null)
                return true;

            if (string.Equals(current[repeatColumName]
                            , after[repeatColumName]) == true)
            {
                return false;
            }

            return true;
        }


        private int _totalAmount;
        private int _totalDiscount;
        private int _total;


        protected override void HeaderMap(DataRow current, DataRow before)
        {
            base.HeaderMap(current, before);

            string _account;
            string _email;
            DateTime _orderDate;
            string _invoice;

            _account = current["account"].ToString();
            _email = current["email"].ToString();
            _orderDate = Convert.ToDateTime(current["orderDate"]);
            _invoice = current["invoice"].ToString();

            this.SetCellByParameter("@account", _account);
            this.SetCellByParameter("@email", _email);
            this.SetCellByParameter("@orderDate", String.Format("{0:yyyy/MM/dd}",_orderDate));
            this.SetCellByParameter("@invoice", _invoice);

            _totalAmount = 0;
            _totalDiscount = 0;
        }

        protected override void RecordMap(DataRow current, DataRow before, DataRow next)
        {
            base.RecordMap(current, before, next);

            int _amount;
            int _discount;

            _amount = Convert.ToInt32(current["Amount"]);
            _discount = Convert.ToInt32(current["Discount"]);

            _totalAmount = _totalAmount + _amount;
            _totalDiscount = _totalDiscount + _discount;
        }

        protected override void FooterMap(DataRow current, DataRow next)
        {
            base.FooterMap(current, next);

            _total = _totalAmount + _totalDiscount;

            this.SetCellByParameter("@TotalAmount", _totalAmount);
            this.SetCellByParameter("@TotalDiscount", _totalDiscount);
            this.SetCellByParameter("@Total", _total);
        }

    }
}
