using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xlsEngine.v2;

namespace xlsEngineApp
{
    public sealed class SalesSimpleXlsEngine : xlsEngineProvider
    {
        protected override void SetOption(xlsEngineOption option)
        {
            option.Author = "報表管理員";
            option.TemplatePath = @"..\..\file\sale.simple.template.xls";
            option.StartIndex = 4;

            option.Detail = new RowTemplateConfig() { SheetIndex = 1, RowIndexStart = 0, RowIndexEnd = 0 };
            option.DocumentFooter = new RowTemplateConfig() { SheetIndex = 2, RowIndexStart = 0, RowIndexEnd = 1 };
        }

        protected override bool InsertHeaderRow(DataRow current, DataRow before)
        {
            return false;
        }

        protected override bool InsertFooterRow(DataRow current, DataRow next)
        {
            return false;
        }

        protected override void RecordMap(DataRow current, DataRow before, DataRow next)
        {
            base.RecordMap(current, before, next);

            this.SetCellByField(current, before, next, 0, "Account", true);
            this.SetCellByField(current, before, next, 1, "email", true);
            this.SetCellByField(current, before, next, 2, "invoice", true);
        }
    }
}
