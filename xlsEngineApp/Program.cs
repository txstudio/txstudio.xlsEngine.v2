using Newtonsoft.Json;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;

namespace xlsEngineApp
{
    class Program
    {
        const string outPath = @"..\..\file\out.xls";

        const string jsonDataPath = @"..\..\file\data.json";

        static void Main(string[] args)
        {
            xlsEngine.v2.xlsEngineProvider _xlsEngine;
            IEnumerable<Sales> _sales;
            SalesFormula _formula;
            DataTable _saleTable;


            Byte[] _buffer;


            _sales = ReadObjects();
            _saleTable = DataMapUtility.MapToTable<Sales>(_sales);


            SaleFormulaManager _formulaManager;

            _formulaManager = new SaleFormulaManager();
            _formulaManager.Calulate(_sales);
            _formula = _formulaManager.Forumla;

            //實作依訂單建立表頭與表尾的報表 EXCEL 檔案
            _xlsEngine = new SalesXlsEngine();

            //實作將資料列進行輸出的報表 EXCEL 檔案
            //_xlsEngine = new SalesSimpleXlsEngine();

            _xlsEngine.AddFormula("@date", _formula.DateDisplay);
            _xlsEngine.AddFormula("@dateFrom", _formula.DateFromDisplay);
            _xlsEngine.AddFormula("@dateTo", _formula.DateToDisplay);

            _xlsEngine.AddFormula("@ReportNo", _formula.ReportNo);
            _xlsEngine.AddFormula("@SaleTotalDiscount", _formula.SaleTotalDiscount);
            _xlsEngine.AddFormula("@SaleTotalAmount", _formula.SaleTotalAmount);

            _xlsEngine.Load(_saleTable);

            _buffer = _xlsEngine.XlsContent;



            File.WriteAllBytes(outPath, _buffer);


            ExecuteFile(outPath);
        }


        static IEnumerable<Sales> ReadObjects()
        {
            IEnumerable<Sales> _sales;
            string _json;

            _json = File.ReadAllText(jsonDataPath);
            _sales = JsonConvert.DeserializeObject<Sales[]>(_json);

            return _sales;
        }

        static void ExecuteFile(string fileName)
        {
            ProcessStartInfo _startInfo;

            _startInfo = new ProcessStartInfo();
            _startInfo.FileName = "cmd.exe";
            _startInfo.Arguments = @"/c " + fileName;
            _startInfo.WindowStyle = ProcessWindowStyle.Hidden;

            Process.Start(_startInfo);
        }

    }
}
