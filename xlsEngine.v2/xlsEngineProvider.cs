using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace xlsEngine.v2
{

    public abstract class xlsEngineProvider
    {
        private byte[] _inBuffer;
        private byte[] _outBuffer;

        private int _index;

        private HSSFWorkbook _workBook;
        private HSSFSheet _sheet;

        private readonly xlsEngineOption _option;
        private readonly Dictionary<string, string> _formulas;


        protected xlsEngineProvider()
        {
            this._formulas = new Dictionary<string, string>();

            this._option = new xlsEngineOption();
            this.SetOption(this._option);


            this._inBuffer = File.ReadAllBytes(this._option.TemplatePath);
        }



        #region abstract/virtual method

        protected abstract void SetOption(xlsEngineOption option);
        protected abstract bool InsertHeaderRow(DataRow current, DataRow before);
        protected abstract bool InsertFooterRow(DataRow current, DataRow next);


        protected virtual void HeaderMap(DataRow current, DataRow before)
        {
            this.SetCellByBindName(current);
        }

        /// <summary>進行表格物件與資料列對應</summary>
        protected virtual void RecordMap(DataRow current, DataRow before, DataRow next)
        {
            //預設的方法：依照原資料列設定名稱從 Current 取得對應的資料欄位內容
            //  設定方式「#ColumnName」
            this.SetCellByBindName(current);
        }

        protected virtual void FooterMap(DataRow current, DataRow next)
        {
            this.SetCellByBindName(current);
        }

        #endregion

        #region protected method

        protected void SetFormula()
        {
            foreach (var item in this._formulas)
            {
                this.SetCellByParameter(item.Key, item.Value);
            }
        }

        protected void SetCellByParameter(string parameterName, int value)
        {
            this.SetCellByParameter(parameterName, value.ToString());
        }

        protected void SetCellByParameter(string parameterName, string value)
        {
            HSSFRow _row;
            HSSFCell _cell;

            string _cellValue;

            _row = this.GetCurrentRow();

            if (_row == null)
                return;

            for (int i = 0; i < _row.LastCellNum; i++)
            {
                _cell = this.GetCell(_row, i);
                _cellValue = this.GetStringValue(_cell);

                if (string.Equals(_cellValue
                                , parameterName
                                , StringComparison.OrdinalIgnoreCase) == true)
                {
                    this.SetCellValue(_cell, value);
                }
            }
        }


        protected void SetCellByField(DataRow current
                                    , DataRow before
                                    , DataRow next
                                    , int cellIndex
                                    , string columnName
                                    , bool isDuplicateHide = false
                                    , ICustomFormula formula = null)
        {

            if (current.Table.Columns.Contains(columnName) == false)
                return;

            HSSFCell _cell;
            HSSFRow _row;

            object _currentValue;
            object _beforeValue;

            _row = this.GetCurrentRow();
            _cell = this.GetCell(_row, cellIndex);
            _currentValue = current[columnName];

            if (_currentValue.Equals(DBNull.Value) == true)
                return;

            if (formula != null)
                _currentValue = formula.GetValue(_currentValue);

            this.SetCellValue(_cell, _currentValue.ToString());

            //與前筆資料相同時不顯示此筆資料
            if (isDuplicateHide == true)
            {
                if (before == null)
                    return;

                _beforeValue = before[columnName];

                if (_currentValue.Equals(_beforeValue) == true)
                    _cell.SetCellValue(string.Empty);
            }
        }
        #endregion

        #region public method

        public void AddFormula(string key, int value)
        {
            this.AddFormula(key, value.ToString());
        }

        /// <summary>新增公式欄位</summary>
        /// <param name="key">變數名稱（開頭為@）</param>
        /// <param name="value">數值</param>
        public void AddFormula(string key, string value)
        {
            if (key.StartsWith("@") == false)
                key = string.Format("@{0}", key);

            this._formulas.Add(key, value);
        }

        public void Load(DataTable table)
        {
            this.SetupWorkbook();
            this.SetupSheet();

            DataRow _beforeDataRow;
            DataRow _currentDataRow;
            DataRow _nextDataRow;

            int _rowIndex;
            int _rowCount;


            _rowIndex = 0;
            _rowCount = table.Rows.Count;

            _currentDataRow = null;
            _beforeDataRow = null;
            _nextDataRow = null;


            //替換起始欄位公式
            for (int i = 0; i < this._option.StartIndex; i++)
            {
                _index = _index + 1;
                this.SetFormula();
            }


            //打印報表資料
            foreach (DataRow _recordRow in table.Rows)
            {
                _currentDataRow = null;
                _nextDataRow = null;

                _currentDataRow = _recordRow;

                if (_rowIndex < (_rowCount - 1))
                    _nextDataRow = table.Rows[_rowIndex + 1];


                if (InsertHeaderRow(_currentDataRow, _beforeDataRow) == true)
                    if (this._option.Header != null)
                        this.SetRowHeader(this._option.Header, _currentDataRow, _beforeDataRow);

                if (this._option.Detail != null)
                    this.SetRowDetail(this._option.Detail, _currentDataRow, _beforeDataRow, _nextDataRow);

                if (InsertFooterRow(_currentDataRow, _nextDataRow) == true)
                    if (this._option.Footer != null)
                        this.SetRowFooter(this._option.Footer, _currentDataRow, _nextDataRow);


                _beforeDataRow = _currentDataRow;
                _rowIndex = _rowIndex + 1;
            }


            if (this._option.DocumentFooter != null)
                this.SetDocumentFooter(this._option.DocumentFooter);

            if (this._option.RemoveTemplateSheet == true)
                this.RemoveTemplateSheet();


            this.SetupInfo();

            this.Save();
        }

        /// <summary>取得完成的 EXCEL 檔案</summary>
        public byte[] XlsContent
        {
            get
            {
                return this._outBuffer;
            }
        }

        #endregion


        private void SetupWorkbook()
        {
            using (MemoryStream _mStream
                = new MemoryStream(this._inBuffer))
            {
                this._workBook = new HSSFWorkbook(_mStream);
            }
        }
        private void SetupSheet()
        {
            this._sheet = this._workBook.GetSheetAt(0) as HSSFSheet;
        }
        private void Save()
        {
            using (MemoryStream _outStream = new MemoryStream())
            {
                this._workBook.Write(_outStream);
                this._outBuffer = _outStream.ToArray();
            }
        }

        private void SetupInfo()
        {
            this._workBook.SummaryInformation.Author = string.Empty;
            this._workBook.SummaryInformation.LastAuthor = string.Empty;

            if(string.IsNullOrWhiteSpace(this._option.Author) == false)
                this._workBook.SummaryInformation.Author = this._option.Author;
        }


        /// <summary>設定資料列表頭</summary>
        private void SetRowHeader(RowTemplateConfig config, DataRow current, DataRow before)
        {
            int _num;
            int _sheetIndex;

            HSSFSheet _sourceSheet;

            _num = config.RowIndexStart;
            _sheetIndex = config.SheetIndex;
            _sourceSheet = this._workBook.GetSheetAt(_sheetIndex) as HSSFSheet;

            while (_num <= this._option.Header.RowIndexEnd)
            {
                this.CopyRow(_sourceSheet, _num);
                this.HeaderMap(current, before);
                _num = _num + 1;
            }
        }

        /// <summary>設定資料列內容</summary>
        private void SetRowDetail(RowTemplateConfig config, DataRow current, DataRow before, DataRow next)
        {
            int _num;
            int _sheetIndex;

            HSSFSheet _sourceSheet;

            _num = config.RowIndexStart;
            _sheetIndex = config.SheetIndex;
            _sourceSheet = this._workBook.GetSheetAt(_sheetIndex) as HSSFSheet;

            while (_num <= config.RowIndexEnd)
            {
                this.CopyRow(_sourceSheet, _num);
                this.RecordMap(current, before, next);
                _num = _num + 1;
            }

        }

        /// <summary>設定資料列表尾</summary>
        private void SetRowFooter(RowTemplateConfig config, DataRow current, DataRow next)
        {
            int _num;
            int _sheetIndex;

            HSSFSheet _sourceSheet;

            _num = config.RowIndexStart;
            _sheetIndex = config.SheetIndex;
            _sourceSheet = this._workBook.GetSheetAt(_sheetIndex) as HSSFSheet;

            while (_num <= config.RowIndexEnd)
            {
                this.CopyRow(_sourceSheet, _num);
                this.FooterMap(current, next);
                _num = _num + 1;
            }
        }

        /// <summary>設定文件表尾</summary>
        private void SetDocumentFooter(RowTemplateConfig config)
        {
            int _num;
            int _sheetIndex;

            HSSFSheet _sourceSheet;

            _num = config.RowIndexStart;
            _sheetIndex = config.SheetIndex;
            _sourceSheet = this._workBook.GetSheetAt(_sheetIndex) as HSSFSheet;

            while (_num <= config.RowIndexEnd)
            {
                this.CopyRow(_sourceSheet, _num);
                _num = _num + 1;
                this.SetFormula();
            }
        }


        /// <summary>移除範本工作表內容</summary>
        private void RemoveTemplateSheet()
        {
            foreach (int sheetIndex in this._option.TemplateSheetIndex)
            {
                this._workBook.RemoveSheetAt(sheetIndex);
            }
        }


        //reference link : http://www.zachhunter.com/2010/05/npoi-copy-row-helper/

        /// <summary>複製指定工作表的指定資料列到指定工作表的指定資料列</summary>
        /// <param name="sourceSheet">來源工作表</param>
        /// <param name="targetSheet">目標工作表</param>
        /// <param name="sourceIndex">來源資料列索引</param>
        /// <param name="targetIndex">目標資料列索引</param>
        private void CopyRow(HSSFSheet sourceSheet
                            , HSSFSheet targetSheet
                            , int sourceIndex
                            , int targetIndex)
        {
            // Get the source / new row
            HSSFRow newRow;
            HSSFRow sourceRow;

            newRow = targetSheet.GetRow(targetIndex) as HSSFRow;
            sourceRow = sourceSheet.GetRow(sourceIndex) as HSSFRow;


            if (newRow == null)
                newRow = targetSheet.CreateRow(targetIndex) as HSSFRow;

            if (sourceRow == null)
                sourceRow = sourceSheet.CreateRow(sourceIndex) as HSSFRow;


            newRow.HeightInPoints = sourceRow.HeightInPoints;

            // Loop through source columns to add to new row

            //複製儲存格內容
            for (int i = 0; i < sourceRow.LastCellNum; i++)
            {
                // Grab a copy of the old/new cell
                HSSFCell oldCell = sourceRow.GetCell(i) as HSSFCell;
                HSSFCell newCell = newRow.CreateCell(i) as HSSFCell;

                // If the old cell is null jump to next cell
                if (oldCell == null)
                {
                    newCell = null;
                    continue;
                }

                // Copy style from old cell and apply to new cell
                newCell.CellStyle = oldCell.CellStyle;

                // If there is a cell comment, copy
                if (newCell.CellComment != null)
                    newCell.CellComment = oldCell.CellComment;

                // If there is a cell hyperlink, copy
                if (oldCell.Hyperlink != null)
                    newCell.Hyperlink = oldCell.Hyperlink;

                // Set the cell data type
                newCell.SetCellType(oldCell.CellType);

                // Set the cell data value
                switch (oldCell.CellType)
                {
                    case CellType.Blank:
                        newCell.SetCellValue(oldCell.StringCellValue);
                        break;
                    case CellType.Boolean:
                        newCell.SetCellValue(oldCell.BooleanCellValue);
                        break;
                    case CellType.Error:
                        newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                        break;
                    case CellType.Formula:
                        newCell.SetCellFormula(oldCell.CellFormula);
                        break;
                    case CellType.Numeric:
                        newCell.SetCellValue(oldCell.NumericCellValue);
                        break;
                    case CellType.String:
                        newCell.SetCellValue(oldCell.RichStringCellValue);
                        break;
                    case CellType.Unknown:
                        newCell.SetCellValue(oldCell.StringCellValue);
                        break;
                }
            }


            //進行合併儲存格設定搬移作業
            CellRangeAddress cellRangeAddress;
            CellRangeAddress newCellRangeAddress;

            for (int i = 0; i < sourceSheet.NumMergedRegions; i++)
            {
                cellRangeAddress = sourceSheet.GetMergedRegion(i);

                if (cellRangeAddress.FirstRow == sourceRow.RowNum)
                {
                    newCellRangeAddress = new CellRangeAddress(newRow.RowNum
                                                                , (cellRangeAddress.LastRow - cellRangeAddress.FirstRow)
                                                                    + newRow.RowNum
                                                                , cellRangeAddress.FirstColumn
                                                                , cellRangeAddress.LastColumn
                                                                );

                    targetSheet.AddMergedRegion(newCellRangeAddress);
                }
            }

            this.MoveCursorPosition();
        }

        /// <summary>複製指定工作表的指定資料列到指定工作表的指定資料列</summary>
        /// <param name="sourceSheet">來源工作表</param>
        /// <param name="sourceIndex">來源資料列索引</param>
        private void CopyRow(HSSFSheet sourceSheet
                            , int sourceIndex)
        {
            this.CopyRow(sourceSheet
                        , this._sheet
                        , sourceIndex
                        , this._index);
        }


        /// <summary>使用指定格式名稱繫結 DataRow 的資料</summary>
        /// <param name="current">指定 DataRow 物件</param>
        /// <param name="prefix">繫結前置詞（預設#）</param>
        private void SetCellByBindName(DataRow current, string prefix = "#")
        {
            HSSFRow _row;
            HSSFCell _cell;

            string _cellValue;
            string _dataValue;
            string _columnName;

            _row = this.GetCurrentRow();

            for (int i = 0; i < _row.LastCellNum; i++)
            {
                _cell = _row.GetCell(i) as HSSFCell;

                if (_cell == null)
                    continue;

                _cellValue = this.GetStringValue(_cell);

                if (_cellValue.StartsWith(prefix) == true)
                {
                    _columnName = _cellValue.Replace(prefix, string.Empty);
                    _dataValue = this.GetStringValue(current, _columnName);

                    this.SetCellValue(_cell, _dataValue);
                }
            }
        }

        /// <summary>取得指定儲存格的文字內容</summary>
        /// <param name="cell">指定儲存格</param>
        private string GetStringValue(HSSFCell cell)
        {
            /*
             * 數字格式將轉換成文字 1 => "1"
             * 其餘格式顯示 string.Empty
             */

            string _value;

            _value = string.Empty;

            switch (cell.CellType)
            {
                case CellType.Blank:
                    break;
                case CellType.Boolean:
                    break;
                case CellType.Error:
                    break;
                case CellType.Formula:
                    break;
                case CellType.Numeric:
                    _value = cell.NumericCellValue.ToString();
                    break;
                case CellType.String:
                    _value = cell.StringCellValue;
                    break;
                case CellType.Unknown:
                    break;
                default:
                    break;
            }

            return _value;
        }

        /// <summary>取得指定欄位名稱的字串</summary>
        /// <param name="row">資料列物件</param>
        /// <param name="columnName">欄位名稱</param>
        private string GetStringValue(DataRow row, string columnName)
        {
            if (row == null)
                return string.Empty;

            if (string.IsNullOrWhiteSpace(columnName) == true)
                return string.Empty;

            if (row.Table.Columns.Contains(columnName) == false)
                return string.Empty;


            return row[columnName].ToString();
        }

        /// <summary>移動存取的資料列指標位置 (預設+1)</summary>
        private void MoveCursorPosition(int position = 1)
        {
            this._index = this._index + position;
        }



        /// <summary>取得存取中的資料列物件</summary>
        private HSSFRow GetCurrentRow()
        {
            HSSFRow _row;

            _row = this._sheet.GetRow(this._index - 1) as HSSFRow;

            return _row;
        }

        /// <summary>取得工作表資料列指定索引的儲存格物件</summary>
        /// <param name="row">工作表資料列</param>
        /// <param name="index">儲存格索引</param>
        /// <returns>儲存格物件</returns>
        private HSSFCell GetCell(HSSFRow row, int index)
        {
            if (row == null)
                return null;

            HSSFCell _cell;

            _cell = row.GetCell(index, MissingCellPolicy.CREATE_NULL_AS_BLANK) as HSSFCell;

            return _cell;
        }

        /// <summary>設定儲存格資料/會自動轉換成數字</summary>
        /// <param name="cell">儲存格物件</param>
        /// <param name="value">字串</param>
        private void SetCellValue(HSSFCell cell, string value)
        {
            double _double;

            if (double.TryParse(value, out _double) == true)
            {
                cell.SetCellValue(_double);
                return;
            }

            cell.SetCellValue(value);
        }

    }
}
