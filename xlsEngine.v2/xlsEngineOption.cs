using System.Collections.Generic;

namespace xlsEngine.v2
{

    /// <summary>EXCEL 範本物件參數設定</summary>
    public sealed class xlsEngineOption
    {
        private bool _removeTemplateSheet;

        public xlsEngineOption()
        {
            this._removeTemplateSheet = true;
        }

        /// <summary>EXCEL 範本檔案路徑</summary>
        public string TemplatePath { get; set; }

        /// <summary>文件表尾範本</summary>
        public RowTemplateConfig DocumentFooter { get; set; }

        /// <summary>文件清單表頭範本</summary>
        public RowTemplateConfig Header { get; set; }
        /// <summary>文件清單範本</summary>
        public RowTemplateConfig Detail { get; set; }
        /// <summary>文件清單表尾範本</summary>
        public RowTemplateConfig Footer { get; set; }


        /// <summary>打印報表內容起始索引</summary>
        public int StartIndex { get; set; }

        /// <summary>報表作者</summary>
        public string Author { get;set; }

        /// <summary>完成後是否移除範本工作表：預設(true)</summary>
        public bool RemoveTemplateSheet
        {
            get
            {
                return this._removeTemplateSheet;
            }
            set
            {
                this._removeTemplateSheet = value;
            }
        }

        /// <summary>取得要移除的範本工作表索引，索引遞減</summary>
        public int[] TemplateSheetIndex
        {
            get
            {
                List<int> _sheets;

                _sheets = new List<int>();


                if (this.DocumentFooter != null)
                    _sheets.Add(this.DocumentFooter.SheetIndex);

                if (this.Header != null)
                    _sheets.Add(this.Header.SheetIndex);

                if (this.Footer != null)
                    _sheets.Add(this.Footer.SheetIndex);

                if (this.Detail != null)
                    _sheets.Add(this.Detail.SheetIndex);


                _sheets.Sort();
                _sheets.Reverse();

                return _sheets.ToArray();
            }
        }
    }

    /// <summary>EXCEL 範本資料的設定物件</summary>
    public sealed class RowTemplateConfig
    {
        /// <summary>範本工作表索引</summary>
        public int SheetIndex { get; set; }

        /// <summary>範本起始資料列索引</summary>
        public int RowIndexStart { get; set; }

        /// <summary>範本結束資料列索引</summary>
        public int RowIndexEnd { get; set; }
    }

}
