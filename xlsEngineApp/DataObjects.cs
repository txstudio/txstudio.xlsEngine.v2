using System;

namespace xlsEngineApp
{
    /*
    資料來源查詢 AdvantureWork 資料庫
    ------------------------------------
    SELECT TOP 5000 e.[LastName]+'.'+ CONVERT(VARCHAR(10),e.[BusinessEntityID]) [Account]
	    ,f.[EmailAddress] [email]
	    ,c.[OrderDate]
	    ,c.[PurchaseOrderNumber] [invoice]
	    ,a.[ProductNumber] [Schema]
	    ,a.[Name] [ProductName]
	    ,b.[UnitPrice]
	    ,b.[OrderQty] [Quantity]
	    ,(0-(b.[UnitPrice] * b.[UnitPriceDiscount])) [Discount]
	    ,(b.[UnitPrice] - (b.[UnitPrice] * b.[UnitPriceDiscount]))*b.[OrderQty] [Amount]
	    ,c.[Comment] [Note]
    FROM [Production].[Product] a
	    INNER JOIN [Sales].[SalesOrderDetail] b ON a.[ProductID] = b.[ProductID]
	    INNER JOIN [Sales].[SalesOrderHeader] c ON b.[SalesOrderID] = c.[SalesOrderID]
	    INNER JOIN [Sales].[Customer] d ON c.[CustomerID] = d.[CustomerID]
	    INNER JOIN [Person].[Person] e ON d.[PersonID] = e.[BusinessEntityID]
	    INNER JOIN [Person].[EmailAddress] f ON e.[BusinessEntityID] = f.[BusinessEntityID]
    WHERE c.[PurchaseOrderNumber] IS NOT NULL
    */

    public sealed class Sales
    {
        public string Account { get; set; }
        public string email { get; set; }
        public Nullable<DateTime> OrderDate { get; set; }
        public string invoice { get; set; }
        public string Schema { get; set; }
        public string ProductName { get; set; }

        public int UnitPrice { get; set; }
        public int Quantity { get; set; }
        public int Discount { get; set; }
        public int Amount { get; set; }
        public string Note { get; set; }
    }

    public sealed class SalesFormula
    {
        private string DateTimeFormat(Nullable<DateTime> datetime)
        {
            if(datetime.HasValue == false)
                return string.Empty;

            return string.Format("{0:yyyy/MM/dd}", datetime.Value);
        }

        public Nullable<DateTime> Date { get;set; }
        public Nullable<DateTime> DateFrom { get;set; }
        public Nullable<DateTime> DateTo { get;set; }

        public string ReportNo { get;set; }
        public int SaleTotalDiscount { get;set; }
        public int SaleTotalAmount { get;set; }

        public string DateDisplay
        {
            get
            {
                return this.DateTimeFormat(Date);
            }
        }
        public string DateFromDisplay
        {
            get
            {
                return this.DateTimeFormat(DateFrom);
            }
        }
        public string DateToDisplay
        {
            get
            {
                return this.DateTimeFormat(DateTo);
            }
        }
    }
}
