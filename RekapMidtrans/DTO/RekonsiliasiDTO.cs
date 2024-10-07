namespace RekapMidtrans.DTO
{
    public class RekonsiliasiDTO
    {
        public string DateAndTime { get; set; }
        public string OrderID { get; set; }
        public string Channel { get; set; }
        public string TransactionType { get; set; }
        public string Amount { get; set; }
        public string TransactionStatus { get; set; }
        public string TransactionTime { get; set; }
        public string TransactionID { get; set; }
        public string CustomerEmail { get; set; }
        public string Note { get; set; }
        public string GroupID { get; set; }
        public string IdOrder { get; set; }
        public string NoInvoice { get; set; }
        public string CustomerName { get; set; }
        public string ProductName { get; set; }
        public int Quantity { get; set; }
        public double QuantityPrice { get; set; }
        public double CbmPrice { get; set; }
        public double SubTotal { get; set; }
        public double LocalShipping { get; set; }
        public double TotalPrice { get; set; }
        public double Selisih { get; set; }
        public double COGS { get; set; }
        public double VoucherDiscount { get; set; }
        public double GrossProfit { get; set; }
        public string Sales { get; set; }
    }
}
