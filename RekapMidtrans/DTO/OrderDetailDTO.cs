namespace RekapMidtrans.DTO
{    
    public class OrderDetailDTO
    {
        public Data data { get; set; }
        public string message { get; set; }
        public int status { get; set; }
    }

    public class Data
    {
        public object adjustment { get; set; }
        public bool airplane { get; set; }
        public int category { get; set; }
        public object cicilan { get; set; }
        public bool custom { get; set; }
        public CustomerDetail customerDetail { get; set; }
        public string etd { get; set; }
        public FinancialInformation financialInformation { get; set; }
        public int id_group { get; set; }
        public string id_so { get; set; }
        public bool is_banned { get; set; }
        public bool is_cicilan { get; set; }
        public int is_installment { get; set; }
        public bool is_packing_kayu { get; set; }
        public string kode_diskon { get; set; }
        public LogisticInformation logisticInformation { get; set; }
        public OrderDetail orderDetail { get; set; }
        public string paymentDate { get; set; }
        public string platform { get; set; }
        public List<PODetail> poDetails { get; set; }
        public string refund_status { get; set; }
        public string rumus_ver { get; set; }
        public string sales { get; set; }
        public string status { get; set; }
        public StatusHistory statusHistory { get; set; }
        public string tanggal { get; set; }
        public int total_adjustment { get; set; }
        public int total_asli { get; set; }
    }

    public class CustomerDetail
    {
        public string address { get; set; }
        public string area { get; set; }
        public string city { get; set; }
        public string courier { get; set; }
        public string email { get; set; }
        public int id_city { get; set; }
        public int id_customer { get; set; }
        public int id_district { get; set; }
        public int id_province { get; set; }
        public int id_subdistrict { get; set; }
        public string level { get; set; }
        public string name { get; set; }
        public string phone { get; set; }
        public string province { get; set; }
        public string sales { get; set; }
        public string service { get; set; }
        public string subdistrict { get; set; }
        public string zip { get; set; }
    }

    public class FinancialInformation
    {
        public Expense expense { get; set; }
        public Income income { get; set; }
        public double percentage { get; set; }
        public double percentage_request { get; set; }
        public double percentage_whchina { get; set; }
        public double profit { get; set; }
        public double profit_request { get; set; }
        public double profit_whchina { get; set; }
    }

    public class Expense
    {
        public object biaya_tambahan { get; set; }
        public int cogs { get; set; }
        public decimal cogs_already_po { get; set; }
        public int cogs_new_orders { get; set; }
        public int cogs_po_paid { get; set; }
        public int discount { get; set; }
        public int others { get; set; }
        public int refund { get; set; }
        public int request_price { get; set; }
        public int shippingChIdn { get; set; }
        public int shippingCustomer { get; set; }
        public double shipping_ch_idn { get; set; }
        public int shipping_forecast_request { get; set; }
        public double shipping_forecast_whchina { get; set; }
        public int shipping_indo { get; set; }
        public decimal tax { get; set; }
        public decimal total { get; set; }
        public decimal total_request { get; set; }
        public double total_whchina { get; set; }
        public int transfer_value { get; set; }
        public int voucher { get; set; }
    }

    public class Income
    {
        public int customerPayment { get; set; }
        public int discount { get; set; }
        public int income { get; set; }
        public int others { get; set; }
        public int processing_fee { get; set; }
        public int proses_fee { get; set; }
        public int shippingCustomer { get; set; }
        public int shipping_international { get; set; }
        public int tax { get; set; }
        public int total { get; set; }
        public int totalraw { get; set; }
    }

    public class LogisticInformation
    {
        public object box { get; set; }
        public ChIdn chIdn { get; set; }
        public object idn { get; set; }
    }

    public class ChIdn
    {
        public int cost { get; set; }
        public string expedition { get; set; }
        public string logisticChannel { get; set; }
        public string totalBox { get; set; }
        public string volume { get; set; }
    }

    public class OrderDetail
    {
        public object bukti { get; set; }
        public int category { get; set; }
        public int discount { get; set; }
        public string eta { get; set; }
        public int expenseIdr { get; set; }
        public double expenseRmb { get; set; }
        public int income { get; set; }
        public string orderNumber { get; set; }
        public string paymentDate { get; set; }
        public List<Product> product { get; set; }
        public int qty { get; set; }
        public int refund { get; set; }
        public string refundStatus { get; set; }
        public string request_status { get; set; }
        public string sales { get; set; }
        public int shipping { get; set; }
        public string status { get; set; }
        public int total { get; set; }
    }

    public class Product
    {
        public int biaya_layanan { get; set; }
        public int category { get; set; }
        public int cbm_calcs { get; set; }
        public int id { get; set; }
        public string id_page { get; set; }
        public string idvariant { get; set; }
        public string image { get; set; }
        public string kat_req { get; set; }
        public string link { get; set; }
        public int moq { get; set; }
        public int ppn { get; set; }
        public int price { get; set; }
        public string product { get; set; }
        public int qty { get; set; }
        public object qty_whchina { get; set; }
        public string request_status { get; set; }
        public string sku { get; set; }
        public double tag { get; set; }
        public string toko { get; set; }
        public int total { get; set; }
        public int volume { get; set; }
    }

    public class PODetail
    {
        public string date { get; set; }
        public string id_po { get; set; }
        public string id_so { get; set; }
        public string link { get; set; }
        public string notes { get; set; }
        public string paymentNumber { get; set; }
        public string payment_number { get; set; }
        public List<POProduct> product { get; set; }
        public int qty { get; set; }
        public string status { get; set; }
        public string supplier { get; set; }
        public double total { get; set; }
        public string trackingNumber { get; set; }
        public string tracking_number { get; set; }
        public string user { get; set; }
    }

    public class POProduct
    {
        public string id_produk { get; set; }
        public string image { get; set; }
        public double price { get; set; }
        public string product { get; set; }
        public int qty { get; set; }
        public int qty_whchina { get; set; }
        public string sku { get; set; }
        public double total { get; set; }
    }

    public class StatusHistory
    {
        public string email { get; set; }
        public string id_so { get; set; }
        public string status { get; set; }
        public List<History> statusHistory { get; set; }
    }

    public class History
    {
        public string date { get; set; }
        public string desc { get; set; }
        public int id_so { get; set; }
        public string keterangan { get; set; }
        public string status { get; set; }
        public string user { get; set; }
    }

}
