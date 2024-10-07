namespace RekapMidtrans.DTO
{
    public class OcistokDTO
    {
    }
    public class ApiResponse
    {
        public ResponseData data { get; set; }
        public string message { get; set; }
        public List<Sales> sales { get; set; }
        public int status { get; set; }
    }

    public class ResponseData
    {
        public List<OrderData> data { get; set; }
        public int dataInPage { get; set; }
        public string email { get; set; }
        public int id_cache { get; set; }
        public int totalData { get; set; }
        public int totalPage { get; set; }
    }

    public class OrderData
    {
        public bool can_manual_payment { get; set; }
        public bool can_transfer { get; set; }
        public string canceled_date { get; set; }
        public string createdAt { get; set; }
        public string customerName { get; set; }
        public string customerPhone { get; set; }
        public string daysAgo { get; set; }
        public string eta { get; set; }
        public int id_group { get; set; }
        public int id_so { get; set; }
        public bool is_expired_order { get; set; }
        public string latestStatus { get; set; }
        public int min_installment { get; set; }
        public bool newCustomer { get; set; }
        public bool notes { get; set; }
        public string orderStatus { get; set; }
        public string payment_date { get; set; }
        public string payment_type { get; set; }
        public string sales { get; set; }
        public string status { get; set; }
        public decimal totalBuy { get; set; }
        public decimal totalPrice { get; set; }
        public int total_notes { get; set; }
    }

    public class Sales
    {
        public string divisi { get; set; }
        public string jabatan { get; set; }
        public string user { get; set; }
    }

}
