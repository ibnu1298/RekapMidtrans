using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using RekapMidtrans.DTO;
using Serilog;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO.Packaging;
using System.Net.Http.Headers;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace RekapMidtrans.Service
{
    public interface IExcel
    {
        Task<DownloadExcelResponse?> DownloadRekapMidtrans(UploadExcelRequest request);
        UploadExcelRequest ExtractUploadRequest(IFormFileCollection streamContents);
    }

    public class ExcelService : IExcel
    {
        private readonly IConfiguration _configuration;

        public ExcelService(IConfiguration Configuration)
        {
            _configuration = Configuration;
        }
        public void ValidateExcelFiletype(MemoryStream stream)
        {
            try
            {
                Package package = Package.Open(stream, FileMode.Open, FileAccess.Read);
            }
            catch (Exception ex)
            {
                //throw new Exception("File has to be in excel format!");
                Log.Error(ex, "Upload error");
                throw;
            }
        }
        private byte[] ReadFullFile(IFormFile uploadedFile)
        {
            using (var memoryStream = new MemoryStream())
            {
                uploadedFile.CopyTo(memoryStream);
                ValidateExcelFiletype(memoryStream);
                return memoryStream.ToArray();
            }
        }
        public UploadExcelRequest ExtractUploadRequest(IFormFileCollection streamContents)
        {
            UploadExcelRequest extractedResult = new UploadExcelRequest();
            TalentAttachment attachmentFile = new TalentAttachment();
            foreach (var content in streamContents)
            {
                switch (content.Name.Replace("\"", string.Empty))
                {
                    //browserfile adalah nama id input upload yang ada di html
                    case "token":
                        extractedResult.token = content.FileName;
                        break;
                    case "browseFile":

                        attachmentFile.name = content.FileName.Replace("\"", string.Empty);
                        attachmentFile.size = Convert.ToInt32(content.Length);
                        attachmentFile.content = ReadFullFile(content);

                        extractedResult.UploadedFile = attachmentFile;
                        break;
                    case "browseFileUpdate":
                        //TalentAttachment attachmentFile = new TalentAttachment();
                        attachmentFile.name = content.FileName.Replace("\"", string.Empty);
                        attachmentFile.size = Convert.ToInt32(content.Length);
                        attachmentFile.content = ReadFullFile(content);

                        extractedResult.UploadedFile = attachmentFile;
                        break;
                }
            }
            return extractedResult;
        }
        private DataTable WorksheetToDataTable(ExcelWorksheet oSheet)
        {
            int totalRows = oSheet.Dimension.End.Row;
            int totalCols = oSheet.Dimension.End.Column;
            DataTable dt = new DataTable(oSheet.Name);
            DataRow dr = dt.NewRow();
            for (int i = 1; i <= totalRows; i++)
            {

                if (i > 1) dr = dt.Rows.Add();
                for (int j = 1; j <= totalCols; j++)
                {
                    if (i == 1)
                        dt.Columns.Add(oSheet.Cells[i, j].Value == null ? "" : oSheet.Cells[i, j].Value.ToString());
                    else
                        dr[j - 1] = (oSheet.Cells[i, j].Value == null) ? "" : oSheet.Cells[i, j].Value.ToString();
                }
            }
            return dt;
        }
        private async Task<OrderDetailDTO> GetOrderDetail(string token, string apiUrl)
        {
            OrderDetailDTO result = new();
            using (var client = new HttpClient())
            {
                // Menambahkan header Authorization
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

                try
                {
                    HttpResponseMessage response = await client.GetAsync(apiUrl);
                    response.EnsureSuccessStatusCode();

                    string responseBody = await response.Content.ReadAsStringAsync();
                    result = JsonSerializer.Deserialize<OrderDetailDTO>(responseBody);

                }
                catch (HttpRequestException e)
                {
                    Console.WriteLine($"Request error: {e.Message}");
                }
            }
            return result;
        }
        private async Task<List<OrderDetailDTO>> GetOrderID(UploadExcelRequest request, string groupID) 
        {
            List<OrderDetailDTO> result = new();
            // URL API yang ingin diakses
            var apiUrl = $"{request.URLgetIdOrder}{groupID}&";
            using (var client = new HttpClient())
            {
                // Menambahkan header Authorization
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", request.token);

                try
                {
                    HttpResponseMessage response = await client.GetAsync(apiUrl);
                    response.EnsureSuccessStatusCode();

                    string responseBody = await response.Content.ReadAsStringAsync();
                    var apiResponse = JsonSerializer.Deserialize<ApiResponse>(responseBody);
                    foreach (var order in apiResponse.data.data)
                    {
                        var responseOrder = await GetOrderDetail(request.token, $"{request.URLgetOrderDetail}{order.id_so}");
                        if (responseOrder.data.id_group.ToString() == groupID)
                        {                            
                            result.Add(responseOrder);
                        }
                    }
                }
                catch (HttpRequestException e)
                {
                    Console.WriteLine($"Request error: {e.Message}");
                }
            }
            return result;
        }
        public async Task<DownloadExcelResponse?> DownloadRekapMidtrans(UploadExcelRequest request)
        {
            try
            {

                List<RekonsiliasiDTO> listRekonsiliasiDTO = [];
                string basePath = _configuration["FileUploadPath"];
                if (basePath == null)
                {
                    throw new Exception("FileUploadPath tidak ditemukan dalam konfigurasi.");
                }
                string fileName = request.UploadedFile.name;
                string pathToTempDir = Path.Combine(basePath);
                string pathToWriteFile = Path.Combine(pathToTempDir, fileName);
                if (!Directory.Exists(pathToTempDir))
                {
                    Directory.CreateDirectory(pathToTempDir);
                }

                BinaryWriter fileWriter = new BinaryWriter(System.IO.File.OpenWrite(pathToWriteFile));
                fileWriter.Write(request.UploadedFile.content);
                fileWriter.Flush();
                fileWriter.Close();

                //region Isi List Grid

                string sheetName = "Reconciliation";
                System.Data.DataTable dt = new System.Data.DataTable();

                using (ExcelPackage pck = new ExcelPackage(new FileInfo(fileName)))
                {
                    using (FileStream stream = new FileStream(pathToWriteFile, FileMode.Open))
                    {
                        pck.Load(stream);
                        ExcelWorksheet oSheet = pck.Workbook.Worksheets[sheetName];

                        //ExcelWorksheet oSheet = pck.Workbook.Worksheets.Add("Template");

                        //ExcelWorksheet oSheet = pck.Workbook.Worksheets[sheetName];
                        dt = WorksheetToDataTable(oSheet);

                        for (int i = 1; i < dt.Rows.Count; i++)
                        {
                            Console.WriteLine($"Processing Row : {i+2}\nOrder ID : {dt.Rows[i].ItemArray.GetValue(1).ToString()}");
                            string[] OrderID = dt.Rows[i].ItemArray.GetValue(1).ToString().Split('-');
                            string paramOrderID = dt.Rows[i].ItemArray.GetValue(1).ToString().Contains("CO")? OrderID[1]: OrderID[0];
                            if (string.IsNullOrEmpty(dt.Rows[i].ItemArray.GetValue(1).ToString())) break;
                            
                            var data = await GetOrderID(request, paramOrderID);
                            foreach (var item in data)
                            {
                                foreach (var detail in item.data.orderDetail.product)
                                {
                                    RekonsiliasiDTO rekonsiliasiDTO = new RekonsiliasiDTO();
                                    rekonsiliasiDTO.DateAndTime = dt.Rows[i].ItemArray.GetValue(0).ToString();
                                    rekonsiliasiDTO.OrderID = dt.Rows[i].ItemArray.GetValue(1).ToString();
                                    rekonsiliasiDTO.Channel = dt.Rows[i].ItemArray.GetValue(2).ToString();
                                    rekonsiliasiDTO.TransactionType = dt.Rows[i].ItemArray.GetValue(3).ToString();
                                    rekonsiliasiDTO.Amount = dt.Rows[i].ItemArray.GetValue(4).ToString();
                                    rekonsiliasiDTO.TransactionStatus = dt.Rows[i].ItemArray.GetValue(5).ToString();
                                    rekonsiliasiDTO.TransactionTime = dt.Rows[i].ItemArray.GetValue(6).ToString();
                                    rekonsiliasiDTO.CustomerEmail = dt.Rows[i].ItemArray.GetValue(8).ToString();
                                    rekonsiliasiDTO.GroupID = OrderID[1];
                                    rekonsiliasiDTO.IdOrder = item.data.id_so;
                                    DateTime dateTime = DateTime.ParseExact(rekonsiliasiDTO.DateAndTime, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);

                                    // Mengambil tanggal, bulan, dan tahun dalam bentuk string
                                    string tanggal = dateTime.ToString("dd");
                                    string bulan = dateTime.ToString("MM");
                                    string tahun = dateTime.ToString("yyyy");
                                    rekonsiliasiDTO.NoInvoice = $"OCI/{tanggal}{bulan}{tahun}/{item.data.id_so}";
                                    rekonsiliasiDTO.CustomerName = item.data.customerDetail.name;
                                    rekonsiliasiDTO.ProductName = detail.product;
                                    rekonsiliasiDTO.Quantity = detail.qty;
                                    rekonsiliasiDTO.QuantityPrice = detail.price;
                                    rekonsiliasiDTO.CbmPrice = detail.cbm_calcs;
                                    rekonsiliasiDTO.SubTotal = (detail.cbm_calcs + detail.price) * detail.qty;
                                    rekonsiliasiDTO.LocalShipping = item.data.financialInformation.expense.shippingCustomer;
                                    rekonsiliasiDTO.COGS = item.data.financialInformation.expense.cogs;
                                    rekonsiliasiDTO.VoucherDiscount = item.data.financialInformation.expense.discount;
                                    rekonsiliasiDTO.GrossProfit = item.data.financialInformation.expense.discount;
                                    rekonsiliasiDTO.Sales = item.data.sales;
                                    listRekonsiliasiDTO.Add(rekonsiliasiDTO);
                                }
                            }
                        }
                    }
                }


                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage pck = new ExcelPackage())
                {
                    Console.WriteLine($"Starting Mapping To Excel");
                    #region Header
                    ExcelWorksheet wsRecon = pck.Workbook.Worksheets.Add("Reconciliation");
                    wsRecon.Cells["A1:J1"].Merge = true;
                    wsRecon.Cells["A1:J1"].Value = "DATA MIDTRANS";
                    wsRecon.Cells["A1:J1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    wsRecon.Cells["A1:J1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#C6E0B4"));
                    wsRecon.Cells["K1:AB1"].Merge = true;
                    wsRecon.Cells["K1:AB1"].Value = "DATA OCISTOK";
                    wsRecon.Cells["K1:AB1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    wsRecon.Cells["K1:AB1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.ColorTranslator.FromHtml("#BDD7EE"));
                    wsRecon.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsRecon.Cells["A2"].Value = "Date & time";
                    wsRecon.Cells["B2"].Value = "Order ID";
                    wsRecon.Cells["C2"].Value = "Channel";
                    wsRecon.Cells["D2"].Value = "Transaction type";
                    wsRecon.Cells["E2"].Value = "Amount";
                    wsRecon.Cells["F2"].Value = "Transaction status";
                    wsRecon.Cells["G2"].Value = "Transaction time";
                    wsRecon.Cells["H2"].Value = "Transaction ID";
                    wsRecon.Cells["I2"].Value = "Customer e-mail";
                    wsRecon.Cells["J2"].Value = "Note";
                    wsRecon.Cells["K2"].Value = "Group ID";
                    wsRecon.Cells["L2"].Value = "Id Order";
                    wsRecon.Cells["M2"].Value = "No. Invoice";
                    wsRecon.Cells["N2"].Value = "Nama Customer";
                    wsRecon.Cells["O2"].Value = "Nama Produk";
                    wsRecon.Cells["P2"].Value = "Qty";
                    wsRecon.Cells["Q2"].Value = "Qty Price";
                    wsRecon.Cells["R2"].Value = "Cbm Price";
                    wsRecon.Cells["S2"].Value = "Sub Total";
                    wsRecon.Cells["T2"].Value = "Local Shipping Indo";
                    wsRecon.Cells["U2"].Value = "Total Price";
                    wsRecon.Cells["V2"].Value = "Selisih";
                    wsRecon.Cells["W2"].Value = "Notes";
                    wsRecon.Cells["X2"].Value = "COGS";
                    wsRecon.Cells["Y2"].Value = "Voucher Diskon";
                    wsRecon.Cells["Z2"].Value = "COGS - Voucher Diskon";
                    wsRecon.Cells["AA2"].Value = "Adjustment Ongkir China";
                    wsRecon.Cells["AB2"].Value = "Gross Profit";
                    wsRecon.Cells["AC2"].Value = "Refund";
                    wsRecon.Cells["AD2"].Value = "Notes";
                    wsRecon.Cells["AE2"].Value = "Sales";
                    wsRecon.Cells["AF2"].Value = "SS Paid OMS";
                    wsRecon.Cells["AG2"].Value = "SS Financial Information";
                    wsRecon.Cells["AH2"].Value = "Payment to vendor";
                    wsRecon.Cells["AI2"].Value = "Delivery to Indonesia";
                    wsRecon.Cells["AJ2"].Value = "Goods in Indonesia";
                    wsRecon.Cells["AK2"].Value = "Goods in warehouse";
                    wsRecon.Cells["AL2"].Value = "Delivery to Customer";
                    wsRecon.Cells["AM2"].Value = "Receive by customer";
                    wsRecon.Row(2).Style.Font.Bold = true;
                    wsRecon.Row(2).Height = 60;
                    wsRecon.Cells["A2:AM2"].AutoFilter = true;                    
                    wsRecon.Row(2).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    wsRecon.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    wsRecon.Cells["A2:AM2"].Style.WrapText = true;
                    #endregion
                    #region Content
                    int Row = 0;
                    int fromRow = 0; int toRow = 0; int fromRowOrder = 0; int toRowOrder = 0;
                    for (int i = 0; i < listRekonsiliasiDTO.Count; i++)
                    {
                        Row = i + 3;                        
                        wsRecon.Cells[$"O{Row}"].Value = listRekonsiliasiDTO[i].ProductName;
                        wsRecon.Cells[$"P{Row}"].Value = listRekonsiliasiDTO[i].Quantity;
                        wsRecon.Cells[$"Q{Row}"].Value = listRekonsiliasiDTO[i].QuantityPrice;
                        wsRecon.Cells[$"Q{Row}"].Style.Numberformat.Format = "#,##0";
                        wsRecon.Cells[$"R{Row}"].Value = listRekonsiliasiDTO[i].CbmPrice;
                        wsRecon.Cells[$"R{Row}"].Style.Numberformat.Format = "#,##0";
                        wsRecon.Cells[$"S{Row}"].Value = listRekonsiliasiDTO[i].SubTotal;
                        wsRecon.Cells[$"S{Row}"].Style.Numberformat.Format = "#,##0";

                        wsRecon.Cells[$"B{Row}"].Value = listRekonsiliasiDTO[i].OrderID;
                        wsRecon.Cells[$"M{Row}"].Value = listRekonsiliasiDTO[i].NoInvoice;
                        wsRecon.Cells[$"I{Row}"].Value = listRekonsiliasiDTO[i].CustomerEmail;
                        wsRecon.Cells[$"C{Row}"].Value = listRekonsiliasiDTO[i].Channel;
                        wsRecon.Cells[$"D{Row}"].Value = listRekonsiliasiDTO[i].TransactionType;
                        wsRecon.Cells[$"F{Row}"].Value = listRekonsiliasiDTO[i].TransactionStatus;
                        wsRecon.Cells.AutoFitColumns();
                        fromRowOrder = (i == 0 || (i > 0 && (listRekonsiliasiDTO[i].IdOrder != listRekonsiliasiDTO[i - 1].IdOrder))) ? Row : fromRowOrder;
                        toRowOrder = (i == listRekonsiliasiDTO.Count - 1 || i < listRekonsiliasiDTO.Count && (listRekonsiliasiDTO[i].IdOrder != listRekonsiliasiDTO[i + 1].IdOrder)) ? Row : toRowOrder;
                        if (fromRowOrder > 0 && toRowOrder > 0)
                        {
                            wsRecon.Cells[$"L{fromRowOrder}:L{toRowOrder}"].Merge = true;
                            wsRecon.Cells[$"L{fromRowOrder}:L{toRowOrder}"].Value = listRekonsiliasiDTO[i].IdOrder;
                            wsRecon.Cells[$"M{fromRowOrder}:M{toRowOrder}"].Merge = true;
                            wsRecon.Cells[$"M{fromRowOrder}:M{toRowOrder}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wsRecon.Cells[$"N{fromRowOrder}:N{toRowOrder}"].Merge = true;
                            wsRecon.Cells[$"N{fromRowOrder}:N{toRowOrder}"].Value = listRekonsiliasiDTO[i].CustomerName;
                            wsRecon.Cells[$"T{fromRowOrder}:T{toRowOrder}"].Merge = true;
                            wsRecon.Cells[$"T{fromRowOrder}:T{toRowOrder}"].Value = listRekonsiliasiDTO[i].LocalShipping;
                            wsRecon.Cells[$"T{fromRowOrder}:T{toRowOrder}"].Style.Numberformat.Format = "#,##0";
                            wsRecon.Cells[$"U{fromRowOrder}:U{toRowOrder}"].Merge = true;
                            wsRecon.Cells[$"U{fromRowOrder}"].Formula = $"=SUM(S{fromRowOrder}:S{toRowOrder})+T{fromRowOrder}";
                            wsRecon.Cells[$"U{fromRowOrder}:U{toRowOrder}"].Style.Numberformat.Format = "#,##0";
                            wsRecon.Cells[$"W{fromRowOrder}:W{toRowOrder}"].Merge = true;
                            wsRecon.Cells[$"W{fromRowOrder}:W{toRowOrder}"].Value = "";
                            wsRecon.Cells[$"X{fromRowOrder}:X{toRowOrder}"].Merge = true;
                            wsRecon.Cells[$"X{fromRowOrder}:X{toRowOrder}"].Value = listRekonsiliasiDTO[i].COGS;
                            wsRecon.Cells[$"X{fromRowOrder}:X{toRowOrder}"].Style.Numberformat.Format = "#,##0";
                            wsRecon.Cells[$"Y{fromRowOrder}:Y{toRowOrder}"].Merge = true;
                            wsRecon.Cells[$"Y{fromRowOrder}:Y{toRowOrder}"].Value = listRekonsiliasiDTO[i].VoucherDiscount;
                            wsRecon.Cells[$"Y{fromRowOrder}:Y{toRowOrder}"].Style.Numberformat.Format = "#,##0";
                            wsRecon.Cells[$"Z{fromRowOrder}:Z{toRowOrder}"].Merge = true;
                            wsRecon.Cells[$"Z{fromRowOrder}:Z{toRowOrder}"].Formula = $"=X{fromRowOrder}-Y{fromRowOrder}";
                            wsRecon.Cells[$"Z{fromRowOrder}:Z{toRowOrder}"].Style.Numberformat.Format = "#,##0";
                            wsRecon.Cells[$"AA{fromRowOrder}:AA{toRowOrder}"].Merge = true;
                            //wsRecon.Cells[$"AA{fromRowOrder}:AA{toRowOrder}"].Value = "";
                            wsRecon.Cells[$"AB{fromRowOrder}:AB{toRowOrder}"].Merge = true;
                            wsRecon.Cells[$"AB{fromRowOrder}:AB{toRowOrder}"].Formula = $"=U{fromRowOrder}-X{fromRowOrder}+Y{fromRowOrder}+AA{fromRowOrder}";
                            wsRecon.Cells[$"AB{fromRowOrder}:AB{toRowOrder}"].Style.Numberformat.Format = "#,##0";
                            wsRecon.Cells[$"AC{fromRowOrder}:AC{toRowOrder}"].Merge = true;
                            //wsRecon.Cells[$"AC{fromRowOrder}:AC{toRowOrder}"].Value = "";
                            wsRecon.Cells[$"AD{fromRowOrder}:AD{toRowOrder}"].Merge = true;
                            //wsRecon.Cells[$"AD{fromRowOrder}:AD{toRowOrder}"].Value = "";
                            wsRecon.Cells[$"AE{fromRowOrder}:AE{toRowOrder}"].Merge = true;
                            wsRecon.Cells[$"AE{fromRowOrder}:AE{toRowOrder}"].Value = listRekonsiliasiDTO[i].Sales;
                            wsRecon.Cells[$"AE{fromRowOrder}:AE{toRowOrder}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                            fromRowOrder = toRowOrder + 1;
                            toRowOrder = 0;
                        }
                        fromRow = (i == 0 || (i > 0 && (listRekonsiliasiDTO[i].OrderID != listRekonsiliasiDTO[i - 1].OrderID))) ? Row : fromRow;
                        toRow = (i == listRekonsiliasiDTO.Count - 1 || i < listRekonsiliasiDTO.Count && (listRekonsiliasiDTO[i].OrderID != listRekonsiliasiDTO[i + 1].OrderID)) ? Row : toRow;
                        if (fromRow > 0 && toRow > 0)
                        {
                            wsRecon.Cells[$"A{fromRow}:A{toRow}"].Merge = true;
                            wsRecon.Cells[$"A{fromRow}:A{toRow}"].Value = DateTime.Parse(listRekonsiliasiDTO[i].DateAndTime);
                            wsRecon.Cells[$"A{fromRow}:A{toRow}"].Style.Numberformat.Format = "dd/MM/yyyy";
                            wsRecon.Cells[$"B{fromRow}:B{toRow}"].Merge = true;
                            Console.WriteLine($"Mapping For Order ID : {listRekonsiliasiDTO[i].OrderID}");
                            wsRecon.Cells[$"C{fromRow}:C{toRow}"].Merge = true;
                            wsRecon.Cells[$"D{fromRow}:D{toRow}"].Merge = true;
                            wsRecon.Cells[$"A{fromRow}:D{toRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;                            
                            wsRecon.Cells[$"E{fromRow}:E{toRow}"].Merge = true;
                            wsRecon.Cells[$"E{fromRow}:E{toRow}"].Value = double.Parse(listRekonsiliasiDTO[i].Amount);
                            wsRecon.Cells[$"E{fromRow}:E{toRow}"].Style.Numberformat.Format = "#,##0";
                            wsRecon.Cells[$"F{fromRow}:F{toRow}"].Merge = true;
                            wsRecon.Cells[$"F{fromRow}:F{toRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wsRecon.Cells[$"G{fromRow}:G{toRow}"].Merge = true;
                            wsRecon.Cells[$"G{fromRow}:G{toRow}"].Value = DateTime.Parse(listRekonsiliasiDTO[i].TransactionTime);
                            wsRecon.Cells[$"G{fromRow}:G{toRow}"].Style.Numberformat.Format = "dd/MM/yyyy";
                            wsRecon.Cells[$"G{fromRow}:G{toRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wsRecon.Cells[$"H{fromRow}:H{toRow}"].Merge = true;
                            wsRecon.Cells[$"H{fromRow}:H{toRow}"].Value = listRekonsiliasiDTO[i].TransactionID;
                            wsRecon.Cells[$"I{fromRow}:I{toRow}"].Merge = true;
                            wsRecon.Cells[$"I{fromRow}:I{toRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wsRecon.Cells[$"J{fromRow}:J{toRow}"].Merge = true;
                            wsRecon.Cells[$"J{fromRow}:J{toRow}"].Value = listRekonsiliasiDTO[i].Note;
                            wsRecon.Cells[$"K{fromRow}:K{toRow}"].Merge = true;
                            wsRecon.Cells[$"K{fromRow}:K{toRow}"].Value = listRekonsiliasiDTO[i].GroupID;
                            wsRecon.Cells[$"V{fromRow}:V{toRow}"].Merge = true;
                            wsRecon.Cells[$"V{fromRow}:V{toRow}"].Formula = $"=IF(E{fromRow}-SUM(U{fromRow}:U{toRow})=0, \"-\", E{fromRow}-SUM(U{fromRow}:U{toRow}))";
                            wsRecon.Cells[$"V{fromRow}:V{toRow}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            wsRecon.Cells[$"V{fromRow}:V{toRow}"].Style.Numberformat.Format = "#,##0";
                            fromRow = toRow + 1;
                            toRow = 0;
                        }
                        
                    }
                    #endregion
                    for (int col = 1; col <= wsRecon.Dimension.End.Column; col++)
                    {
                        wsRecon.Column(col).Width += 5;
                        wsRecon.Column(col).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    }
                    var border = wsRecon.Cells[$"A1:AM{Row}"].Style.Border;

                    border.Top.Style = ExcelBorderStyle.Thin;
                    border.Bottom.Style = ExcelBorderStyle.Thin;
                    border.Left.Style = ExcelBorderStyle.Thin;
                    border.Right.Style = ExcelBorderStyle.Thin;

                    return new DownloadExcelResponse()
                    {
                        FileContents = pck.GetAsByteArray(),
                        ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        FileDownloadName = $"Rekonsiliasi_Midtrans.xlsx"
                    };
                }
            }
            catch (Exception ex) 
            {
                Console.WriteLine(ex.Message);
                throw ex;
            }
        }
    }
}
