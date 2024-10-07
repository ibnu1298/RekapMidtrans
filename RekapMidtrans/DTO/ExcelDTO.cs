namespace RekapMidtrans.DTO
{
    public class ExcelDTO
    {
    }
    public class DownloadExcelResponse
    {
        public byte[] FileContents { get; set; }
        public string ContentType { get; set; }
        public string FileDownloadName { get; set; }
    }
    public class UploadExcelRequest
    {
        public string URLgetOrderDetail { get; set; }        
        public string URLgetIdOrder { get; set; }        
        public string token { get; set; }        
        public TalentAttachment? UploadedFile { get; set; }
    }
    public class TalentAttachment
    {
        public string fileName { get; set; }
        public String name { get; set; }
        public String extension { get; set; }
        public Int32 size { get; set; }
        public Byte[] content { get; set; }

        public TalentAttachment()
        {

        }

        public TalentAttachment(string name, int contentLength, byte[] content)
        {
            this.name = name;
            this.size = contentLength;
            this.content = content;
        }
    }
    public class ListOrderID
    {
        public List<int> OrderID { get; set; } 
    }
}
