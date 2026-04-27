namespace WarehouseToolTracking.Models
{
    public class TrackingRecord
    {
        public DateTime ThoiGian { get; set; } = DateTime.Now;
        public string Ngay { get; set; } = string.Empty;
        public string CaLamViec { get; set; } = string.Empty;
        public string TenNVTracking { get; set; } = string.Empty;
        public string ListID { get; set; } = string.Empty;
        public string SKU { get; set; } = string.Empty;
        public string ViTriThieu { get; set; } = string.Empty;
        public int SLThieu { get; set; }
        public string ViTriLayBu { get; set; } = string.Empty;
        public int SLLayBu { get; set; }
        public string GhiChu { get; set; } = string.Empty;
    }
}