namespace WarehouseToolTracking.Models
{
    public class TrackingModel
    {
        public string ListID { get; set; } = string.Empty;
        public string SKU { get; set; } = string.Empty;
        public string TenSanPham { get; set; } = string.Empty;
        public string ViTriThieu { get; set; } = string.Empty;
        public int SLThieu { get; set; }
        public string ViTriLayBu { get; set; } = string.Empty;     // Có thể chứa nhiều vị trí cách nhau bằng dấu |
        public int SLLayBu { get; set; }
        public DateTime ThoiGian { get; set; } = DateTime.Now;
        public string NhanVien { get; set; } = string.Empty;
        public string CaLamViec { get; set; } = string.Empty;
        public string Note { get; set; } = string.Empty;
    }
}
