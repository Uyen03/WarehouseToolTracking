namespace WarehouseToolTracking.Models
{
    public class ShiftSelectionModel
    {
        public string CaLamViec { get; set; } = string.Empty;
        public string TenNhanVien { get; set; } = string.Empty;
        public DateTime NgayLamViec { get; set; } = DateTime.Today;
    }
}
