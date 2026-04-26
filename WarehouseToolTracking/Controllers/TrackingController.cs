using ExcelDataReader;
using Microsoft.AspNetCore.Mvc;
using System.Data;
using WarehouseToolTracking.Models;

namespace WarehouseToolTracking.Controllers
{
    public class TrackingController : Controller
    {

        private static DataTable dtExcel;

        public IActionResult Index()
        {
            if (dtExcel == null)
                LoadExcelFile();

            return View(new TrackingModel());
        }

        [HttpPost]
        public IActionResult SearchBySKU(string sku)
        {
            if (string.IsNullOrWhiteSpace(sku))
                return Json(new { success = false, message = "Vui lòng nhập SKU" });

            if (dtExcel == null)
                LoadExcelFile();

            // Kiểm tra lại lần nữa sau khi load
            if (dtExcel == null)
                return Json(new { success = false, message = "Không load được file Excel. Kiểm tra đường dẫn file!" });

            try
            {
                DataView dv = new DataView(dtExcel);
                dv.RowFilter = $"[Mã sản phẩm] = '{sku.Trim()}'";

                if (dv.Count == 0)
                {
                    return Json(new { success = false, message = $"Không tìm thấy SKU: {sku}" });
                }

                var model = new TrackingModel
                {
                    SKU = sku,
                    TenSanPham = dv[0]["Tên sản phẩm"]?.ToString() ?? ""
                };

                var positions = dv.Cast<DataRowView>()
                    .Select(row => new
                    {
                        Barcode = row["Barcode"]?.ToString(),
                        SKU = row["Mã sản phẩm"]?.ToString(),
                        TenSanPham = row["Tên sản phẩm"]?.ToString(),
                        ViTri = row["Vị trí"]?.ToString(),
                        OnHand = Convert.ToInt32(row["On Hand"] ?? 0),
                        Allocated = Convert.ToInt32(row["Allocated"] ?? 0)
                    })
                    .ToList();

                return Json(new { success = true, model = model, positions = positions });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = "Lỗi: " + ex.Message });
            }
        }

        private void LoadExcelFile()
        {
            try
            {
                // ==================== SỬA ĐƯỜNG DẪN FILE Ở ĐÂY ====================
                string filePath = @"E:\Project\16.04.2026.M.xlsm"; 

                using (var stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var conf = new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                        {
                            UseHeaderRow = true,
                            ReadHeaderRow = (rowReader) =>
                            {
                                // Bỏ qua 7 dòng trống để dòng 8 là header
                                for (int i = 0; i < 7; i++)
                                    rowReader.Read();
                            }
                        }
                    };

                    var result = reader.AsDataSet(conf);
                    dtExcel = result.Tables["Count"];
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Lỗi load Excel: " + ex.Message);
                dtExcel = null; // Đảm bảo rõ ràng là null khi lỗi
            }
        }
    }
}


