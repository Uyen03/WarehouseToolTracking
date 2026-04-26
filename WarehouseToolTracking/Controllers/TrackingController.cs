using ExcelDataReader;
using Microsoft.AspNetCore.Mvc;
using System.Data;
using WarehouseToolTracking.Models;

namespace WarehouseToolTracking.Controllers
{
    public class TrackingController : Controller
    {
        private static DataTable dtDSNV;

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
                string filePath = @"E:\Project\16.04.2026.M.xlsm";   // ← Đường dẫn của bạn

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
                                // Skip 7 dòng cho sheet Count (header ở dòng 8)
                                for (int i = 0; i < 7; i++) rowReader.Read();
                            }
                        }
                    };

                    var result = reader.AsDataSet(conf);

                    dtExcel = result.Tables["Count"];

                    // === LOAD SHEET DSNV (header ở dòng 2) ===
                    if (result.Tables.Contains("DSNV"))
                    {
                        var dtTemp = result.Tables["DSNV"];
                        dtDSNV = dtTemp.Clone();

                        // Bắt đầu từ dòng 2 (index 1) để lấy header
                        for (int i = 1; i < dtTemp.Rows.Count; i++)
                        {
                            dtDSNV.ImportRow(dtTemp.Rows[i]);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Lỗi load Excel: " + ex.Message);
            }
        }

        public IActionResult Start()
        {
            if (dtDSNV == null)
                LoadExcelFile();

            if (dtDSNV == null || dtDSNV.Columns.Count == 0)
            {
                ViewBag.Error = "Không load được sheet DSNV!";
                return View(new ShiftSelectionModel());
            }

            // DEBUG: Hiển thị tất cả tên cột
            var columnList = string.Join(", ", dtDSNV.Columns.Cast<DataColumn>().Select(c => $"[{c.ColumnName}]"));
            ViewBag.ColumnDebug = "Các cột trong DSNV: " + columnList;

            // Load danh sách nhân viên
            ViewBag.DSNV = dtDSNV.AsEnumerable()
                .Select(row => row.Field<string>(4)?.Trim())   
                .Where(x => !string.IsNullOrEmpty(x))
                .Distinct()
                .OrderBy(x => x)
                .ToList();

            return View(new ShiftSelectionModel());
        }

        [HttpPost]
        public IActionResult Start(ShiftSelectionModel model)
        {
            if (string.IsNullOrEmpty(model.CaLamViec) || string.IsNullOrEmpty(model.TenNhanVien))
            {
                ViewBag.Error = "Vui lòng chọn đầy đủ Ca làm việc và Tên nhân viên";
                return View(model);
            }

            // Lưu thông tin vào Session để Form Tracking sử dụng sau
            HttpContext.Session.SetString("CaLamViec", model.CaLamViec);
            HttpContext.Session.SetString("TenNhanVien", model.TenNhanVien);
            HttpContext.Session.SetString("NgayLamViec", model.NgayLamViec.ToString("dd/MM/yyyy"));

            return RedirectToAction("Index", "Tracking");
        }
    }
}


