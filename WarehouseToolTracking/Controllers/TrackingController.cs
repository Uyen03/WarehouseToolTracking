using ExcelDataReader;
using Microsoft.AspNetCore.Mvc;
using System.Data;
using WarehouseToolTracking.Models;
using ClosedXML.Excel;    
using System.IO;

namespace WarehouseToolTracking.Controllers
{
    public class TrackingController : Controller
    {
        private static readonly string BaoCaoFolder = @"E:\Project\WarehouseToolTracking\BaoCaoTracking\";
        private static DataTable dtDSNV;

        private static DataTable dtExcel;
        static TrackingController()
        {
            if (!Directory.Exists(BaoCaoFolder))
                Directory.CreateDirectory(BaoCaoFolder);
        }

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
        //[HttpPost]
        //public IActionResult SaveRecord(TrackingRecord record)
        //{
        //    try
        //    {
        //        string fileName = $"BaoCao_Tracking_{DateTime.Today:dd-MM-yyyy}.xlsx";
        //        string fullPath = Path.Combine(BaoCaoFolder, fileName);

        //        // SỬA LỖI Ở ĐÂY: Dùng System.IO.File.Exists
        //        using (var workbook = System.IO.File.Exists(fullPath)
        //            ? new XLWorkbook(fullPath)
        //            : new XLWorkbook())
        //        {
        //            var worksheet = workbook.Worksheets.FirstOrDefault(w => w.Name == "DonDaTra")
        //                         ?? workbook.Worksheets.Add("DonDaTra");

        //            // Tạo header nếu file mới
        //            if (worksheet.Cell(1, 1).Value.ToString() == "")
        //            {
        //                worksheet.Cell(1, 1).Value = "Thời gian";
        //                worksheet.Cell(1, 2).Value = "Ngày";
        //                worksheet.Cell(1, 3).Value = "Ca làm việc";
        //                worksheet.Cell(1, 4).Value = "Tên NV Tracking";
        //                worksheet.Cell(1, 5).Value = "List ID";
        //                worksheet.Cell(1, 6).Value = "SKU";
        //                worksheet.Cell(1, 7).Value = "Vị trí Thiếu";
        //                worksheet.Cell(1, 8).Value = "SL Thiếu";
        //                worksheet.Cell(1, 9).Value = "Vị trí Lấy Bù";
        //                worksheet.Cell(1, 10).Value = "SL Lấy Bù";
        //                worksheet.Cell(1, 11).Value = "Ghi chú";
        //            }

        //            int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
        //            lastRow++;

        //            worksheet.Cell(lastRow, 1).Value = record.ThoiGian;
        //            worksheet.Cell(lastRow, 2).Value = record.Ngay;
        //            worksheet.Cell(lastRow, 3).Value = record.CaLamViec;
        //            worksheet.Cell(lastRow, 4).Value = record.TenNVTracking;
        //            worksheet.Cell(lastRow, 5).Value = record.ListID;
        //            worksheet.Cell(lastRow, 6).Value = record.SKU;
        //            worksheet.Cell(lastRow, 7).Value = record.ViTriThieu;
        //            worksheet.Cell(lastRow, 8).Value = record.SLThieu;
        //            worksheet.Cell(lastRow, 9).Value = record.ViTriLayBu;
        //            worksheet.Cell(lastRow, 10).Value = record.SLLayBu;
        //            worksheet.Cell(lastRow, 11).Value = record.GhiChu;

        //            workbook.SaveAs(fullPath);
        //        }

        //        return Json(new { success = true });
        //    }
        //    catch (Exception ex)
        //    {
        //        return Json(new { success = false, message = ex.Message });
        //    }
        //}
        [HttpPost]
        public IActionResult SaveRecord(TrackingRecord record)
        {
            try
            {
                // Backend tự lấy thời gian chính xác (giờ Việt Nam)
                record.ThoiGian = DateTime.Now;

                string dbPath = Path.Combine(BaoCaoFolder, "TrackingData.db");

                using (var connection = new Microsoft.Data.Sqlite.SqliteConnection($"Data Source={dbPath}"))
                {
                    connection.Open();

                    string createTable = @"
                CREATE TABLE IF NOT EXISTS DonDaTra (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    ThoiGian TEXT,
                    Ngay TEXT,
                    CaLamViec TEXT,
                    TenNVTracking TEXT,
                    ListID TEXT,
                    SKU TEXT,
                    ViTriThieu TEXT,
                    SLThieu INTEGER,
                    ViTriLayBu TEXT,
                    SLLayBu INTEGER,
                    GhiChu TEXT
                )";

                    using (var cmd = new Microsoft.Data.Sqlite.SqliteCommand(createTable, connection))
                        cmd.ExecuteNonQuery();

                    string insert = @"
                INSERT INTO DonDaTra 
                (ThoiGian, Ngay, CaLamViec, TenNVTracking, ListID, SKU, ViTriThieu, SLThieu, ViTriLayBu, SLLayBu, GhiChu)
                VALUES 
                (@ThoiGian, @Ngay, @CaLamViec, @TenNVTracking, @ListID, @SKU, @ViTriThieu, @SLThieu, @ViTriLayBu, @SLLayBu, @GhiChu)";

                    using (var cmd = new Microsoft.Data.Sqlite.SqliteCommand(insert, connection))
                    {
                        cmd.Parameters.AddWithValue("@ThoiGian", record.ThoiGian.ToString("yyyy-MM-dd HH:mm:ss"));
                        cmd.Parameters.AddWithValue("@Ngay", record.Ngay);
                        cmd.Parameters.AddWithValue("@CaLamViec", record.CaLamViec);
                        cmd.Parameters.AddWithValue("@TenNVTracking", record.TenNVTracking);
                        cmd.Parameters.AddWithValue("@ListID", record.ListID);
                        cmd.Parameters.AddWithValue("@SKU", record.SKU);
                        cmd.Parameters.AddWithValue("@ViTriThieu", record.ViTriThieu);
                        cmd.Parameters.AddWithValue("@SLThieu", record.SLThieu);
                        cmd.Parameters.AddWithValue("@ViTriLayBu", record.ViTriLayBu);
                        cmd.Parameters.AddWithValue("@SLLayBu", record.SLLayBu);
                        cmd.Parameters.AddWithValue("@GhiChu", record.GhiChu);

                        cmd.ExecuteNonQuery();
                    }
                }

                // Tự động xuất Excel sau khi lưu
                ExportToExcel();

                return Json(new { success = true });
            }
            catch (Exception ex)
            {
                return Json(new { success = false, message = ex.Message });
            }
        }

        private void ExportToExcel()
        {
            try
            {
                string dbPath = Path.Combine(BaoCaoFolder, "TrackingData.db");
                string excelPath = Path.Combine(BaoCaoFolder, $"BaoCao_Tracking_{DateTime.Today:dd-MM-yyyy}.xlsx");

                using (var connection = new Microsoft.Data.Sqlite.SqliteConnection($"Data Source={dbPath}"))
                {
                    connection.Open();
                    using (var cmd = new Microsoft.Data.Sqlite.SqliteCommand("SELECT * FROM DonDaTra ORDER BY Id ASC", connection))
                    using (var reader = cmd.ExecuteReader())
                    {
                        using (var workbook = new XLWorkbook())
                        {
                            var ws = workbook.Worksheets.Add("DonDaTra");

                            // Font mặc định size 18 cho toàn sheet
                            ws.Style.Font.FontSize = 18;

                            // ==================== TIÊU ĐỀ CHÍNH ====================
                            ws.Cell(1, 1).Value = "TRACKING ĐƠN THIẾU KHO";
                            ws.Range("A1:J1").Merge();
                            ws.Cell(1, 1).Style.Font.Bold = true;
                            ws.Cell(1, 1).Style.Font.FontSize = 20;
                            ws.Cell(1, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                            // ==================== HEADER CHÍNH ====================
                            // ĐƠN THIẾU (màu đỏ) - chỉ từ A3 đến G3
                            ws.Cell(3, 1).Value = "ĐƠN THIẾU";
                            ws.Range("A3:G3").Merge();
                            ws.Cell(3, 1).Style.Fill.BackgroundColor = XLColor.Red;
                            ws.Cell(3, 1).Style.Font.FontColor = XLColor.White;
                            ws.Cell(3, 1).Style.Font.Bold = true;
                            ws.Cell(3, 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                            // VỊ TRÍ THAY THẾ (màu vàng) - chỉ từ H3 đến I3
                            ws.Cell(3, 8).Value = "VỊ TRÍ THAY THẾ";
                            ws.Range("H3:I3").Merge();
                            ws.Cell(3, 8).Style.Fill.BackgroundColor = XLColor.Yellow;
                            ws.Cell(3, 8).Style.Font.Bold = true;
                            ws.Cell(3, 8).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                            // Header chi tiết dòng 4
                            string[] headers = { "DAY", "CA", "TÊN NGƯỜI TRA ĐƠN", "LIST ID", "SKU",
                                       "VỊ TRÍ THIẾU", "SỐ LƯỢNG THIẾU",
                                       "VỊ TRÍ LẤY BÙ", "SỐ LƯỢNG LẤY BÙ", "NOTE" };

                            for (int i = 0; i < headers.Length; i++)
                            {
                                ws.Cell(4, i + 1).Value = headers[i];
                            }

                            // Format header dòng 4
                            var headerRange = ws.Range("A4:J4");
                            headerRange.Style.Font.Bold = true;
                            headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                            headerRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                            headerRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                            // ==================== ĐỔ DỮ LIỆU ====================
                            int row = 5;
                            while (reader.Read())
                            {
                                ws.Cell(row, 1).Value = reader.GetString(2);   // DAY
                                ws.Cell(row, 2).Value = reader.GetString(3);   // CA
                                ws.Cell(row, 3).Value = reader.GetString(4);   // TÊN NGƯỜI TRA ĐƠN
                                ws.Cell(row, 4).Value = reader.GetString(5);   // LIST ID
                                ws.Cell(row, 5).Value = reader.GetString(6);   // SKU
                                ws.Cell(row, 6).Value = reader.GetString(7);   // VỊ TRÍ THIẾU
                                ws.Cell(row, 7).Value = reader.GetInt32(8);    // SỐ LƯỢNG THIẾU
                                ws.Cell(row, 8).Value = reader.GetString(9);   // VỊ TRÍ LẤY BÙ
                                ws.Cell(row, 9).Value = reader.GetInt32(10);   // SỐ LƯỢNG LẤY BÙ
                                ws.Cell(row, 10).Value = reader.GetString(11); // NOTE

                                // Viền cho từng dòng dữ liệu
                                var dataRange = ws.Range(row, 1, row, 10);
                                dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                                dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

                                row++;
                            }

                            // Auto fit cột và viền toàn bảng
                            ws.Columns().AdjustToContents();
                            ws.Range("A3:J" + (row - 1)).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;

                            workbook.SaveAs(excelPath);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Lỗi xuất Excel: " + ex.Message);
            }
        }
    }
}



