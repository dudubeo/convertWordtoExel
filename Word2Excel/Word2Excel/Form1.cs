using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Mammoth;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace Word2Excel
{
    public partial class Form1 : Form
    {
        private Microsoft.Office.Interop.Excel.Application _xlApp;
        private Microsoft.Office.Interop.Excel.Workbook _xlWorkbook;
        private string _fileName = string.Empty;
        private string _tempdata = string.Empty;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
        }

        private void ExportExcel(string fileName, List<Data> list)
        {
            using (var workbook = SpreadsheetDocument.Create(fileName, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook))
            {
                // Create Workbook
                var workbookPart = workbook.AddWorkbookPart();
                workbook.WorkbookPart.Workbook = new Workbook();
                workbook.WorkbookPart.Workbook.Sheets = new Sheets();

                progressBar1.Maximum = list.Count + 1;
                progressBar1.Value = 0;

                // Add Sheet
                var sheetPart = workbook.WorkbookPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                sheetPart.Worksheet = new Worksheet(sheetData);

                var sheets = workbook.WorkbookPart.Workbook.GetFirstChild<Sheets>();
                string relationshipId = workbook.WorkbookPart.GetIdOfPart(sheetPart);

                uint sheetId = 1;
                if (sheets.Elements<Sheet>().Count() > 0)
                {
                    sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }

                var sheet = new Sheet()
                {
                    Id = relationshipId,
                    SheetId = sheetId,
                    Name = "Exported Data"
                };
                sheets.Append(sheet);

                // Add Header
                var headerRow = new Row();

                var columns = new List<string>()
                    {
                        "Số tờ",
                        "Số thửa",
                        "Diện tích",
                        "Mục đích sử dụng đất",
                        "Địa chỉ thửa đất (1), hoặc địa chỉ sau khi thay đổi (2). (nếu có (2) thay (1)",
                        "Họ và tên chủ hộ",
                        "Năm sinh",
                        "Giới tính",
                        "Loại giấy tờ",
                        "Số giấy tờ",
                        "Ngày cấp",
                        "Nơi cấp",
                        "Địa chỉ thường trú (1), hoặc địa chỉ sau khi thay đổi (2). (nếu có (2) thay (1)",
                        "Họ và tên vợ/chồng",
                        "Năm sinh",
                        "Giới tính",
                        "Loại giấy tờ",
                        "Số giấy tờ",
                        "Ngày cấp",
                        "Nơi cấp",
                        "Địa chỉ thường trú (1), hoặc địa chỉ sau khi thay đổi (2). (nếu có (2) thay (1)",
                        "Họ và tên chủ hộ",
                        "Năm sinh",
                        "Giới tính",
                        "Loại giấy tờ",
                        "Số giấy tờ",
                        "Ngày cấp",
                        "Nơi cấp",
                        "Địa chỉ thường trú (1), hoặc địa chỉ sau khi thay đổi (2). (nếu có (2) thay (1)",
                        "Họ và tên vợ/chồng",
                        "Năm sinh",
                        "Giới tính",
                        "Loại giấy tờ",
                        "Số giấy tờ",
                        "Ngày cấp",
                        "Nơi cấp",
                        "Địa chỉ thường trú",
                        "Hình thức",
                        "Hợp đồng",
                        "Giấy chứng nhận số (Seria)",
                        "Mã vạch",
                        "Số vào sổ cấp GCN",
                        "Ngày/tháng/năm cấp GCN",
                        "Ngày/tháng/năm nhận chuyển nhượng GCN"
                    };

                foreach (var column in columns)
                {
                    headerRow.AppendChild(CreateCell(column));
                }

                sheetData.AppendChild(headerRow);

                // Add Rows
                foreach (var line in list)
                {
                    progressBar1.Value++;
                    var cells = new List<string>()
                        {
                            line.ThongTinDat.SoTo,
                            line.ThongTinDat.SoThua,
                            line.ThongTinDat.DienTich,
                            line.ThongTinDat.MucDichSuDung,
                            line.ThongTinDat.DiaChiThuaDat,
                            line.ChuSuDung.ChuHo.HoTen,
                            line.ChuSuDung.ChuHo.NamSinh,
                            line.ChuSuDung.ChuHo.GioiTinh,
                            line.ChuSuDung.ChuHo.LoaiGiayTo,
                            line.ChuSuDung.ChuHo.SoGiayTo,
                            line.ChuSuDung.ChuHo.NgayCap,
                            line.ChuSuDung.ChuHo.NoiCap,
                            line.ChuSuDung.ChuHo.DiaChiThuongTru,
                            line.ChuSuDung.VoChong.HoTen,
                            line.ChuSuDung.VoChong.NamSinh,
                            line.ChuSuDung.VoChong.GioiTinh,
                            line.ChuSuDung.VoChong.LoaiGiayTo,
                            line.ChuSuDung.VoChong.SoGiayTo,
                            line.ChuSuDung.VoChong.NgayCap,
                            line.ChuSuDung.VoChong.NoiCap,
                            line.ChuSuDung.VoChong.DiaChiThuongTru,
                            line.ChuyenNhuong.ChuHo.HoTen,
                            line.ChuyenNhuong.ChuHo.NamSinh,
                            line.ChuyenNhuong.ChuHo.GioiTinh,
                            line.ChuyenNhuong.ChuHo.LoaiGiayTo,
                            line.ChuyenNhuong.ChuHo.SoGiayTo,
                            line.ChuyenNhuong.ChuHo.NgayCap,
                            line.ChuyenNhuong.ChuHo.NoiCap,
                            line.ChuyenNhuong.ChuHo.DiaChiThuongTru,
                            line.ChuyenNhuong.VoChong.HoTen,
                            line.ChuyenNhuong.VoChong.NamSinh,
                            line.ChuyenNhuong.VoChong.GioiTinh,
                            line.ChuyenNhuong.VoChong.LoaiGiayTo,
                            line.ChuyenNhuong.VoChong.SoGiayTo,
                            line.ChuyenNhuong.VoChong.NgayCap,
                            line.ChuyenNhuong.VoChong.NoiCap,
                            line.ChuyenNhuong.VoChong.DiaChiThuongTru,
                            line.ChuyenNhuong.HinhThuc,
                            line.ChuyenNhuong.HopDong,
                            line.GCN.Serial,
                            line.GCN.MaVach,
                            line.GCN.SoGCN,
                            line.GCN.NgayCap,
                            line.GCN.NgayChuyenNhuong
                        };

                    var newRow = new Row();
                    foreach (var cell in cells)
                    {
                        newRow.AppendChild(CreateCell(cell));
                    }

                    sheetData.AppendChild(newRow);
                }
            }
        }

        private Cell CreateCell(string value)
        {
            var cell = new Cell
            {
                DataType = CellValues.String,
                CellValue = new CellValue(value),
            };
            return cell;
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            var response = openFileDialog1.ShowDialog();
            if (response == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
                _fileName = openFileDialog1.SafeFileName.Replace(".docx", "");
            }
            var converter = new DocumentConverter();
            var result = converter.ExtractRawText(textBox1.Text);
            _tempdata = result.Value; // The raw text
                                      // var warnings = result.Warnings;
                                      //File.WriteAllText("b.txt", html);
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox1.Text))
            {
                return;
            }
            try
            {
                //string test = System.IO.Directory.GetCurrentDirectory();
                string newfile = Guid.NewGuid() + "tem.xls";
                string template = Path.GetFullPath("Template") + "\\Template.xls";
                string tempFile = Path.GetFullPath("Template\\") + newfile;
                File.Copy(template, tempFile, true);
                _xlApp = new Microsoft.Office.Interop.Excel.Application();
                _xlWorkbook = _xlApp.Workbooks.Open(tempFile);
                //Xu ly data
                Microsoft.Office.Interop.Excel._Worksheet sheet = _xlWorkbook.Worksheets[1];
                // sheet.Range
                // sheet.Rows
                _xlApp.DisplayAlerts = false;
                //SAVE
                _xlWorkbook.SaveAs(Path.GetFullPath("Output\\") + _fileName + ".xls");
                _xlApp.DisplayAlerts = true;
                File.Delete(tempFile);
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                if (_xlWorkbook != null)
                {
                    _xlWorkbook.Close();
                }
                if (_xlApp != null)
                {
                    _xlApp.Quit();
                }
            }
        }
    }

    internal class Data
    {
        public LandInfo ThongTinDat { get; set; } = new LandInfo();
        public Owner ChuSuDung { get; set; } = new Owner();
        public Transfer ChuyenNhuong { get; set; } = new Transfer();
        public GCN GCN { get; set; } = new GCN();
    }

    internal class LandInfo
    {
        public string SoTo { get; set; }
        public string SoThua { get; set; }
        public string DienTich { get; set; }
        public string MucDichSuDung { get; set; }
        public string DiaChiThuaDat { get; set; }
    }

    internal class Person
    {
        public string HoTen { get; set; }
        public string NamSinh { get; set; }
        public string GioiTinh { get; set; }
        public string LoaiGiayTo { get; set; }
        public string SoGiayTo { get; set; }
        public string NgayCap { get; set; }
        public string NoiCap { get; set; }
        public string DiaChiThuongTru { get; set; }
    }

    internal class Owner
    {
        public Person ChuHo { get; set; } = new Person();
        public Person VoChong { get; set; } = new Person();
    }

    internal class Transfer
    {
        public Person ChuHo { get; set; } = new Person();
        public Person VoChong { get; set; } = new Person();
        public string HinhThuc { get; set; }
        public string HopDong { get; set; }
    }

    internal class GCN
    {
        public string Serial { get; set; }
        public string MaVach { get; set; }
        public string SoGCN { get; set; }
        public string NgayCap { get; set; }
        public string NgayChuyenNhuong { get; set; }
    }
}