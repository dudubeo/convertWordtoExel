using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.IO.Packaging;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Newtonsoft.Json;
using System.Collections;
using DocumentFormat.OpenXml.Spreadsheet;
using Mammoth;

namespace Word2Excel
{
    public partial class Form1 : Form
    {

        public Hashtable vietnamese = new Hashtable()
        {
            { "a", "á|à|ả|ã|ạ|ă|ắ|ặ|ằ|ẳ|ẵ|â|ấ|ầ|ẩ|ẫ|ậ|Á|À|Ả|Ã|Ạ|Ă|Ắ|Ặ|Ằ|Ẳ|Ẵ|Â|Ấ|Ầ|Ẩ|Ẫ|Ậ" },
            { "d", "đ|Đ" },
            { "e", "é|è|ẻ|ẽ|ẹ|ê|ế|ề|ể|ễ|ệ|É|È|Ẻ|Ẽ|Ẹ|Ê|Ế|Ề|Ể|Ễ" },
            { "i", "í|ì|ỉ|ĩ|ị|Í|Ì|Ỉ|Ĩ|Ị" },
            { "o", "ó|ò|ỏ|õ|ọ|ô|ố|ồ|ổ|ỗ|ộ|ơ|ớ|ờ|ở|ỡ|ợ|Ó|Ò|Ỏ|Õ|Ọ|Ô|Ố|Ồ|Ổ|Ỗ|Ộ|Ơ|Ớ|Ờ|Ở|Ỡ|Ợ" },
            { "u", "ú|ù|ủ|ũ|ụ|ư|ứ|ừ|ử|ữ|ự|Ú|Ù|Ủ|Ũ|Ụ|Ư|Ứ|Ừ|Ử|Ữ|Ự" },
            { "y", "ý|ỳ|ỷ|ỹ|ỵ|Ý|Ỳ|Ỷ|Ỹ|Ỵ" }
        };


        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var response = openFileDialog1.ShowDialog();
            if (response == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
            var converter = new DocumentConverter();
            var result = converter.ExtractRawText(textBox1.Text);
            var html = result.Value; // The raw text
            var warnings = result.Warnings;
            File.WriteAllText("b.txt", html);



        }

        private void button2_Click(object sender, EventArgs e)
        {
            var word = new Microsoft.Office.Interop.Word.Application();
            if (string.IsNullOrEmpty(textBox1.Text))
            {
                return;
            }

            var reader = new List<string>();

            Stream stream = File.Open(textBox1.Text, FileMode.Open);
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(stream, false))
            {
                Body body = wordDocument.MainDocumentPart.Document.Body;
                var elements = body.ChildElements;

                reader.Clear();
                progressBar1.Maximum = elements.Count;

                for (int i = 0; i < elements.Count; i++)
                {
                    var element = elements.GetItem(i);
                    if (element is DocumentFormat.OpenXml.Wordprocessing.Table)
                    {
                        try
                        {
                            var d = VietnameseDecode(element.InnerText);
                            if (d.Replace(" ", "").Replace(":", "").Contains("mucdichsudungthoihansudung"))
                            {
                                foreach (var row in ((DocumentFormat.OpenXml.Wordprocessing.Table)element).ChildElements)
                                {
                                    var vi = VietnameseDecode(row.InnerText);
                                    if (vi.Replace(" ", "").Contains("congnhanquyen"))
                                    {
                                        var cells = ((TableRow)row).ChildElements;
                                        reader.Add("thua dat so: " + ((TableCell)cells[2]).InnerText + " to so: " + ((TableCell)cells[3]).InnerText);
                                        reader.Add("dien tich: " + ((TableCell)cells[4]).InnerText + " m2");
                                        reader.Add("muc dich su dung: " + ((TableCell)cells[8]).InnerText);
                                    }
                                }
                            }
                        }
                        catch { }
                    }

                    var line = elements.GetItem(i).InnerText;
                    line = line.Trim().Replace("a)", "^")
                        .Replace("b)", "^")
                        .Replace("c)", "^")
                        .Replace("d)", "^")
                        .Replace("đ)", "^")
                        .Replace("e)", "^")
                        .Replace("g)", "^")
                        .Replace("'", "")
                        .Replace("Địa chỉ", ";Địa chỉ")
                        .Replace("VỢ", ";Vợ")
                        .Replace("Vợ", ";Vợ");
                    var decoded = VietnameseDecode(line);
                    if (Regex.Matches(decoded, "\\^").Count > 1)
                    {
                        var split = line.Split('^');
                        reader.AddRange(split);
                    }
                    else if (line.Contains(";"))
                    {
                        reader.AddRange(line.Split(';'));
                    }
                    else
                    {
                        reader.Add(line);
                    }

                    progressBar1.Value++;
                }

                wordDocument.Close();
            }


            var output = new List<Data>();
            var data = new Data();
            var ong = "";
            var ba = "";
            var ongTruoc = false;
            var namsinh = new List<string>();
            var cmnd = new List<string>();
            var diachi = new List<string>();
            var loaigiayto = new List<string>();
            var noicap = new List<string>();
            var ngaycap = new List<string>();
            var sothua = "";
            var soto = "";
            var diachidat = "";
            var mucdich = "";
            var dientich = "";
            var sohopdong = "";
            var serial = "";
            var mavach = "";
            var sogcn = "";
            var ngaycapso = "";
            var ngaychuyennhuong = "";
            var dirty = false;

            for (int i = 0; i < reader.Count; i++)
            {
                var line = reader[i];
                var decoded = VietnameseDecode(line);
                if (string.IsNullOrEmpty(decoded)) continue;

                if (decoded.Contains("nguoi su dung dat"))
                {
                    if (ongTruoc)
                    {
                        data.ChuSuDung.ChuHo.HoTen = ong;
                        data.ChuSuDung.ChuHo.GioiTinh = "1";

                        if (!string.IsNullOrEmpty(ba))
                        {
                            data.ChuSuDung.VoChong.HoTen = ba;
                            data.ChuSuDung.VoChong.GioiTinh = "0";
                        }
                    }
                    else
                    {
                        data.ChuSuDung.ChuHo.HoTen = ba;
                        data.ChuSuDung.ChuHo.GioiTinh = "0";

                        if (!string.IsNullOrEmpty(ba))
                        {
                            data.ChuSuDung.VoChong.HoTen = ong;
                            data.ChuSuDung.VoChong.GioiTinh = "1";
                        }
                    }

                    if (loaigiayto.Count > 0)
                    {
                        data.ChuSuDung.ChuHo.LoaiGiayTo = loaigiayto.First();

                        if (loaigiayto.Count > 1)
                        {
                            data.ChuSuDung.VoChong.LoaiGiayTo = loaigiayto.Last();
                        }
                    }

                    if (namsinh.Count > 0)
                    {
                        data.ChuSuDung.ChuHo.NamSinh = namsinh.First();
                        if (namsinh.Count > 1)
                        {
                            data.ChuSuDung.VoChong.NamSinh = namsinh.Last();
                        }
                    }

                    if (cmnd.Count > 0)
                    {
                        data.ChuSuDung.ChuHo.SoGiayTo = cmnd.First();
                        if (cmnd.Count > 1)
                        {
                            data.ChuSuDung.VoChong.SoGiayTo = cmnd.Last();
                        }
                    }

                    if (diachi.Count > 0)
                    {
                        data.ChuSuDung.ChuHo.DiaChiThuongTru = diachi.First();
                        if (diachi.Count > 1)
                        {
                            data.ChuSuDung.VoChong.DiaChiThuongTru = diachi.Last();
                        }
                    }

                    if (noicap.Count > 0)
                    {
                        data.ChuSuDung.ChuHo.NoiCap = noicap.First();

                        if (noicap.Count > 1)
                        {
                            data.ChuSuDung.VoChong.NoiCap = noicap.Last();
                        }
                    }

                    if (ngaycap.Count > 0)
                    {
                        data.ChuSuDung.ChuHo.NgayCap = ngaycap.First();

                        if (ngaycap.Count > 1)
                        {
                            data.ChuSuDung.VoChong.NgayCap = ngaycap.Last();
                        }
                    }

                    data.ThongTinDat.SoTo = soto;
                    data.ThongTinDat.SoThua = sothua;
                    data.ThongTinDat.DienTich = dientich;
                    data.ThongTinDat.MucDichSuDung = mucdich;
                    data.ThongTinDat.DiaChiThuaDat = diachidat;

                    data.GCN.Serial = serial;
                    data.GCN.NgayCap = ngaycapso;

                    if (dirty) output.Add(data);
                    data = new Data();
                    ong = "";
                    ba = "";
                    ongTruoc = false;
                    namsinh.Clear();
                    cmnd.Clear();
                    diachi.Clear();
                    loaigiayto.Clear();
                    noicap.Clear();
                    ngaycap.Clear();
                    sothua = "";
                    soto = "";
                    diachidat = "";
                    dientich = "";
                    mucdich = "";
                    sohopdong = "";
                    serial = "";
                    mavach = "";
                    sogcn = "";
                    ngaycapso = "";
                    ngaychuyennhuong = "";
                }

                if (decoded.StartsWith("ong:") || decoded.StartsWith("ho ong:") || decoded.StartsWith("chong:"))
                {
                    dirty = true;
                    ong = line.Split(':').Last().Trim();
                    if (string.IsNullOrEmpty(ba)) ongTruoc = true;

                    if (decoded.Contains("sinh nam:"))
                    {
                        ong = line.Split(';').First().Split(':').Last();
                        namsinh.Add(line.Split(';').Last().Split(':').Last());
                    }
                }

                if (decoded.StartsWith("ba:") || decoded.StartsWith("ho ba:") || decoded.StartsWith("vo:"))
                {
                    dirty = true;
                    ba = line.Split(':').Last().Trim();
                    if (string.IsNullOrEmpty(ong)) ongTruoc = false;

                    if (decoded.Contains("sinh nam:"))
                    {
                        ba = line.Split(';').First().Split(':').Last();
                        namsinh.Add(line.Split(';').Last().Split(':').Last());
                    }
                }

                if (decoded.StartsWith("nam sinh:") || decoded.StartsWith("sinh nam:"))
                {
                    dirty = true;
                    namsinh.Add(line.Split(':').Last().Trim());
                }

                if ((decoded.StartsWith("so") && decoded.Contains("cmnd:")) || decoded.StartsWith("cmnd so:"))
                {
                    dirty = true;
                    loaigiayto.Add("CMND");

                    if (decoded.Contains("cap ngay"))
                    {
                        try
                        {
                            var regex = new Regex(@"[0-9]\w+");
                            var matches = regex.Matches(decoded);
                            cmnd.Add(matches[0].Value.Trim());
                        }
                        catch
                        {
                            cmnd.Add(line.Split(':').Last().Trim());
                        }
                    }
                    else
                    {
                        cmnd.Add(line.Split(':').Last().Trim());
                    }
                }

                if (decoded.StartsWith("so") && decoded.Contains("cccd:"))
                {
                    dirty = true;
                    loaigiayto.Add("CCCD");
                    cmnd.Add(line.Split(':').Last().Trim());
                }

                if (decoded.Contains(", do") && decoded.Contains("cap ngay"))
                {
                    try
                    {
                        var splits = line.Split(',');
                        cmnd.Add(splits[0].Split(':').Last().Trim());
                        noicap.Add(splits[1].Replace("do ", "").Trim());
                        ngaycap.Add(splits[2].Split(':').Last().Trim());
                    }
                    catch { }
                }

                if (decoded.StartsWith("dia chi thuong tru:") || decoded.StartsWith("dia chi thuong uu:"))
                {
                    dirty = true;
                    diachi.Add(line.Split(':').Last().Trim());
                }

                if ((decoded.Contains("thua dat so:") || decoded.Contains("to ban do so:") || decoded.Contains("to ban do")
                    || decoded.Contains("thuadatso:") || decoded.Contains("tobandoso:"))
                    && string.IsNullOrEmpty(sothua))
                {
                    dirty = true;
                    try
                    {
                        var regex = new Regex(@"([0-9])\w+");
                        var matches = regex.Matches(decoded);
                        sothua = matches[0].Value.Trim();
                        soto = matches[1].Value.Trim();
                    }
                    catch
                    {
                        sothua = decoded.Split(',').First().Split(':').Last().Trim();
                        soto = decoded.Split(':').Last().Trim();
                    }
                }

                if (decoded.Contains("dia chi:") && string.IsNullOrEmpty(diachidat))
                {
                    dirty = true;
                    try
                    {
                        diachidat = line.Split(':')[1].Split('.').First().Trim();
                    }
                    catch
                    {
                        diachidat = line.Split(':').Last().Trim();
                    }
                }

                if (decoded.Contains("dien tich:"))
                {
                    dirty = true;
                    dientich = line.Split('(').First().Split(':').Last().Trim();
                }

                if (decoded.Contains("muc dich su dung:"))
                {
                    dirty = true;
                    var a = line.Replace("ng:", "^");
                    mucdich = a.Split('^').Last().Trim();
                }

                if (decoded.Contains("giay chung nhan so") || decoded.Contains("giay chung nhon so"))
                {
                    dirty = true;
                    try
                    {
                        var pattern = @"[(0-9)]+\/[(0-9)]+\/[(0-9)]+";
                        var regex = new Regex(pattern);
                        var matches = regex.Matches(decoded);
                        ngaycapso = matches[0].Value.Trim();
                    }
                    catch
                    {
                        ngaycapso = decoded.Split('.').First().Split(' ').Last();
                    }
                }
            }

            File.WriteAllLines(@"C:\Users\gh057\Desktop\output.txt", reader);
            progressBar1.Value = 0;

            var save = saveFileDialog1.ShowDialog();
            if (save == DialogResult.OK)
            {
                //File.WriteAllText(saveFileDialog1.FileName, JsonConvert.SerializeObject(output), Encoding.UTF8);
                var vi = new List<string>();
                reader.ForEach(item => vi.Add(VietnameseDecode(item)));
                File.WriteAllLines(@"C:\Users\gh057\Desktop\output.txt", vi);
                ExportExcel(saveFileDialog1.FileName, output);
                MessageBox.Show("Done!");
                progressBar1.Value = 0;
            }
        }

        private string VietnameseDecode(string input)
        {
            string output = input;
            foreach (DictionaryEntry entry in vietnamese)
            {
                foreach (string key in entry.Value.ToString().Split('|'))
                {
                    output = output.Replace(key, entry.Key.ToString());
                }
            }
            return output.ToLower().Trim();
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
    }

    class Data
    {
        public LandInfo ThongTinDat { get; set; } = new LandInfo();
        public Owner ChuSuDung { get; set; } = new Owner();
        public Transfer ChuyenNhuong { get; set; } = new Transfer();
        public GCN GCN { get; set; } = new GCN();
    }

    class LandInfo
    {
        public string SoTo { get; set; }
        public string SoThua { get; set; }
        public string DienTich { get; set; }
        public string MucDichSuDung { get; set; }
        public string DiaChiThuaDat { get; set; }
    }

    class Person
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

    class Owner
    {
        public Person ChuHo { get; set; } = new Person();
        public Person VoChong { get; set; } = new Person();
    }

    class Transfer
    {
        public Person ChuHo { get; set; } = new Person();
        public Person VoChong { get; set; } = new Person();
        public string HinhThuc { get; set; }
        public string HopDong { get; set; }
    }

    class GCN
    {
        public string Serial { get; set; }
        public string MaVach { get; set; }
        public string SoGCN { get; set; }
        public string NgayCap { get; set; }
        public string NgayChuyenNhuong { get; set; }
    }
}
