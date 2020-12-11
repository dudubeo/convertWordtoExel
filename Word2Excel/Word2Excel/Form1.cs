//using DocumentFormat.OpenXml.Spreadsheet;
using Aspose.Cells;
using Mammoth;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Word2Excel
{
    public partial class Form1 : Form
    {
        private string _fileName = string.Empty;

        public Form1()
        {
            InitializeComponent();
        }

        private static string RemoveVietNamese(string text)
        {
            var normalizedString = text.Normalize(NormalizationForm.FormD);
            var stringBuilder = new StringBuilder();

            foreach (var c in normalizedString)
            {
                var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != UnicodeCategory.NonSpacingMark)
                {
                    stringBuilder.Append(c);
                }
            }

            return stringBuilder.ToString().Normalize(NormalizationForm.FormC);
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
            var value = result.Value;
            if (value == null)
            {
                MessageBox.Show("File docx not data");
            }
            var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "temp.txt");
            File.WriteAllText(filePath, value);
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox1.Text))
            {
                return;
            }
            var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "temp.txt");
            string readText2 = File.ReadAllText(filePath);

            var query = File.ReadAllLines(filePath);
            var listWord = query.Where(s => s != string.Empty).ToList();
            //var listWord = query.Select(x => x.Trim()).Where(s => s != string.Empty).ToList();
            var listData = new List<DataModel>();
            var data = new DataModel();

            process_raw_data(listWord, listData, data);
            string newfile = Guid.NewGuid() + "tem.xls";
            string template = Path.GetFullPath("Template") + "\\Template.xls";
            string tempFile = Path.GetFullPath("Template\\") + newfile;
            Workbook book = new Workbook(template);
            try
            {
                //string test = System.IO.Directory.GetCurrentDirectory();
                //xu ly data
                // cell put value https://www.codeproject.com/Articles/5253765/Using-Excel-2019-Features-in-Aspose-Cells
                // add data Workbook https://docs.aspose.com/cells/net/different-ways-to-open-files/
                // set sytle https://www.csharpcodi.com/csharp-examples/Aspose.Cells.Cell.GetStyle()/
                book.Save(Path.GetFullPath("Output\\") + _fileName + ".xls", SaveFormat.Excel97To2003);
                MessageBox.Show("Done!");
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                book.Dispose();
            }
        }

        private void process_raw_data(List<string> listWord, List<DataModel> listData, DataModel data)
        {
            foreach (var list in listWord)
            {
                var item = RemoveVietNamese(list).ToUpper();

                if (item.StartsWith("ong:") || item.StartsWith("ho ong:") || item.StartsWith("chong:"))
                {
                    data.ChuSuDung_ChuHo_HoTen = list.Split(':').Last().Trim();
                }

                if (item.Contains("nam sinh:"))
                {
                    data.ChuSuDung_ChuHo_NamSinh = list.Split(':').Last().Trim();
                }

                if (item.Contains("so cmnd:"))
                {
                    data.ChuSuDung_ChuHo_LoaiGiayTo = "CMND";
                    data.ChuSuDung_ChuHo_SoGiayTo = list.Split(':').Last().Trim();
                }

                if (item.StartsWith("so") && item.Contains("cccd:"))
                {
                    data.ChuSuDung_ChuHo_LoaiGiayTo = "CCCD";
                    data.ChuSuDung_ChuHo_SoGiayTo = list.Split(':').Last().Trim();
                }
                if (item.StartsWith("dia chi thuong tru:") || item.StartsWith("dia chi thuong uu:"))
                {
                    data.ChuSuDung_ChuHo_DiaChiThuongTru = list.Split(':').Last().Trim();
                }

                if (item.StartsWith("ba:") || item.StartsWith("ho ba:") || item.StartsWith("vo:"))
                {
                    data.ChuSuDung_VoChong_HoTen = list.Split(':').Last().Trim();
                }

                if (item.StartsWith("nam sinh:") || item.StartsWith("sinh nam:"))
                {
                    data.ChuSuDung_VoChong_NamSinh = list.Split(':').Last().Trim();
                }

                if (item.Contains("so cmnd:") || item.Contains("socmnd:"))
                {
                    data.ChuSuDung_VoChong_LoaiGiayTo = "CMND";
                    data.ChuSuDung_VoChong_SoGiayTo = list.Split(':').Last().Trim();
                }

                if (item.StartsWith("so") && item.Contains("cccd:"))
                {
                    data.ChuSuDung_VoChong_LoaiGiayTo = "CCCD";
                    data.ChuSuDung_VoChong_SoGiayTo = list.Split(':').Last().Trim();
                }

                if (item.StartsWith("dia chi thuong tru:") || item.StartsWith("dia chi thuong uu:"))
                {
                    data.ChuSuDung_VoChong_DiaChiThuongTru = list.Split(':').Last().Trim();
                }

                if ((item.Contains("thua dat so:") || item.Contains("to ban do so:") || item.Contains("to ban do") || item.Contains("thuadatso:") || item.Contains("tobandoso:")))
                {
                    var regex = new Regex(@"([0-9])\w+");
                    var matches = regex.Matches(item);
                    data.ThongTinDat_SoThua = matches[0].Value.Trim();
                    data.ThongTinDat_SoTo = matches[1].Value.Trim();
                }

                if (item.Contains("dia chi:"))
                {
                    data.ThongTinDat_DiaChiThuaDat = list.Split(':').Last().Trim();
                }

                if (item.Contains("dien tich:"))
                {
                    data.ThongTinDat_DienTich = list.Split('(').First().Split(':').Last().Trim();
                }

                if (item.Contains("muc dich su dung:"))
                {
                    int index = list.IndexOf(':');
                    data.ThongTinDat_MucDichSuDung = list.Substring(index + 1);
                }

                // 0
                if (item.Contains("giay chung nhan so") || item.Contains("giay chung nhon so"))
                {
                    var regex = new Regex(@"[(0-9)]+\/[(0-9)]+\/[(0-9)]+");
                    var matches = regex.Matches(list);

                    if (regex.IsMatch(list))
                    {
                        data.ChuSuDung_ChuHo_NgayCap = matches[0].Value.Trim();
                    }
                    else
                    {
                        data.ChuSuDung_ChuHo_NgayCap = item.Split('.').First().Split(' ').Last();
                    }
                }
                // 0
                if (item.Contains("chuyen nhuong cho"))
                {
                    if (item.Contains("chuyen nhuong cho ong:"))
                    {
                        data.ChuyenNhuong_ChuHo_GioiTinh = "Nam";
                        data.ChuyenNhuong_ChuHo_HoTen = list.Split(':').Last().Trim();
                    }
                    if (item.Contains("chuyen nhuong cho ba"))
                    {
                        data.ChuyenNhuong_ChuHo_GioiTinh = "Nữ";
                    }
                }
                if (item.Contains("dia chi:"))
                {
                    data.ChuyenNhuong_ChuHo_DiaChiThuongTru = list.Split(':').Last().Trim();
                }

                if (item.Contains("cmnd:"))
                {
                    data.ChuyenNhuong_ChuHo_LoaiGiayTo = "CMND";
                    var temp = item.Replace(" ", string.Empty);
                    int index = temp.IndexOf(':');
                    data.ChuyenNhuong_ChuHo_SoGiayTo = temp.Substring(index + 1, 9);
                }

                if (item.Contains("can cuoc cong dan"))
                {
                    data.ChuyenNhuong_ChuHo_LoaiGiayTo = "CCCD";
                    var temp = item.Replace(" ", string.Empty);
                    int index = temp.IndexOf(':');
                    data.ChuyenNhuong_ChuHo_SoGiayTo = temp.Substring(index + 1, 12);
                }

                if (item.Contains("theo ho so so"))
                {
                    data.GCN_SoGCN = list.Split(':').Last().Trim();
                }

                //listWord.Remove(item);
                listData.Add(data);
            }
        }

        internal class DataModel
        {
            public string ThongTinDat_SoTo { get; set; }
            public string ThongTinDat_SoThua { get; set; }
            public string ThongTinDat_DienTich { get; set; }
            public string ThongTinDat_MucDichSuDung { get; set; }
            public string ThongTinDat_DiaChiThuaDat { get; set; }
            public string ChuSuDung_ChuHo_HoTen { get; set; }
            public string ChuSuDung_ChuHo_NamSinh { get; set; }
            public string ChuSuDung_ChuHo_GioiTinh { get; set; }
            public string ChuSuDung_ChuHo_LoaiGiayTo { get; set; }
            public string ChuSuDung_ChuHo_SoGiayTo { get; set; }
            public string ChuSuDung_ChuHo_NgayCap { get; set; }
            public string ChuSuDung_ChuHo_NoiCap { get; set; }
            public string ChuSuDung_ChuHo_DiaChiThuongTru { get; set; }
            public string ChuSuDung_VoChong_HoTen { get; set; }
            public string ChuSuDung_VoChong_NamSinh { get; set; }
            public string ChuSuDung_VoChong_GioiTinh { get; set; }
            public string ChuSuDung_VoChong_LoaiGiayTo { get; set; }
            public string ChuSuDung_VoChong_SoGiayTo { get; set; }
            public string ChuSuDung_VoChong_NgayCap { get; set; }
            public string ChuSuDung_VoChong_NoiCap { get; set; }
            public string ChuSuDung_VoChong_DiaChiThuongTru { get; set; }
            public string ChuyenNhuong_ChuHo_HoTen { get; set; }
            public string ChuyenNhuong_ChuHo_NamSinh { get; set; }
            public string ChuyenNhuong_ChuHo_GioiTinh { get; set; }
            public string ChuyenNhuong_ChuHo_LoaiGiayTo { get; set; }
            public string ChuyenNhuong_ChuHo_SoGiayTo { get; set; }
            public string ChuyenNhuong_ChuHo_NgayCap { get; set; }
            public string ChuyenNhuong_ChuHo_NoiCap { get; set; }
            public string ChuyenNhuong_ChuHo_DiaChiThuongTru { get; set; }
            public string ChuyenNhuong_VoChong_HoTen { get; set; }
            public string ChuyenNhuong_VoChong_NamSinh { get; set; }
            public string ChuyenNhuong_VoChong_GioiTinh { get; set; }
            public string ChuyenNhuong_VoChong_LoaiGiayTo { get; set; }
            public string ChuyenNhuong_VoChong_SoGiayTo { get; set; }
            public string ChuyenNhuong_VoChong_NgayCap { get; set; }
            public string ChuyenNhuong_VoChong_NoiCap { get; set; }
            public string ChuyenNhuong_VoChong_DiaChiThuongTru { get; set; }
            public string ChuyenNhuong_HinhThuc { get; set; }
            public string ChuyenNhuong_HopDong { get; set; }
            public string GCN_Serial { get; set; }
            public string GCN_MaVach { get; set; }
            public string GCN_SoGCN { get; set; }
            public string GCN_NgayCap { get; set; }
            public string GCN_NgayChuyenNhuong { get; set; }
        }
    }
}