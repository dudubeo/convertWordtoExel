//using DocumentFormat.OpenXml.Spreadsheet;
using Aspose.Cells;
using Mammoth;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
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
            //string readText2 = File.ReadAllText(filePath);
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
                Worksheet ws = book.Worksheets[0];
                int row = 3;
                foreach (var item in listData)
                {
                    IList<PropertyInfo> properties = new List<PropertyInfo>(item.GetType().GetProperties())
                            .Where(prop => prop.IsDefined(typeof(ExcelMapperAttribute), false))
                            .ToList();
                    if (properties == null || properties.Count < 1)
                    {
                        break;
                    }
                    foreach (PropertyInfo property in properties)
                    {
                        ExcelMapperAttribute mapperColumn = property.GetCustomAttribute<ExcelMapperAttribute>();

                        if (!string.IsNullOrWhiteSpace(mapperColumn.ColumnName) && property.GetValue(item, null) != null)
                        {
                            ws.Cells[string.Format("{0}{1}", mapperColumn.ColumnName, row)].Value = property.GetValue(item, null);
                        }
                    }
                    row++;
                }
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
            [ExcelMapper(columname: "A")]
            public string ThongTinDat_SoTo { get; set; }

            [ExcelMapper(columname: "B")]
            public string ThongTinDat_SoThua { get; set; }

            [ExcelMapper(columname: "C")]
            public string ThongTinDat_DienTich { get; set; }

            [ExcelMapper(columname: "D")]
            public string ThongTinDat_MucDichSuDung { get; set; }

            [ExcelMapper(columname: "E")]
            public string ThongTinDat_DiaChiThuaDat { get; set; }

            [ExcelMapper(columname: "F")]
            public string ChuSuDung_ChuHo_HoTen { get; set; }

            [ExcelMapper(columname: "G")]
            public string ChuSuDung_ChuHo_NamSinh { get; set; }

            [ExcelMapper(columname: "H")]
            public string ChuSuDung_ChuHo_GioiTinh { get; set; }

            [ExcelMapper(columname: "I")]
            public string ChuSuDung_ChuHo_LoaiGiayTo { get; set; }

            [ExcelMapper(columname: "J")]
            public string ChuSuDung_ChuHo_SoGiayTo { get; set; }

            [ExcelMapper(columname: "K")]
            public string ChuSuDung_ChuHo_NgayCap { get; set; }

            [ExcelMapper(columname: "L")]
            public string ChuSuDung_ChuHo_NoiCap { get; set; }

            [ExcelMapper(columname: "M")]
            public string ChuSuDung_ChuHo_DiaChiThuongTru { get; set; }

            [ExcelMapper(columname: "N")]
            public string ChuSuDung_VoChong_HoTen { get; set; }

            [ExcelMapper(columname: "O")]
            public string ChuSuDung_VoChong_NamSinh { get; set; }

            [ExcelMapper(columname: "P")]
            public string ChuSuDung_VoChong_GioiTinh { get; set; }

            [ExcelMapper(columname: "Q")]
            public string ChuSuDung_VoChong_LoaiGiayTo { get; set; }

            [ExcelMapper(columname: "R")]
            public string ChuSuDung_VoChong_SoGiayTo { get; set; }

            [ExcelMapper(columname: "S")]
            public string ChuSuDung_VoChong_NgayCap { get; set; }

            [ExcelMapper(columname: "T")]
            public string ChuSuDung_VoChong_NoiCap { get; set; }

            [ExcelMapper(columname: "U")]
            public string ChuSuDung_VoChong_DiaChiThuongTru { get; set; }

            [ExcelMapper(columname: "V")]
            public string ChuyenNhuong_ChuHo_HoTen { get; set; }

            [ExcelMapper(columname: "W")]
            public string ChuyenNhuong_ChuHo_NamSinh { get; set; }

            [ExcelMapper(columname: "X")]
            public string ChuyenNhuong_ChuHo_GioiTinh { get; set; }

            [ExcelMapper(columname: "Y")]
            public string ChuyenNhuong_ChuHo_LoaiGiayTo { get; set; }

            [ExcelMapper(columname: "Z")]
            public string ChuyenNhuong_ChuHo_SoGiayTo { get; set; }

            [ExcelMapper(columname: "AA")]
            public string ChuyenNhuong_ChuHo_NgayCap { get; set; }

            [ExcelMapper(columname: "AB")]
            public string ChuyenNhuong_ChuHo_NoiCap { get; set; }

            [ExcelMapper(columname: "AC")]
            public string ChuyenNhuong_ChuHo_DiaChiThuongTru { get; set; }

            [ExcelMapper(columname: "AD")]
            public string ChuyenNhuong_VoChong_HoTen { get; set; }

            [ExcelMapper(columname: "AE")]
            public string ChuyenNhuong_VoChong_NamSinh { get; set; }

            [ExcelMapper(columname: "AF")]
            public string ChuyenNhuong_VoChong_GioiTinh { get; set; }

            [ExcelMapper(columname: "AG")]
            public string ChuyenNhuong_VoChong_LoaiGiayTo { get; set; }

            [ExcelMapper(columname: "AH")]
            public string ChuyenNhuong_VoChong_SoGiayTo { get; set; }

            [ExcelMapper(columname: "AI")]
            public string ChuyenNhuong_VoChong_NgayCap { get; set; }

            [ExcelMapper(columname: "AJ")]
            public string ChuyenNhuong_VoChong_NoiCap { get; set; }

            [ExcelMapper(columname: "AK")]
            public string ChuyenNhuong_VoChong_DiaChiThuongTru { get; set; }

            [ExcelMapper(columname: "AL")]
            public string ChuyenNhuong_HinhThuc { get; set; }

            [ExcelMapper(columname: "AM")]
            public string ChuyenNhuong_HopDong { get; set; }

            [ExcelMapper(columname: "AN")]
            public string GCN_Serial { get; set; }

            [ExcelMapper(columname: "AO")]
            public string GCN_MaVach { get; set; }

            [ExcelMapper(columname: "AP")]
            public string GCN_SoGCN { get; set; }

            [ExcelMapper(columname: "AQ")]
            public string GCN_NgayCap { get; set; }

            [ExcelMapper(columname: "AR")]
            public string GCN_NgayChuyenNhuong { get; set; }
        }

        public class ExcelMapperAttribute : Attribute
        {
            public ExcelMapperAttribute()
            {
            }

            public ExcelMapperAttribute(string columname)
            {
                ColumnName = columname;
            }

            public string ColumnName { get; private set; }
        }
    }
}