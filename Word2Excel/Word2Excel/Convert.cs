//using DocumentFormat.OpenXml.Spreadsheet;
using Aspose.Cells;
using Mammoth;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace Word2Excel
{
    public partial class Convert : Form
    {
        private const string _p_sothua = @"(?<sothua>\d{2,4})\D{1,5}T. b.n ..\D*(?<soto>\d{2,3})";
        private const string _p_ThongTinDat_DienTich = @".i.n t.ch:(?<dientich>[\d\.]*).*";
        private const string _p_ThongTinDat_MucDichSuDung = @"M.c ..ch s. d.ng[^:\n]*:(?<mucdichsudungdat>.*)\n";
        private const string _p_ThongTinDat_DiaChiThuaDat = @"..a ch.:(?<diachi>.*)\n";
        private const string _p_DiachiThuongtru = @"\n..a ch. th..ng tr.:(?<diachithuogntru>.*)\n";
        private const string _p_ChuSuDung_ChuHo_HoTen = @"\n(|...)(.ng|B.|Ch.ng|V.): (?<hoten>[^:;,.\n\d]{8,})";
        private const string _p_ChuSuDung_ChuHo_NamSinh = @"(N.m sinh|Sinh n.m): (?<namsinh>\d{4})(\n|;|,)";
        private const string _p_ChuSuDung_ChuHo_LoaiGiayTo = @"(CMND|CCCD)\D*(?<socmnd>\d{9})(\D*(?<noicap>C.ng An \D*), c.p ng.y: (?<ngaycap>[\d\/]*)|)";

        //private const string _p_ChuSuDung_ChuHo_DiaChiThuongTru = @"..a ch. th..ng tr.:(?<diachithuogntru>.*)\n";
        //private const string _p_ChuSuDung_VoChong_HoTen = @"^(.ng|B.|Ch.ng|V.): (?<hoten>.*)$";
        //private const string _p_ChuSuDung_VoChong_NamSinh = @"^(N.m sinh|Sinh n.m): (?<namsinh>\d{4})($|;|,)";
        private const string _p_tachtrang = @"(C.NG H.A X. H.I CH. NGH.A VI.T NAM|..c l.p - T. do - H.nh ph.c|\nGI.Y CH.NG NH.N\n|\nQUY.N S. D.NG ..T\n)";

        private const string _p_GCN_Serial = @"\n(?<seria>\S{2} \d{6})";
        private const string _p_GCN_MaVach = @"(?<mavach>[\d \.]{15,25})";
        private const string _p_chimuc = @":.\)";
        private const string _p_bo_dientich_trongdiachi = @".i.n t.ch:";
        private const string _p_tach_chuyennhuong = @"\nN.i dung thay ..i v. c. s. ph.p l.\n";
        private const string _p_chuyennhuong = @"(T.ng cho|Chuy.n nh..ng cho)";
        private const string _p_chuyennhuong_hoten = @"(\n| )(.ng|b.)\W{1,2}(?<hoten>[^:;,.\n\d]{8,})";
        private const string _p_chuyennhuong_cmnd = @"\D(?<socmnd>\d{9,12})\D";
        private const string _p_chuyennhuong_loaigiayto = @"(CMND|CCCD)";
        private const string _p_chuyennhuong_diachi = @"\n..a ch.: (?<daichi>[^\n\.\:]*)";
        private const string _p_chuyennhuong_hopdong = @"Theo h. s. s.: (?<sohopdong>.*)";
        private const string _p_chuyennhuong_tangcho = @"(T.ng cho)";
        private const string _p_chuyennhuong_cn = @"(Chuy.n nh..ng cho)";

        private List<string> _temdata;

        private string _fileName = string.Empty;

        public Convert()
        {
            InitializeComponent();
        }

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

        private string RemoveVietNamese(string text)
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

            return VietnameseDecode(stringBuilder.ToString()).Normalize(NormalizationForm.FormC);
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            var response = openFileDialog1.ShowDialog();
            if (response == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
                _fileName = openFileDialog1.SafeFileName.Replace(".docx", "");
            }
            DateTime date1 = new DateTime(2020, 12, 13, 0, 0, 0);

            if (DateTime.Now.Year == date1.Year)
            {
                if (DateTime.Now.Month == date1.Month)
                {
                    if (DateTime.Now.Day == date1.Day || DateTime.Now.Day == (int)(date1.Day + 1))
                    {
                        var converter = new DocumentConverter();
                        var result = converter.ExtractRawText(textBox1.Text);
                        var value = result.Value;
                        _temdata = new List<string>();
                        _temdata = (Regex.Split(value, _p_tachtrang)).Where(x => x.Length > 500).ToList();

                        if (value == null)
                        {
                            MessageBox.Show("File docx not data");
                        }
                        var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "temp.txt");
                        //File.WriteAllText(filePath, value);
                    }
                    else
                    {
                        MessageBox.Show("Hết hạn dùng thử!");
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("Hết hạn dùng thử!");
                    return;
                }
            }
            else
            {
                MessageBox.Show("Hết hạn dùng thử!");
                return;
            }
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox1.Text))
            {
                MessageBox.Show("Vui lòng chọn file trước để thực hiện !");
                return;
            }
            // var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "temp.txt");
            //string readText2 = File.ReadAllText(filePath);
            // var query = File.ReadAllLines(filePath);
            List<DataModel> listData = new List<DataModel>();

            if (_temdata.Count > 0)
            {
                int a = 1;
                int first = 0;
                //Bien de danh dau chay het text khong se loi mat 1 dong du lieu
                int index = 0;
                foreach (var tempLine in _temdata)
                {
                    DataModel data = new DataModel();
                    index = index + 1;
                    MatchCollection value;

                    string line = Regex.Replace(tempLine, @":\n\n", ":").Replace("'", string.Empty);
                    if (index == 140)
                    {
                        int b = 1;
                    }
                    //Thông tin về đất
                    if (IsMatch_Pattern(line, out value, _p_sothua))
                    {
                        var regex = new Regex(@"([0-9])\w+");
                        var matches = regex.Matches(value[0].Value);
                        // data.ThongTinDat_SoThua = string.IsNullOrEmpty(matches[0].Value) ? null : matches[0].Value.Trim();
                        data.ThongTinDat_SoThua = value[0].Groups["sothua"].Value;
                        data.ThongTinDat_SoTo = value[0].Groups["soto"].Value;
                    }

                    if (IsMatch_Pattern(line, out value, _p_ThongTinDat_DienTich))
                    {
                        //data.ThongTinDat_DienTich = value[0].Value.Substring(value[0].Value.IndexOf(':') + 1).Trim();
                        data.ThongTinDat_DienTich = value[0].Groups["dientich"].Value;

                        //if (!string.IsNullOrEmpty(value[0].Groups["dientich"].Value) && IsMatch_Pattern(value[0].Groups["dientich"].Value, out value, _p_chimuc))
                        //{
                        //    data.ThongTinDat_DienTich = string.Empty;
                        //}
                    }
                    if (IsMatch_Pattern(line, out value, _p_ThongTinDat_MucDichSuDung))
                    {
                        //data.ThongTinDat_MucDichSuDung = value[0].Value.Substring(value[0].Value.IndexOf(':') + 1).Trim();
                        data.ThongTinDat_MucDichSuDung = value[0].Groups["mucdichsudungdat"].Value;
                        if (!string.IsNullOrEmpty(value[0].Groups["mucdichsudungdat"].Value) && IsMatch_Pattern(value[0].Groups["mucdichsudungdat"].Value, out value, _p_chimuc))
                        {
                            data.ThongTinDat_MucDichSuDung = string.Empty;
                        }
                    }
                    if (IsMatch_Pattern(line, out value, _p_ThongTinDat_DiaChiThuaDat))
                    {
                        //data.ThongTinDat_DiaChiThuaDat = value[0].Value.Substring(value[0].Value.IndexOf(':') + 1).Trim();
                        data.ThongTinDat_DiaChiThuaDat = value[0].Groups["diachi"].Value;
                        if (!string.IsNullOrEmpty(value[0].Groups["diachi"].Value) && IsMatch_Pattern(value[0].Groups["diachi"].Value, out value, _p_bo_dientich_trongdiachi))
                        {
                            data.ThongTinDat_DiaChiThuaDat = string.Empty;
                        }
                    }
                    if (IsMatch_Pattern(line, out value, _p_DiachiThuongtru))
                    {
                        data.ChuSuDung_ChuHo_DiaChiThuongTru = value[0].Groups["diachithuogntru"].Value;

                        if (value.Count > 1)
                        {
                            data.ChuSuDung_VoChong_DiaChiThuongTru = value[1].Groups["diachithuogntru"].Value;
                        }
                    }
                    if (IsMatch_Pattern(line, out value, _p_ChuSuDung_ChuHo_HoTen))
                    {
                        // data.ChuSuDung_ChuHo_HoTen = value[0].Value.Substring(value[0].Value.IndexOf(':') + 1).Trim();
                        data.ChuSuDung_ChuHo_HoTen = value[0].Groups["hoten"].Value;
                        if (VietnameseDecode(value[0].Value).Contains("ong"))
                        {
                            data.ChuSuDung_ChuHo_GioiTinh = "1";
                        }
                        else
                        {
                            if (VietnameseDecode(value[0].Value).Contains("ba"))
                            {
                                data.ChuSuDung_ChuHo_GioiTinh = "0";
                            }
                        }
                        if (value.Count > 1)
                        {
                            data.ChuSuDung_VoChong_HoTen = value[1].Groups["hoten"].Value;
                            if (VietnameseDecode(value[1].Value).Contains("ong"))
                            {
                                data.ChuSuDung_VoChong_GioiTinh = "1";
                            }
                            else
                            {
                                if (VietnameseDecode(value[1].Value).Contains("ba"))
                                {
                                    data.ChuSuDung_VoChong_GioiTinh = "0";
                                }
                            }
                        }
                    }
                    if (IsMatch_Pattern(line, out value, _p_ChuSuDung_ChuHo_NamSinh))
                    {
                        data.ChuSuDung_ChuHo_NamSinh = value[0].Groups["namsinh"].Value;

                        if (value.Count > 1)
                        {
                            // data.ChuSuDung_VoChong_NamSinh = value[1].Value.Substring(value[1].Value.IndexOf(':') + 1).Trim();
                            data.ChuSuDung_VoChong_NamSinh = value[1].Groups["namsinh"].Value;
                        }
                    }
                    if (IsMatch_Pattern(line, out value, _p_ChuSuDung_ChuHo_LoaiGiayTo))
                    {
                        data.ChuSuDung_ChuHo_SoGiayTo = value[0].Groups["socmnd"].Value;
                        data.ChuSuDung_ChuHo_NoiCap = value[0].Groups["noicap"].Value;
                        data.ChuSuDung_ChuHo_NgayCap = value[0].Groups["ngaycap"].Value;

                        if (value[0].Value.Contains("CMND"))
                        {
                            data.ChuSuDung_ChuHo_LoaiGiayTo = "CMND";
                        }
                        if (value[0].Value.Contains("CCCD"))
                        {
                            data.ChuSuDung_ChuHo_LoaiGiayTo = "CCCD";
                        }
                        if (value.Count > 1)
                        {
                            data.ChuSuDung_VoChong_SoGiayTo = value[1].Groups["socmnd"].Value;
                            data.ChuSuDung_VoChong_NoiCap = value[1].Groups["noicap"].Value;
                            data.ChuSuDung_VoChong_NgayCap = value[1].Groups["ngaycap"].Value;

                            if (value[1].Value.Contains("CMND"))
                            {
                                data.ChuSuDung_VoChong_LoaiGiayTo = "CMND";
                            }
                            if (value[1].Value.Contains("CCCD"))
                            {
                                data.ChuSuDung_VoChong_LoaiGiayTo = "CCCD";
                            }
                        }
                    }
                    if (IsMatch_Pattern(line, out value, _p_GCN_Serial))
                    {
                        data.GCN_Serial = value[0].Groups["seria"].Value;
                    }
                    if (IsMatch_Pattern(line, out value, _p_GCN_MaVach))
                    {
                        data.GCN_MaVach = value[0].Groups["mavach"].Value;
                    }

                    //if(IsMatch_Pattern(line,out value,_p_chuyennhuong))
                    //{
                    //    a++;
                    //}
                    List<string> lst_cn = (Regex.Split(line, _p_tach_chuyennhuong)).Where(x => x.Length > 200).ToList();

                    if (lst_cn.Count > 1)
                    {
                        var line_cn = Regex.Replace(lst_cn[1], @":\n\n", ":").Replace("'", string.Empty); ;
                        if (IsMatch_Pattern(line_cn, out value, _p_chuyennhuong))
                        {
                            if (IsMatch_Pattern(line_cn, out value, _p_chuyennhuong_hoten))
                            {
                                data.ChuyenNhuong_ChuHo_HoTen = value[0].Groups["hoten"].Value;
                                if (VietnameseDecode(value[0].Value).Contains("ong"))
                                {
                                    data.ChuyenNhuong_ChuHo_GioiTinh = "1";
                                }
                                else
                                {
                                    if (VietnameseDecode(value[0].Value).Contains("ba"))
                                    {
                                        data.ChuyenNhuong_ChuHo_GioiTinh = "0";
                                    }
                                }
                                if (value.Count > 1)
                                {
                                    data.ChuyenNhuong_VoChong_HoTen = value[1].Groups["hoten"].Value;
                                    if (VietnameseDecode(value[1].Value).Contains("ong"))
                                    {
                                        data.ChuyenNhuong_VoChong_GioiTinh = "1";
                                    }
                                    else
                                    {
                                        if (VietnameseDecode(value[1].Value).Contains("ba"))
                                        {
                                            data.ChuyenNhuong_VoChong_GioiTinh = "0";
                                        }
                                    }
                                }
                            }
                            if (IsMatch_Pattern(line_cn, out value, _p_ChuSuDung_ChuHo_NamSinh))
                            {
                                data.ChuyenNhuong_ChuHo_NamSinh = value[0].Groups["namsinh"].Value;
                                if (value.Count > 1)
                                {
                                    data.ChuyenNhuong_VoChong_NamSinh = value[1].Groups["namsinh"].Value;
                                }
                            }
                            if (IsMatch_Pattern(line_cn, out value, _p_chuyennhuong_cmnd))
                            {
                                data.ChuyenNhuong_ChuHo_SoGiayTo = value[0].Groups["socmnd"].Value;
                                data.ChuyenNhuong_ChuHo_NoiCap = value[0].Groups["noicap"].Value;
                                data.ChuyenNhuong_ChuHo_NgayCap = value[0].Groups["ngaycap"].Value;

                                if (value.Count > 1)
                                {
                                    data.ChuyenNhuong_VoChong_SoGiayTo = value[1].Groups["socmnd"].Value;
                                    data.ChuyenNhuong_VoChong_NoiCap = value[1].Groups["noicap"].Value;
                                    data.ChuyenNhuong_VoChong_NgayCap = value[1].Groups["ngaycap"].Value;
                                }
                            }
                            if (IsMatch_Pattern(line_cn, out value, _p_chuyennhuong_loaigiayto))
                            {
                                if (value[0].Value.Contains("CMND"))
                                {
                                    data.ChuyenNhuong_ChuHo_LoaiGiayTo = "CMND";
                                }
                                if (value[0].Value.Contains("CCCD"))
                                {
                                    data.ChuyenNhuong_ChuHo_LoaiGiayTo = "CCCD";
                                }
                                if (value.Count > 1)
                                {
                                    if (value[1].Value.Contains("CMND"))
                                    {
                                        data.ChuyenNhuong_VoChong_LoaiGiayTo = "CMND";
                                    }
                                    if (value[1].Value.Contains("CCCD"))
                                    {
                                        data.ChuyenNhuong_VoChong_LoaiGiayTo = "CCCD";
                                    }
                                }
                            }
                            if (IsMatch_Pattern(line, out value, _p_chuyennhuong_diachi))
                            {
                                data.ChuyenNhuong_ChuHo_DiaChiThuongTru = value[0].Groups["diachi"].Value;
                                if (!string.IsNullOrEmpty(value[0].Groups["diachi"].Value) && IsMatch_Pattern(value[0].Groups["diachi"].Value, out value, _p_bo_dientich_trongdiachi))
                                {
                                    data.ChuyenNhuong_ChuHo_DiaChiThuongTru = string.Empty;
                                }
                                if (value.Count > 1)
                                {
                                    data.ChuyenNhuong_VoChong_DiaChiThuongTru = value[1].Groups["diachi"].Value;
                                }
                            }
                            if (IsMatch_Pattern(line, out value, _p_chuyennhuong_tangcho))
                            {
                                data.ChuyenNhuong_HinhThuc = "Tặng cho";
                            }
                            else
                            {
                                if (IsMatch_Pattern(line, out value, _p_chuyennhuong_cn))
                                {
                                    data.ChuyenNhuong_HinhThuc = "Chuyển nhượng";
                                }
                            }
                            if (IsMatch_Pattern(line, out value, _p_chuyennhuong_hopdong))
                            {
                                data.ChuyenNhuong_HopDong = value[0].Groups["sohopdong"].Value;
                            }
                        }
                    }

                    //Xu ly chuyen nhuong
                    if (checkObjectNotEmty(data))
                    {
                        listData.Add(data);
                    }
                }
                // MessageBox.Show(a.ToString());
            }

            // var listWord = query.Where(s => s != string.Empty).ToList();
            //var listWord = query.Select(x => x.Trim()).Where(s => s != string.Empty).ToList();

            // process_raw_data(listWord, listData);
            //File.Delete(filePath);

            string template = Path.GetFullPath("Template") + "\\Template.xls";
            Workbook book = new Workbook(template);
            progressBar1.Maximum = listData.Count + 1;
            progressBar1.Value = 0;
            try
            {
                Worksheet ws = book.Worksheets[0];
                int row = 3;
                Style style = new Style();
                style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                style.Borders[BorderType.TopBorder].Color = Color.Black;
                style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                style.Borders[BorderType.BottomBorder].Color = Color.Black;
                style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                style.Borders[BorderType.LeftBorder].Color = Color.Black;
                style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                style.Borders[BorderType.RightBorder].Color = Color.Black;
                style.HorizontalAlignment = TextAlignmentType.Left;
                Range rb = ws.Cells.CreateRange(string.Format("A3:AR{0}", (listData.Count + 2).ToString()));
                rb.SetStyle(style);
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
                            //ws.Cells[string.Format("{0}{1}", mapperColumn.ColumnName, row)].SetStyle(style);
                        }
                    }
                    row++;
                    progressBar1.Value++;
                }

                bool exists = System.IO.Directory.Exists(Path.GetFullPath("Output1\\"));
                if (!exists)
                    System.IO.Directory.CreateDirectory(Path.GetFullPath("Output\\"));
                book.Save(Path.GetFullPath("Output\\") + _fileName + ".xls", SaveFormat.Excel97To2003);

                MessageBox.Show("Done!");
                progressBar1.Value = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                book.Dispose();
            }
        }

        private bool checkObjectNotEmty(DataModel item)
        {
            bool check = false;
            IList<PropertyInfo> properties = new List<PropertyInfo>(item.GetType().GetProperties())
                           .Where(prop => prop.IsDefined(typeof(ExcelMapperAttribute), false))
                           .ToList();

            foreach (PropertyInfo property in properties)
            {
                if (property.GetValue(item, null) != null && !string.IsNullOrEmpty(property.GetValue(item, null).ToString()))
                {
                    check = true;
                    break;
                }
            }
            return check;
        }

        private bool IsMatch_Pattern(string line, out MatchCollection value, string pattern)
        {
            var regex = new Regex(pattern);
            value = regex.Matches(line);
            return regex.IsMatch(line);
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