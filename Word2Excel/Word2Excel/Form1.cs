//using DocumentFormat.OpenXml.Spreadsheet;
using Aspose.Cells;
using Mammoth;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
            //  _tempdata = result.Value; // The raw text
            // var warnings = result.Warnings;
            //File.WriteAllText("b.txt", html);
            var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "myfile.txt");
            File.WriteAllText(filePath, result.Value);

            string[] lines = File.ReadAllLines(filePath);
            List<string> list = lines.Where(i => i != string.Empty).ToList();

            foreach (var item in list)
            {
            }
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox1.Text))
            {
                return;
            }
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
    }

    public class DataModel
    {
        public string SoTo { get; set; }
        public string SoThua { get; set; }
        public string DienTich { get; set; }
        public string MucDichSuDung { get; set; }
        public string DiaChiThuaDat { get; set; }
        public string HoTen { get; set; }
        public string NamSinh { get; set; }
        public string GioiTinh { get; set; }
        public string LoaiGiayTo { get; set; }
        public string SoGiayTo { get; set; }
        public string NgayCap { get; set; }
        public string NoiCap { get; set; }
        public string DiaChiThuongTru { get; set; }
        public string HinhThuc { get; set; }
        public string HopDong { get; set; }

        public string Serial { get; set; }

        public string MaVach { get; set; }

        public string SoGCN { get; set; }

        public string NgayCapMoi { get; set; }

        public string NgayChuyenNhuong { get; set; }
    }
}