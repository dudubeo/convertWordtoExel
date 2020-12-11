using DocumentFormat.OpenXml.Spreadsheet;
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