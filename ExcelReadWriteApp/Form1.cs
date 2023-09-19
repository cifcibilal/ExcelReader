using System;
using System.Windows.Forms;
using ExcelAcmaVeOkumaEgitim;

namespace ExcelReadWriteApp
{
    public partial class Form1 : Form
    {
        OpenFileDialog fileSelect;
        Excel excel;
        public Form1()
        {
            InitializeComponent();
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            fileSelect = new OpenFileDialog();
            fileSelect.Title = "Excel dosyası seç";
            fileSelect.Filter = "Excel Dosyaları|*.xls;*.xlsx|Tüm Dosyalar|*.*";
            if (fileSelect.ShowDialog() == DialogResult.OK)
            {

                txtFilePath.Text = fileSelect.FileName;
                txtFileName.Text = fileSelect.SafeFileName;
            }
            else
            {
                MessageBox.Show("Dosya seçilemedi.");
            }
            excel = new Excel(txtFilePath.Text, 1);
            int satirSayisi = excel.ws.UsedRange.Rows.Count;
            int sutunSayisi = excel.ws.UsedRange.Columns.Count;

            string[,] read = excel.ReadRange(1, 1, satirSayisi, sutunSayisi);
            excel.Close();
            //12 satır sayısını ifade ediyor
            for (int i = 1; i < satirSayisi; i++) //i 1 olmasının sebebi ilk satırlar sütun ismi
            {
                listBox1.Items.Add("{");
                if (i <= sutunSayisi)
                {
                    comboBoxSutun.Items.Add(read[0, i - 1]);
                }
                for (int p = 0; p < sutunSayisi; p++)
                {
                    listBox1.Items.Add(read[0, p] + " : " + read[i, p]);
                }
                listBox1.Items.Add("}");
                listBox1.Items.Add("");
            }
        }
        private void comboBoxSutun_SelectedIndexChanged(object sender, EventArgs e)
        {
            excel = new Excel(txtFilePath.Text, 1);
            int satirSayisi = excel.ws.UsedRange.Rows.Count;
            int sutunSayisi = excel.ws.UsedRange.Columns.Count;
            string[,] read = excel.ReadRange(1, 1, satirSayisi, sutunSayisi);
            excel.Close();

            listBox1.Items.Clear();

            int secilenBox = comboBoxSutun.SelectedIndex;

            for (int i = 0; i < satirSayisi; i++)
            {
                listBox1.Items.Add(read[i, secilenBox]);
            }
        }
    }
}
