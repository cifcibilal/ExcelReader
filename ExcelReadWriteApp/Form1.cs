using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelAcmaVeOkumaEgitim;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Data.Odbc;
using System.Data.OleDb;

namespace ExcelReadWriteApp
{
    public partial class Form1 : Form
    {
        //private string[,] read = new string[12, 4];
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
                if (i<=sutunSayisi)
                {
                    comboBoxSutun.Items.Add(read[0, i - 1]);
                    comboBox2.Items.Add(read[0, i - 1]);
                    comboBox1.Items.Add(read[0, i - 1]);

                }
                for (int p = 0; p < sutunSayisi; p++)
                {
                    listBox1.Items.Add(read[0,p]+" : " +read[i, p]);
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
                string[,] read = excel.ReadRange(1, 1,satirSayisi, sutunSayisi);
                excel.Close();

            listBox1.Items.Clear();

                int secilenBox = comboBoxSutun.SelectedIndex;
            /*
                if (secilenBox == 0) 
                {
                    listBox1.Items.Clear();
                    for (int i = 0; i < satirSayisi; i++)
                    {      
                        listBox1.Items.Add(read[i, 0]);
                    }
                }
                else if (secilenBox == 1)
                {
                    listBox1.Items.Clear();
                    for (int i = 0; i < satirSayisi; i++)
                    {
                        listBox1.Items.Add(read[i, 1]);
                    }
                }
                else if (secilenBox == 2)
                {
                    listBox1.Items.Clear();
                    for (int i = 0; i < satirSayisi; i++)
                    {
                        listBox1.Items.Add(read[i, 2]);
                    }
                }
                else if (secilenBox == 3)
                {
                    listBox1.Items.Clear();
                    for (int i = 0; i < satirSayisi; i++)
                    {
                        listBox1.Items.Add(read[i, 3]);
                    }
                }
                else
                {
                    MessageBox.Show("Item seçerken hata oluştur");
                }
            */
            for (int i = 0; i < satirSayisi; i++)
            {
                listBox1.Items.Add(read[i, secilenBox]);
            }

            
            }

        private void btnSaveAs_Click(object sender, EventArgs e)
        {
            excel = new Excel(txtFilePath.Text, 1);
            int satirSayisi = excel.ws.UsedRange.Rows.Count;
            int sutunSayisi = excel.ws.UsedRange.Columns.Count;
            string[,] read = excel.ReadRange(1, 1, satirSayisi, sutunSayisi);
            excel.Close();

            SaveFileDialog saveFileDialog = new SaveFileDialog();

            saveFileDialog.Filter = "Excel Dosyaları|*.xls;*.xlsx|Tüm Dosyalar|*.*";
            saveFileDialog.DefaultExt = "xlsx";

            Excel excelSaveAs = new Excel(@"C:\Users\muham\Desktop\KM.xlsx",1);

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string savePath = saveFileDialog.FileName;

                excelSaveAs.WriteRange(1, 1, satirSayisi, 1, read);
                excelSaveAs.SaveAs(@savePath);
                excelSaveAs.Close();

                MessageBox.Show("Dosya başarılıyla kaydedildi.");
            }
            else
            {
                MessageBox.Show("Dosya kaydedilemedi.");
            }

        }
    }
}
