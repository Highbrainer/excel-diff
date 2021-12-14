using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelDiff
{
    public partial class Form1 : Form
    {
        private List<Column> Columns { get; set; } = new List<Column>();

        private List<string> Sheets;

        private ExcelComparator excelComparator;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            textBoxA.Text = openFileDialog1.FileName;

        }

        private void UpdateColumnsAndSheets()
        {
            if (!string.IsNullOrWhiteSpace(textBoxA.Text) && !string.IsNullOrWhiteSpace(textBoxB.Text))
            {
                excelComparator = new ExcelComparator(textBoxA.Text, textBoxB.Text);
                this.Sheets = excelComparator.Sheets();
                this.sheetsComboBox.Items.Clear();
                this.sheetsComboBox.Items.AddRange(Sheets.ToArray());
                this.sheetsComboBox.SelectedIndex = 0;
                UpdateColumns();
            }
        }

        private void UpdateColumns()
        {
            if (string.IsNullOrWhiteSpace(sheetsComboBox.Text))
            {
                return;
            }
            Columns = excelComparator.Columns(sheetsComboBox.Text);
            this.columnsComboBox.Items.Clear();
            this.columnsComboBox.Items.AddRange(Columns.ToArray());
            this.columnsComboBox.SelectedIndex = 0;
        }


        private void buttonLaunch_Click(object sender, EventArgs e)
        {
            
            progressBarA.Maximum = 0;
            progressBarA.Value = 0;
            excelComparator.Compare(sheetsComboBox.Text, ((Column)columnsComboBox.SelectedItem).index, this.progressBarA);
            MessageBox.Show("C'est fait !");
        }

        private void buttonB_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            textBoxB.Text = openFileDialog1.FileName;
        }

        private void textBoxA_TextChanged(object sender, EventArgs e)
        {
            UpdateColumnsAndSheets();
        }

        private void textBoxB_TextChanged(object sender, EventArgs e)
        {
            UpdateColumnsAndSheets();
        }

        private void sheetsComboBox_TextChanged(object sender, EventArgs e)
        {
            if (this.excelComparator != null) UpdateColumns();
        }
    }
}
