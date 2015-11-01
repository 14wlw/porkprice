using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using Tools.Excel;


namespace porkprice
{
    public partial class Form1 : Form
    {
        List<pork> porklist = new List<pork>();
        public Form1()
        {
            InitializeComponent();
        }
        private void show()
        {
            Series data = new Series();
            data.ChartType = SeriesChartType.Line;
            int i = 0;
            for (i = 1; i <= 50; i++)
            { data.Points.Add(i, i); }
            chart1.Series.Add(data);
        }
        private void loadinfo()
        {
            int i = 0;
            clsExcelReader er = new clsExcelReader();
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = "c:\\";//初始路径
            openFileDialog.Filter = "文本文件|*.*|C#文件|*.cs|所有文件|*.*";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.FilterIndex = 1;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                er.FileName = openFileDialog.FileName;
            }
            er.SheetNumber = 1;

            if (er.OpenFileContinuously() == false)
            {
                label1.Text = er.ErrorString;
                return;
            }
            int iCount = er.RowCount;
            for (i = 0; i < iCount - 2; i++)
            {
                pork Pork = new pork();
                Pork.name = er.getTextInOneCell(i + 3, 20);
                string priceA = er.getTextInOneCell(i + 3, 7);
                if (priceA.IndexOf('-') == -1)
                { Pork.price = Convert.ToDouble(er.getTextInOneCell(i + 3, 7)); }

                else

                { Pork.price = Convert.ToDouble(priceA.Substring(0, priceA.IndexOf('-'))); }

                Pork.date = er.getTextInOneCell(i + 3, 18).Substring(0, 9);
                porklist.Add(Pork);
            }
            label1.Text = porklist.Count.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            show();
            loadinfo();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            int i = 0;
            for (i = 0; i <= 10; i++)
            { label1.Text += porklist[i].date.ToString(); }

        }
    }
}
