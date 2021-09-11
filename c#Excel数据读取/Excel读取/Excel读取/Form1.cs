using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using ClosedXML.Excel;

namespace Excel读取
{
    public partial class Form1 : Form
    {
        int row1 = 1;
        int column1 = 0;
        XLWorkbook g_wb = new XLWorkbook();

        public Form1()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            g_wb = new XLWorkbook(@"E:\c#Excel数据读取\练习.xlsx");
            IXLWorksheet sheet = g_wb.Worksheet(1);
            if (column1<8)
            {
                column1++;
                richTextBoxRead.AppendText(sheet.Cell(row1, column1).GetString()+" ");
            }
            else
            {
                column1 = 0;
                richTextBoxRead.AppendText("\n");
                if (row1 < 4) row1++;
                else row1 = 1;
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            g_wb = new XLWorkbook(@"E:\c#Excel数据读取\练习.xlsx");
            IXLWorksheet sheet = g_wb.Worksheet(1);
            if (column1 < 8)
            {
                column1++;
                richTextBoxRead.AppendText(sheet.Cell(row1, column1).GetString() + " ");
            }
            else
            {
                column1 = 0;
                richTextBoxRead.AppendText("\n");
                if (row1 < 4) row1++;
                else row1 = 1;
            }
        }
    }
}
