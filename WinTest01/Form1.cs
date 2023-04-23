using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WinTest01
{
    public partial class Form1 : Form
    {
        ExcelClass Ex1 = new ExcelClass();
        
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            bool res = Ex1.openExcel();
            //Ex1.writeExcel(this.textBox1.Text);

            string[,] Data = { { "Peter", "John", "Thomas" }, { "David", "Christian", "Maria" } };
            Ex1.writeExcelData(Data,"Sheet2",3,3);

            if (res) Ex1.closeExcel();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Ex1.killExcelProcesses();
        }
    }
}
