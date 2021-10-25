using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using OfficeOpenXml;
namespace MergeTools
{
    public partial class Form1 : Form
    {
        public static Form1 frm1 = null;//建立一個自身的靜態物件
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }


        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            file.ShowDialog();
            this.textBox1.Text = file.FileName;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            file.ShowDialog();
            this.textBox2.Text = file.FileName;
        }

        private void button3_Click(object sender, EventArgs e)
        {

            this.label1.Text = "執行中";
            this.label1.BackColor = Color.White;
            this.Refresh();

            string bomPath = this.textBox1.Text;
            string locationPath = this.textBox2.Text;

            try
            {
                if (bomPath != "" && locationPath != "")
                {
                    List<BomModel> bom = Bom.GetBomModel(bomPath);
                    List<LocationModel> locations = LocationC.GetLocationModel(locationPath);

                    bool isFinish = Merge.MergeData(locationPath, locations, bom);
                    if (isFinish)
                    {
                        this.label1.Text = "完成";
                        this.label1.BackColor = Color.AliceBlue;
                    }
                }
                else
                {
                    this.label1.Text = "您未選擇檔案";
                    this.label1.BackColor = Color.Red;
                }
            }
            catch (Exception ex)
            {
                this.label1.Text = ex.Message;
                this.label1.BackColor = Color.Red;
            }

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {
         
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.Show();
        }
    }
}
