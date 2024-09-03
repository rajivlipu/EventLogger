using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EventLogger
{
    public partial class Setting : Form
    {
        public Setting()
        {
            InitializeComponent();
        }
        public  decimal CPU = 0;
        public decimal HDD = 0;
        public decimal WIFI = 0;
        public decimal WU = 0;

        public string Type="";
        public decimal val = 0;

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            RadioButton rbtn=(RadioButton)sender;
            if (rbtn.Checked) {
                
                Type= rbtn.Text;
                val = numericUpDown1.Value;
                if (rbtn.Text == "CPU" || rbtn.Text == "HDD" || rbtn.Text == "WIFI")
                    label1.Text = "Usage Percentage";
                if (rbtn.Text == "WU")
                    label1.Text = "Pending Days";
                if (rbtn.Text == "APP CRASH")
                {
                    label1.Visible = false;
                    numericUpDown1.Visible = false;
                }
                else {
                    label1.Visible = true;
                    numericUpDown1.Visible = true;
                }
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (Type == "") { Type = radioButton1.Text; }
            val = numericUpDown1.Value;
            DialogResult = DialogResult.OK;
            this.Close();

        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Setting_Load(object sender, EventArgs e)
        {

        }
    }
}
