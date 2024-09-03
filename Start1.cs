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
    public partial class Start1 : Form
    {
        public Start1()
        {
            InitializeComponent();
            
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            this.Close();
            //if (label1.Text == "SmartOPS Agent Loading.....") {
            //    
            //}
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
