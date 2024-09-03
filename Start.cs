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
    public partial class Start : Form
    {
        public Start()
        {
            InitializeComponent();
            
        }
        Form1 frm = new Form1();
        Start1 Start1 = new Start1();   
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
                Application.Exit();
            Application.ExitThread();
         
        }
        bool isVis = false;
        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (isVis == false)
            {
                frm.Show();
                isVis = true;
                openToolStripMenuItem.Text = "Close";
            }
            else {
                frm.Hide();
                isVis = false;
                openToolStripMenuItem.Text = "Open";
            }            
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (isVis == false)
            {
                frm.Show();
                isVis = true;
                openToolStripMenuItem.Text = "Close";
            }
            else
            {
                frm.Hide();
                isVis = false;
                openToolStripMenuItem.Text = "Open";
            }
        }

        private void Start_Load(object sender, EventArgs e)
        {
            Start1.Show();
            notifyIcon1.ShowBalloonTip(1000, "SmartOPS Agent Notification", "Smart OPS is monitoring this system.", ToolTipIcon.Info);

        }

        private void timer1_Tick(object sender, EventArgs e)
        {

            if (frm.notify != "")
            {
                notifyIcon1.ShowBalloonTip(1000, "SmartOPS Notification", frm.notify, frm.notifytype == "warn" ? ToolTipIcon.Warning : frm.notifytype  == "err" ? ToolTipIcon.Error : ToolTipIcon.Info);
                frm.notify = "";
            }

        }
    }
}
