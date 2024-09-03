using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Forms;

namespace EventLogger
{
    public partial class ProcessCPU : Form
    {
        public ProcessCPU()
        {
            InitializeComponent();
            this.StartPosition = FormStartPosition.Manual;
            this.Location = new System.Drawing.Point(Screen.PrimaryScreen.WorkingArea.Width - this.Width, Screen.PrimaryScreen.WorkingArea.Height - this.Height);
            timer1.Enabled = true;  

        }

        private void CalcCpu()
        {

        }

        private void ProcessCPU_Load(object sender, EventArgs e)
        {

            Dictionary<string, double> proc = new Dictionary<string, double>();

            List<Process> list = new List<Process>();
            list = Process.GetProcesses().ToList();// GetTopProc(5);
            foreach (Process p in list)
            {
                try
                {
                    if (proc.ContainsKey(p.ProcessName))
                    {
                        proc[p.ProcessName] = proc[p.ProcessName] + p.TotalProcessorTime.TotalSeconds;
                    }
                    else
                    {
                        proc.Add(p.ProcessName, p.TotalProcessorTime.TotalSeconds);
                    }
                }
                catch { }
            }

            foreach (var a in proc.OrderByDescending(key => key.Value).Take(3))
            {
                listBox1.Items.Add(a.Key);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedItems.Count > 0) {
                List<Process> list = new List<Process>();
                list = Process.GetProcesses().ToList();// GetTopProc(5);
                foreach (Process p in list)
                {
                    if (p.ProcessName == listBox1.SelectedItems[0].ToString())
                    {
                        p.Kill();
                    }
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            this.TopMost = true;    
        }
    }
}
