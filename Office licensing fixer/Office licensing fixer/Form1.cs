using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Security.Principal;

namespace Office_licensing_fixer
{
	public partial class Form1 : Form
	{
        

        public Form1()
		{
			InitializeComponent();
		}

        private void button1_Click(object sender, EventArgs e)
        {
            Process cmd = new Process();
            cmd.StartInfo.WorkingDirectory = @"C:\Program Files\Microsoft Office\Office16\";
            cmd.StartInfo.FileName = "cmd.exe";
            cmd.StartInfo.RedirectStandardInput = true;
            cmd.StartInfo.RedirectStandardOutput = true;
            cmd.StartInfo.CreateNoWindow = true;
            cmd.StartInfo.UseShellExecute = false;
            cmd.StartInfo.Verb = "runas";
            cmd.Start();

            cmd.StandardInput.WriteLine("cscript ospp.vbs /dstatus");
            cmd.StandardInput.Flush();
            cmd.StandardInput.Close();
            cmd.WaitForExit();
            

            string statusoutput = cmd.StandardOutput.ReadToEnd();

            //declar the string to search for
            string s = "key:";

            List<string> keys = new List<string>();

            for (int i = 0; i < statusoutput.Length - s.Length + 1; i++)
            {
                if (statusoutput.Substring(i, s.Length).Equals(s))
                {
                    string key = statusoutput.Substring(i + 5, 5);
                    listBox1.Items.Add(key);
                    keys.Add(key);
                    
                }
            }

            foreach (string key in keys) 
            {
                cmd.StartInfo.WorkingDirectory = @"C:\Program Files\Microsoft Office\Office16\";
                cmd.StartInfo.FileName = "cmd.exe";
                cmd.StartInfo.RedirectStandardInput = true;
                cmd.StartInfo.RedirectStandardOutput = true;
                cmd.StartInfo.CreateNoWindow = true;
                cmd.StartInfo.UseShellExecute = false;
                cmd.StartInfo.Verb = "runas";
                cmd.Start();
                listBox1.Items.Add("Removing: " + key);
                string command_torun = "cscript ospp.vbs /unpkey:";
                cmd.StandardInput.WriteLine(command_torun+key);
                cmd.StandardInput.Flush();
                cmd.StandardInput.Close();
                cmd.WaitForExit();
                listBox1.Items.Add("Removed: " + key);
            }

            listBox1.Items.Add("All keys removed - restart office and log in!");

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //check if run as admin and save as bool
            var identity = WindowsIdentity.GetCurrent();
            var principal = new WindowsPrincipal(identity);
            bool isadmin = principal.IsInRole(WindowsBuiltInRole.Administrator);
            

            FormBorderStyle = FormBorderStyle.FixedSingle;
            button1.Enabled = false;
            
            if (isadmin == false)
            {
                listBox1.Items.Add("ERROR: Please run this as Admin");
            }
            else
            {
                listBox1.Items.Add("Program running as admin");
            }

            string office16filelocation = @"C:\Program Files\Microsoft Office\Office16\ospp.vbs";
            

            

            if (File.Exists(office16filelocation) && isadmin == true)
            {
                listBox1.Items.Add("Office 16 detected");
                button1.Enabled = true;
                
            }
            else 
            {
                listBox1.Items.Add("Office 16 NOT detected");
            }
            

        }
        

    }
}
