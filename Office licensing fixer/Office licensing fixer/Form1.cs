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
        private Process cmd = new Process();


        public Form1()
		{
			InitializeComponent();
		}

        private void button1_Click(object sender, EventArgs e)
        {
            //run the /dstatus command
            string statusoutput = Runcommand("cscript ospp.vbs /dstatus");

            //dec the string to search for
            string s = "key:";

            //dec keys array
            List<string> keys = new List<string>();

            //itterate through the /dstatus command find all the "key:" strings and grab the next 5 chars afer it and put it into the keys array
            for (int i = 0; i < statusoutput.Length - s.Length + 1; i++)
            {
                if (statusoutput.Substring(i, s.Length).Equals(s))
                {
                    string key = statusoutput.Substring(i + 5, 5);
                    listBox1.Items.Add(key);
                    keys.Add(key);
                    
                }
            }
            //run the command for each key and print
            foreach (string key in keys) 
            {
                
                string output = Runcommand("cscript ospp.vbs /unpkey:"+key);
                listBox1.Items.Add("Removed: " + key);
            }
            //notify the user of completion
            listBox1.Items.Add("All keys removed - restart office and log in!");

        }

        public void Form1_Load(object sender, EventArgs e)
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
            

            

            if (File.Exists(office16filelocation))
            {
                listBox1.Items.Add("Office 16 detected");
                if(isadmin == true)
                {
                    button1.Enabled = true;
                }
  
            }
            else 
            {
                listBox1.Items.Add("Office 16 NOT detected");
            }
            

        }

        //function that takes a cmd command to be run as a string and returns its output as a string
        private string Runcommand(string CommandToRun)
        {
            cmd.StartInfo.WorkingDirectory = @"C:\Program Files\Microsoft Office\Office16\";
            cmd.StartInfo.FileName = "cmd.exe";
            cmd.StartInfo.RedirectStandardInput = true;
            cmd.StartInfo.RedirectStandardOutput = true;
            cmd.StartInfo.CreateNoWindow = true;
            cmd.StartInfo.UseShellExecute = false;
            cmd.StartInfo.Verb = "runas";
            cmd.Start();

            cmd.StandardInput.WriteLine(CommandToRun);
            cmd.StandardInput.Flush();
            cmd.StandardInput.Close();
            cmd.WaitForExit();
            string cmdreturn = cmd.StandardOutput.ReadToEnd();

            return cmdreturn;
        }
        

    }
}
