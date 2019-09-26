using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MigrationUtility
{
    public partial class Login : System.Windows.Forms.Form
    {
        common cmn = new common();
        public Login()
        {
            InitializeComponent();
        }

        private void Login_Load(object sender, EventArgs e)
        {
            TaskCreation tskfrm = new TaskCreation();
        }

        private void btn_sourceLogin_Click(object sender, EventArgs e)
        {
            timer1.Enabled = true;
            common.OutPutKey = null;
            backgroundWorkerLogin.RunWorkerAsync();
        }

        private void backgroundWorkerLogin_DoWork(object sender, DoWorkEventArgs e)
        {

            var LoginDetails = cmn.Auth(txt_SourceUName.Text, txt_SourceUPass.Text, txt_SourceURL.Text);
            common.SourceProjContext = LoginDetails.Value;
            common.OutPutKey = LoginDetails.Key;
        }


        private void timer1_Tick(object sender, EventArgs e)
        {
            progressBarLogin.Value = progressBarLogin.Value + 1;
            if (common.OutPutKey == "True")
            {
                TaskCreation frm = new TaskCreation();
               
                frm.Show();
                this.Hide();
                timer1.Enabled = false;
            }
            else if (common.OutPutKey == "False")
            {
                lbl_SourceErrorMessage.Text = "Invalid Credentials!!";
                progressBarLogin.Value = 0;
                timer1.Enabled = false;
            }
            else if (progressBarLogin.Value>98)
            {
                progressBarLogin.Value = 94;
            }
        }
    }
}
