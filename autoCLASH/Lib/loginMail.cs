using ClashCheck;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace autoCLASH.UI
{
    public partial class loginMail : Form
    {
        public string IDT { get; set; }
        public string PASST { get; set; }
        public string excelpath { get; set; }
        public string company { get; set; }
        public loginMail(string s)
        {
            InitializeComponent();
            company = s;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (ID.Text.Trim() == "")
            {
                MessageBox.Show("사원번호를 입력해주세요");
                return;
            }
            if (PASS.Text.Trim() == "")
            {
                MessageBox.Show("비밀번호를 입력해주세요");
                return;
            }
            IDT = ID.Text;
            PASST = PASS.Text;
            this.Close();
            string topath = "jms@infoget.co.kr";
            if (company!="IFG")
            {
                topath = IDT + "@hmd.co.kr";
            }
            else
            {
                topath = IDT + "@gmail.co.kr";
            }
            SendMails amail = new SendMails(excelpath, topath, IDT,PASST,company);
            amail.Show();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        public List<string> getID()
        {
            List<string> test = new List<string>();
            test.Add(IDT);
            test.Add(PASST);
            return test;
        }
    }
}
