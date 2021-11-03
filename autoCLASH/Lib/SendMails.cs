using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using autoCLASH.UI;

namespace ClashCheck
{
    public partial class SendMails : Form
    {
        public string id { get; set; }
        public string pass { get; set; }
        string to = "";
        public string company { get; set; }
        public SendMails(string addfile,string topath,string a,string b,string s)
        {
            InitializeComponent();
            if (addfile != null&&addfile!="")
            {
                attachFile.Add(addfile);
            }
            to = topath;
            id = a; pass = b; company = s;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            sendmail();
            this.Close();
        }
        public void sendmail()
        {
            try
            {
                MailMessage mail = new MailMessage();

                //set the address  
                //mail.To.Add(to);
                foreach (var a in txt_toAddress.Text.Split(',').ToList())
                {
                    mail.To.Add(a);
                }

                //set the contents  
                mail.Subject = txt_subject.Text;
                mail.Body = Message.Text;//Message.Text;
                mail.BodyEncoding = System.Text.Encoding.UTF8;
                mail.SubjectEncoding = System.Text.Encoding.UTF8;

                //set attachment  
                System.Net.Mail.Attachment attachment;
                foreach (var a in attachFile)
                {
                    attachment = new System.Net.Mail.Attachment(a);
                    mail.Attachments.Add(attachment);
                }
                SmtpClient SmtpServer;
                //setting smtpServer  


                if (true)
                {
                    mail.From = new MailAddress(id+"@gmail.com", "관리자", Encoding.UTF8);
                    SmtpServer = new SmtpClient("smtp.gmail.com",587);
                    SmtpServer.Credentials = new System.Net.NetworkCredential(id, pass);
                    //SmtpServer.UseDefaultCredentials = false;//구글용
                    SmtpServer.EnableSsl = true; //구글용
                    //SmtpServer.DeliveryMethod = SmtpDeliveryMethod.Network;//구글용
                }
                else
                {
                    mail.From = new MailAddress("jmsa45@hmd.co.kr", "LBD Clash Check", Encoding.UTF8);
                    SmtpServer = new SmtpClient("hmd.mipoware.co.kr", 25);
                    //mail.From = new MailAddress(userId + "@hmd.co.kr", "Naviswork Clashcheck", Encoding.UTF8);
                    //SmtpServer = new SmtpClient("hmd.mipoware.co.kr", 25); //SMTP서버를 만들고 미포
                    //SmtpServer.Port = 25;
                    SmtpServer.Credentials = new System.Net.NetworkCredential("jmsa45", "wkdans562@");
                    //SmtpServer.Credentials = new NetworkCredential(userId, userPass);
                }

                //SmtpServer.Send(mail);  
                // MessageBox.Show("mail Send");  

                //async send mail  
                object userState = mail;
                SmtpServer.SendCompleted += new SendCompletedEventHandler(client_SendCompleted);
                SmtpServer.SendAsync(mail, userState);

            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
            }
        }
        void client_SendCompleted(object sender, System.ComponentModel.AsyncCompletedEventArgs e)
        {  
            //Get the Original MailMessage object  
            MailMessage mail = (MailMessage)e.UserState;  
  
            //write out the subject  
            string subject = mail.Subject;  
  
            if (e.Error == null)  
            {  
                // Send succeeded   
               string message = string.Format("Send Mail with subject [{0}]", subject);  
                MessageBox.Show(message);  
            }  
            else if (e.Cancelled == true)  
            {  
                string message = string.Format("Send canceld for mail with subject [{0}]", subject);  
                MessageBox.Show(message);  
            }  
            else  
            {  
                // log exception   
                string message = string.Format("Send Mail Fail - Error: [{0}]", e.Error.ToString());  
                MessageBox.Show(message);  
            }
        }
        List<string> attachFile = new List<string>();
        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.Multiselect = true;
            DialogResult dialog = openFileDialog1.ShowDialog();
            if (dialog == DialogResult.OK)
            {
                listBox1.Items.Clear();
                attachFile.AddRange(openFileDialog1.FileNames.ToList());
                foreach (string fileName in attachFile)
                    this.listBox1.Items.Add(fileName.Split('\\').Last());
            }
            else
            {

            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            attachFile.Clear();
        }

        private void txt_toAddress_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
