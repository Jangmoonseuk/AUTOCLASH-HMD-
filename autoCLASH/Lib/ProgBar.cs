/******************************************************************************************
 ******************************************************************************************/

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace autoCLASH.Lib
{
    public partial class ProgBar : Form
    {
        // parentForm이 있을때
        public ProgBar(int maxNum, Form parentForm, string title)
        {
            InitializeComponent();
            this.Text = title;
            this.StartPosition = FormStartPosition.Manual;

            if (parentForm == null)
                SetLocation(this);
            else
                SetLocation(parentForm);

            this.Size = new Size(298, 53);
            this.progressBar1.Value = 0;
            this.progressBar1.Maximum = maxNum;
        }

        // 기본
        public ProgBar(int maxNum, string title)
        {
            InitializeComponent();
            this.Text = title;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.Size = new Size(298, 53);
            this.progressBar1.Value = 0;
            this.progressBar1.Maximum = maxNum;
        }

        void SetLocation(Form parentForm)
        {
            Point loc = new Point(parentForm.Location.X + parentForm.Width / 2, parentForm.Location.Y + parentForm.Height / 2);
            this.Location = loc;
        }

        public void IncreaseValue()
        {
            this.progressBar1.Value++;
        }

        public void IncreaseValue_ChangeTitle(string title)
        {
            this.Text = title;
            this.progressBar1.Value++;
        }
    }
}
