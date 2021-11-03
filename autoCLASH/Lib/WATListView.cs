using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace autoCLASH.Lib
{
    // 이벤트에사용할아규먼트 

    public class WATListComboEventArgs : EventArgs

    {

        int iLastX = 0;

        public int LastX

        {

            get { return iLastX; }

            set { iLastX = value; }

        }



        int iLastY = 0;

        public int LastY

        {

            get { return iLastY; }

            set { iLastY = value; }

        }



        int iSelectSubItem = 0;

        public int SelectSubItem

        {

            get { return iSelectSubItem; }

            set { iSelectSubItem = value; }

        }



        System.Windows.Forms.ListViewItem m_SelectedItem;

        public System.Windows.Forms.ListViewItem SelectedItem

        {

            get { return m_SelectedItem; }

            set { m_SelectedItem = value; }

        }



        ComboBox m_combobox;

        public System.Windows.Forms.ComboBox Combobox

        {

            get { return m_combobox; }

            set { m_combobox = value; }

        }



        bool bCanceled = true;

        public bool Canceled

        {

            get { return bCanceled; }

            set { bCanceled = value; }

        }



    }



    public class WATListView : System.Windows.Forms.ListView

    {

        WATListComboEventArgs m_lvwComboArgs = new WATListComboEventArgs();



        public delegate void eventComboChanged(WATListComboEventArgs e);

        public event eventComboChanged ComboChanged;



        readonly int MAX_COMBO_COUNT = 10;

        string subItemText = "";

        /// <summary> 콤보상자데이터가변경되었을때리스트값도변경할것인가? /// </summary> 

        bool bComboToList = false;

        public bool ComboToList

        {

            get { return bComboToList; }

            set { bComboToList = value; }

        }



        private System.Windows.Forms.ComboBox[] cboData;//= new System.Windows.Forms.ComboBox[MAX_COMBO_COUNT]; 



        public WATListView()

        {

            cboData = new System.Windows.Forms.ComboBox[MAX_COMBO_COUNT];

            for (int iTemp = 0; iTemp < MAX_COMBO_COUNT; iTemp++)

            {

                this.cboData[iTemp] = new System.Windows.Forms.ComboBox();

                this.cboData[iTemp].Hide();

                this.cboData[iTemp].SelectedIndexChanged += new EventHandler(WATListView_SelectedIndexChanged);

                this.cboData[iTemp].LostFocus += new EventHandler(WATListView_LostFocus);





                this.cboData[iTemp].KeyDown += new KeyEventHandler(WATListView_KeyDown);

                this.cboData[iTemp].DropDownStyle = ComboBoxStyle.DropDownList;

                this.Controls.Add(cboData[iTemp]);

            }



            this.FullRowSelect = true;

            this.Click += new EventHandler(WATListView_DoubleClick);

            this.MouseDown += new System.Windows.Forms.MouseEventHandler(WATListView_MouseDown);

        }



        void WATListView_KeyDown(object sender, KeyEventArgs e)

        {

            ComboBox cb = sender as ComboBox;



            if (e.KeyValue == 13)

            {

                m_lvwComboArgs.Canceled = false;

                if (null != ComboChanged)

                    ComboChanged(m_lvwComboArgs);

            }

            else

            {

                m_lvwComboArgs.Canceled = true;

            }



            if (e.KeyValue == 13 || e.KeyValue == 27)

            {

                cb.Hide(); // Lost Focus 가호출된다. 

            }

        }



        void WATListView_LostFocus(object sender, EventArgs e)

        {

            ComboBox cb = sender as ComboBox;

            cb.Hide(); // 이거없애도되나? 

            if (!m_lvwComboArgs.Canceled)

            {

                if (this.ComboToList)

                {

                    string str = m_lvwComboArgs.Combobox.Text;

                    m_lvwComboArgs.SelectedItem.SubItems[m_lvwComboArgs.SelectSubItem].Text = str;



                }



                if (null != ComboChanged)

                    ComboChanged(m_lvwComboArgs);

            }

        }



        void WATListView_SelectedIndexChanged(object sender, EventArgs e)

        {

            m_lvwComboArgs.Canceled = false;

            m_lvwComboArgs.Combobox = sender as ComboBox;

        }



        void WATListView_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)

        {

            m_lvwComboArgs.SelectedItem = this.GetItemAt(e.X, e.Y);

            m_lvwComboArgs.LastX = e.X;

            m_lvwComboArgs.LastY = e.Y;

        }



        void WATListView_DoubleClick(object sender, EventArgs e)

        {
            int start = m_lvwComboArgs.LastX;
            int position = 0;

            int end = 0; //this.Columns[0].Width; 
            for (int i = 0; i < this.Columns.Count; i++)
           {
                end += this.Columns[i].Width;
                if (start > position && start < end)
                {
                    m_lvwComboArgs.SelectSubItem = i;
                    break;
                }
                position = end;
                //end += this.Columns[i].Width;
            }
            if (this.cboData[m_lvwComboArgs.SelectSubItem].Items.Count <= 0) return;





            subItemText = this.m_lvwComboArgs.SelectedItem.SubItems[m_lvwComboArgs.SelectSubItem].Text;





            Trace.Write("SelectSubItem : " + m_lvwComboArgs.SelectSubItem.ToString() + subItemText);





            this.cboData[m_lvwComboArgs.SelectSubItem].Size = new System.Drawing.Size(end - position, this.m_lvwComboArgs.SelectedItem.Bounds.Bottom - this.m_lvwComboArgs.SelectedItem.Bounds.Top);

            this.cboData[m_lvwComboArgs.SelectSubItem].Location = new System.Drawing.Point(position, m_lvwComboArgs.SelectedItem.Bounds.Y);

            this.cboData[m_lvwComboArgs.SelectSubItem].Show();

            this.cboData[m_lvwComboArgs.SelectSubItem].Text = subItemText;

            this.cboData[m_lvwComboArgs.SelectSubItem].SelectAll();

            this.cboData[m_lvwComboArgs.SelectSubItem].Focus();

        }



        public void AddString(int iSubItem, string[] strData)

        {

            cboData[iSubItem].Items.AddRange(strData);

        }



    }

}

