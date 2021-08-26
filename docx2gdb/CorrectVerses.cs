using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace docx2gdb
{
    public partial class CorrectVerses : Form
    {
        public CorrectVerses()
        {
            InitializeComponent();
        }

        private void txtRightOrLeft_TextChanged(object sender, EventArgs e)
        {
            btnOK.Enabled = txtRight.Lines.Length == txtLeft.Lines.Length;
        }

        public string[] RightVerses
        {
            get
            {
                return txtRight.Lines;
            }
            set
            {
                txtRight.Lines = value;
            }
        }
        public string[] LeftVerses
        {
            get
            {
                return txtLeft.Lines;
            }
            set
            {
                txtLeft.Lines = value;
            }
        }

    }
}
