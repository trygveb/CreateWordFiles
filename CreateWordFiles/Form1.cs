using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CreateWordFiles
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void buttonOk_Click(object sender, EventArgs e)
        {
            String file = @"d:\Development\CreateWordFiles\CreateWordFiles\Documents\Test.docx";
            Creator.CreateWordprocessingDocument(file, this.textBoxDanceName.Text, this.textBoxDates.Text);
            this.Close();
        }

    }
}
