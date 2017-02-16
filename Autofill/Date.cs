using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Autofill
{
    public partial class Date : Form
    {
        Properties.Settings ps = Properties.Settings.Default;

        public Date()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text == "Город")
                ps.is_city = true;
            else ps.is_city = false;

            this.Close();
        }

       
    }
}
