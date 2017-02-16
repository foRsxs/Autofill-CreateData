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
using System.Xml;


namespace Autofill
{
    public partial class Options : Form
    {
        System.Windows.Forms.Timer timerClicker; //таймер, просто для примера
        Properties.Settings ps = Properties.Settings.Default;
        Form1 frm;
        DataTable dt = new DataTable("zl");
        String XMLFileName;

        public Options()
        {
            InitializeComponent();

            timerClicker = new System.Windows.Forms.Timer();
            timerClicker.Interval = 1000;
            timerClicker.Tick += new EventHandler(timerClicker_Tick);
            timerClicker.Start();

            this.KeyPreview = true;
        }

        private void Options_Load(object sender, EventArgs e)
        {
            frm = this.Owner as Form1;

            dataGridView1.DataSource = dt.DefaultView;
            XMLFileName = frm.XMLFileName;
            

            textBox9.Text = ps.url;
            textBox1.Text = ps.w1.ToString();
            textBox2.Text = ps.w2.ToString();
            textBox4.Text = ps.w3.ToString();
            textBox6.Text = ps.medorg;
            textBox7.Text = ps.medotd;

        
            textBox12.Text = ps.bro_version;
            checkBox1.Checked = ps.status;
                    


            label4.Text = ps.login_text;


            dt.Columns.Add(new DataColumn("org", typeof(System.String)));
            dt.Columns.Add(new DataColumn("login", typeof(System.String)));
            dt.Columns.Add(new DataColumn("pass", typeof(System.String)));

            if (File.Exists(XMLFileName))
            {

                try
                {
                    using (Stream stream = new FileStream(XMLFileName, FileMode.Open, FileAccess.Read))
                    {
                        dt.ReadXml(stream);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }

            }


            dataGridView1.DataSource = dt.DefaultView;


            dataGridView1.Columns[0].HeaderText = "Организация";
            dataGridView1.Columns[1].HeaderText = "Логин";
            dataGridView1.Columns[2].HeaderText = "Пароль";

            dataGridView1.AutoResizeColumns();
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;


            foreach (DataGridViewRow dr in dataGridView1.Rows)
            {
                if (dr.Cells[1].Value != null)
                {
                    string iin = dr.Cells[1].Value.ToString();

                    if (iin == ps.login_text)
                        dataGridView1.CurrentCell = dr.Cells[1];
                }

            }


        }

        void timerClicker_Tick(object sender, EventArgs e)
        {
            //Вызов импортируемой функции с текущей позиции курсора
            uint X = (uint)Cursor.Position.X;
            uint Y = (uint)Cursor.Position.Y;

            textBox11.Text = X.ToString();
            textBox10.Text = Y.ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
        
            Properties.Settings.Default.Save();
            MessageBox.Show("Сохранено");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default["w1"] = Convert.ToDouble(textBox1.Text);
            Properties.Settings.Default["w2"] = Convert.ToDouble(textBox2.Text);
            Properties.Settings.Default["w3"] = Convert.ToDouble(textBox4.Text);
            Properties.Settings.Default["medorg"] = textBox6.Text;
            Properties.Settings.Default["medotd"] = textBox7.Text;
            Properties.Settings.Default["bro_version"] = textBox12.Text;
            Properties.Settings.Default["status"] = checkBox1.Checked;
            Properties.Settings.Default["url"] = textBox9.Text;


            Properties.Settings.Default.Save();
            MessageBox.Show("Сохранено");

            this.Close();
        }


        private void Options_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
               
                try
                {
                    Control ctl;
                    ctl = (Control)sender;
                    var check_txt =  groupBox2.Controls.OfType<TextBox>().FirstOrDefault(r => r.Focused);
                    //var check_txt_1 =  groupBox3.Controls.OfType<TextBox>().FirstOrDefault(r => r.Focused);
                    if (check_txt != null)
                    {
                        check_txt.Text = Cursor.Position.X.ToString() + ":" + Cursor.Position.Y.ToString();
                        ctl.SelectNextControl(ActiveControl, true, true, true, true);
                    }
                   
                    else MessageBox.Show("Выделите область для вставки координат!");
                }
                catch
                {
                    MessageBox.Show("Error Textbox");
                }
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)  frm.statusStrip2.Visible = true;
            else frm.statusStrip2.Visible = false;
        }

        private void button2_Click_1(object sender, EventArgs e)
        {

            save_xml(XMLFileName);

           // Properties.Settings.Default.Save();
            MessageBox.Show("Сохранено");

            this.Close();
           
        }



        //Сохранение в файл XML
        public void save_xml(string path)
        {
            XmlWriterSettings settings = new XmlWriterSettings();
            settings.Encoding = Encoding.Default;
            settings.OmitXmlDeclaration = false;
            settings.Indent = true;
            settings.IndentChars = "  ";
            settings.NewLineOnAttributes = true;
            try
            {
                using (XmlWriter xmlWriter = XmlWriter.Create(path, settings))
                {
                    xmlWriter.WriteStartElement("zl_list");

                    foreach (DataRow dataRow in dt.Rows)
                    {
                        xmlWriter.WriteStartElement("zl");

                        foreach (DataColumn dataColumn in dt.Columns)
                        {
                            if (dataColumn.ColumnName != "v")
                                xmlWriter.WriteElementString(dataColumn.ColumnName, dataRow[dataColumn].ToString());
                        }

                        xmlWriter.WriteEndElement();
                    }


                    xmlWriter.WriteEndElement();

                    xmlWriter.Close();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }

        }


        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox19_TextChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {

            if (dataGridView1.SelectedRows.Count > 0)
            {
                Properties.Settings.Default["login_text"] = dataGridView1.SelectedRows[0].Cells[1].Value.ToString();
                Properties.Settings.Default["pass_text"] = dataGridView1.SelectedRows[0].Cells[2].Value.ToString();
                Properties.Settings.Default.Save();

                label4.Text = dataGridView1.SelectedRows[0].Cells[1].Value.ToString();

                MessageBox.Show("Сохранено");
            }
            else MessageBox.Show("Выберите профиль");
        }

        private void dataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
  
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }
    }
}
