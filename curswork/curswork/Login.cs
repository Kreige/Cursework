using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace curswork
{
    public partial class Login : Form
    {
       
      
        BindingSource binso = new BindingSource();
        bdkursachDataSetTableAdapters.sotrudnikiTableAdapter sotr = new bdkursachDataSetTableAdapters.sotrudnikiTableAdapter();
        bdkursachDataSet.sotrudnikiDataTable sot = new bdkursachDataSet.sotrudnikiDataTable();
        
      
        public Login()
        {
            
            InitializeComponent();
            sotr.Fill(sot);
        }

        private void button1_Click(object sender, EventArgs e)
        {binso.DataSource=sot;
        
        //MessageBox.Show(sot.Rows[binso.Find("Login", textBox1.Text)]["pass"].ToString());
        try
        {
            if (sot.Rows[binso.Find("Логин", textBox1.Text)]["Пароль"].ToString().Contains(textBox2.Text))
            {
                MessageBox.Show("ok");
                Form1 f1 = new Form1(binso.Find("Логин", textBox1.Text));
                // f1.ShowDialog();
                this.Hide();
                f1.Show();
                //  Close();
            }
        }
        catch { MessageBox.Show("Логин и/или пароль не существуют"); }
       
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
