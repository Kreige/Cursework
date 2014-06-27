using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.ServiceModel;
using System.Runtime.Serialization;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Windows.Forms.DataVisualization.Charting;
using Excel = Microsoft.Office.Interop.Excel;

namespace curswork
{

    [ServiceContract]
    interface serverbd
    {
        [OperationContract]
        DataSet senddata();
    }
    public partial class Form1 : Form
    {
        bool statpd=false;
        int rowred;
        public bdkursachDataSetTableAdapters.sotrudnikiTableAdapter tas = new bdkursachDataSetTableAdapters.sotrudnikiTableAdapter();
      public bdkursachDataSetTableAdapters.ProectTableAdapter tab = new bdkursachDataSetTableAdapters.ProectTableAdapter();
      public bdkursachDataSetTableAdapters.zadachiTableAdapter zadtb = new bdkursachDataSetTableAdapters.zadachiTableAdapter();
      public bdkursachDataSetTableAdapters.plantimeTableAdapter palttb = new bdkursachDataSetTableAdapters.plantimeTableAdapter();
      BindingSource binstr = new BindingSource();
     BindingSource binpro=new BindingSource();
     BindingSource binzad=new BindingSource ();
     BindingSource binmet = new BindingSource();
      public bdkursachDataSet.sotrudnikiDataTable str = new bdkursachDataSet.sotrudnikiDataTable();
      public bdkursachDataSet.ProectDataTable pt = new bdkursachDataSet.ProectDataTable();
      public bdkursachDataSet.zadachiDataTable zaddt = new bdkursachDataSet.zadachiDataTable();
      public bdkursachDataSet.plantimeDataTable plandt = new bdkursachDataSet.plantimeDataTable();
   
            Uri tcpUri = new Uri("http://localhost:1020/TestService");
            EndpointAddress address;
            BasicHttpBinding binding;
            ChannelFactory<serverbd> factory;
            serverbd service;
            public Form1(int idsot)
            {
                InitializeComponent();
               
                tab.Fill(pt);
                tas.Fill(str);
                palttb.Fill(plandt);
                zadtb.Fill(zaddt);
              
                dataGridView1.DataSource = zaddt;
                textBox1.Text = str.Rows[idsot]["Имя"].ToString() + str.Rows[idsot]["Фамилия"].ToString();
               
               
                textBox2.Text = str.Rows[idsot]["Статус"].ToString();
                binstr.DataSource =str;
                binmet.DataSource = plandt;
                binpro.DataSource = pt;
                binzad.DataSource = zaddt;
                
                for(int t=0; t<plandt.Rows.Count; t++)
                {
                    cbmetrasch.Items.Add(plandt.Rows[t][1].ToString());
                }
                for (int t = 0; t < pt.Rows.Count; t++)
                {
                    tbishproj.Items.Add(pt.Rows[t][1]);
                    projlist.Items.Add(pt.Rows[t][1]);
                }
                for(int t=0; t<str.Rows.Count; t++)
                {
                    ispolcom.Items.Add(str.Rows[t][1] + " " + str.Rows[t][2]);
                }
                timer1.Interval = 30000;
                timer1.Start();
                chart2.Legends.Add("Выполнено заданий");
            }
            

        private void button1_Click(object sender, EventArgs e)
        {var serializationFormatter = new BinaryFormatter();
            object t = new object();
            t = pt;
             }
      
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            tab.Update(pt);
            tas.Update(str);
            palttb.Update(plandt);
            zadtb.Update(zaddt);
            Application.Exit();
        }

       

        private void Form1_Load(object sender, EventArgs e)
        {
            address = new EndpointAddress(tcpUri);
            binding = new BasicHttpBinding();
            factory = new ChannelFactory<serverbd>(binding, address);
            service = factory.CreateChannel();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = pt;
            button1.Text = "Добавить проект";
            button2.Text = "Удалить проект";
        }

      
       private void выход_Click_1(object sender, EventArgs e)
       {
           binzad.EndEdit();
           binstr.EndEdit();
           binpro.EndEdit();
           binmet.EndEdit();
           tab.Update(pt);
           tas.Update(str);
           palttb.Update(plandt);
           zadtb.Update(zaddt);
           

           
           Close();
       }

      
       private void toolStripSplitButton1_Click(object sender, EventArgs e)
       {
           this.Text = toolStripSplitButton1.Text;
           панельпоиск.Visible = true;
           dataGridView1.DataSource = str;
           button1.Text = "Добавить сотрудника";
           button2.Text = "Удалить сотрудника";
          


       }

       private void toolStripButton2_Click(object sender, EventArgs e)
       {
           this.Text = toolStripButton2.Text;
           панельпоиск.Visible = false;
           dataGridView1.DataSource = zaddt;
           button1.Text = "Добавить задачу";
           button2.Text = "Удалить задачу";
           obproc();
               
       }

       private void toolStripButton3_Click(object sender, EventArgs e)
       {
           this.Text = toolStripButton3.Text;
           панельпоиск.Visible = false;
           dataGridView1.DataSource = plandt;
           button1.Text = "Добавить метод";
           button2.Text = "Удалить метод";
          
          
       }

       private void Проекты_ButtonClick(object sender, EventArgs e)
       {
           this.Text = Проекты.Text;
           панельпоиск.Visible = false;
       }
       private void button1_Click_1(object sender, EventArgs e)
       {
           панельпоиск.Visible = false;
           if (textBox2.Text.Contains("admin"))
           {
               if (dataGridView1.DataSource.ToString().Contains("zadachi"))
               {
                   statpd = false;
                   дорезад.Visible = true;

               }
               if (dataGridView1.DataSource.ToString().Contains("Proect"))
               {
                   adredproj.Visible = true;
                   statpd = false;

               }
               if (dataGridView1.DataSource.ToString().Contains("plantime"))
               {
                   statpd = false;
                   дорерасчср.Visible = true; 
               }
               if (dataGridView1.DataSource.ToString().Contains("sotrudniki"))
               {
                   adresotr.Visible = true;
                  
                  statpd = false;
               }
           }
           else
           { MessageBox.Show("Вы не обладаете правами для данного действия"); }
       }

       private void button2_Click_1(object sender, EventArgs e)
       {
           if (textBox2.Text.Contains("admin"))
           {
               if (dataGridView1.DataSource.ToString().Contains("zadachi"))
               {
                   pt.Rows[tbishproj.Items.IndexOf(zaddt.Rows[dataGridView1.CurrentRow.Index][3].ToString()) + 1][2] = Convert.ToInt16(pt.Rows[tbishproj.Items.IndexOf(zaddt.Rows[dataGridView1.CurrentRow.Index][3]) + 1][2]) - 1;
                   foreach (var r in str)
                   {
                       if (r.Задача.Contains(zaddt.Rows[dataGridView1.CurrentRow.Index][1].ToString()))
                       { r.Задача = ""; }
                   }
                   binzad.RemoveAt(dataGridView1.CurrentRow.Index);
                   zadtb.Update(zaddt);

               }
               if (dataGridView1.DataSource.ToString().Contains("Proect"))
               {
                   
                  
                   for (int v = 0; v < zaddt.Rows.Count; v++ )
                       {
                           if (zaddt.Rows[v][2].ToString().Contains(pt.Rows[dataGridView1.CurrentRow.Index][1].ToString()))
                               zaddt.Rows[v].Delete();
                       }
                   binpro.RemoveAt(dataGridView1.CurrentRow.Index);
 
               }
               if (dataGridView1.DataSource.ToString().Contains("plantime"))
               {cbmetrasch.Items.RemoveAt(dataGridView1.CurrentRow.Index);
                   binmet.RemoveAt(dataGridView1.CurrentRow.Index);
                   
             
               }
               if (dataGridView1.DataSource.ToString().Contains("sotrudniki"))
               {
                   foreach (var v in zaddt)
                   {
                       if (v.Исполнитель.Contains(str.Rows[dataGridView1.CurrentRow.Index][1] + " " + str.Rows[dataGridView1.CurrentRow.Index][2]))
                       { v.Исполнитель = ""; }
                       ispolcom.Items.Remove(ispolcom.Items[dataGridView1.CurrentRow.Index]);
                      
                   }
                   binstr.RemoveAt(dataGridView1.CurrentRow.Index);

               }
              
           }
           else
           { MessageBox.Show("Вы не обладаете правами для данного действия"); }
       }

       private void button3_Click_1(object sender, EventArgs e)
       {
           панельпоиск.Visible = false;
           if (textBox2.Text.Contains("admin"))
           {
               if (dataGridView1.DataSource.ToString().Contains("zadachi"))
               {
                   statpd = true;
                   дорезад.Visible = true;
                   rowred = dataGridView1.CurrentRow.Index;
                   tbzadname.Text = zaddt.Rows[rowred][1].ToString();
                   tbishproj.Text = zaddt.Rows[rowred][2].ToString();
                   cbzadsost.Text = zaddt.Rows[rowred][3].ToString();
                   tbobrab.Text = zaddt.Rows[rowred][4].ToString();


               }
               if (dataGridView1.DataSource.ToString().Contains("Proect"))
               {
                   statpd = true;
                   adredproj.Visible = true;
                   tbnamepro.Text = pt.Rows[dataGridView1.CurrentRow.Index][1].ToString();
                   cbstatpro.Text = pt.Rows[dataGridView1.CurrentRow.Index][5].ToString();
               }
               if (dataGridView1.DataSource.ToString().Contains("plantime"))
               {
                   дорерасчср.Visible = true;
                   tbnamemet.Text = plandt.Rows[dataGridView1.CurrentRow.Index][1].ToString();
                   tbedrasch.Text = plandt.Rows[dataGridView1.CurrentRow.Index][2].ToString();
                   tbmnoj.Text = plandt.Rows[dataGridView1.CurrentRow.Index][3].ToString();
                   statpd = true;
                   rowred = dataGridView1.CurrentRow.Index;
               }
               if (dataGridView1.DataSource.ToString().Contains("sotrudniki"))
               {
                   adresotr.Visible = true;
                   statpd = true;
                   tbnamesotr.Text = str.Rows[dataGridView1.CurrentRow.Index][1].ToString();
                   tbfamsotr.Text = str.Rows[dataGridView1.CurrentRow.Index][2].ToString();
                  
               }
           }
           else
           { MessageBox.Show("Вы не обладаете правами для данного действия"); }
       }

       


       #region Методы расчета др
       private void button5_Click(object sender, EventArgs e)
       {
           try
           {
               if (statpd == false)
               {
                   plandt.AddplantimeRow(plandt.Rows.Count + 1, tbnamemet.Text, tbedrasch.Text, Convert.ToInt16(tbmnoj.Text));
                   cbmetrasch.Items.Add(tbnamemet.Text);
               }
               else
               {
                   cbmetrasch.Items[rowred] = tbnamemet.Text;
                   plandt.Rows[rowred][1] = tbnamemet.Text;
                   plandt.Rows[rowred][2] = tbedrasch.Text;
                   plandt.Rows[rowred][3] = Convert.ToInt16(tbmnoj.Text);
               }
               tbnamemet.Clear();
               tbedrasch.Clear();
               tbmnoj.Clear();
               дорерасчср.Visible = false;
           }
           catch { MessageBox.Show("Не все поля заполнены или введен неправильный тип данных"); }
       }

       private void button6_Click(object sender, EventArgs e)
       {
           дорерасчср.Visible = false;
       }
       #endregion
       #region Задачи др
       private void button4_Click_1(object sender, EventArgs e)
       {
           try
           {
               if (statpd == false)
               {
                  zaddt.AddzadachiRow(zaddt.Rows.Count + 1, tbzadname.Text, tbishproj.Text, ispolcom.Text, Convert.ToInt16(tbobrab.Text), DateTime.Now, DateTime.Now.AddMinutes(Convert.ToInt16(tbobrab.Text)), 0, cbzadsost.Text);

                  if (tbishproj.SelectedIndex!=0)
                  { 
                      pt.Rows[tbishproj.SelectedIndex][3] = Convert.ToInt16(pt.Rows[tbishproj.SelectedIndex-1][3]) + 1;
                   pt.Rows[tbishproj.SelectedIndex][4] = Convert.ToInt16(tbobrab.Text) + Convert.ToInt16(pt.Rows[tbishproj.SelectedIndex][3]);
                  }
                      str.Rows[ispolcom.SelectedIndex][4] = tbzadname.Text;
               }
               else
               {
                   zaddt.Rows[rowred][1] = tbzadname.Text;
                   zaddt.Rows[rowred][2] = tbishproj.Text;
                   zaddt.Rows[rowred][3] = cbzadsost.Text;
                   zaddt.Rows[rowred][4] = Convert.ToInt16(tbobrab.Text);

               }

               дорезад.Visible = false;
               tbzadname.Clear();
               tbishproj.Text = "";
               cbzadsost.Text = "";
               tbobrab.Clear();
           }
           catch { MessageBox.Show("Не все поля заполнены или введен неправильный тип данных"); }
       }


       private void button7_Click(object sender, EventArgs e)
       {
           дорезад.Visible = false;

       }

       private void cbmetrasch_SelectedIndexChanged(object sender, EventArgs e)
       {
           label12.Text=plandt.Rows[cbmetrasch.SelectedIndex][2].ToString();
           if(tbcoled.Text!="")
           tbobrab.Text = (Convert.ToInt16(plandt.Rows[cbmetrasch.SelectedIndex][3]) * Convert.ToInt16(tbcoled.Text)).ToString();
       }

       private void tbcoled_TextChanged(object sender, EventArgs e)
       {
           if (tbcoled.Text != "" & cbmetrasch.Text!="")
           tbobrab.Text = (Convert.ToInt16(plandt.Rows[cbmetrasch.SelectedIndex][3]) * Convert.ToInt16(tbcoled.Text)).ToString();
       }
       #endregion
       #region Сотрудники др
       private void button8_Click(object sender, EventArgs e)
       {
           try
           {
               if (statpd == false)
               {
                   str.AddsotrudnikiRow(str.Rows.Count + 1, tbnamesotr.Text, tbfamsotr.Text, 100, "", tbnamesotr.Text, tbfamsotr.Text, "Работник");
                   ispolcom.Items.Add(tbnamesotr.Text + " " + tbfamsotr.Text);
               }
               else
               {
                   str.Rows[dataGridView1.CurrentRow.Index][1] = tbnamesotr.Text;
                   str.Rows[dataGridView1.CurrentRow.Index][2] = tbfamsotr.Text;

                   ispolcom.Items[dataGridView1.CurrentRow.Index] = tbnamesotr.Text + " " + tbfamsotr.Text;
               }

               adresotr.Visible = false;
               панельпоиск.Visible = true;
           }
           catch { MessageBox.Show("Не все поля заполнены или введен неправильный тип данных"); }
       }

       private void button9_Click(object sender, EventArgs e)
       {
           adresotr.Visible = false;
           панельпоиск.Visible = true;
       }

       private void button13_Click(object sender, EventArgs e)
       {

           string filt = "";
           int t = 0;

           if (tbid.Text != "")
           {
               filt += "id=" + tbid.Text;
               t++;
           }

           if (imyatb.Text != "")
           {
               if (t != 0)
               { filt += " and "; }
               filt += " Имя='" + imyatb.Text + "'";
               t++;
           }
           if (famtb.Text != "")
           {
               if (t != 0)
               {
                   filt += " and ";
               }
               filt += "Фамилия='" + famtb.Text + "'";
               t++;
           }
           if (tbzad.Text != "")
           {
               if (t != 0)
               {
                   filt += " and ";
               }
               filt += "Задача='" + tbzad.Text + "'";
               t++;
           }
           if (zpmintb.Text != "")
           {
               if (t != 0)
               { filt += " and "; }
               filt += "Заплата >= " + zpmintb.Text;
               t++;
           }
           if (zpmaxtb.Text != "")
           {
               if (t != 0)
               { filt += " and "; }
               filt += "Заплата<" + zpmaxtb.Text;
               t++;
           }

           binstr.Filter = filt;
           dataGridView1.DataSource = binstr;
       }

       private void button14_Click(object sender, EventArgs e)
       {
           binstr.RemoveFilter();
       }

      
#endregion
       #region Проекты др
       private void button10_Click(object sender, EventArgs e)
       {try
       {
           if (statpd == false)
           {
               pt.AddProectRow(pt.Rows.Count+1, tbnamepro.Text, 0, 0, 0, cbstatpro.Text);
               tbishproj.Items.Add(tbnamepro.Text);
               projlist.Items.Add(tbnamepro.Text);
               
           }
           else
           {
              pt.Rows[dataGridView1.CurrentRow.Index][1]= tbnamepro.Text;
              pt.Rows[dataGridView1.CurrentRow.Index][5]= cbstatpro.Text;
           }
           adredproj.Visible = false;
       }
           catch { MessageBox.Show("Не все поля заполнены или введен неправильный тип данных"); }
       }

       private void button11_Click(object sender, EventArgs e)
       {
           adredproj.Visible = false;
       }
        #endregion
       #region Процент выполнения
       private void timer1_Tick(object sender, EventArgs e)
       {
           obproc();
       }
        
        public void obproc()
       {
           
            if(dataGridView1.DataSource.ToString().Contains("zadachi"))
       {
          
           foreach (var r in zaddt)
           {
               if (r.Статус.Contains("Завершена") == false)
               {
                   if (DateTime.Now < r.Дата_завершения)
                   {
                       r.Процент_выполнения = Convert.ToInt16((Convert.ToDouble((DateTime.Now - r.Дата_старта).Minutes) / Convert.ToDouble(r.Объем_работы)) * 100);
                     
                   }
                   else
                   {
                       r.Процент_выполнения = Convert.ToInt16((Convert.ToDouble((DateTime.Now - r.Дата_старта).Minutes) / Convert.ToDouble(r.Объем_работы)) * 100);
                     
                       r.Статус = "Завершена";
                      
                   }

               }
        }
       }
       }
#endregion
       #region Графики
       private void графикРеализацииToolStripMenuItem_Click(object sender, EventArgs e)
       {
           панельпоиск.Visible = false;
           панельграф.Visible = true;
       }

       private void button12_Click(object sender, EventArgs e)
       {


           chart2.Series.Clear();
           chart2.Titles.Clear();
           chart2.Titles.Add(projlist.Text);
           chart2.ChartAreas.FindByName("ChartArea1").AxisY.Title = "Ось Y";


          

           chart2.ChartAreas.FindByName("ChartArea1").AxisX.Title = "Ось X";
         chart2.Series.Add("Выполнено заданий");
chart2.Series.Add("Осталось заданий");

chart2.Series["Выполнено заданий"].XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.String;
chart2.Series["Выполнено заданий"].YValueType=System.Windows.Forms.DataVisualization.Charting.ChartValueType.Time;
chart2.Series["Осталось заданий"].XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.String;
chart2.Series["Осталось заданий"].YValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.Time;


           foreach (var v in zaddt)
           {
               if (v.Проект == projlist.Text)
               {

                   if (v.Процент_выполнения != 100)
                   {
                       chart2.Series["Выполнено заданий"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.RangeColumn;
                       chart2.Series[0].Points.AddXY(v.Название_задачи, v.Дата_старта.ToOADate(), DateTime.Now.ToOADate());

                       chart2.Series["Осталось заданий"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.RangeColumn;
                       chart2.Series[1].Points.AddXY(v.Название_задачи, DateTime.Now.ToOADate(), v.Дата_завершения.ToOADate());
                   }
                   else
                   {
                       chart2.Series["Выполнено заданий"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.RangeColumn;
                       chart2.Series[0].Points.AddXY(v.Название_задачи, v.Дата_старта.ToOADate(), v.Дата_завершения.ToOADate());
                   }
               }
           }
       }

       private void Закрыть_Click(object sender, EventArgs e)
       {
           панельграф.Visible = false;
       }
       #endregion

       private void button15_Click(object sender, EventArgs e)
       {
           new toExcel(zaddt, typedia.SelectedIndex, projlist.Text, typedia.Text);
       }

       
    }
   public class toExcel
   {public toExcel(DataTable vh,int index,string proj,string typediaa)
       {
           Excel.Application xlapp = new Excel.Application();
           xlapp.Visible = true;
           Excel.Workbook xlappwb = xlapp.Workbooks.Add(Type.Missing);
           Excel.Worksheet xlappwsh = (Excel.Worksheet)xlappwb.Worksheets.get_Item(1);
           object misValue = System.Reflection.Missing.Value;
           int colz = 1;
           for (int i = 0; i <= vh.Rows.Count - 1; i++)
           {
               if (vh.Rows[i][2].ToString().Contains(proj))
               {

                   if (index == 0)
                   {
                       xlappwsh.Cells[1,i+4] = vh.Rows[i][1];
                       xlappwsh.Cells[3, i + 4] = vh.Rows[i][7];
                       xlappwsh.Cells[4,i+4] =100-Convert.ToInt16(vh.Rows[i][7]);
                       colz++;
                   }
                   else
                   {
                       xlappwsh.Cells[1, i+4] =vh.Rows[i][1];
                       xlappwsh.Cells[3, i+4] = Convert.ToDateTime(vh.Rows[i][5]);
                       xlappwsh.Cells[4, i+4] = Convert.ToDateTime(vh.Rows[i][6]);
                       colz++;
                   
                   }
               }
               
           }
           xlappwsh.Cells[3, 3] = "Выполнено заданий";
           xlappwsh.Cells[4, 3] = "Всего надо выполнить";

           Excel.Range chartRange;
           Excel.ChartObjects xlCharts = (Excel.ChartObjects)xlappwsh.ChartObjects(Type.Missing);
           Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(0, 0, 400, 350);
           Excel.Chart chartPage = myChart.Chart;

           chartRange = xlappwsh.get_Range("C1", "H4");
           colz = 1;
           chartPage.SetSourceData(chartRange, misValue);
           chartPage.HasTitle = true;
           chartPage.ChartTitle.Text = proj + " " + typediaa;
           chartPage.ChartTitle.Font.Size = 13;
           chartPage.ChartTitle.Font.Color = 254;

           chartPage.ChartTitle.Shadow = true;
           chartPage.ChartTitle.Border.LineStyle = Excel.Constants.xlSolid;

          // chartPage.ChartType = Microsoft.Office.Interop.Excel.XlChartType.xlColumnStacked;
           xlappwb.Close(true, Type.Missing, false);
           xlapp.Quit();
       }
       }
   }

