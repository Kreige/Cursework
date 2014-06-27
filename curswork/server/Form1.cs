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
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
namespace server
{
  
    public partial class Form1 : Form
    {ServiceHost host = new ServiceHost(typeof(servmetod), new Uri("http://localhost:1020/TestService"));
        public Form1()
        {
            InitializeComponent();
            
            host.AddServiceEndpoint(typeof(serverbd), new BasicHttpBinding(), "");
            host.Open();

            #region Output dispatchers listening
            foreach (Uri uri in host.BaseAddresses)
            { Console.WriteLine("\t{0}", uri.ToString()); }
            Console.WriteLine();
            textBox1.Text="Count and list of listening :"+host.ChannelDispatchers.Count.ToString();
            foreach (System.ServiceModel.Dispatcher.ChannelDispatcher dispatcher in host.ChannelDispatchers)
            {
                textBox1.Text+=(dispatcher.Listener.Uri.ToString()+dispatcher.BindingName.ToString()+Environment.NewLine);
            }
           
         
           
             #endregion

        }

        private void button1_Click(object sender, EventArgs e)
        {
            host.Close();
        }
    



    }
    [System.Runtime.Serialization.DataContract]
    class servmetod:serverbd
    {
        bdkursachDataSet bdcurs = new bdkursachDataSet();
         bdkursachDataSetTableAdapters.TableAdapterManager tabmen = new bdkursachDataSetTableAdapters.TableAdapterManager();
         bdkursachDataSetTableAdapters.ProectTableAdapter tabpro = new bdkursachDataSetTableAdapters.ProectTableAdapter();
         bdkursachDataSetTableAdapters.sotrudnikiTableAdapter sotr = new bdkursachDataSetTableAdapters.sotrudnikiTableAdapter();
         DataSet ta;
         public DataSet senddata()
         {
             BinaryFormatter serializationFormatter = new BinaryFormatter();
             MemoryStream buffer = new MemoryStream();

             Dictionary<string,object> listdata = new Dictionary<string,object>();
            
           tabpro.Fill(bdcurs.Proect);
            listdata.Add("Proect", bdcurs.Proect);
            sotr.Fill(bdcurs.sotrudniki);
            listdata.Add("Sotrudniki", bdcurs.sotrudniki);
            serializationFormatter.Serialize(buffer, bdcurs.sotrudniki);
             
           //List<sotrudniki.
            ta=bdcurs.sotrudniki.DataSet;
           // buffer.Position = 0;
                                  return ta; }
       
    }
    [ServiceContract]
    interface serverbd
    {
        [OperationContract]
       DataSet senddata();
    }

    //   [DataContract]
    //public class MyObject
    //{

    //   DataTable tablic;
      
    //    [DataMember]
    //    public DataTable Tabl
    //    {
    //        set { tablic = value; }
    //        get { return tablic; }
    //    }
       

    //    public MyObject(DataTable tabl)
    //    {
    //        tablic =tabl;
           
    //    }
    //     public DataTable dat()
    //    { return Tabl; }
        
    //}
      

   
}
