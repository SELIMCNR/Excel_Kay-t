using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace excel_test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        //C:\Users\btk02\OneDrive\Desktop

        OleDbConnection bgl = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source= C:\Users\btk02\OneDrive\Desktop\CinarYazilim.xlsx;Extend Properties='Excel 12.0 Xml;HDR=YES;'");

        void listele()
        {
            bgl.Open();
            OleDbDataAdapter da = new OleDbDataAdapter("Select * From [Sayfa1$]",bgl);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            bgl.Close();
        }

        private void btnListele_Click(object sender, EventArgs e)
        {
            listele();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            listele();
        }

        private void btnKaydet_Click(object sender, EventArgs e)
        {
            bgl.Open();
            OleDbCommand kmt = new OleDbCommand("insert into [Sayfa1$] (Saat,Ders) values (@p1,@p2)",bgl);
            kmt.Parameters.AddWithValue("@p1", TxtSaat.Text);
            kmt.Parameters.AddWithValue("@p2", TxtDers.Text);
            kmt.ExecuteNonQuery();
            bgl.Close();
            MessageBox.Show("Yeni Ders Bilgisi Eklendi");
            listele();

        }
    }
}
