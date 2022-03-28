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
using System.Globalization;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        private DateTime ParseDate(string date)
        {
            DateTime convertedDate;

            if (!DateTime.TryParseExact(date, "ddMMyyyy", new CultureInfo("en-US"),
                 DateTimeStyles.None, out convertedDate))

                throw new FormatException(string.Format("Unable to format date:{0}", date));

            return convertedDate;

        }
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            OleDbConnectionStringBuilder excelAyar = new OleDbConnectionStringBuilder();
            excelAyar.DataSource = @"D:\KOPYA_AFM_OE.xlsx"; // excel kitabının tam yol adı
            excelAyar.Provider = "Microsoft.ACE.OLEDB.12.0";
            excelAyar["Extended Properties"] = "Excel 12.0 Xml;HDR=YES";

            string excelSayfaAdi = "PLAN"; // verileri alacağınız Excel sayfasının adı

            OleDbConnection excelBag = new OleDbConnection(excelAyar.ConnectionString); excelBag.Open();
            OleDbDataAdapter adap = new OleDbDataAdapter("select * from [" + excelSayfaAdi + "$]", excelBag);
            
            DataTable dt = new DataTable(); 
            adap.Fill(dt);
            dataGridView1.DataSource = dt;
            int ThisYear = DateTime.Now.Year;
            int DayNumber = DateTime.Now.DayOfYear;
            int Bugun = DayNumber + 275;
            int satir = (DayNumber * 3) + 32;
            satir = satir + 78;
            int cycle;
            // cycle = Convert.ToInt32(dt.Rows[satir][0]);
            // textBox1.Text = cycle.ToString();
            /*
            var rows = dt.Select($"Tarih = {gun}");
            // Assuming that you have only one row with  valueToSearchFor,
            // otherwise you need a loop over the rows array
            if (rows.Length == 1)
            {
                Gun1Hedef = Convert.ToDouble(rows[0]["Pres1Hedef"].ToString()) + Convert.ToDouble(rows[0]["Pres2Hedef"].ToString()) + Convert.ToDouble(rows[0]["Pres3Hedef"].ToString()) + Convert.ToDouble(rows[0]["Pres4Hedef"].ToString());
            }
            textBox1.Text = Convert.ToInt32(Gun1Hedef).ToString();


            // 28774
            rows = dt.Select($"Tarih = {Gun2.ToString()}");

            // Assuming that you have only one row with  valueToSearchFor,
            // otherwise you need a loop over the rows array
            if (rows.Length == 1)
            {
                Gun2Hedef = Convert.ToDouble(rows[0]["Pres1Hedef"].ToString()) + Convert.ToDouble(rows[0]["Pres2Hedef"].ToString()) + Convert.ToDouble(rows[0]["Pres3Hedef"].ToString()) + Convert.ToDouble(rows[0]["Pres4Hedef"].ToString());
            }
            textBox2.Text = Convert.ToInt32(Gun2Hedef).ToString();

            rows = dt.Select($"Tarih = {Gun3.ToString()}");

            // Assuming that you have only one row with  valueToSearchFor,
            // otherwise you need a loop over the rows array
            if (rows.Length == 1)
            {
                Gun3Hedef = Convert.ToDouble(rows[0]["Pres1Hedef"].ToString()) + Convert.ToDouble(rows[0]["Pres2Hedef"].ToString()) + Convert.ToDouble(rows[0]["Pres3Hedef"].ToString()) + Convert.ToDouble(rows[0]["Pres4Hedef"].ToString());
            }
            textBox3.Text = Convert.ToInt32(Gun3Hedef).ToString();

            rows = dt.Select($"Tarih = {Gun4.ToString()}");

            // Assuming that you have only one row with  valueToSearchFor,
            // otherwise you need a loop over the rows array
            if (rows.Length == 1)
            {
                Gun4Hedef = Convert.ToDouble(rows[0]["Pres1Hedef"].ToString()) + Convert.ToDouble(rows[0]["Pres2Hedef"].ToString()) + Convert.ToDouble(rows[0]["Pres3Hedef"].ToString()) + Convert.ToDouble(rows[0]["Pres4Hedef"].ToString());
            }
            textBox4.Text = Convert.ToInt32(Gun4Hedef).ToString();

            rows = dt.Select($"Tarih = {Gun5.ToString()}");

            // Assuming that you have only one row with  valueToSearchFor,
            // otherwise you need a loop over the rows array
            if (rows.Length == 1)
            {
                Gun5Hedef = Convert.ToDouble(rows[0]["Pres1Hedef"].ToString()) + Convert.ToDouble(rows[0]["Pres2Hedef"].ToString()) + Convert.ToDouble(rows[0]["Pres3Hedef"].ToString()) + Convert.ToDouble(rows[0]["Pres4Hedef"].ToString());
            }
            textBox5.Text = Convert.ToInt32(Gun5Hedef).ToString();

            rows = dt.Select($"Tarih = {Gun6.ToString()}");
            if (rows.Length == 1)
            {
                Gun6Hedef = Convert.ToDouble(rows[0]["Pres1Hedef"].ToString()) + Convert.ToDouble(rows[0]["Pres2Hedef"].ToString()) + Convert.ToDouble(rows[0]["Pres3Hedef"].ToString()) + Convert.ToDouble(rows[0]["Pres4Hedef"].ToString());
            }

            textBox6.Text = Convert.ToInt32(Gun6Hedef).ToString();
            rows = dt.Select($"Tarih = {Gun7.ToString()}");
            if (rows.Length == 1)
            {
                Gun7Hedef = Convert.ToDouble(rows[0]["Pres1Hedef"].ToString()) + Convert.ToDouble(rows[0]["Pres2Hedef"].ToString()) + Convert.ToDouble(rows[0]["Pres3Hedef"].ToString()) + Convert.ToDouble(rows[0]["Pres4Hedef"].ToString());
            }

            textBox7.Text = Convert.ToInt32(Gun7Hedef).ToString();*/
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
}
