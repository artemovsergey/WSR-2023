using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data;
using MySql.Data.MySqlClient;

namespace Lopushok

{ 
    public partial class Products : Form
    {
        public Products()
        {
           InitializeComponent();
        }

        public void Form1_Load(object sender, EventArgs e)
        {
            //LoadData();

            string mysql = @"server=localhost;userid=root;password=root;database=test_data";
            MySqlConnection connect = new MySqlConnection(mysql);
            connect.Open();
            
            // MessageBox.Show("Подключение прошло успешно!");

            // Загрузка таблицы данными:

            DataSet ds = new DataSet();
            MySqlCommand cmd = connect.CreateCommand();
            cmd.CommandText = "select * from materials";

            MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
            adapter.Fill(ds);

            DataView dw = new DataView(ds.Tables[0]);

            //dw.Sort = "material";

            dw.Sort = comboBox1.Text;
            dw.RowFilter = comboBox2.Text;

            dataGridView1.DataSource = dw;
            adapter.Update(ds.Tables[0]);
            

            connect.Close();

        }

        private void LoadData()
        {
            try
            {
                // create connection
                string mysql = @"server=localhost;userid=root;password=root;database=test_data";
                MySqlConnection connect = new MySqlConnection(mysql);
                connect.Open();

                //MessageBox.Show("Подключение прошло успешно!");
                connect.Close();

            }

            catch
            {
                //MessageBox.Show("Не удалось подключиться!");
            }

            finally
            {
                //MessageBox.Show("Сеанс работы завершен");    
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Materials materials = new Materials();
            materials.Show();
        }

        private void button5_Click(object sender, EventArgs e)
        {

        }
    }
}
