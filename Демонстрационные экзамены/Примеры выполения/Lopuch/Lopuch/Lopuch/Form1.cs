using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Lopuch
{
    public partial class Form1 : Form
    {
        // создаем объект класса контекста Model1
        Model1 db = new Model1();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // передаем промежуточному объекту данные из модели
            productBindingSource.DataSource = db.Product.ToList();
        }

        // обработчик собятия для кнопки "Поиск"
        private void button1_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                dataGridView1.Rows[i].Selected = false;
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    if (dataGridView1.Rows[i].Cells[j].Value != null)
                        if (dataGridView1.Rows[i].Cells[j].Value.ToString().Contains(textBox1.Text))
                        {
                            dataGridView1.Rows[i].Selected = true;
                            break;
                        }
            }
        }

        private void button1_KeyPress(object sender, KeyPressEventArgs e)
        {           
        }

        // обработчик собятия для кнопки "Показать исходные данныне"
        private void button3_Click(object sender, EventArgs e)
        {
            List<Product> prod = (List<Product>)productBindingSource.List;
            
        }

        // обработчик собятия для кнопки "Добавить продукт"
        private void button4_Click(object sender, EventArgs e)
        {

            Form2NewProduct fnp = new Form2NewProduct();
            fnp.db = db;
            fnp.Show();

        }
        
        //Обработчик события для кнопки "Выйти"
        private void button7_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void р_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Запись удалена");
        }
    }
}
