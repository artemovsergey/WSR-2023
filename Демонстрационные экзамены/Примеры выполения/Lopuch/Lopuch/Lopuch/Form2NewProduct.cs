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
    public partial class Form2NewProduct : Form
    {
        //Передаем объект модели
        public Model1 db { get; set; }

        public Form2NewProduct()
        {
            InitializeComponent();
        }

        private void Form2NewProduct_Load(object sender, EventArgs e)
        {

        }

        private void productionPersonCountLabel_Click(object sender, EventArgs e)
        {

        }

        private void productionWorkshopNumberLabel_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Product p1 = new Product();
            p1.ArticleNumber = maskedTextBox2.Text;

            db.Product.Add(p1);
            db.SaveChanges();
        }
    }
}
    
