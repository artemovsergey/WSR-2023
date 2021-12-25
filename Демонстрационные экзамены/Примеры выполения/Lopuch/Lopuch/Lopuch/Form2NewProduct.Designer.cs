
namespace Lopuch
{
    partial class Form2NewProduct
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.Label articleNumberLabel;
            System.Windows.Forms.Label minCostForAgentLabel;
            System.Windows.Forms.Label productionPersonCountLabel;
            System.Windows.Forms.Label productionWorkshopNumberLabel;
            System.Windows.Forms.Label productTypeIDLabel;
            System.Windows.Forms.Label titleLabel;
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form2NewProduct));
            this.titleTextBox = new System.Windows.Forms.TextBox();
            this.maskedTextBox1 = new System.Windows.Forms.MaskedTextBox();
            this.maskedTextBox2 = new System.Windows.Forms.MaskedTextBox();
            this.maskedTextBox3 = new System.Windows.Forms.MaskedTextBox();
            this.maskedTextBox4 = new System.Windows.Forms.MaskedTextBox();
            this.maskedTextBox5 = new System.Windows.Forms.MaskedTextBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label2 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.button7 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.productBindingSource = new System.Windows.Forms.BindingSource(this.components);
            articleNumberLabel = new System.Windows.Forms.Label();
            minCostForAgentLabel = new System.Windows.Forms.Label();
            productionPersonCountLabel = new System.Windows.Forms.Label();
            productionWorkshopNumberLabel = new System.Windows.Forms.Label();
            productTypeIDLabel = new System.Windows.Forms.Label();
            titleLabel = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.productBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // articleNumberLabel
            // 
            articleNumberLabel.AutoSize = true;
            articleNumberLabel.Location = new System.Drawing.Point(43, 97);
            articleNumberLabel.Name = "articleNumberLabel";
            articleNumberLabel.Size = new System.Drawing.Size(64, 29);
            articleNumberLabel.TabIndex = 1;
            articleNumberLabel.Text = "Артикль:";
            // 
            // minCostForAgentLabel
            // 
            minCostForAgentLabel.AutoSize = true;
            minCostForAgentLabel.Location = new System.Drawing.Point(42, 152);
            minCostForAgentLabel.Name = "minCostForAgentLabel";
            minCostForAgentLabel.Size = new System.Drawing.Size(151, 58);
            minCostForAgentLabel.TabIndex = 9;
            minCostForAgentLabel.Text = "Минимальная стоимость\r\nдля агента";
            // 
            // productionPersonCountLabel
            // 
            productionPersonCountLabel.AutoSize = true;
            productionPersonCountLabel.Location = new System.Drawing.Point(43, 231);
            productionPersonCountLabel.Name = "productionPersonCountLabel";
            productionPersonCountLabel.Size = new System.Drawing.Size(186, 58);
            productionPersonCountLabel.TabIndex = 11;
            productionPersonCountLabel.Text = "Количество производственного\r\nперсонала";
            productionPersonCountLabel.Click += new System.EventHandler(this.productionPersonCountLabel_Click);
            // 
            // productionWorkshopNumberLabel
            // 
            productionWorkshopNumberLabel.AutoSize = true;
            productionWorkshopNumberLabel.Location = new System.Drawing.Point(41, 314);
            productionWorkshopNumberLabel.Name = "productionWorkshopNumberLabel";
            productionWorkshopNumberLabel.Size = new System.Drawing.Size(75, 29);
            productionWorkshopNumberLabel.TabIndex = 13;
            productionWorkshopNumberLabel.Text = "Номер цеха";
            productionWorkshopNumberLabel.Click += new System.EventHandler(this.productionWorkshopNumberLabel_Click);
            // 
            // productTypeIDLabel
            // 
            productTypeIDLabel.AutoSize = true;
            productTypeIDLabel.Location = new System.Drawing.Point(40, 362);
            productTypeIDLabel.Name = "productTypeIDLabel";
            productTypeIDLabel.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            productTypeIDLabel.Size = new System.Drawing.Size(103, 29);
            productTypeIDLabel.TabIndex = 15;
            productTypeIDLabel.Text = "Тип продукта ID";
            // 
            // titleLabel
            // 
            titleLabel.AutoSize = true;
            titleLabel.Location = new System.Drawing.Point(40, 421);
            titleLabel.Name = "titleLabel";
            titleLabel.Size = new System.Drawing.Size(64, 29);
            titleLabel.TabIndex = 17;
            titleLabel.Text = "Название";
            // 
            // titleTextBox
            // 
            this.titleTextBox.DataBindings.Add(new System.Windows.Forms.Binding("Text", this.productBindingSource, "Title", true));
            this.titleTextBox.Location = new System.Drawing.Point(282, 430);
            this.titleTextBox.Name = "titleTextBox";
            this.titleTextBox.Size = new System.Drawing.Size(415, 35);
            this.titleTextBox.TabIndex = 18;
            // 
            // maskedTextBox1
            // 
            this.maskedTextBox1.Location = new System.Drawing.Point(282, 313);
            this.maskedTextBox1.Name = "maskedTextBox1";
            this.maskedTextBox1.Size = new System.Drawing.Size(417, 35);
            this.maskedTextBox1.TabIndex = 19;
            // 
            // maskedTextBox2
            // 
            this.maskedTextBox2.Location = new System.Drawing.Point(284, 96);
            this.maskedTextBox2.Name = "maskedTextBox2";
            this.maskedTextBox2.Size = new System.Drawing.Size(415, 35);
            this.maskedTextBox2.TabIndex = 20;
            // 
            // maskedTextBox3
            // 
            this.maskedTextBox3.Location = new System.Drawing.Point(284, 166);
            this.maskedTextBox3.Name = "maskedTextBox3";
            this.maskedTextBox3.Size = new System.Drawing.Size(413, 35);
            this.maskedTextBox3.TabIndex = 21;
            // 
            // maskedTextBox4
            // 
            this.maskedTextBox4.Location = new System.Drawing.Point(284, 239);
            this.maskedTextBox4.Name = "maskedTextBox4";
            this.maskedTextBox4.Size = new System.Drawing.Size(415, 35);
            this.maskedTextBox4.TabIndex = 22;
            // 
            // maskedTextBox5
            // 
            this.maskedTextBox5.Location = new System.Drawing.Point(282, 372);
            this.maskedTextBox5.Name = "maskedTextBox5";
            this.maskedTextBox5.Size = new System.Drawing.Size(417, 35);
            this.maskedTextBox5.TabIndex = 23;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.pictureBox1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(778, 72);
            this.panel1.TabIndex = 24;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.White;
            this.label2.Font = new System.Drawing.Font("Gabriola", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(321, 19);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(196, 39);
            this.label2.TabIndex = 2;
            this.label2.Text = "Добавить новый продукт";
            this.label2.Click += new System.EventHandler(this.label2_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Lopuch.Properties.Resources.Лопушок;
            this.pictureBox1.Location = new System.Drawing.Point(3, 2);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(68, 67);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 0;
            this.pictureBox1.TabStop = false;
            // 
            // button7
            // 
            this.button7.Location = new System.Drawing.Point(0, 0);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(75, 23);
            this.button7.TabIndex = 0;
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(206)))), ((int)(((byte)(255)))), ((int)(((byte)(249)))));
            this.button2.Location = new System.Drawing.Point(48, 490);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(149, 76);
            this.button2.TabIndex = 25;
            this.button2.Text = "Добавить продукт";
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(206)))), ((int)(((byte)(255)))), ((int)(((byte)(249)))));
            this.button1.Location = new System.Drawing.Point(521, 490);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(149, 76);
            this.button1.TabIndex = 26;
            this.button1.Text = "Выйти";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // productBindingSource
            // 
            this.productBindingSource.DataSource = typeof(Lopuch.Product);
            // 
            // Form2NewProduct
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 29F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(778, 578);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.button7);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.maskedTextBox5);
            this.Controls.Add(this.maskedTextBox4);
            this.Controls.Add(this.maskedTextBox3);
            this.Controls.Add(this.maskedTextBox2);
            this.Controls.Add(this.maskedTextBox1);
            this.Controls.Add(articleNumberLabel);
            this.Controls.Add(minCostForAgentLabel);
            this.Controls.Add(productionPersonCountLabel);
            this.Controls.Add(productionWorkshopNumberLabel);
            this.Controls.Add(productTypeIDLabel);
            this.Controls.Add(titleLabel);
            this.Controls.Add(this.titleTextBox);
            this.Font = new System.Drawing.Font("Gabriola", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 7, 3, 7);
            this.Name = "Form2NewProduct";
            this.Text = "Форма добавления нового продукта";
            this.Load += new System.EventHandler(this.Form2NewProduct_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.productBindingSource)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.BindingSource productBindingSource;
        private System.Windows.Forms.TextBox titleTextBox;
        private System.Windows.Forms.MaskedTextBox maskedTextBox1;
        private System.Windows.Forms.MaskedTextBox maskedTextBox2;
        private System.Windows.Forms.MaskedTextBox maskedTextBox3;
        private System.Windows.Forms.MaskedTextBox maskedTextBox4;
        private System.Windows.Forms.MaskedTextBox maskedTextBox5;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button1;
    }
}