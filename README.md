# WSR 2022  🚀

---

##  🛠️ Технологии

![C#](https://img.shields.io/badge/c%23-%23239120.svg?style=for-the-badge&logo=c-sharp&logoColor=white)
![Markdown](https://img.shields.io/badge/markdown-%23000000.svg?style=for-the-badge&logo=markdown&logoColor=white)
![.Net](https://img.shields.io/badge/.NET-5C2D91?style=for-the-badge&logo=.net&logoColor=white)
![Visual Studio Code](https://img.shields.io/badge/Visual%20Studio%20Code-0078d7.svg?style=for-the-badge&logo=visual-studio-code&logoColor=white)
![Visual Studio](https://img.shields.io/badge/Visual%20Studio-5C2D91.svg?style=for-the-badge&logo=visual-studio&logoColor=white)
![Git](https://img.shields.io/badge/git-%23F05033.svg?style=for-the-badge&logo=git&logoColor=white)
![GitHub](https://img.shields.io/badge/github-%23121011.svg?style=for-the-badge&logo=github&logoColor=white)
![MySQL](https://img.shields.io/badge/mysql-%2300f.svg?style=for-the-badge&logo=mysql&logoColor=white)
 
---


## Основные команды Git по работе с ветками
```git

git config --global user.name "YOUR NAME"
git config --global user.email "YOUR_EMAIL@EXAMPLE.COM"

git log — история коммитов.
git status — измененные файлы (показывает добавлены в коммит или нет).
git add file — добавить файл в коммит.
git add . — добавить все изменённые файлы в коммит.
git commit — m ‘text’ - добавить подпись коммитов.
git commit --amend — изменения сообщение последнего коммита.
git branch — посмотреть ветки.
git branch -v — просмотре веток с последним в ней коммитом.
git branch –a - просмотре всех веток: удаленных и локальных
git branch -m my-bug-fix изменение имени ветки
git branch -m my-new-feature my-new-feature (из-под другой ветки) - изменение имени ветки

git branch -d название ветки — удалить ветку. 
git checkout название ветки — переключиться в ветку.
git checkout -b название ветки — создать новую ветку и сразу в неё переключиться.
git push сервер ветка – залить изменения на сервер в указанную ветку.
git push -f  — залить изменения на сервер в режиме force, то есть с возможностью переписать уже имеющиеся коммиты на сервере. Будьте очень аккуратны с этой командой, а лучше минимизируйте её использование, ведь вы будете переписывать серверные файлы.

git fetch получаем актуальную информацию о репозитории на сервере. Извлечём все последние изменения удалённого репозитория. Извлекаются последние актуальные метаданные.

git pull origin dev (ключевое слово origin указывает на удаленный репозиторий). Извлечь и скопировать все изменения из удалённого репозитория

git merge name_new_brahch
```

## Порядок работы c Git
```
git checkout master - переключаемся на мастер

git pull -p - обновляем ветку мастер с сервера. В общих чертах pull состоит из двух компонентов: команды fetch, которая подтягивает изменения с удаленного репозитория, и команды merge, которая совмещает эти изменения с вашей локальной веткой.

git checkout -b [название-ветки]

```
## Принципы работы c Git

- ваш комментарий к коммиту должен полностью описывать (в настоящем времени) его содержимое, например "add About section to navbar on static pages". Комментарий должен быть на английском языке, вне зависимости от того, какой язык используете вы или ваша команда разработчиков. Если вы хотите использовать точку или "и", то, возможно, вы включили в коммит слишком много, и его следовало бы делать более модульным и независимым.

- У нас есть две основные ветки -- master и dev. master используется для готового к работе в production кода. Любой код, попадающий сюда, будет протестирован и перейдет в production. Вы же будете работать с отдельной веткой и создавать пулл реквесты в ветку dev. Представьте, что ветки master вообще не существует.
  
1 Создайте ветку для любого функционала, используя $ git checkout -b your_feature_name.

2 Напишите код, закоммитьте его, напишите еще, повторите коммит (видите шаблон?).

3 Когда вы сделали работу, есть шанс, что кто-то уже внес изменения в "upstream". Это значит, что ветки master и dev уже устарели. Выполните $ git fetch upstream для получения свежих данных.

4 Используйте $ git branch --all, чтобы получить список всех веток, в том числе тех, которые обычно скрыты (например удаленных веток, которые вы только что получили). Среди них вы должны увидеть upstream/master и upstream/dev.

5 Теперь слейте изменения из upstream в локальную версию dev, используя $ git merge. В данном случае, будет необходимо выполнить $ git checkout dev, чтобы попасть на ветку dev, и затем выполнить $ git merge upstream/dev. На этом перенос изменений в dev будет закончен.

6 Замечу, что команды $ git fetch upstream и $ git merge upstream/some_branch - это то же самое, что и одна команда $ git pull upstream/some_branch, просто я предпочитаю разделять ее на два шага.

7 Теперь, когда ветка dev содержит актуальные данные, необходимо слить ее с вашей веткой нового функционала. Хотя это звучит странно, тем не менее это так и есть. Хотите вместо этого слить ветку нового функционала с dev? Да хотите, но не сейчас. Ваша ветка с новым функционалом "грязная". Вы не знаете, содержит ли она потенциальные конфликты. Каждый раз, когда вы сливаетесь со "старшими" ветками (например ветку с вашим функционалом в dev, или же dev в master), вам необходимо, чтобы слияние было неконфликтным. Поэтому, сначала необходимо слить "старшую" ветку с "младшей" для разрешения конфликтов. Так что, мы выполняем $ git checkout your_feature_name, чтобы перейти на вашу ветку с новым функционалом, а затем $ git merge dev, чтобы слить в нее dev.

8 Возможно вы получите конфликт слияния... разрешите его с помощью $ git mergetool, или просто откройте вручную конфликтующие файлы. В основном, в таких случаях, в файлы вносятся маркеры, обозначающие, какие строки относятся к новому коду, а какие - к существующему. Необходимо один за одним отредактировать эти файлы (включая удаление маркированного текста), а затем пересохранить их. После этого, необходимо провести коммит, чтобы закончить слияние.

## Создание Pull Request (ПР)

1 После того, как ветка с вашим функционалом проверена, и вы знаете, что она сольется с dev без ошибок, осталось пройти пару шагов. Выполните слияние с dev командами $ git checkout dev и $ git merge your_feature_name.

2 Теперь необходимо отослать локальную ветку dev в origin (напомним, это ваш форк на Github). Вы не можете отослать изменения сразу на upstream, так как у вас нет туда доступа, поэтому вам необходимо выполнить Pull Request (ПР). Команда $ git push origin dev отправит ветку dev в origin.

3 И, наконец, создайте ПР, чтобы отправить свою версию dev в ветку dev репозитория upstream. Это может быть сделано с использованием пользовательского интерфейса Github. Здесь вы должны быть уверены в том, что отправляете ПР в ветку dev, а не в master.

## Пункты для размышления
Постарайтесь ответить на предложенные вопросы. После выполнения задания попробуйте ответить на них ещё раз

- Как часто необходимо выполнять коммит?
- Как велики они должны быть?
- Что должен содержать комментарий?
- Можно ли коммитить незаконченный функционал?
- Какой метод работы вы должны использовать? (Merge? Topic Branches? Git-Flow? Rebase?). Подсказка: здесь нет правильного ответа.

## Быстрое добавление изменений на сервер Git

```bash
git add . && git commit -m "%1" && git push origin master
```
Далее в командной строке Git Bash можно просто запустить файл fast.bat с параметром message

```
./fast.bat new_message
```


# Справка по командам

## Чтобы посмотреть в MySQL все текущие кодировки и настройки
```SQL
show variables like "character_set_%";
```

## Резервное копирование базы данных

```sql
mysqldump -u username -p password -h localhost databasename > "C:\script.sql"
```
## Восстановление базы данных

```sql
mysql -u username -p password -h localhost databasename < "C:\script.sql"
```


## Работа с DataReader и ExecuteScalar

```C#
                string mysql = @"server=localhost;userid=root;password=root;database=test_data";
                MySqlConnection connect = new MySqlConnection(mysql);
                connect.Open();

                string sql = "select * from materials";
                MySqlCommand query = new MySqlCommand(sql, connect);
                MySqlDataReader result1 = query.ExecuteReader();
                while (result1.Read())
                {
                    richTextBox1.Text = ($"{result1[0]} \n {result1[1]}");
                }

                // Alternate 

                /* while (result1.Read())
                {
                    string thisrow = "";
                    for (int i = 0; i < result1.FieldCount; i++)
                        thisrow += result1.GetValue(i).ToString() + ",";
                    richTextBox1.Text = thisrow;
                } */

```

```C#

                string mysql = @"server=localhost;userid=root;password=root;database=test_data";
                MySqlConnection connect = new MySqlConnection(mysql);
                connect.Open();
                string sql = "select * from materials limit 1";
                MySqlCommand query = new MySqlCommand(sql, connect);
                string result = query.ExecuteScalar().ToString();
                MySqlDataReader result1 = query.ExecuteReader();
                richTextBox1.Text = ($"{result}");
                connect.Close();

```

## Валидация полей

```C#
private void textBox1_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            if (char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
                MessageBox.Show("Поле Name не может содержать цифры");

                ErrorProvider errorProvider1 = new ErrorProvider();
                
                errorProvider1.SetError(textBox1, "Введите только буквы");
            }
        }

        // Событие валидации поля
        private void textBox2_Validating_1(object sender, CancelEventArgs e)
        {
            if (textBox2.Text == "")
            {
                e.Cancel = false;
            }
            else
            {
                try
                {
                    double.Parse(textBox2.Text);
                    e.Cancel = false;
                }
                catch
                {
                    e.Cancel = true;
                    MessageBox.Show("Поле Password не может содержать буквы");

                }
            }
        }
```

 ## Сохранение и открытие файла

 ```C#
 namespace SaveFile
{
    public partial class SafeFile : Form
    {
        public SafeFile()
        {
            InitializeComponent();
        }


        // Open Button
        private void button1_Click(object sender, EventArgs e)
        {
            Stream myStream = null;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.InitialDirectory = @"c:\";
            openFileDialog1.Filter = "txt files (*.txt)|*.txt|Allfiles(*.*) | *.* ";
            openFileDialog1.FilterIndex = 2;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if ((myStream = openFileDialog1.OpenFile()) != null)
                    {
                        using (myStream)
                        {

                            richTextBox1.LoadFile(openFileDialog1.FileName,
                             RichTextBoxStreamType.PlainText);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk: "
                   + ex.Message);
                }
            }
        }


        // Save Button
        private void button2_Click(object sender, EventArgs e)
        {

            SaveFileDialog saveFileDialog1 = new SaveFileDialog(); // or create visual in form
            saveFileDialog1.Filter = "txt files (*.txt)|*.txt";
            if (saveFileDialog1.ShowDialog() ==
            System.Windows.Forms.DialogResult.OK
             && saveFileDialog1.FileName.Length > 0)
            {
                richTextBox1.SaveFile(saveFileDialog1.FileName, RichTextBoxStreamType.PlainText);
            }
        }
    }
}
 ```

 ## Печать файла

```C#
        string s;
        string[] strings;
        int ArrayCounter = 0;
        OpenFileDialog openFileDialog1 = new OpenFileDialog();

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            //Font myFont = new Font("Tahoma", 12, FontStyle.Regular, GraphicsUnit.Pixel);
            //string Hello = "Hello World!";
            //e.Graphics.DrawString(Hello, myFont, Brushes.Black, 20, 20);

            float LeftMargin = e.MarginBounds.Left;
            float TopMargin = e.MarginBounds.Top;
            float MyLines = 0;
            float YPosition = 0;
            int Counter = 0;
            string CurrentLine;
            MyLines = e.MarginBounds.Height /
            this.Font.GetHeight(e.Graphics);
            while (Counter < MyLines && ArrayCounter <=
           strings.Length - 1)
            {
                CurrentLine = strings[ArrayCounter];
                YPosition = TopMargin + Counter *
               this.Font.GetHeight(e.Graphics);
                e.Graphics.DrawString(CurrentLine, this.Font,
               Brushes.Black, LeftMargin, YPosition, new StringFormat());
                Counter++;
                ArrayCounter++;
            }
            if (!(ArrayCounter >= strings.GetLength(0) - 1))
                e.HasMorePages = true;
            else
                e.HasMorePages = false;


        }


        // Настройка печати
        private void button1_Click_1(object sender, EventArgs e)
        {
            PageSetupDialog pageSetupDialog1 = new PageSetupDialog();
            pageSetupDialog1.Document = printDocument1;
            pageSetupDialog1.ShowDialog();
        }

        // Print Dialog
        private void button2_Click_1(object sender, EventArgs e)
        {
            PrintDialog printDialog1 = new PrintDialog();
            printDialog1.Document = printDocument1;
            if (printDialog1.ShowDialog() == DialogResult.OK)


                printDocument1.Print();
        }

        // Предварительный просмотр
        private void button3_Click_1(object sender, EventArgs e)
        {
            PrintPreviewDialog printPreviewDialog1 = new PrintPreviewDialog();
            printPreviewDialog1.Document = printDocument1;
            printPreviewDialog1.ShowDialog();
        }

        private void button4_Click_1(object sender, EventArgs e)
        {
            System.Windows.Forms.DialogResult aResult;
            aResult = openFileDialog1.ShowDialog();
            if (aResult == System.Windows.Forms.DialogResult.OK)
            {
                System.IO.StreamReader aReader =
                new System.IO.StreamReader(openFileDialog1.FileName);
                s = aReader.ReadToEnd();
                aReader.Close();
                strings = s.Split('\n');
                //printDocument1.DocumentName = s;

            }
        }
```

 ## Открытие справки или документации

 ```C#
        private void button1_Click(object sender, EventArgs e)
        {
            HelpProvider helpProvider1 = new HelpProvider();
            helpProvider1.HelpNamespace = "C:/Users/artik/Desktop/help.txt";
            System.Windows.Forms.Help.ShowHelp(this, helpProvider1.HelpNamespace);   
        }
 ```

## Импорт данных в MySQL с помощью команды LOAD DATA

```SQL
SET SQL_SAFE_UPDATES = 0;
set global local_infile = 1;

SET FOREIGN_KEY_CHECKS=0;

LOAD DATA LOCAL INFILE 'C:/Users/artik/Desktop/productmaterial_b_import.csv' INTO TABLE materials_has_products
  FIELDS TERMINATED BY ';' ENCLOSED BY '"'
  LINES TERMINATED BY '\n'
  IGNORE 1 LINES;
```

---

## Работа с DataAdapter

```C#

private void DataAdapter_Load(object sender, EventArgs e)
        {

            string connection = @"server=localhost;userid=root;password=root;database=test_data;Character Set=utf8";
            MySqlConnection connect = new MySqlConnection(connection);
            connect.Open();
            MySqlCommand cmd = connect.CreateCommand();
            cmd.CommandText = "select * from users";
            MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
            DataSet dataset = new DataSet();
            adapter.Fill(dataset);
            dataGridView1.DataSource = dataset.Tables[0];

        }

        // Добавление столбцов и строк, но НЕ ДОЛЖНО БЫТЬ ПРИВЯЗКИ ДАННЫХ ИЗ БАЗЫ,
        // т.е. dataGridView1.DataSource = dTable;
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
            dataGridView1.ColumnCount = 3;
            dataGridView1.Columns[0].Name = "Product ID";
            dataGridView1.Columns[1].Name = "Product Name";
            dataGridView1.Columns[2].Name = "Product Price";

            string[] row = new string[] { "1", "Product 1", "1000" };
            dataGridView1.Rows.Add(row);

            row = new string[] { "2", "Product 2", "2000" };
            dataGridView1.Rows.Add(row);
            row = new string[] { "3", "Product 3", "3000" };
            dataGridView1.Rows.Add(row);
            row = new string[] { "4", "Product 4", "4000" };
            dataGridView1.Rows.Add(row);

            }
            catch
            {
                MessageBox.Show("Возможно есть привязка данных из базы данных. Проверьте");
            }
        }

        // Cкрытие элементов и запрет на чтение
        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.Columns[2].Visible = false;
            dataGridView1.Rows[1].Visible = false;

            dataGridView1.Rows[0].ReadOnly = true;
            
                MessageBox.Show("Теперь первая строка только для чтения!");
            
        }

        // Добавление кнопки и чекбокс в столбец DataGridView
        private void button3_Click(object sender, EventArgs e)
        {

            try
            {
                DataGridViewButtonColumn btn = new DataGridViewButtonColumn();
                dataGridView1.Columns.Add(btn);
                btn.HeaderText = "Click Data";
                btn.Text = "Click Here";
                btn.Name = "btn";
                btn.UseColumnTextForButtonValue = true;

                //


                DataGridViewCheckBoxColumn chk = new DataGridViewCheckBoxColumn();
                dataGridView1.Columns.Add(chk);
                chk.HeaderText = "Check Data";
                chk.Name = "chk";
                dataGridView1.Rows[2].Cells[3].Value = true;

                //

                DataGridViewComboBoxColumn cmb = new DataGridViewComboBoxColumn();
                cmb.HeaderText = "Select Data";
                cmb.Name = "cmb";
                cmb.MaxDropDownItems = 4;
                cmb.Items.Add("True");
                cmb.Items.Add("False");
                dataGridView1.Columns.Add(cmb);
            }
            catch
            {
                MessageBox.Show("Возможно ощибка DataError");
            }


        }

        // Add Image Надо прописывать полный путь до изображения
        private void button4_Click(object sender, EventArgs e)
        {
            DataGridViewImageColumn img = new DataGridViewImageColumn();
            Image image = Image.FromFile("F:/Dropbox/WSR_2022/RoadToWordskills/materials/image_1.jpeg");
            img.Image = image;
            try
            {
                dataGridView1.Columns.Add(img);
            }
            catch
            {
                MessageBox.Show("Возможно ошибка DataError");
                img.HeaderText = "Image";
                img.Name = "img";
            }

        }

        // Add Link Column
        private void button5_Click(object sender, EventArgs e)
        {
            DataGridViewLinkColumn lnk = new DataGridViewLinkColumn();
            dataGridView1.Columns.Add(lnk);
            lnk.HeaderText = "Link Data";
            lnk.Name = "Http://csharp.net-informations.com";
            lnk.Text = "Http://csharp.net-informations.com";
            lnk.UseColumnTextForLinkValue = true;
        }

        // Insert Into
        private void button6_Click(object sender, EventArgs e)
        {

            string connection = @"server=localhost;userid=root;password=root;database=test_data;Character Set=utf8";
            MySqlConnection connect = new MySqlConnection(connection);
            connect.Open();

            MySqlCommand cmd = connect.CreateCommand();
            cmd.CommandText = "select * from users";

            // Изменяется адаптер и изменяется база данных

            MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);

            adapter.InsertCommand = new MySqlCommand("insert into users(name,age) values('Коля',23)", connect);
            adapter.InsertCommand.ExecuteNonQuery();
            DataSet dataset = new DataSet();
            adapter.Fill(dataset);
            dataGridView1.DataSource = dataset.Tables[0];
            connect.Close();
        }

        // Sort
        private void SortButton_Click(object sender, EventArgs e)
        {
            string connection = @"server=localhost;userid=root;password=root;database=test_data;Character Set=utf8";
            MySqlConnection connect = new MySqlConnection(connection);
            connect.Open();

            MySqlCommand cmd = connect.CreateCommand();
            cmd.CommandText = "select * from users";
            MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
            DataSet dataset = new DataSet();
            adapter.Fill(dataset);
            DataView materialDataView = new DataView(dataset.Tables[0]);
            DataView materialDataVIew = dataset.Tables[0].DefaultView;

            
            materialDataView = new DataView(dataset.Tables[0], "id > 10 and id < 15", "name Desc", DataViewRowState.CurrentRows);
            dataGridView1.DataSource = materialDataView;

            connect.Close();
        }

        // Filter
        private void FilterButton_Click(object sender, EventArgs e)
        {
            string connection = @"server=localhost;userid=root;password=root;database=test_data;Character Set=utf8";
            MySqlConnection connect = new MySqlConnection(connection);
            connect.Open();

            MySqlCommand cmd = connect.CreateCommand();
            cmd.CommandText = "select * from users";
            MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
            DataSet dataset = new DataSet();
            adapter.Fill(dataset);
            DataView materialDataView = new DataView(dataset.Tables[0]);
            DataView materialDataVIew = dataset.Tables[0].DefaultView;

            // Filter
            materialDataView = new DataView(dataset.Tables[0], "age < = 23", "age asc", DataViewRowState.CurrentRows);
            dataGridView1.DataSource = materialDataView;

            connect.Close();
        }

        // Search
        private void SearchButton_Click(object sender, EventArgs e)
        {
            string connection = @"server=localhost;userid=root;password=root;database=test_data;Character Set=utf8";
            MySqlConnection connect = new MySqlConnection(connection);
            connect.Open();

            MySqlCommand cmd = connect.CreateCommand();
            cmd.CommandText = "select * from users";
            MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
            DataSet dataset = new DataSet();
            adapter.Fill(dataset);
            DataView materialDataView = new DataView(dataset.Tables[0]);
            DataView materialDataVIew = dataset.Tables[0].DefaultView;

            // Search
            materialDataView = new DataView(dataset.Tables[0]);
            materialDataView.Sort = "name";

            int index = materialDataView.Find(textBox1.Text);
            if (index == -1)
            {
                MessageBox.Show("Отсутствие результатов поиска!");
            }
            else
            {
                MessageBox.Show(materialDataView[index]["id"].ToString() + " " + materialDataView[index]["name"].ToString());
            }


            // C определенного места

            /*            adapter.Fill(dataset, 5, 3, "users");
                        for (int i = 0; i <= dataset.Tables[0].Rows.Count - 1; i++)
                        {
                            MessageBox.Show(dataset.Tables[0].Rows[i].ItemArray[1].ToString());
                        }*/

            connect.Close();
        }

        // Добавление строки только в DataGrid
        private void button7_Click(object sender, EventArgs e)
        {
            string connection = @"server=localhost;userid=root;password=root;database=test_data;Character Set=utf8";
            MySqlConnection connect = new MySqlConnection(connection);
            connect.Open();
            MySqlCommand cmd = connect.CreateCommand();
            cmd.CommandText = "select * from users";
            MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
            DataSet dataset = new DataSet();
            adapter.Fill(dataset);
            dataGridView1.DataSource = dataset.Tables[0];

            // Добавление новой строки в DataView
            var materialDataView = new DataView(dataset.Tables[0]);
            DataRowView newRow = materialDataView.AddNew();
            newRow["id"] = 72;
            newRow["name"] = "Сергей";
            newRow["age"] = 32;
            newRow.EndEdit();
            materialDataView.Sort = "name";
            dataGridView1.DataSource = materialDataView;

            /* var changes = dataset.GetChanges();
                if (changes != null)
                {
                 adapter.Update(changes);
            }
            */

        }

        // Изменение значения поля записи и сохранение
        private void EditButton_Click(object sender, EventArgs e)
        {

            string connection = @"server=localhost;userid=root;password=root;database=test_data;Character Set=utf8";
            MySqlConnection connect = new MySqlConnection(connection);
            connect.Open();
            MySqlCommand cmd = connect.CreateCommand();
            cmd.CommandText = "select * from users";
            MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
            DataSet dataset = new DataSet();
            adapter.Fill(dataset);
            

            // Редактирование
            var materialDataView = new DataView(dataset.Tables[0], "", "name", DataViewRowState.CurrentRows);
            int index = materialDataView.Find("Коля");

            if (index == -1)
            {
                MessageBox.Show("Не найдено!");
            }
            else
            {
                adapter.UpdateCommand = new MySqlCommand("Update users set name = 'Коля_new' where name = 'Коля'", connect);
                adapter.UpdateCommand.ExecuteNonQuery();
                materialDataView[index]["name"] = "Коля_new";
                MessageBox.Show("Запись обновлена!");
            }

            dataGridView1.DataSource = materialDataView;


        }

        // Удаление из DataView
        private void DeleteButton_Click(object sender, EventArgs e)
        {
            string connection = @"server=localhost;userid=root;password=root;database=test_data;Character Set=utf8";
            MySqlConnection connect = new MySqlConnection(connection);
            connect.Open();
            MySqlCommand cmd = connect.CreateCommand();
            cmd.CommandText = "select * from users";
            MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
            DataSet dataset = new DataSet();
            adapter.Fill(dataset);

            // Удаление строки из DataView
            var materialDataView = new DataView(dataset.Tables[0], "", "", DataViewRowState.CurrentRows);
            int count = materialDataView.Table.Rows.Count;
            materialDataView.Table.Rows[count - 1].Delete();
            MessageBox.Show("Последняя строка удалена!");

            dataGridView1.DataSource = materialDataView;

        }

```

### Экспорт в Exсel

```C#

private void button1_Click(object sender, EventArgs e)
        {

            //

            string connection = @"server=localhost;userid=root;password=root;database=test_data;Character Set=utf8";
            MySqlConnection connect = new MySqlConnection(connection);
            connect.Open();
            MySqlCommand cmd = connect.CreateCommand();
            cmd.CommandText = "select * from users";
            MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
            DataSet dataset = new DataSet();
            adapter.Fill(dataset);

            dataGridView1.DataSource = dataset.Tables[0];


            //
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            Int16 i, j;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);

            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            for (i = 0; i <= dataGridView1.RowCount - 2; i++)
            {
                for (j = 0; j <= dataGridView1.ColumnCount - 1; j++)
                {
                    xlWorkSheet.Cells[i + 1, j + 1] = dataGridView1[j, i].Value.ToString();
                }
            }

            xlWorkBook.SaveAs(@"c:\Users\artik\Desktop\ExportExcel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }


```

### Загрузка данных из Excel

```C#

private void button1_Click(object sender, EventArgs e)
        {
            //

            string connection = @"server=localhost;userid=root;password=root;database=test_data;Character Set=utf8";
            MySqlConnection connect = new MySqlConnection(connection);
            connect.Open();
            MySqlCommand cmd = connect.CreateCommand();
            cmd.CommandText = "select * from users";
            MySqlDataAdapter adapter = new MySqlDataAdapter(cmd);
            DataSet dataset = new DataSet();
            adapter.Fill(dataset);

            dataGridView1.DataSource = dataset.Tables[0];


            //

            System.Data.OleDb.OleDbConnection MyConnection;
            System.Data.DataSet DtSet;
            System.Data.OleDb.OleDbDataAdapter MyCommand;
            MyConnection = new System.Data.OleDb.OleDbConnection(@"provider=Microsoft.Jet.OLEDB.4.0;Data Source='c:\Users\artik\Desktop\ExportExcel.xls';Extended Properties=Excel 8.0;");
            MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [Лист1$]", MyConnection);
            MyCommand.TableMappings.Add("users", "Net-informations.com");
            DtSet = new System.Data.DataSet();
            MyCommand.Fill(DtSet);
            dataGridView1.DataSource = DtSet.Tables[0];
            MyConnection.Close();

        }


```

### CRUD in DataGridView

```C#

public partial class CRUD : Form
    {

        MySqlCommand sCommand;
        MySqlDataAdapter sAdapter;
        MySqlCommandBuilder sBuilder;
        DataSet sDs;
        DataTable sTable;


        public CRUD()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string connectionString = @"server=localhost;userid=root;password=root;database=test_data;charset=utf8";
            string sql = "SELECT * FROM users";
            MySqlConnection connection = new MySqlConnection(connectionString);
            connection.Open();
            sCommand = new MySqlCommand(sql, connection);
            sAdapter = new MySqlDataAdapter(sCommand);
            sBuilder = new MySqlCommandBuilder(sAdapter);
            sDs = new DataSet();
            sAdapter.Fill(sDs, "users");
            sTable = sDs.Tables["users"];
            connection.Close();
            dataGridView1.DataSource = sDs.Tables["users"];
            dataGridView1.ReadOnly = true;
            save_btn.Enabled = false;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            dataGridView1.RowsDefaultCellStyle.BackColor = Color.Bisque;
            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.Beige;
        }

        private void new_btn_Click(object sender, EventArgs e)
        {
            dataGridView1.ReadOnly = false;
            save_btn.Enabled = true;
            new_btn.Enabled = false;
            delete_btn.Enabled = false;
        }

        private void delete_btn_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Do you want to delete this row ?", "Delete", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                dataGridView1.Rows.RemoveAt(dataGridView1.SelectedRows[0].Index);
                sAdapter.Update(sTable);
            }
        }

        private void save_btn_Click(object sender, EventArgs e)
        {
            sAdapter.Update(sTable);
            dataGridView1.ReadOnly = true;
            save_btn.Enabled = false;
            new_btn.Enabled = true;
            delete_btn.Enabled = true;
        }


        // Событие по выделению строки в DataGridView
        private void dataGridView1_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
           
                dataGridView1.Rows[e.RowIndex].Selected = true;
                rowIndex = e.RowIndex;
                dataGridView1.CurrentCell = dataGridView1.Rows[e.RowIndex].Cells[1];

                if (!dataGridView1.Rows[rowIndex].IsNewRow)
                {
                    dataGridView1.Rows.RemoveAt(rowIndex);
                }
        }
    }

```

### Автозаполнение поля в DataGridView


```C#

public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.ColumnCount = 3;
            dataGridView1.Columns[0].Name = "Product ID";
            dataGridView1.Columns[1].Name = "Product Name";
            dataGridView1.Columns[2].Name = "Product Price";

            string[] row = new string[] { "1", "Product 1", "1000" };
            dataGridView1.Rows.Add(row);
            row = new string[] { "2", "Product 2", "2000" };
            dataGridView1.Rows.Add(row);
            row = new string[] { "3", "Product 3", "3000" };
            dataGridView1.Rows.Add(row);
            row = new string[] { "4", "Product 4", "4000" };
            dataGridView1.Rows.Add(row);

        }

        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            string titleText = dataGridView1.Columns[1].HeaderText;
            if (titleText.Equals("Product Name"))
            {
                TextBox autoText = e.Control as TextBox;
                if (autoText != null)
                {
                    autoText.AutoCompleteMode = AutoCompleteMode.Suggest;
                    autoText.AutoCompleteSource = AutoCompleteSource.CustomSource;
                    AutoCompleteStringCollection DataCollection = new AutoCompleteStringCollection();
                    addItems(DataCollection);
                    autoText.AutoCompleteCustomSource = DataCollection;       
                }
            }
         }

        public void addItems(AutoCompleteStringCollection col)
        {
            col.Add("Product 1");
            col.Add("Product 2");
            col.Add("Product 3");
            col.Add("Product 4");
            col.Add("Product 5");
            col.Add("Product 6");
        }

    }

```

 ### 🔗 Дополнительные материалы

---
 ## Инструменты для работы

 ### Редактор README.md https://dillinger.io/
 ### Git https://www.atlassian.com/ru/git/tutorials/learn-git-with-bitbucket-cloud   
 ### C# http://csharp.net-informations.com/

### Ресурсы для работы с readme.md, оформление профиля на GitHub, Markdown 

* https://github.com/abhisheknaiidu/awesome-github-profile-readme#github-actions-

* https://arturssmirnovs.github.io/github-profile-readme-generator/

* https://www.markdownguide.org/cheat-sheet/

* https://simpleicons.org/

* https://gist.github.com/rxaviers/7360908

* https://github.com/anuraghazra/github-readme-stats

* https://shields.io/

* https://www.webfx.com/tools/emoji-cheat-sheet/

### Ресурсы по работе с Git

- https://www.gitkraken.com/

- https://www.sourcetreeapp.com/7https://desktop.

- [github.com/](https://desktop.github.com/)

- https://learngitbranching.js.org/?locale=ru_RU

- https://www.git-tower.com/learn/git/ebook/en/command-line/appendix/best-practices#start

- https://sethrobertson.github.io/GitBestPractices/

- https://softwareengineering.stackexchange.com/questions/165725/git-branching-and-tagging-best-practices

- https://nvie.com/posts/a-successful-git-branching-model/

- https://www.lullabot.com/articles/git-best-practices-workflow-guidelines

- https://services.github.com/

- https://www.sbf5.com/~cduan/technical/git/

- https://learngitbranching.js.org/?locale=ru_RU

- https://www.webfx.com/blog/web-design/git-tutorials-beginners/

- https://gitimmersion.com/lab_08.html

- Тренажер для визуализации работы с ветками https://git-school.github.io/visualizing-git/
- https://github.com/git-school/visualizing-git
- Badges https://github.com/Ileriayo/markdown-badges

---

## Проблемы

### Отображение русских символов через командную строку в MySQL
#### Решение:   в командной строке

 ```sql 
 set names "cp866" 
 ```


### MySQL connector не работает с Visual Studio 2022 на начало 2022 года
#### Решение: подключение через NuGet пакет MySQL.Data в каждый проект или подключать локально

<!-- <details> 
  <summary> Problems </summary>
   A1: JavaScript 
</details> -->
****