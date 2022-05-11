using Excel = Microsoft.Office.Interop.Excel;

namespace WinFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()  // формирование формы
        {
            InitializeComponent();

            Text = "Меню входа";

            //считывание данных из excel
            ReadDataExcel();

            // начальные параметры для формы и панелей
            Size = new Size(600, 450);
            panel1.Visible = true;
            panel1.Location = new Point(0, 0);

            panel2.Visible = false;
            panel3.Visible = false;
            panel4.Visible = false;
            panel5.Visible = false;
        }

        // закрытие приложения
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Close();
        }

        // поиск файла Excel
        static public string path()
        {
            string path = "err";
            path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            for (int i = 0; i < 5; i++)
                path = Path.GetDirectoryName(path);
            path = path + "\\Data.xlsx";
            return path;
        }

        // кнопка "Продолжить" в меню входа
        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("Вы не выбрали режим", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                // проверки на пустые поля в меню входа
                if (textBox1.Text == "")
                {
                    MessageBox.Show("Вы не ввели значение в поле \"Логин\" ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (textBox2.Text == "")
                {
                    MessageBox.Show("Вы не ввели значение в поле \"Пароль\" ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (checkLogin(textBox1.Text) != "Ok")
                {
                    MessageBox.Show(checkLogin(textBox1.Text), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (checkPassword(textBox2.Text) != "Ok")
                {
                    MessageBox.Show(checkPassword(textBox2.Text), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    // считывание логина и пароля
                    string login = textBox1.Text;
                    string password = textBox2.Text;

                    // если выбран режим "Вход"
                    if (comboBox1.SelectedIndex == 0) 
                    {
                        int uniqueNumber = Person.FindUniqueNumber(login);

                        if (uniqueNumber == -1)
                        {
                            MessageBox.Show("Неверный логин или пароль", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        else if (uniqueNumber != -1 && StaticClass.persons[uniqueNumber].Password != password)
                        {
                            MessageBox.Show("Неверный логин или пароль", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        else
                        {
                            int levelAccess = Person.FindLevelAccess(login);
                            if (levelAccess == 0)
                            {
                                SetDisign(2);
                            }
                            else if (levelAccess == 1)
                            {
                                SetDisign(3);
                            }
                        }
                    }

                    // если выбран режим "Регистрация"
                    else if (comboBox1.SelectedIndex == 1) 
                    {
                        int uniqueNumber = Person.FindUniqueNumber(login);
                        if (uniqueNumber == -1)
                        {
                            SetDisign(1);
                        }
                        else
                        {
                            MessageBox.Show("Такой пользователь существует", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
        }

        // метод для дизайна формы, определяет какая панель будет активной
        void SetDisign(int number)
        {
            switch (number)
            {
                // меню входа
                case 0: 
                    Size = new Size(600, 450);
                    Text = "Меню входа";
                    panel1.Visible = true;
                    panel1.Location = new Point(0, 0);

                    textBox1.Clear();
                    textBox2.Clear();
                    comboBox1.SelectedIndex = -1;

                    panel2.Visible = false;
                    panel3.Visible = false;
                    panel4.Visible = false;
                    panel5.Visible = false;

                    break;

                // регистрация пользователя
                case 1: 
                    Size = new Size(550, 600);
                    Text = "Меню регистрации";
                    panel2.Visible = true;
                    panel2.Location = new Point(0, 0);

                    button4.Enabled = true;
                    textBox3.Enabled = true;
                    textBox4.Enabled = true;
                    textBox5.Enabled = true;
                    textBox6.Enabled = true;
                    textBox7.Enabled = true;

                    textBox3.Clear();
                    textBox4.Clear();
                    textBox5.Clear();
                    textBox6.Clear();
                    textBox7.Clear();

                    panel1.Visible = false;
                    panel3.Visible = false;
                    panel4.Visible = false;
                    panel5.Visible = false;

                    break;

                // панель пользователя
                case 2: 
                    Size = new Size(700, 550);
                    Text = "Панель пользователя";
                    panel3.Visible = true;
                    panel3.Location = new Point(0, 0);

                    label10.Text = "Вы вошли как пользователь: " + textBox1.Text;

                    button15.Visible = false;
                    button7.Enabled = true;
                    comboBox2.Enabled = true;

                    comboBox2.SelectedIndex = -1;
                    textBox8.Clear();
                    textBox9.Clear();
                    textBox10.Clear();
                    textBox11.Clear();
                    textBox12.Clear();

                    textBox8.Enabled = false;
                    textBox9.Enabled = false;
                    textBox10.Enabled = false;
                    textBox11.Enabled = false;
                    textBox12.Enabled = false;

                    panel1.Visible = false;
                    panel2.Visible = false;
                    panel4.Visible = false;
                    panel5.Visible = false;

                    break;

                // панель админа
                case 3: 
                    Size = new Size(700, 550);
                    Text = "Панель администратора";
                    panel4.Visible = true;
                    panel4.Location = new Point(0, 0);

                    label25.Text = "Вы вошли как admin: " + textBox1.Text;

                    comboBox5.Items.Clear();
                    for (int i = 0; i < StaticClass.counterPersons; i++)
                    {
                        comboBox5.Items.Add(i);
                    }

                    button16.Visible = false;
                    button11.Enabled = true;

                    comboBox3.SelectedIndex = -1;
                    textBox13.Clear();
                    textBox14.Clear();
                    textBox15.Clear();
                    textBox16.Clear();
                    textBox17.Clear();

                    textBox13.Enabled = false;
                    textBox14.Enabled = false;
                    textBox15.Enabled = false;
                    textBox16.Enabled = false;
                    textBox17.Enabled = false;

                    panel1.Visible = false;
                    panel2.Visible = false;
                    panel3.Visible = false;
                    panel5.Visible = false;

                    break;

                // панель админа дял поиска пользователей
                case 4:
                    Size = new Size(950, 700);
                    Text = "Панель администратора";
                    panel5.Visible = true;
                    panel5.Location = new Point(0, 0);
                    dataGridView1.Rows.Clear();

                    comboBox4.SelectedIndex = -1;
                    textBox19.Clear();

                    panel1.Visible = false;
                    panel2.Visible = false;
                    panel3.Visible = false;
                    panel4.Visible = false;

                    break;
            }
        }

        // проверка на правильный ввод логина
        string checkLogin(string login)
        {
            string a = "Ok";
            if (login.Length < 5)
            {
                a = "Логин должен содержать не менее 5 символов";
            }
            else if (login.Length > 20)
            {
                a = "Логин не может содержать более 20 символов";
            }
            else
            {
                bool b = true;
                foreach (char c in login)
                {
                    if (b == true && ((int)c > 64 && (int)c < 91 || (int)c > 96 && (int)c < 123 || (int)c > 47 && (int)c < 58))
                    {
                        b = true;
                    }
                    else
                    {
                        a = "Логин должен содержать только буквы латиницы или арабские цифры";
                        b = false;
                    }
                }
            }
            return a;
        }

        // проверка на правильный ввод пароля
        string checkPassword(string password)
        {
            string a = "Ok";
            if (password.Length < 5)
            {
                a = "Пароль должен содержать не менее 5 символов";
            }
            else if (password.Length > 20)
            {
                a = "Пароль не может содержать более 20 символов";
            }
            else
            {
                bool b = true;
                foreach (char c in password)
                {
                    if (b == true && !(c >= 'А' && c <= 'Ю' || c >= 'а' && c <= 'ю'))
                    {
                        b = true;

                    }
                    else
                    {
                        a = "Пароль не может содержать буквы кириллицы";
                        b = false;
                    }
                }
            }
            return a;
        }

        // кнопка "Закончить регистрацию" в меню регистрации
        private void button4_Click(object sender, EventArgs e)
        {
            if (textBox3.Text == "")
            {
                MessageBox.Show("Вы не ввели значение в поле \"Имя\" ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (textBox4.Text == "")
            {
                MessageBox.Show("Вы не ввели значение в поле \"Фамилия\" ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (textBox5.Text == "")
            {
                MessageBox.Show("Вы не ввели значение в поле \"e-mail\" ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            else if (textBox6.Text == "")
            {
                MessageBox.Show("Вы не ввели значение в поле \"Контактный телефон\" ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            else if (textBox7.Text == "")
            {
                MessageBox.Show("Вы не ввели значение в поле \"Подвердите пароль\" ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            else
            {
                if (textBox2.Text != textBox7.Text)
                {
                    MessageBox.Show("Пароль не совпадает", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    // добавление пользователя
                    User.AddUser(textBox1.Text, textBox2.Text, textBox3.Text, textBox4.Text, textBox5.Text, textBox6.Text);
                    WriteInExcel(StaticClass.persons[StaticClass.counterPersons], StaticClass.counterPersons);

                    button4.Enabled = false;
                    textBox3.Enabled = false;
                    textBox4.Enabled = false;
                    textBox5.Enabled = false;
                    textBox6.Enabled = false;
                    textBox7.Enabled = false;
                }
            }
        }

        // кнопка "Перейти в главное меню" в меню регистрации
        private void button5_Click(object sender, EventArgs e)
        {
            SetDisign(0);
        }

        // кнопка "Подвердить" в меню пользователя
        private void button7_Click(object sender, EventArgs e)
        {
            string login = textBox1.Text;
            int uniqueNumber = Person.FindUniqueNumber(login);

            if (comboBox2.SelectedIndex == -1)
            {
                MessageBox.Show("Вы не выбрали действие", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (comboBox2.SelectedIndex == 0)
            {
                // вывод информации  в панель пользователя
                textBox11.Text = StaticClass.persons[uniqueNumber].FirstName;
                textBox12.Text = StaticClass.persons[uniqueNumber].FullName;
                textBox10.Text = StaticClass.persons[uniqueNumber].Email;
                textBox9.Text = StaticClass.persons[uniqueNumber].NumberPhone;
                textBox8.Text = StaticClass.persons[uniqueNumber].Password;

            }
            else if (comboBox2.SelectedIndex == 1)
            {
                // вывод информации  в панель пользователя
                textBox11.Text = StaticClass.persons[uniqueNumber].FirstName;
                textBox12.Text = StaticClass.persons[uniqueNumber].FullName;
                textBox10.Text = StaticClass.persons[uniqueNumber].Email;
                textBox9.Text = StaticClass.persons[uniqueNumber].NumberPhone;
                textBox8.Text = StaticClass.persons[uniqueNumber].Password;

                textBox8.Enabled = true;
                textBox9.Enabled = true;
                textBox10.Enabled = true;
                textBox11.Enabled = true;
                textBox12.Enabled = true;

                button15.Visible = true;
                button7.Enabled = false;
                comboBox2.Enabled = false;
            }
        }

        // кнопка "Сохранить информацию" в меню пользователя
        private void button15_Click(object sender, EventArgs e)
        {
            if (textBox11.Text == "")
            {
                MessageBox.Show("Вы не ввели значение в поле \"Имя\" ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (textBox12.Text == "")
            {
                MessageBox.Show("Вы не ввели значение в поле \"Фамилия\" ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (textBox10.Text == "")
            {
                MessageBox.Show("Вы не ввели значение в поле \"e-mail\" ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            else if (textBox9.Text == "")
            {
                MessageBox.Show("Вы не ввели значение в поле \"Контактный телефон\" ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            else if (textBox8.Text == "")
            {
                MessageBox.Show("Вы не ввели значение в поле \"Подвердите пароль\" ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            else if (checkPassword(textBox8.Text) != "Ok")
            {
                MessageBox.Show(checkPassword(textBox8.Text), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                string login = textBox1.Text;
                int uniqueNumber = Person.FindUniqueNumber(login);
                // изменение информации о пользователе
                Admin.ChangeUserPanel3(uniqueNumber, textBox11.Text, textBox12.Text, textBox10.Text, textBox9.Text, textBox8.Text);
                // изменение информации о пользователе в Excel
                ChangeUserinExcel(uniqueNumber);

                button15.Visible = false;
                button7.Enabled = true;
                comboBox2.Enabled = true;

                comboBox2.SelectedIndex = -1;
                textBox8.Clear();
                textBox9.Clear();
                textBox10.Clear();
                textBox11.Clear();
                textBox12.Clear();

                textBox8.Enabled = false;
                textBox9.Enabled = false;
                textBox10.Enabled = false;
                textBox11.Enabled = false;
                textBox12.Enabled = false;
            }
        }

        // кнопка "Перейти в главное меню" в меню пользователя
        private void button9_Click(object sender, EventArgs e)
        {
            SetDisign(0);
        }

        // кнопка "Подвердить" в меню админа
        private void button11_Click(object sender, EventArgs e)
        {
            int uniqueNumber = comboBox5.SelectedIndex;

            if (comboBox3.SelectedIndex == -1)
            {
                MessageBox.Show("Вы не выбрали действие", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (comboBox3.SelectedIndex == 0)
            {
                if (comboBox5.SelectedIndex == -1)
                {
                    MessageBox.Show("Выберите пользователя", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {

                    if (StaticClass.persons[uniqueNumber].LevelAccess == 1)
                    {
                        MessageBox.Show("У этого пользователя есть привилегии администратора", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    else
                    {
                        StaticClass.persons[uniqueNumber].LevelAccess = 1;
                        comboBox3.SelectedIndex = -1;
                        comboBox5.SelectedIndex = -1;

                        textBox17.Clear();
                        textBox16.Clear();
                        textBox15.Clear();
                        textBox14.Clear();
                        textBox13.Clear();

                        ChangeUserinExcel(uniqueNumber);
                    }
                }
            }
            else if (comboBox3.SelectedIndex == 1)
            {
                if (comboBox5.SelectedIndex == -1)
                {
                    MessageBox.Show("Выберите пользователя", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    textBox17.Text = StaticClass.persons[uniqueNumber].FirstName;
                    textBox16.Text = StaticClass.persons[uniqueNumber].FullName;
                    textBox15.Text = StaticClass.persons[uniqueNumber].Email;
                    textBox14.Text = StaticClass.persons[uniqueNumber].NumberPhone;
                    textBox13.Text = Convert.ToString(StaticClass.persons[uniqueNumber].LevelAccess);
                }

            }
            else if (comboBox3.SelectedIndex == 2)
            {
                if (comboBox5.SelectedIndex == -1)
                {
                    MessageBox.Show("Выберите пользователя", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    textBox17.Text = StaticClass.persons[uniqueNumber].FirstName;
                    textBox16.Text = StaticClass.persons[uniqueNumber].FullName;
                    textBox15.Text = StaticClass.persons[uniqueNumber].Email;
                    textBox14.Text = StaticClass.persons[uniqueNumber].NumberPhone;
                    textBox13.Text = Convert.ToString(StaticClass.persons[uniqueNumber].LevelAccess);

                    //textBox13.Enabled = true;
                    textBox14.Enabled = true;
                    textBox15.Enabled = true;
                    textBox16.Enabled = true;
                    textBox17.Enabled = true;

                    button16.Visible = true;
                    button11.Enabled = false;
                }
            }
            else if (comboBox3.SelectedIndex == 3)
            {
                if (comboBox5.SelectedIndex == -1)
                {
                    MessageBox.Show("Выберите пользователя", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    textBox17.Text = StaticClass.persons[uniqueNumber].FirstName;
                    textBox16.Text = StaticClass.persons[uniqueNumber].FullName;
                    textBox15.Text = StaticClass.persons[uniqueNumber].Email;
                    textBox14.Text = StaticClass.persons[uniqueNumber].NumberPhone;
                    textBox13.Text = Convert.ToString(StaticClass.persons[uniqueNumber].LevelAccess);

                    textBox17.Enabled = false;
                    textBox16.Enabled = false;
                    textBox15.Enabled = false;
                    textBox14.Enabled = false;
                    textBox13.Enabled = false;

                    button17.Visible = true;
                    button17.Location = new System.Drawing.Point(430, 330);
                    button11.Enabled = false;
                }
            }
        }

        // кнопка "Перейти в главное меню" в меню администратора
        private void button8_Click(object sender, EventArgs e)
        {
            SetDisign(0);
        }

        // кнопка "Перейти к поиску" в меню администратора
        private void button2_Click(object sender, EventArgs e)
        {
            SetDisign(4);
        }

        // кнопка "Сохранить информацию" в меню администратора
        private void button16_Click(object sender, EventArgs e)
        {
            if (textBox17.Text == "")
            {
                MessageBox.Show("Вы не ввели значение в поле \"Имя\" ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (textBox16.Text == "")
            {
                MessageBox.Show("Вы не ввели значение в поле \"Фамилия\" ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (textBox15.Text == "")
            {
                MessageBox.Show("Вы не ввели значение в поле \"e-mail\" ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            else if (textBox14.Text == "")
            {
                MessageBox.Show("Вы не ввели значение в поле \"Контактный телефон\" ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            else if (textBox13.Text == "")
            {
                MessageBox.Show("Вы не ввели значение в поле \"Уровень доступа\" ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            else
            {
                int uniqueNumber = comboBox5.SelectedIndex;
                // изенение информации о пользователе в меню админа
                Admin.ChangeUserPanel4(uniqueNumber, textBox17.Text, textBox16.Text, textBox15.Text, textBox14.Text, textBox13.Text);
                // изенение информации о пользователе в Excel
                ChangeUserinExcel(uniqueNumber);

                button16.Visible = false;
                button11.Enabled = true;

                comboBox3.SelectedIndex = -1;
                comboBox5.SelectedIndex = -1;

                textBox13.Clear();
                textBox14.Clear();
                textBox15.Clear();
                textBox16.Clear();
                textBox17.Clear();

                textBox13.Enabled = false;
                textBox14.Enabled = false;
                textBox15.Enabled = false;
                textBox16.Enabled = false;
                textBox17.Enabled = false;
            }
        }

        // кнопка "Удалить пользователя" в меню администратора
        private void button17_Click(object sender, EventArgs e)
        {
            int uniqueNumber = comboBox5.SelectedIndex;
            if (StaticClass.persons[uniqueNumber].LevelAccess == 1)
            {
                MessageBox.Show("Вы не можете удалить администратора", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                // удаление информации о пользователе
                for (int i = uniqueNumber; i < StaticClass.counterPersons- 1; i++)
                {
                    StaticClass.persons[i].UniqueNumber = StaticClass.persons[i + 1].UniqueNumber;
                    StaticClass.persons[i].Login = StaticClass.persons[i + 1].Login;
                    StaticClass.persons[i].Password = StaticClass.persons[i + 1].Password;
                    StaticClass.persons[i].FirstName = StaticClass.persons[i + 1].FirstName;
                    StaticClass.persons[i].FullName = StaticClass.persons[i + 1].FullName;
                    StaticClass.persons[i].Email = StaticClass.persons[i + 1].Email;
                    StaticClass.persons[i].NumberPhone = StaticClass.persons[i + 1].NumberPhone;
                    StaticClass.persons[i].LevelAccess = StaticClass.persons[i + 1].LevelAccess;
                }

                StaticClass.counterPersons--;

                // удаление информации о пользователе в Excel
                DeleteUserInExcel(StaticClass.counterPersons, uniqueNumber);
            }

            textBox17.Enabled = true;
            textBox16.Enabled = true;
            textBox15.Enabled = true;
            textBox14.Enabled = true;
            textBox13.Enabled = true;

            button17.Visible = false;
            button11.Enabled = true;

            comboBox3.SelectedIndex = -1;

            comboBox5.Items.Clear();
            for (int i = 0; i < StaticClass.counterPersons; i++)
            {
                comboBox5.Items.Add(i);
            }
            textBox13.Clear();
            textBox14.Clear();
            textBox15.Clear();
            textBox16.Clear();
            textBox17.Clear();
        }

        // кнопка "Поиск" в меню поиска для администратора
        private void button12_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            if (comboBox4.SelectedIndex == -1)
            {
                MessageBox.Show("Выберите параметр поиска", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                if (comboBox4.SelectedIndex == 0)
                {
                    // вывод всех пользователей в таблицу
                    for (int i = 0; i < StaticClass.counterPersons; i++)
                    {

                        dataGridView1.Rows.Add();
                        dataGridView1.Rows[i].Cells[0].Value = StaticClass.persons[i].UniqueNumber;
                        dataGridView1.Rows[i].Cells[1].Value = StaticClass.persons[i].LevelAccess;
                        dataGridView1.Rows[i].Cells[2].Value = StaticClass.persons[i].Login;
                        dataGridView1.Rows[i].Cells[3].Value = StaticClass.persons[i].Password;
                        dataGridView1.Rows[i].Cells[4].Value = StaticClass.persons[i].Email;
                        dataGridView1.Rows[i].Cells[5].Value = StaticClass.persons[i].NumberPhone;
                        dataGridView1.Rows[i].Cells[6].Value = StaticClass.persons[i].DataRegistration;
                    }
                    comboBox4.SelectedIndex = -1;
                    textBox19.Clear();
                }
                else if (textBox19.Text == "")
                {
                    MessageBox.Show("Вы не ввели значение в поле \"Поиска\" ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    if (comboBox4.SelectedIndex == 1)
                    {
                        // вывод всех пользователей в таблицу по параметру "логина"
                        int k = 0;
                        for (int i = 0; i < StaticClass.counterPersons; i++)
                        {
                            if (StaticClass.persons[i].Login == textBox19.Text)
                            {
                                dataGridView1.Rows.Add();
                                dataGridView1.Rows[k].Cells[0].Value = StaticClass.persons[i].UniqueNumber;
                                dataGridView1.Rows[k].Cells[1].Value = StaticClass.persons[i].LevelAccess;
                                dataGridView1.Rows[k].Cells[2].Value = StaticClass.persons[i].Login;
                                dataGridView1.Rows[k].Cells[3].Value = StaticClass.persons[i].Password;
                                dataGridView1.Rows[k].Cells[4].Value = StaticClass.persons[i].Email;
                                dataGridView1.Rows[k].Cells[5].Value = StaticClass.persons[i].NumberPhone;
                                dataGridView1.Rows[k].Cells[6].Value = StaticClass.persons[i].DataRegistration;
                                k++;
                            }

                        }
                        if (k == 0)
                        {
                            dataGridView1.Rows.Add();
                            dataGridView1.Rows[0].Cells[0].Value = "Такий пользователей нет";
                        }
                    }

                    if (comboBox4.SelectedIndex == 2)
                    {
                        int num;
                        bool isNum = int.TryParse(textBox19.Text, out num);
                        if (isNum is true)
                        {
                            // вывод всех пользователей в таблицу по параметру "Уровень доступа"
                            int k = 0;
                            for (int i = 0; i < StaticClass.counterPersons; i++)
                            {
                                if (StaticClass.persons[i].LevelAccess == num)
                                {
                                    dataGridView1.Rows.Add();
                                    dataGridView1.Rows[k].Cells[0].Value = StaticClass.persons[i].UniqueNumber;
                                    dataGridView1.Rows[k].Cells[1].Value = StaticClass.persons[i].LevelAccess;
                                    dataGridView1.Rows[k].Cells[2].Value = StaticClass.persons[i].Login;
                                    dataGridView1.Rows[k].Cells[3].Value = StaticClass.persons[i].Password;
                                    dataGridView1.Rows[k].Cells[4].Value = StaticClass.persons[i].Email;
                                    dataGridView1.Rows[k].Cells[5].Value = StaticClass.persons[i].NumberPhone;
                                    dataGridView1.Rows[k].Cells[6].Value = StaticClass.persons[i].DataRegistration;
                                    k++;
                                }

                            }
                            if (k == 0)
                            {
                                dataGridView1.Rows.Add();
                                dataGridView1.Rows[0].Cells[0].Value = "Такий пользователей нет";
                            }
                        }
                        else
                        {
                            MessageBox.Show("Введите целое число в поле \"Поиска\" ", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }

                    if (comboBox4.SelectedIndex == 3)
                    {
                        // вывод всех пользователей в таблицу по параметру "E-mail"
                        int k = 0;
                        for (int i = 0; i < StaticClass.counterPersons; i++)
                        {
                            if (StaticClass.persons[i].Email == textBox19.Text)
                            {
                                dataGridView1.Rows.Add();
                                dataGridView1.Rows[k].Cells[0].Value = StaticClass.persons[i].UniqueNumber;
                                dataGridView1.Rows[k].Cells[1].Value = StaticClass.persons[i].LevelAccess;
                                dataGridView1.Rows[k].Cells[2].Value = StaticClass.persons[i].Login;
                                dataGridView1.Rows[k].Cells[3].Value = StaticClass.persons[i].Password;
                                dataGridView1.Rows[k].Cells[4].Value = StaticClass.persons[i].Email;
                                dataGridView1.Rows[k].Cells[5].Value = StaticClass.persons[i].NumberPhone;
                                dataGridView1.Rows[k].Cells[6].Value = StaticClass.persons[i].UniqueNumber;
                                dataGridView1.Rows[k].Cells[7].Value = StaticClass.persons[i].DataRegistration;
                                k++;
                            }

                        }
                        if (k == 0)
                        {
                            dataGridView1.Rows.Add();
                            dataGridView1.Rows[0].Cells[0].Value = "Такий пользователей нет";
                        }
                    }

                    if (comboBox4.SelectedIndex == 4)
                    {
                        // вывод всех пользователей в таблицу по параметру "Номер телефона"
                        int k = 0;
                        for (int i = 0; i < StaticClass.counterPersons; i++)
                        {
                            if (StaticClass.persons[i].NumberPhone == textBox19.Text)
                            {
                                dataGridView1.Rows.Add();
                                dataGridView1.Rows[k].Cells[0].Value = StaticClass.persons[i].UniqueNumber;
                                dataGridView1.Rows[k].Cells[1].Value = StaticClass.persons[i].LevelAccess;
                                dataGridView1.Rows[k].Cells[2].Value = StaticClass.persons[i].Login;
                                dataGridView1.Rows[k].Cells[3].Value = StaticClass.persons[i].Password;
                                dataGridView1.Rows[k].Cells[4].Value = StaticClass.persons[i].Email;
                                dataGridView1.Rows[k].Cells[5].Value = StaticClass.persons[i].NumberPhone;
                                dataGridView1.Rows[k].Cells[6].Value = StaticClass.persons[i].UniqueNumber;
                                dataGridView1.Rows[k].Cells[7].Value = StaticClass.persons[i].DataRegistration;
                                k++;
                            }

                        }
                        if (k == 0)
                        {
                            dataGridView1.Rows.Add();
                            dataGridView1.Rows[0].Cells[0].Value = "Такий пользователей нет";
                        }
                    }
                    comboBox4.SelectedIndex = -1;
                    textBox19.Clear();
                }
            }
        }

        // кнопка "Перейти в главное меню" в меню поиска для администратора
        private void button13_Click(object sender, EventArgs e)
        {
            SetDisign(0);
        }

        // кнопка "Вернуться к панели администратора" в меню поиска для администратора
        private void button14_Click(object sender, EventArgs e)
        {
            SetDisign(3);
        }

        // функция для считывание данных из Excel
        void ReadDataExcel()
        {
            string path = Form1.path();
            Excel.Application excelApp = new Excel.Application();                                  // создание ссылки на COM-оъект
            Excel.Workbook excelBook = excelApp.Workbooks.Open(path);        // открываем excel файл
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

           StaticClass.counterPersons = Convert.ToInt16(workSheet.Cells[1, "A"].Text.ToString());

            for (int i = 0; i < StaticClass.counterPersons; i++)
            {
                if (Convert.ToInt16(workSheet.Cells[i + 3, "B"].Text.ToString()) == 1)
                {
                    StaticClass.persons[i] = new User();
                    StaticClass.persons[i].UniqueNumber = Convert.ToInt16(workSheet.Cells[i + 3, "A"].Text.ToString());
                    StaticClass.persons[i].LevelAccess = Convert.ToInt16(workSheet.Cells[i + 3, "B"].Text.ToString());
                    StaticClass.persons[i].Login = workSheet.Cells[i + 3, "C"].Text.ToString();
                    StaticClass.persons[i].Password = workSheet.Cells[i + 3, "D"].Text.ToString();
                    StaticClass.persons[i].FirstName = workSheet.Cells[i + 3, "E"].Text.ToString();
                    StaticClass.persons[i].FullName = workSheet.Cells[i + 3, "F"].Text.ToString();
                    StaticClass.persons[i].Email = workSheet.Cells[i + 3, "G"].Text.ToString();
                    StaticClass.persons[i].NumberPhone = workSheet.Cells[i + 3, "H"].Text.ToString();
                    StaticClass.persons[i].DataRegistration = workSheet.Cells[i + 3, "I"].Text.ToString();
                }
                else
                {
                    StaticClass.persons[i] = new Admin();
                    StaticClass.persons[i].UniqueNumber = Convert.ToInt16(workSheet.Cells[i + 3, "A"].Text.ToString());
                    StaticClass.persons[i].LevelAccess = Convert.ToInt16(workSheet.Cells[i + 3, "B"].Text.ToString());
                    StaticClass.persons[i].Login = workSheet.Cells[i + 3, "C"].Text.ToString();
                    StaticClass.persons[i].Password = workSheet.Cells[i + 3, "D"].Text.ToString();
                    StaticClass.persons[i].FirstName = workSheet.Cells[i + 3, "E"].Text.ToString();
                    StaticClass.persons[i].FullName = workSheet.Cells[i + 3, "F"].Text.ToString();
                    StaticClass.persons[i].Email = workSheet.Cells[i + 3, "G"].Text.ToString();
                    StaticClass.persons[i].NumberPhone = workSheet.Cells[i + 3, "H"].Text.ToString();
                    StaticClass.persons[i].DataRegistration = workSheet.Cells[i + 3, "I"].Text.ToString();
                }
            }            

            excelApp.Workbooks.Close();
            excelApp.Quit();
            MessageBox.Show("Считывание данных прошло успешно, пользователей в системе: " + StaticClass.counterPersons, "Проверка", MessageBoxButtons.OK);
        }

        // функция для записывания информации в Excel
        void WriteInExcel(Person user, int counterUsers)
        {
            string path = Form1.path();
            Excel.Application excelApp = new Excel.Application();                   // создание ссылки на COM-оъект
            Excel.Workbook excelBook = excelApp.Workbooks.Open(path);        // открываем excel файл
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            workSheet.Cells[1, "A"] = Convert.ToString(counterUsers);
            workSheet.Cells[counterUsers + 2, "A"] = Convert.ToString(StaticClass.persons[counterUsers - 1].UniqueNumber);
            workSheet.Cells[counterUsers + 2, "B"] = Convert.ToString(StaticClass.persons[counterUsers - 1].LevelAccess);
            workSheet.Cells[counterUsers + 2, "C"] = Convert.ToString(StaticClass.persons[counterUsers - 1].Login);
            workSheet.Cells[counterUsers + 2, "D"] = Convert.ToString(StaticClass.persons[counterUsers - 1].Password);
            workSheet.Cells[counterUsers + 2, "E"] = Convert.ToString(StaticClass.persons[counterUsers - 1].FirstName);
            workSheet.Cells[counterUsers + 2, "F"] = Convert.ToString(StaticClass.persons[counterUsers - 1].FullName);
            workSheet.Cells[counterUsers + 2, "G"] = Convert.ToString(StaticClass.persons[counterUsers - 1].Email);
            workSheet.Cells[counterUsers + 2, "H"] = Convert.ToString(StaticClass.persons[counterUsers - 1].NumberPhone);
            workSheet.Cells[counterUsers + 2, "I"] = Convert.ToString(StaticClass.persons[counterUsers - 1].DataRegistration);

            excelApp.Workbooks.Close();
            excelApp.Quit();
        }

        // функция для изменения информации о пользователе в Excel
        void ChangeUserinExcel(int uniqueNumber)
        {
            string path = Form1.path();
            Excel.Application excelApp = new Excel.Application();                   // создание ссылки на COM-оъект
            Excel.Workbook excelBook = excelApp.Workbooks.Open(path);       // открываем excel файл
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            workSheet.Cells[uniqueNumber + 3, "A"] = Convert.ToString(StaticClass.persons[uniqueNumber].UniqueNumber);
            workSheet.Cells[uniqueNumber + 3, "B"] = Convert.ToString(StaticClass.persons[uniqueNumber].LevelAccess);
            workSheet.Cells[uniqueNumber + 3, "C"] = Convert.ToString(StaticClass.persons[uniqueNumber].Login);
            workSheet.Cells[uniqueNumber + 3, "D"] = Convert.ToString(StaticClass.persons[uniqueNumber].Password);
            workSheet.Cells[uniqueNumber + 3, "E"] = Convert.ToString(StaticClass.persons[uniqueNumber].FirstName);
            workSheet.Cells[uniqueNumber + 3, "F"] = Convert.ToString(StaticClass.persons[uniqueNumber].FullName);
            workSheet.Cells[uniqueNumber + 3, "G"] = Convert.ToString(StaticClass.persons[uniqueNumber].Email);
            workSheet.Cells[uniqueNumber + 3, "H"] = Convert.ToString(StaticClass.persons[uniqueNumber].NumberPhone);
            workSheet.Cells[uniqueNumber + 3, "I"] = Convert.ToString(StaticClass.persons[uniqueNumber].DataRegistration);

            excelApp.Workbooks.Close();
            excelApp.Quit();
        }

        // функция для удаления информации о пользователе в Excel
        void DeleteUserInExcel(int counterUser, int uniqueNumber)
        {
            string path = Form1.path();
            Excel.Application excelApp = new Excel.Application();                   // создание ссылки на COM-оъект
            Excel.Workbook excelBook = excelApp.Workbooks.Open(path);       // открываем excel файл
            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            workSheet.Cells[1, "A"] = Convert.ToString(counterUser);

            for (int i = uniqueNumber; i <= counterUser; i++)
            {
                workSheet.Cells[i + 3, "A"] = i;
                workSheet.Cells[i + 3, "B"] = workSheet.Cells[i + 4, "B"];
                workSheet.Cells[i + 3, "C"] = workSheet.Cells[i + 4, "C"];
                workSheet.Cells[i + 3, "D"] = workSheet.Cells[i + 4, "D"];
                workSheet.Cells[i + 3, "E"] = workSheet.Cells[i + 4, "E"];
                workSheet.Cells[i + 3, "F"] = workSheet.Cells[i + 4, "F"];
                workSheet.Cells[i + 3, "G"] = workSheet.Cells[i + 4, "G"];
                workSheet.Cells[i + 3, "H"] = workSheet.Cells[i + 4, "H"];
                workSheet.Cells[i + 3, "I"] = workSheet.Cells[i + 4, "I"];
            }

            workSheet.Cells[counterUser + 3, "A"].Clear();
            excelApp.Workbooks.Close();
            excelApp.Quit();
        }
    }
}