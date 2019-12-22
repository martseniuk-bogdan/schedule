using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop;


namespace Schedule
{
    public partial class Form1 : Form
    {
        OleDbConnection con = new OleDbConnection();
        DataSet ds = new DataSet();
        static OleDbConnection myOleDbConnection = new OleDbConnection("Provider = Microsoft.Jet.OLEDB.4.0; Data Source = BD.mdb");
        static OleDbCommand myOleDbCommand = myOleDbConnection.CreateCommand();
        bool tmpBool = false;
        public DataSet DS
        {
            get { return ds; }
        }

        public OleDbConnection Con
        {
            get { return con; }
        }
        public Form1()
        {
            InitializeComponent();
            button2.Enabled = false;
            button3.Enabled = false;
            groupBox1.Visible = false;
            dataGridView2.Visible = false;
            button5.Enabled = false;
            groupBox2.Visible = false;
            groupBox3.Visible = false;
            textBox1.Visible = false;

        }

        //
        // Загрузка формы
        //
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {

                con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=BD.mdb";
                OleDbDataAdapter ad = new OleDbDataAdapter("SELECT * FROM groupp", con);
                ad.Fill(ds, "groupp");
                ad = new OleDbDataAdapter("SELECT * FROM cabinet", con);
                ad.Fill(ds, "cabinet");
                ad = new OleDbDataAdapter("SELECT * FROM dayOfTheWeek", con);
                ad.Fill(ds, "dayOfTheWeek");
                ad = new OleDbDataAdapter("SELECT * FROM lessons", con);
                ad.Fill(ds, "lessons");
                ad = new OleDbDataAdapter("SELECT * FROM lessonTime", con);
                ad.Fill(ds, "lessonTime");
                ad = new OleDbDataAdapter("SELECT * FROM schedule", con);
                ad.Fill(ds, "schedule");

                foreach (DataRow dr in DS.Tables["groupp"].Rows)
                    comboBoxGr.Items.Add(dr[1].ToString());
                foreach (DataRow dr in DS.Tables["dayOfTheWeek"].Rows)
                    comboBoxDW.Items.Add(dr[1].ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //
        // Запрос на выполнение
        //
        private void query(string s)
        {
            try
            {
                OleDbCommand com = new OleDbCommand(s, con);
                con.Open();
                OleDbDataReader read = com.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(read);
                dataGridView1.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally { if (con.State == ConnectionState.Open) con.Close(); }

        }


        // 
        // Просмотр рассписания группы
        // 
        private void button1_Click(object sender, EventArgs e)
        {
            string tmp = comboBoxGr.SelectedItem.ToString();
            string s = "SELECT dayOfTheWeek.dayoftheweek as 'День недели', lessons.name_of_lesson as 'Лекция', lessons.teachers as 'Преподователь', cabinet.number_of_cab as 'Кабинет', lessonTime.start_time as 'Начало' FROM lessonTime INNER JOIN(lessons INNER JOIN (groupp INNER JOIN (dayOfTheWeek INNER JOIN (cabinet INNER JOIN schedule ON cabinet.id_cab = schedule.id_cabinet) ON dayOfTheWeek.id_days = schedule.id_dayOfTheWeek) ON groupp.id_gr = schedule.id_groupp) ON lessons.id_less = schedule.id_lessons) ON lessonTime.id_lessTime = schedule.id_lessonstime WHERE (((groupp.name_of_group)='" + tmp + "'))";
            query(s);

            button2.Enabled = true;
            button5.Enabled = true;
        }


        //
        // Генерировать отчет рассписания группы на неделю
        // 
        private void button2_Click(object sender, EventArgs e)
        {

            string tmp = comboBoxGr.SelectedItem.ToString();

            try
            {
                object ms = Missing.Value;
                //
                // Создание приложения Excel
                //
                Excel.Application app = new Excel.Application();
                app.Visible = false;
                //
                // Создание нового документа
                //
                Excel.Workbook book = app.Workbooks.Add();
                // Страница рассписания группы
                Excel.Worksheet sheetFlats = book.Worksheets[1];
                sheetFlats.Name = "Рассписание группы № " + tmp;

                DataTable Schedule = new DataTable("Рассписание выбранной группы");
                // День недели
                DataColumn flatsColumn = new DataColumn();
                flatsColumn.DataType = Type.GetType("System.String");
                flatsColumn.ColumnName = "День недели";
                Schedule.Columns.Add(flatsColumn);
                // Лекция
                flatsColumn = new DataColumn();
                flatsColumn.DataType = Type.GetType("System.String");
                flatsColumn.ColumnName = "Лекция";
                Schedule.Columns.Add(flatsColumn);
                // Преподователь
                flatsColumn = new DataColumn();
                flatsColumn.DataType = Type.GetType("System.String");
                flatsColumn.ColumnName = "Преподователь";
                Schedule.Columns.Add(flatsColumn);
                // Кабинет
                flatsColumn = new DataColumn();
                flatsColumn.DataType = Type.GetType("System.Int32");
                flatsColumn.ColumnName = "Кабинет";
                Schedule.Columns.Add(flatsColumn);
                // Начало
                flatsColumn = new DataColumn();
                flatsColumn.DataType = Type.GetType("System.String");
                flatsColumn.ColumnName = "Начало";
                Schedule.Columns.Add(flatsColumn);

                //
                // Очистка таблицы
                //
                Schedule.Clear();
                //
                // Запрос к БД на выборку данных о рассписании за неделю
                //
                myOleDbCommand.CommandText = "SELECT dayOfTheWeek.dayoftheweek as 'День недели', lessons.name_of_lesson as 'Лекция', lessons.teachers as 'Преподователь', cabinet.number_of_cab as 'Кабинет', lessonTime.start_time as 'Начало' FROM lessonTime INNER JOIN(lessons INNER JOIN (groupp INNER JOIN (dayOfTheWeek INNER JOIN (cabinet INNER JOIN schedule ON cabinet.id_cab = schedule.id_cabinet) ON dayOfTheWeek.id_days = schedule.id_dayOfTheWeek) ON groupp.id_gr = schedule.id_groupp) ON lessons.id_less = schedule.id_lessons) ON lessonTime.id_lessTime = schedule.id_lessonstime WHERE (((groupp.name_of_group)='" + tmp + "'))";
                myOleDbConnection.Open();
                OleDbDataReader myOleDbDataReader = myOleDbCommand.ExecuteReader();
                while (myOleDbDataReader.Read())
                {
                    Schedule.Rows.Add(myOleDbDataReader.GetString(0), myOleDbDataReader.GetString(1), myOleDbDataReader.GetString(2), myOleDbDataReader.GetInt32(3), myOleDbDataReader.GetString(4));
                }
                myOleDbDataReader.Close();
                myOleDbConnection.Close();
                //
                // Вывод информации о рассписании
                //
                sheetFlats.Range["A1"].Value = "День недели";
                sheetFlats.Range["B1"].Value = "Лекция";
                sheetFlats.Range["C1"].Value = "Преподователь";
                sheetFlats.Range["D1"].Value = "Кабинет";
                sheetFlats.Range["E1"].Value = "Начало";

                sheetFlats.get_Range("A1:E1").Font.Color = Color.Green;
                for (int i = 0; i < Schedule.Rows.Count; i++)
                {
                    for (int j = 0; j < Schedule.Columns.Count; j++)
                    {
                        sheetFlats.Cells[i + 2, j + 1] = Schedule.Rows[i][j].ToString();
                        sheetFlats.Columns.EntireColumn.AutoFit();
                    }

                }
                sheetFlats.Columns.EntireColumn.AutoFit();
                // Видимость документа Excel
                app.Visible = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }


        //
        // Просмотр рассписания на день недели
        //
        private void button4_Click(object sender, EventArgs e)
        {

            string tmp2 = comboBoxDW.SelectedItem.ToString();
            string s = "SELECT groupp.name_of_group as Группа, lessons.name_of_lesson as Предмет, lessons.teachers as Преподователь, cabinet.number_of_cab as Кабинет, lessonTime.start_time as Начало FROM lessonTime INNER JOIN(lessons INNER JOIN (groupp INNER JOIN (dayOfTheWeek INNER JOIN (cabinet INNER JOIN schedule ON cabinet.id_cab = schedule.id_cabinet) ON dayOfTheWeek.id_days = schedule.id_dayOfTheWeek) ON groupp.id_gr = schedule.id_groupp) ON lessons.id_less = schedule.id_lessons) ON lessonTime.id_lessTime = schedule.id_lessonstime WHERE(((dayOfTheWeek.dayoftheweek) ='" + tmp2 + "'))";
            query(s);

            button3.Enabled = true;
            button5.Enabled = false;
            groupBox1.Visible = false;
        }


        //
        // Генерировать отчет рассписания на день недели
        //
        private void button3_Click(object sender, EventArgs e)
        {

            string tmp = comboBoxDW.SelectedItem.ToString();

            try
            {
                object ms = Missing.Value;
                //
                // Создание приложения Excel
                //
                Excel.Application app = new Excel.Application();
                app.Visible = false;
                //
                // Создание нового документа
                //
                Excel.Workbook book = app.Workbooks.Add();
                // Страница рассписания группы
                Excel.Worksheet sheetFlats = book.Worksheets[1];
                sheetFlats.Name = "Рассписание дня " + tmp;

                DataTable Schedule = new DataTable("Рассписание дня недели");
                // День недели
                DataColumn flatsColumn = new DataColumn();
                flatsColumn.DataType = Type.GetType("System.String");
                flatsColumn.ColumnName = "Группа";
                Schedule.Columns.Add(flatsColumn);
                // Лекция
                flatsColumn = new DataColumn();
                flatsColumn.DataType = Type.GetType("System.String");
                flatsColumn.ColumnName = "Лекция";
                Schedule.Columns.Add(flatsColumn);
                // Преподователь
                flatsColumn = new DataColumn();
                flatsColumn.DataType = Type.GetType("System.String");
                flatsColumn.ColumnName = "Преподователь";
                Schedule.Columns.Add(flatsColumn);
                // Кабинет
                flatsColumn = new DataColumn();
                flatsColumn.DataType = Type.GetType("System.Int32");
                flatsColumn.ColumnName = "Кабинет";
                Schedule.Columns.Add(flatsColumn);
                // Начало
                flatsColumn = new DataColumn();
                flatsColumn.DataType = Type.GetType("System.String");
                flatsColumn.ColumnName = "Начало";
                Schedule.Columns.Add(flatsColumn);

                //
                // Очистка таблицы
                //
                Schedule.Clear();
                //
                // Запрос к БД на выборку данных о рассписании за день недели
                //
                myOleDbCommand.CommandText = "SELECT groupp.name_of_group as Группа, lessons.name_of_lesson as Предмет, lessons.teachers as Преподователь, cabinet.number_of_cab as Кабинет, lessonTime.start_time as Начало FROM lessonTime INNER JOIN(lessons INNER JOIN (groupp INNER JOIN (dayOfTheWeek INNER JOIN (cabinet INNER JOIN schedule ON cabinet.id_cab = schedule.id_cabinet) ON dayOfTheWeek.id_days = schedule.id_dayOfTheWeek) ON groupp.id_gr = schedule.id_groupp) ON lessons.id_less = schedule.id_lessons) ON lessonTime.id_lessTime = schedule.id_lessonstime WHERE(((dayOfTheWeek.dayoftheweek) ='" + tmp + "'))";
                myOleDbConnection.Open();
                OleDbDataReader myOleDbDataReader = myOleDbCommand.ExecuteReader();
                while (myOleDbDataReader.Read())
                {
                    Schedule.Rows.Add(myOleDbDataReader.GetString(0), myOleDbDataReader.GetString(1), myOleDbDataReader.GetString(2), myOleDbDataReader.GetInt32(3), myOleDbDataReader.GetString(4));
                }
                myOleDbDataReader.Close();
                myOleDbConnection.Close();

                //
                // Вывод информации 
                //
                sheetFlats.Range["A1"].Value = "Группа";
                sheetFlats.Range["B1"].Value = "Лекция";
                sheetFlats.Range["C1"].Value = "Преподователь";
                sheetFlats.Range["D1"].Value = "Кабинет";
                sheetFlats.Range["E1"].Value = "Начало";

                sheetFlats.get_Range("A1:E1").Font.Color = Color.Green;
                for (int i = 0; i < Schedule.Rows.Count; i++)
                {
                    for (int j = 0; j < Schedule.Columns.Count; j++)
                    {
                        sheetFlats.Cells[i + 2, j + 1] = Schedule.Rows[i][j].ToString();
                        sheetFlats.Columns.EntireColumn.AutoFit();
                    }

                }
                sheetFlats.Columns.EntireColumn.AutoFit();
                // Видимость документа Excel
                app.Visible = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }
        
        
        //
        // Кнопка открытия меню редактирования
        //
        private void button5_Click(object sender, EventArgs e)
        {
            groupBox1.Visible = true;
            groupBox2.Visible = false;
            groupBox3.Visible = false;
            comboBox1.Items.Clear();
            comboBox2.Items.Clear();
            comboBox3.Items.Clear();
            comboBox4.Items.Clear();

            string Cells1 = Convert.ToString(dataGridView1.CurrentRow.Cells[0].Value);
            string Cells2 = Convert.ToString(dataGridView1.CurrentRow.Cells[1].Value);
            string Cells3 = Convert.ToString(dataGridView1.CurrentRow.Cells[2].Value);
            string Cells4 = Convert.ToString(dataGridView1.CurrentRow.Cells[3].Value);
            string Cells5 = Convert.ToString(dataGridView1.CurrentRow.Cells[4].Value);
            

            label3.Text = Cells1;
            foreach (DataRow dr in DS.Tables["dayOfTheWeek"].Rows)
                comboBox1.Items.Add(dr[1].ToString());

            label4.Text = "Предмет " + Cells2;
            foreach (DataRow dr in DS.Tables["lessons"].Rows)
                comboBox2.Items.Add(dr[1].ToString());


            label6.Text = "Кабинет " + Cells4;
            foreach (DataRow dr in DS.Tables["cabinet"].Rows)
                comboBox3.Items.Add(dr[1].ToString());

            label7.Text = "Начало " + Cells5;
            foreach (DataRow dr in DS.Tables["lessonTime"].Rows)
                comboBox4.Items.Add(dr[1].ToString());

            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;
            comboBox4.SelectedIndex = 0;
        }      

        //
        // Кнопка отмены редактирование
        //
        private void button10_Click(object sender, EventArgs e)
        {
            groupBox1.Visible = false;
        }


        //
        // Кнопка применения редактирования
        //
        private void button6_Click(object sender, EventArgs e)
        {
            tmpBool = false; //пеменная для проверки "можно ли создать/отредактировать данное поле"

            //
            // Исходные уникальные ключи
            //
            queryFor2("Select * From dayOfTheWeek Where dayoftheweek='" + Convert.ToString(dataGridView1.CurrentRow.Cells[0].Value) + "'");
            int dayOfTheWeekSource = Convert.ToInt32(dataGridView2[0, 0].Value);

            queryFor2("Select * From lessons Where name_of_lesson='" + Convert.ToString(dataGridView1.CurrentRow.Cells[1].Value) + "'");
            int lessonsSource = Convert.ToInt32(dataGridView2[0, 0].Value);

            queryFor2("Select * From cabinet Where number_of_cab=" + Convert.ToString(dataGridView1.CurrentRow.Cells[3].Value) + "");
            int cabinetSource = Convert.ToInt32(dataGridView2[0, 0].Value);

            queryFor2("Select * From lessonTime Where start_time='" + Convert.ToString(dataGridView1.CurrentRow.Cells[4].Value) + "'");
            int startSource = Convert.ToInt32(dataGridView2[0, 0].Value);



            //
            // Получаем уникальные ключи выбранных полей, на которые будем менять 
            //
            queryFor2("Select * From dayOfTheWeek Where dayoftheweek='" + comboBox1.SelectedItem.ToString() + "'");
            int dayOfTheWeek = Convert.ToInt32(dataGridView2[0, 0].Value);

            queryFor2("Select * From lessons Where name_of_lesson='" + comboBox2.SelectedItem.ToString() + "'");
            int lessons = Convert.ToInt32(dataGridView2[0, 0].Value);

            queryFor2("Select * From cabinet Where number_of_cab=" + comboBox3.SelectedItem.ToString() + "");
            int cabinet = Convert.ToInt32(dataGridView2[0, 0].Value);

            queryFor2("Select * From lessonTime Where start_time='" + comboBox4.SelectedItem.ToString() + "'");
            int start = Convert.ToInt32(dataGridView2[0, 0].Value);

            queryFor2("Select * From groupp Where name_of_group='" + comboBoxGr.SelectedItem.ToString() + "'");
            int name_of_group = Convert.ToInt32(dataGridView2[0, 0].Value);

            string findID = "SELECT id_schedule FROM schedule WHERE (id_groupp=" + name_of_group.ToString() + " AND id_dayOfTheWeek=" + dayOfTheWeekSource.ToString() + " AND id_lessonstime=" + startSource.ToString() + " AND id_lessons = " + lessonsSource.ToString() + " AND id_cabinet=" + cabinetSource.ToString() + ")";
            queryFor2(findID);
            String ID= Convert.ToString(dataGridView2[0, 0].Value);
            textBox1.Text = ID;


            //
            // Для вызова метода проверки "свободна ли аудитория в выбраное время в выбранный день недели"
            //
            string s = "SELECT count(*) FROM dayOfTheWeek INNER JOIN(lessonTime INNER JOIN(cabinet INNER JOIN schedule ON cabinet.id_cab = schedule.id_cabinet) ON lessonTime.id_lessTime = schedule.id_lessonstime) ON dayOfTheWeek.id_days = schedule.id_dayOfTheWeek WHERE(((lessonTime.start_time) ='" + comboBox4.SelectedItem.ToString() + "') AND((cabinet.number_of_cab) =" + comboBox3.SelectedItem.ToString() + ") AND((dayOfTheWeek.dayoftheweek) ='" + comboBox1.SelectedItem.ToString() + "'))";
            string mbS = "В этом кабинете уже есть занятие в это время! ";
            NewMethod(s, mbS);

            //
            // Проверить нет ли такой записи в рассписании
            //
            string s2 = "Select count(*) From schedule WHERE (id_groupp=" + name_of_group.ToString() + " AND id_dayOfTheWeek=" + dayOfTheWeek.ToString() + " AND id_lessonstime=" + start.ToString() + " AND id_lessons=" + lessons.ToString() + " AND id_cabinet=" + cabinet.ToString() + ")";
            string mbS2 = "Такая запись в рассписании присутствует!";
            NewMethod(s2, mbS2);

            //
            // Проверить нет ли этого предмета в это время у другой группы
            //
            string s3 = "Select count(*) From schedule WHERE (id_lessons=" + lessons.ToString() + " AND id_lessonstime=" + start.ToString() + " AND id_dayOfTheWeek=" + dayOfTheWeek.ToString() + ")";
            string mbS3 = "Этот предмет у другой группы в это время!";
            NewMethod(s3, mbS3);

            try
            {
                if (!tmpBool)
                {

                    Con.Open();
                    string query1 = "UPDATE schedule SET [id_groupp]=" + name_of_group.ToString()
                           + ", [id_dayOfTheWeek]=" + dayOfTheWeek.ToString()
                           + ", [id_lessonstime]=" + start.ToString()
                           + ", [id_lessons]=" + lessons.ToString()
                           + ", [id_cabinet]=" + cabinet.ToString()
                           + " WHERE [id_schedule]=" + ID + "";
                    OleDbCommand com = new OleDbCommand(query1, Con);
                    com.ExecuteNonQuery();
                    MessageBox.Show("Запись обновлена!");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally { if (Con.State == ConnectionState.Open) Con.Close();}
        }
       
        
        //
        // Запрос на выполнение
        //
        private void queryFor2(string s)
        {
            try
            {
                OleDbCommand com = new OleDbCommand(s, con);
                con.Open();
                OleDbDataReader read = com.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(read);
                dataGridView2.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally { if (con.State == ConnectionState.Open) con.Close(); }

        }

        //
        // Метод проверки 
        //
        private bool NewMethod(string s, string mbS)
        {
            try
            {
                OleDbCommand com = new OleDbCommand(s, con);

                con.Open();

                OleDbDataReader read = com.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(read);
                dataGridView2.DataSource = dt;
                if (Convert.ToInt32(dataGridView2[0, 0].Value) != 0)
                {
                    MessageBox.Show(mbS);
                    tmpBool = true;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally { if (con.State == ConnectionState.Open) con.Close(); }
            return tmpBool;
        }


        //
        // Открыть меню добавления предмета
        //
        private void button7_Click(object sender, EventArgs e)
        {
            groupBox2.Visible = true;
            groupBox1.Visible = false;
            groupBox3.Visible = false;
        }        
        
        //
        // Отмена добавления предмета
        //
        private void button9_Click(object sender, EventArgs e)
        {
            textBox2.Text = "";
            textBox3.Text = "";
            groupBox2.Visible = false;
        }

        //
        // Добавление предмета
        //
        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                if (textBox2.TextLength==0 || textBox3.TextLength == 0)
                {
                    MessageBox.Show("Пожалуйста, введите название и ФИО!");
                }
                else
                {
                    Con.Open();
                    string query = "INSERT INTO lessons ([name_of_lesson], [teachers]) "
                    + "VALUES"
                    + "('" + textBox2.Text
                    + "', '" + textBox3.Text
                    + "')";
                    OleDbCommand com = new OleDbCommand(query, Con);
                    com.ExecuteNonQuery();
                    if (Con.State == ConnectionState.Open) Con.Close();
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally { if (Con.State == ConnectionState.Open) Con.Close(); }
        }

        //
        // Открыть меню - Добавить запись
        //
        private void button11_Click(object sender, EventArgs e)
        {
            comboBox5.Items.Clear();
            comboBox6.Items.Clear();
            comboBox7.Items.Clear();
            comboBox8.Items.Clear();
            comboBox9.Items.Clear();
            groupBox3.Visible = true;
            groupBox1.Visible = false;
            groupBox2.Visible = false;
            // Заполняем комбобоксы
            foreach (DataRow dr in DS.Tables["groupp"].Rows)
                comboBox5.Items.Add(dr[1].ToString());

            foreach (DataRow dr in DS.Tables["dayOfTheWeek"].Rows)
                comboBox6.Items.Add(dr[1].ToString());

            foreach (DataRow dr in DS.Tables["lessons"].Rows)
                comboBox7.Items.Add(dr[1].ToString());

            foreach (DataRow dr in DS.Tables["cabinet"].Rows)
                comboBox8.Items.Add(dr[1].ToString());

            foreach (DataRow dr in DS.Tables["lessonTime"].Rows)
                comboBox9.Items.Add(dr[1].ToString());

            // Присваиваем значение по умолчанию
            comboBox5.SelectedIndex = 0;
            comboBox6.SelectedIndex = 0;
            comboBox7.SelectedIndex = 0;
            comboBox8.SelectedIndex = 0;
            comboBox9.SelectedIndex = 0;
        }

        //
        // Кнопка для создания новой записи
        //
        private void button12_Click(object sender, EventArgs e)
        {
            tmpBool = false; //пеменная для проверки "можно ли создать/отредактировать данное поле"
            //
            // Получаем уникальные ключи выбранных полей
            //
            queryFor2("Select * From groupp Where name_of_group='" + comboBox5.SelectedItem.ToString() + "'");
            int name_of_group = Convert.ToInt32(dataGridView2[0, 0].Value);

            queryFor2("Select * From dayOfTheWeek Where dayoftheweek='" + comboBox6.SelectedItem.ToString() + "'");
            int dayOfTheWeek = Convert.ToInt32(dataGridView2[0, 0].Value);

            queryFor2("Select * From lessons Where name_of_lesson='" + comboBox7.SelectedItem.ToString() + "'");
            int lessons = Convert.ToInt32(dataGridView2[0, 0].Value);

            queryFor2("Select * From cabinet Where number_of_cab=" + comboBox8.SelectedItem.ToString() + "");
            int cabinet = Convert.ToInt32(dataGridView2[0, 0].Value);

            queryFor2("Select * From lessonTime Where start_time='" + comboBox9.SelectedItem.ToString() + "'");
            int start = Convert.ToInt32(dataGridView2[0, 0].Value);
           
            //
            // Проверить нет ли занятия в этом кабинете в это время
            //
            string s = "SELECT count(*) FROM dayOfTheWeek INNER JOIN(lessonTime INNER JOIN(cabinet INNER JOIN schedule ON cabinet.id_cab = schedule.id_cabinet) ON lessonTime.id_lessTime = schedule.id_lessonstime) ON dayOfTheWeek.id_days = schedule.id_dayOfTheWeek WHERE(((lessonTime.start_time) ='" + comboBox9.SelectedItem.ToString() + "') AND((cabinet.number_of_cab) =" + comboBox8.SelectedItem.ToString() + ") AND((dayOfTheWeek.dayoftheweek) ='" + comboBox6.SelectedItem.ToString() + "'))";
            string mbS = "В этом кабинете уже есть занятие в это время! ";
            NewMethod(s, mbS);

            //
            // Проверить нет ли такой записи в рассписании
            //
            string s2 = "Select count(*) From schedule WHERE (id_groupp=" + name_of_group.ToString() + " AND id_dayOfTheWeek=" + dayOfTheWeek.ToString() + " AND id_lessonstime=" + start.ToString() + " AND id_lessons=" + lessons.ToString() + " AND id_cabinet=" + cabinet.ToString() + ")";
            string mbS2 = "Такая запись в рассписании присутствует!";
            NewMethod(s2, mbS2);


            //
            // Проверить нет ли у группы в это время урока в другом кабинете
            //
            string s4 = "Select count(*) From schedule WHERE (id_groupp=" + name_of_group.ToString() + " AND id_lessonstime=" + start.ToString()+ " AND id_dayOfTheWeek=" + dayOfTheWeek.ToString() + ")";
            string mbS4 = "У этой группы в это время другой урок!";
            NewMethod(s4, mbS4);

            try
            {
                    if (!tmpBool)
                    {
                        Con.Open();
                        string query = "INSERT INTO schedule ([id_groupp], [id_dayOfTheWeek], [id_lessonstime], [id_lessons], [id_cabinet]) "
                        + "VALUES"
                        + "('" + name_of_group.ToString()
                        + "', '" + dayOfTheWeek.ToString()
                        + "', '" + start.ToString()
                        + "', '" + lessons.ToString()
                        + "', '" + cabinet.ToString()
                        + "')";
                        OleDbCommand com = new OleDbCommand(query, Con);
                        com.ExecuteNonQuery();
                        if (Con.State == ConnectionState.Open) Con.Close();
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally { if (Con.State == ConnectionState.Open) Con.Close(); }

        }

        //
        // Отменить создание новой записи
        //
        private void button13_Click(object sender, EventArgs e)
        {
            groupBox3.Visible = false;
        }
    }
}

