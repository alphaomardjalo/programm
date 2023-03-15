using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.IO;
//using Stergen_Model;


namespace Variant3
{
    public partial class Form2 : Form
    {
       // Model M = new Model();
        IniF Fini = new IniF("config.ini");
        public Form2()
        {
            InitializeComponent();
            if (File.Exists("config.ini")) //Если фаил конфигурации существует
            {
                textBox1.Text = Fini.ReadINI("TextBox1", "Save"); //Читает параметры
                textBox2.Text = Fini.ReadINI("TextBox2", "Save"); //Читает параметры
                textBox3.Text = Fini.ReadINI("TextBox3", "Save");
                textBox4.Text = Fini.ReadINI("TextBox4", "Save");
                textBox5.Text = Fini.ReadINI("TextBox5", "Save");
                textBox6.Text = Fini.ReadINI("TextBox6", "Save");
                textBox7.Text = Fini.ReadINI("TextBox7", "Save");
                textBox8.Text = Fini.ReadINI("TextBox8", "Save");
                textBox9.Text = Fini.ReadINI("TextBox9", "Save");
                textBox10.Text = Fini.ReadINI("TextBox10", "Save");
                textBox11.Text = Fini.ReadINI("TextBox11", "Save");
            }

        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Fini.WriteINI("TextBox1", "Save", textBox1.Text);
            Fini.WriteINI("TextBox2", "Save", textBox2.Text);
            Fini.WriteINI("TextBox3", "Save", textBox3.Text);
            Fini.WriteINI("TextBox4", "Save", textBox4.Text);
            Fini.WriteINI("TextBox5", "Save", textBox5.Text);
            Fini.WriteINI("TextBox6", "Save", textBox6.Text);
            Fini.WriteINI("TextBox7", "Save", textBox7.Text);
            Fini.WriteINI("TextBox8", "Save", textBox8.Text);
            Fini.WriteINI("TextBox9", "Save", textBox9.Text);
            Fini.WriteINI("TextBox10", "Save", textBox10.Text);
            Fini.WriteINI("TextBox11", "Save", textBox11.Text);

        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            Model_St Ob=new Model_St();
            Ob.L=Convert.ToDouble(textBox1.Text);
            Ob.B = Convert.ToDouble(textBox2.Text);
            Ob.H = Convert.ToDouble(textBox3.Text);
            Ob.V = Convert.ToDouble(textBox4.Text);
            Ob.E = Convert.ToDouble(textBox5.Text);
            Ob.Ro = Convert.ToDouble(textBox6.Text);
            Ob.N = Convert.ToDouble(textBox7.Text);
            Ob.T = Convert.ToDouble(textBox8.Text);
            double t,x,x1;
            x=Convert.ToDouble(textBox10.Text);
            x1 = Convert.ToDouble(textBox11.Text);
            t=Convert.ToDouble(textBox8.Text);
            int n;
            n=Convert.ToInt32(t/0.1);
            double[,] rez = new double[n, 2];
            rez = Ob.Rez(n, x);
            double[,] rez1= new double[n, 2];
            rez1 = Ob.Rez(n, x1);
            for(int i=0;i<n;i++)
            {
                dataGridView1.Rows.Add(rez[i,0], rez[i,1], rez1[i,1]);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //Построение графика
            chart1.Visible = true;
            
            double[] points = new double[500];
            chart1.Legends.Clear();
            chart1.Series[0].ChartType = SeriesChartType.Spline;
            Title title = new Title("График");
            title.Alignment = System.Drawing.ContentAlignment.BottomCenter;
            chart1.Titles.Add(title);
            chart1.ChartAreas[0].AxisX.Minimum = 0;
            chart1.ChartAreas[0].AxisX.Maximum = 50;//points.Length - 1;
            chart1.Series[0].Points.Clear();
            chart1.Series[0].Color = System.Drawing.Color.Blue;


            //------------------------------------

            Model_St Ob = new Model_St();
            Ob.L = Convert.ToDouble(textBox1.Text);
            Ob.B = Convert.ToDouble(textBox2.Text);
            Ob.H = Convert.ToDouble(textBox3.Text);
            Ob.V = Convert.ToDouble(textBox4.Text);
            Ob.E = Convert.ToDouble(textBox5.Text);
            Ob.Ro = Convert.ToDouble(textBox6.Text);
            Ob.N = Convert.ToDouble(textBox7.Text);
            Ob.T = Convert.ToDouble(textBox8.Text);
            double t, x, x1;
            x = Convert.ToDouble(textBox10.Text);
            x1 = Convert.ToDouble(textBox11.Text);
            t = Convert.ToDouble(textBox8.Text);
            int n;
            n = Convert.ToInt32(t / 0.1);
            double[,] rez = new double[n, 2];
            rez = Ob.Rez(n, x);
            for (int i = 0; i < n; i++)
            {
                    points[i] = rez[i,1];
                    chart1.Series[0].Points.AddXY(i, points[i]);                 
                }
            //--------------------------------------------
          
            double[,] rez1 = new double[n, 2];
            rez1 = Ob.Rez(n, x1);
            for (int i = 0; i < n; i++)
            {
                points[i] = rez1[i, 1];
                chart1.Series[0].Points[i].Color = Color.Red;
                chart1.Series[0].Points.AddXY(i, points[i]);
             }


            }

        private void button4_Click(object sender, EventArgs e)
        {
            Model_St Ob = new Model_St();
            Ob.L = Convert.ToDouble(textBox1.Text);
            Ob.B = Convert.ToDouble(textBox2.Text);
            Ob.H = Convert.ToDouble(textBox3.Text);
            Ob.V = Convert.ToDouble(textBox4.Text);
            Ob.E = Convert.ToDouble(textBox5.Text);
            Ob.Ro = Convert.ToDouble(textBox6.Text);
            Ob.N = Convert.ToDouble(textBox7.Text);
            Ob.T = Convert.ToDouble(textBox8.Text);
            double t, x;
            t = Convert.ToDouble(textBox8.Text);
            x = Convert.ToDouble(textBox10.Text);
            int n;
            n = Convert.ToInt32(t/0.1);
            Ob.WordDocument(n, x);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Model_St Ob = new Model_St();
            Ob.L = Convert.ToDouble(textBox1.Text);
            Ob.B = Convert.ToDouble(textBox2.Text);
            Ob.H = Convert.ToDouble(textBox3.Text);
            Ob.V = Convert.ToDouble(textBox4.Text);
            Ob.E = Convert.ToDouble(textBox5.Text);
            Ob.Ro = Convert.ToDouble(textBox6.Text);
            Ob.N = Convert.ToDouble(textBox7.Text);
            Ob.T = Convert.ToDouble(textBox8.Text);
            double t, x;
            t = Convert.ToDouble(textBox8.Text);
            x = Convert.ToDouble(textBox10.Text);
            int n;
            n = Convert.ToInt32(t / 0.1);
            Ob.ExcelTab(n, x);
        }

        private void button6_Click(object sender, EventArgs e)
        {
           /* Model_St Ob = new Model_St();
            Ob.L = Convert.ToDouble(textBox1.Text);
            Ob.B = Convert.ToDouble(textBox2.Text);
            Ob.H = Convert.ToDouble(textBox3.Text);
            Ob.V = Convert.ToDouble(textBox4.Text);
            Ob.E = Convert.ToDouble(textBox5.Text);
            Ob.Ro = Convert.ToDouble(textBox6.Text);
            Ob.N = Convert.ToDouble(textBox7.Text);
            Ob.T = Convert.ToDouble(textBox8.Text);
            double t, x;
            t = Convert.ToDouble(textBox8.Text);
            x = Convert.ToDouble(textBox10.Text);
            int n;
            n = Convert.ToInt32(t / 0.1);
            Ob.PowerP(n, x);*/
        }

        private void оПрограммеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Назначение программы - моделирование колебаний стержня");
        }

        private void оРазработчикеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Автор программы - студент университета");
        }

        private void назначениеКнопокToolStripMenuItem_Click(object sender, EventArgs e)
        {
            String Mes;
            Mes = "" + "Назначение кнопок на формах\n <Сохранить данные> - сохранить исходные данные в ini-файле\n" +
                " <Выполнить расчет> - рассчитать данные и  результат представить в таблице\n" +
                "<Построить график> - построить график по рассчитанным данным\n";
            MessageBox.Show(Mes);

        }

        private void button7_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            /*Form3 fm = new Form3();
            fm.Show();*/
        }

        private void расчетСтержневойСистемыToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Model_St Ob = new Model_St();
            Ob.L = Convert.ToDouble(textBox1.Text);
            Ob.B = Convert.ToDouble(textBox2.Text);
            Ob.H = Convert.ToDouble(textBox3.Text);
            Ob.V = Convert.ToDouble(textBox4.Text);
            Ob.E = Convert.ToDouble(textBox5.Text);
            Ob.Ro = Convert.ToDouble(textBox6.Text);
            Ob.N = Convert.ToDouble(textBox7.Text);
            Ob.T = Convert.ToDouble(textBox8.Text);
            double t, x, x1;
            x = Convert.ToDouble(textBox10.Text);
            x1 = Convert.ToDouble(textBox11.Text);
            t = Convert.ToDouble(textBox8.Text);
            int n;
            n = Convert.ToInt32(t / 0.1);
            double[,] rez = new double[n, 2];
            rez = Ob.Rez(n, x);
            double[,] rez1 = new double[n, 2];
            rez1 = Ob.Rez(n, x1);
            for (int i = 0; i < n; i++)
            {
                dataGridView1.Rows.Add(rez[i, 0], rez[i, 1], rez1[i, 1]);
            }
        }

        private void построитьГрафикToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            //Построение графика
            chart1.Visible = true;

            double[] points = new double[500];//{ 0.067, 0.045, 0.043, 0.040, 0.026, 0.015, 0.01, 0.008, 0.005, 0 };
            chart1.Legends.Clear();
            chart1.Series[0].ChartType = SeriesChartType.Spline;
            Title title = new Title("График");
            title.Alignment = System.Drawing.ContentAlignment.BottomCenter;
            chart1.Titles.Add(title);
            chart1.ChartAreas[0].AxisX.Minimum = 0;
            chart1.ChartAreas[0].AxisX.Maximum = 50;//points.Length - 1;
            chart1.Series[0].Points.Clear();
            chart1.Series[0].Color = System.Drawing.Color.Blue;


            //------------------------------------

            Model_St Ob = new Model_St();
            Ob.L = Convert.ToDouble(textBox1.Text);
            Ob.B = Convert.ToDouble(textBox2.Text);
            Ob.H = Convert.ToDouble(textBox3.Text);
            Ob.V = Convert.ToDouble(textBox4.Text);
            Ob.E = Convert.ToDouble(textBox5.Text);
            Ob.Ro = Convert.ToDouble(textBox6.Text);
            Ob.N = Convert.ToDouble(textBox7.Text);
            Ob.T = Convert.ToDouble(textBox8.Text);
            double t, x, x1;
            x = Convert.ToDouble(textBox10.Text);
            x1 = Convert.ToDouble(textBox11.Text);
            t = Convert.ToDouble(textBox8.Text);
            int n;
            n = Convert.ToInt32(t / 0.1);
            double[,] rez = new double[n, 2];
            rez = Ob.Rez(n, x);
            for (int i = 0; i < n; i++)
            {
                points[i] = rez[i, 1];
                chart1.Series[0].Points.AddXY(i, points[i]);
            }
            //--------------------------------------------

            double[,] rez1 = new double[n, 2];
            rez1 = Ob.Rez(n, x1);
            for (int i = 0; i < n; i++)
            {
                points[i] = rez1[i, 1];
                chart1.Series[0].Points[i].Color = Color.Red;
                chart1.Series[0].Points.AddXY(i, points[i]);
            }


        }

        private void сформироватьWordдокументToolStripMenuItem_Click(object sender, EventArgs e)
        {
           /* Model_St Ob = new Model_St();
            Ob.L = Convert.ToDouble(textBox1.Text);
            Ob.B = Convert.ToDouble(textBox2.Text);
            Ob.H = Convert.ToDouble(textBox3.Text);
            Ob.V = Convert.ToDouble(textBox4.Text);
            Ob.E = Convert.ToDouble(textBox5.Text);
            Ob.Ro = Convert.ToDouble(textBox6.Text);
            Ob.N = Convert.ToDouble(textBox7.Text);
            Ob.T = Convert.ToDouble(textBox8.Text);
            double t, x;
            t = Convert.ToDouble(textBox8.Text);
            x = Convert.ToDouble(textBox10.Text);
            int n;
            n = Convert.ToInt32(t / 0.1);
            Ob.WordDocument(n, x);*/
        }

        private void сформироватьExcelтаблицуToolStripMenuItem_Click(object sender, EventArgs e)
        {
          /*  Model_St Ob = new Model_St();
            Ob.L = Convert.ToDouble(textBox1.Text);
            Ob.B = Convert.ToDouble(textBox2.Text);
            Ob.H = Convert.ToDouble(textBox3.Text);
            Ob.V = Convert.ToDouble(textBox4.Text);
            Ob.E = Convert.ToDouble(textBox5.Text);
            Ob.Ro = Convert.ToDouble(textBox6.Text);
            Ob.N = Convert.ToDouble(textBox7.Text);
            Ob.T = Convert.ToDouble(textBox8.Text);
            double t, x;
            t = Convert.ToDouble(textBox8.Text);
            x = Convert.ToDouble(textBox10.Text);
            int n;
            n = Convert.ToInt32(t / 0.1);
            Ob.ExcelTab(n, x);*/
        }

        private void сформироватьПрезентациюToolStripMenuItem_Click(object sender, EventArgs e)
        {
            /*Model_St Ob = new Model_St();
            Ob.L = Convert.ToDouble(textBox1.Text);
            Ob.B = Convert.ToDouble(textBox2.Text);
            Ob.H = Convert.ToDouble(textBox3.Text);
            Ob.V = Convert.ToDouble(textBox4.Text);
            Ob.E = Convert.ToDouble(textBox5.Text);
            Ob.Ro = Convert.ToDouble(textBox6.Text);
            Ob.N = Convert.ToDouble(textBox7.Text);
            Ob.T = Convert.ToDouble(textBox8.Text);
            double t, x;
            t = Convert.ToDouble(textBox8.Text);
            x = Convert.ToDouble(textBox10.Text);
            int n;
            n = Convert.ToInt32(t / 0.1);
            Ob.PowerP(n, x);*/
        }

        private void анимацияToolStripMenuItem1_Click(object sender, EventArgs e)
        {
          /*  Form3 fm = new Form3();
            fm.Show();*/
       
        }

        private void выходToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void справкаToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
        }
    }

