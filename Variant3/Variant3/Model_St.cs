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
using PowerPoint=Microsoft.Office.Interop.PowerPoint;
using System.IO;



namespace Variant3
{
    class Model_St
    {
        public double v; //Исходные данные скорость
        public double l;// длина стержня
        public double b;//ширина стержня
        public double h;//высота стержня
        public double e;//модуль упругости
        public double ro;//плотность материала
        public double n;// количество точек
        public double t;//время моделирования
        //public double x1;//первая точка


        //public double t;
        //-----------------------------
        //Специальные методы доступа- свойства
        public double V { set { v = value; } }
        public double L { set { l = value; } }
        public double B { set { b = value; } }
        public double H { set { h = value; } }
        public double E { set { e = value; } }
        public double Ro { set { ro = value; } }
        public double N { set { n = value; } }
        public double T { set { t = value; } }
        //public double X1 { set { x1 = value; } }




        //------------------------------
        public double Y(double t, double x) //функция  расчета y 
        {
            double y;
            double p, I, f;
            double sum = 0;
            f = b * h;
            I = b * h * h * h / 12;
            for (int i = 1; i < n; i++)
            {
                p = (i * i * Math.PI * Math.PI / (l * l)) * Math.Sqrt(e * I / (ro * f));//расчет напряжения
                sum += (1 / (i * p)) * Math.Sin(i * Math.PI * x / l) * Math.Sin(p * t);
            }

            y = (4 * v / Math.PI) * sum;//расчет перемещения

            return y;
        }

        public double[,] Rez(int n1, double x)
        {
            double h;
            h = 0.1;
            double[,] rez = new double[n1, 2];
            int i = 0;
            for (double tt = 0; tt < t; tt = tt + h)
            {
                rez[i, 0] = tt;
                rez[i, 1] = Y(tt, x);
                i++;

            }
            return rez;
        }


        public void WordDocument(int n1, double x)
        {
            // Создание документ Word.
            Word.Application word_app = new Word.Application();
            word_app.Visible = true;
            // Создаем документ Word.
            object missing = Type.Missing;//Создание документа 
            Word._Document word_doc = word_app.Documents.Add(ref missing, ref missing, ref missing, ref missing);
            Word.Paragraph para = word_doc.Paragraphs.Add(ref missing);// Создание абзаца заголовка
            para.Range.Text = "Моделирование колебаний стержня";
            object style_name = "Заголовок 1";
            para.Range.set_Style(ref style_name);
            para.Range.InsertParagraphAfter();
            para.Range.Text = "t            x";// Добавление текста
            para.Range.InsertParagraphAfter();
            string old_font = para.Range.Font.Name;//Установление шрифта
            para.Range.Font.Name = "Courier New";
            // Расчет перемещений и добавление их к документу
            double h;
            h = 0.1;
            double[,] rez = new double[n1, 2];
            int i = 0;
            for (double tt = 0; tt < t; tt = tt + h)
            {
                rez[i, 0] = tt;
                rez[i, 1] = Y(tt, x);

                Math.Round(tt, 1);
                para.Range.Text = tt.ToString() + "   " + rez[i, 1].ToString();
                para.Range.InsertParagraphAfter();// Перейти новому абзацу
                i++;
            }
            para.Range.Font.Name = old_font; //усттановит исходный шрифт


        }

        public void ExcelTab(int n1, double x)
        {
            // Загрузка Excel, создание новой  книги
            Excel.Application excelApp = new Excel.Application();
            // Сделать приложение Excel видимым
            excelApp.Visible = true;
            excelApp.Workbooks.Add();
            Excel._Worksheet workSheet = excelApp.ActiveSheet;
            // Установить заголовки столбцов в ячейках
            bool retry = true;
            do
            {
                try
                {
                    //Заполнить ячейки  в ексель - шапка таблицы  
                    workSheet.Cells[1, "A"] = "t";
                    workSheet.Cells[1, "B"] = "y";

                    double  h;
                    h = 0.1;
                    int i;
                    i = 2;
                          
                    double[,] rez = new double[n1, 2];
                    for (double tt = 0; tt < 10; tt = tt + h)
                    {
                        rez[i, 0] = tt;
                        rez[i, 1] = Y(tt, x);

                        Math.Round(tt, 1); ;
                        
                        //заполнение ячеек таблицы
                        workSheet.Cells[i, "A"] = tt;
                        workSheet.Cells[i, "B"] = rez[i, 1];
                        i++;
                    }
                    retry = false;
                }
                catch (Exception exp)
                {
                    System.Threading.Thread.Sleep(10);
                }
            } while (retry);
        }
        //----------------------------------------------
        public void PowerP(int n1, double x)
        {

            PowerPoint.Application pptApp = new PowerPoint.Application();
            pptApp = new PowerPoint.Application();
            PowerPoint.Presentation presentation;
            PowerPoint.Presentations presentations;
            Microsoft.Office.Interop.PowerPoint.Slides slides;
            Microsoft.Office.Interop.PowerPoint._Slide slide;
            Microsoft.Office.Interop.PowerPoint.TextRange objText;


            // Создание  File презентации
            PowerPoint.Presentation pptPresentation = pptApp.Presentations.Add();
            Microsoft.Office.Interop.PowerPoint.CustomLayout customLayout = pptPresentation.SlideMaster.CustomLayouts[Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutText];
            // Создание слайда
            slides = pptPresentation.Slides;
            slide = slides.AddSlide(1, customLayout);
            // Добавление заголовка слайда
            objText = slide.Shapes[1].TextFrame.TextRange;
            objText.Text = "Моделирование колебаний стержня";
            objText.Font.Name = "Arial";
            objText.Font.Size = 30;
            double h;
            h = 0.1;
            int i;
            i = 2;
            string Res;
            Res = "";
            double[,] rez = new double[n1, 2];
            for (double tt = 0; tt < 0.5; tt = tt + h)
            {
                // rez[i, 0] = tt;
                rez[i, 1] = Y(tt, x);

                Math.Round(tt, 1); ; Math.Round(t, 1);
                Res = Res + "\n" + tt.ToString() + "   " + rez[i, 1].ToString();
                i++;
            }
            objText = slide.Shapes[2].TextFrame.TextRange;
            objText.Text = "t     x" + Res;
        }
        



    }
}