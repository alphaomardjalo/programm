using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Variant3
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

                 
            Model_St Ob = new Model_St();
            Ob.L =3.5;
            Ob.B = 0.05;
            Ob.H = 0.05;
            Ob.V = 2;
            Ob.E = 2E+11;
            Ob.Ro =5000;
            Ob.N =25;
            Ob.T =50;
            double t, x;
            t = 50;
            x = 1.5;
            int n;
            n = Convert.ToInt32(t / 0.1);

            //Нарисовать квадрат
            int xx, y, xy;
            xx = 250; y = 150;
            double  h;
            h = 0.1;
            Graphics graphicsObj;
            graphicsObj =this.CreateGraphics();
            Pen myPen = new Pen(Color.Black,5);
            Rectangle myRectangle = new Rectangle(xx, y, 10, 10);
            graphicsObj.DrawRectangle(myPen, myRectangle);
            //------------------------------------------

            double[,] rez = new double[n, 2];
            rez=Ob.Rez(n,x);

            int j; j = 1;
            for (int i=0; i<n; i++)
            {
                
               xy=Convert.ToInt32(rez[i,1]*100);
               
               //myPen = new Pen(this.BackColor);
                myRectangle = new Rectangle(xx, y - xy, 10, 10);
                graphicsObj.DrawRectangle(myPen, myRectangle);
                
                System.Threading.Thread.Sleep(50);
               myPen = new Pen(this.BackColor, 5);
               myRectangle = new Rectangle(xx, y - xy, 10, 10);
                graphicsObj.DrawRectangle(myPen, myRectangle);
                myPen = new Pen(Color.Black, 5);
                j++;

            }
            myRectangle = new Rectangle(xx, y, 10, 10);
            graphicsObj.DrawRectangle(myPen, myRectangle);
         }

        private void button2_Click(object sender, EventArgs e)
        {
            Close();
        }

        }
    }

