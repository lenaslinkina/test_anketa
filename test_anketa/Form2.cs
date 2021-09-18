using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace test_anketa
{
    public partial class Form2 : Form
    {
        int n = 0;
        int[] answer;
       
        double incorr1 = 0, incorr2 = 0, incorr3 = 0, incorr4 = 0, incorr5 = 0, incorr6 = 0;
        double corr1 = 0, corr2 = 0, corr3 = 0, corr4 = 0, corr5 = 0, corr6;

        

        

        public Form2()
        {
            InitializeComponent();
            answer = new int[6];

        }
        public void questions (int n)
        {
            switch (answer[n])
            {
                case 0:
                    radioButton1.Checked = false;
                    radioButton2.Checked = false;
                    radioButton3.Checked = false;
                    break;
                case 1:
                    radioButton1.Checked = true;
                    radioButton2.Checked = false;
                    radioButton3.Checked = false;
                    break;
                case 2:
                    radioButton1.Checked = false;
                    radioButton2.Checked = true;
                    radioButton3.Checked = false;
                    break;
                case 3:
                    radioButton1.Checked = false;
                    radioButton2.Checked = false;
                    radioButton3.Checked = true;
                    break;
            }




            switch (n)
            {
                case 0:
                    label1.Text = "Вторая планета Солнечной системы";
                    radioButton1.Text = "Венера";
                    radioButton2.Text = "Меркурий";
                    radioButton3.Text = "Земля";
                    //checkBox1.Text = "Венера";
                    //checkBox2.Text = "Меркурий";
                    //checkBox3.Text = "Земля";

                    break;

                case 1:
                    label1.Text = "Число 27 в двоичной системе исчисления";
                    radioButton1.Text = "111000";
                    radioButton2.Text = "101010";
                    radioButton3.Text = "11011";
                    //checkBox1.Text = "111000";
                    //checkBox2.Text = "101010";
                    //checkBox3.Text = "11011";

                    break;
                case 2:
                    label1.Text = "Примерное количество людей на Земле";
                    radioButton1.Text = "7 млрд.";
                    radioButton2.Text = "10 млрд.";
                    radioButton3.Text = "5 млрд.";
                    //checkBox1.Text = "7 млрд.";
                    //checkBox2.Text = "10 млрд.";
                    //checkBox3.Text = "5 млрд.";

                    break;
                case 3:
                    label1.Text = "Кто написал «Сказка о царе Салтане»";
                    radioButton1.Text = "Лермонтов";
                    radioButton2.Text = "Пушкин";
                    radioButton3.Text = "Некрасов";
                    //checkBox1.Text = "Лермонтов";
                    //checkBox2.Text = "Пушкин";
                    //checkBox3.Text = "Некрасов";

                    break;
                case 4:
                    label1.Text = "Сколько граней у куба?";
                    radioButton1.Text = "6";
                    radioButton2.Text = "8";
                    radioButton3.Text = "12";
                    //checkBox1.Text = "6";
                    //checkBox2.Text = "8";
                    //checkBox3.Text = "12";

                    break;
                case 5:
                    label1.Text = "2*2=?";
                    radioButton1.Text = "2";
                    radioButton2.Text = "4";
                    radioButton3.Text = "8";
                    //checkBox1.Text = "2";
                    //checkBox2.Text = "4";
                    //checkBox3.Text = "8";

                    break;
            }

        }
        private void button1_Click(object sender, EventArgs e)
        {
            n--;
            if (n < 0) n=0;
            questions(n);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            n++;
            if (n > 5) n = 5;
            questions(n);

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            answer[n] = 1;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            answer[n] = 2;
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            answer[n] = 3;
        }

        public void WriteData ()
        {
           Excel excel = new Excel(@"ts.xlsx", 1);
            excel.WriteToCell(0, 1, excel.ReadCell(0, 0));
            excel.Save();
            excel.SaveAs(@"test3.xlsx");
            excel.Close();

        }

       
        private void button3_Click(object sender, EventArgs e)
        {
            int correct=0;
           
            if (answer[0] == 1)
            {
                correct++;
                corr1=1;
            }
            else
            {
                incorr1=1;
            }
            if (answer[1] == 3)
            {
                correct++;
                corr2=1;
            }
            else
            {
                incorr2=1;
            }
            if (answer[2] == 1)
            { 
                correct++;
                corr3=1;
            }
            else
            {
                incorr3=1;
            }
            if (answer[3] == 2)
            {
                correct++;
                corr4=1;
            }
            else
            {
                incorr4=1;
            }
            if (answer[4] == 1)
            {
                correct++;
                corr5=1;
            }
            else
            {
                incorr5=1;
            }
            if (answer[5] == 2)
            {
                correct++;
                corr6=1;
            }
            else
            {
                incorr6=1;
            }


            if (File.Exists(@"test.xlsx"))
            {
                ;
            }
            else
            {
                Excel ex = new Excel();
                ex.CreateNewFile();
                ex.SaveAs(@"test.xlsx");
                ex.Close();
            }

            
            Excel excel = new Excel(@"test.xlsx", 1);
            
            excel.WriteToCell(0, 0, "Количество участников");
            excel.WriteToCell(1, 1, "Правильно");
            excel.WriteToCell(1, 2, "Неправильно");
            excel.WriteToCell(2, 0, "1");
            excel.WriteToCell(2, 1, Convert.ToString(corr1 + excel.CellNull(excel, 2, 1)));
            excel.WriteToCell(2, 2, Convert.ToString(incorr1 + excel.CellNull(excel, 2, 2)));
            excel.WriteToCell(3, 0, "2");
            excel.WriteToCell(3, 1, Convert.ToString(corr2 + excel.CellNull(excel, 3, 1)));
            excel.WriteToCell(3, 2, Convert.ToString(incorr2 + excel.CellNull(excel,3, 2)));
            excel.WriteToCell(4, 0, "3");
            excel.WriteToCell(4, 1, Convert.ToString(corr3 + excel.CellNull(excel, 4, 1)));
            excel.WriteToCell(4, 2, Convert.ToString(incorr3 + excel.CellNull(excel,4, 2)));
            excel.WriteToCell(5, 0, "4");
            excel.WriteToCell(5, 1, Convert.ToString(corr4 + excel.CellNull(excel,5, 1)));
            excel.WriteToCell(5, 2, Convert.ToString(incorr4 + excel.CellNull(excel,5, 2)));
            excel.WriteToCell(6, 0, "5");
            excel.WriteToCell(6, 1, Convert.ToString(corr5 + excel.CellNull(excel,6, 1)));
            excel.WriteToCell(6, 2, Convert.ToString(incorr5 + excel.CellNull(excel,6, 2)));
            excel.WriteToCell(7, 0, "6");
            excel.WriteToCell(7, 1, Convert.ToString(corr6 + excel.CellNull(excel,7, 1)));
            excel.WriteToCell(7, 2, Convert.ToString(incorr6 + excel.CellNull(excel,7, 2)));
            
            excel.WriteToCell(0, 1, Convert.ToString(1+ excel.CellNull(excel, 0, 1)));
            excel.range("A2", "C8");
            excel.Save();
            excel.Close();
            n = 0;
            questions(n);
            radioButton1.Checked = false;
            radioButton2.Checked = false;
            radioButton3.Checked = false;
            answer[0] = 0;
            answer[1] = 0;
            answer[2] = 0;
            answer[3] = 0;
            answer[4] = 0;
            answer[5] = 0;
            MessageBox.Show(Convert.ToString(textBox1.Text) + " " + Convert.ToString(textBox2.Text) + ", Вы дали " +
                 Convert.ToString(correct) + " из 6 правильных ответов", "Результаты");

            textBox1.Text = "";
            textBox2.Text = "";
        }

        private void button4_Click(object sender, EventArgs e)
        {

            Excel excel = new Excel(@"test.xlsx", 1);
         
            excel.addchart();
            excel.Save();
            //excel.Close();
            
        }
    }
}
