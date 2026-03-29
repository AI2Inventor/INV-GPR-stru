using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace c_1114
{

    public partial class Form7 : Form
    {
        public static string Disatershelter = @"I:\Wcode\VStudio\c_1114\bin\Debug\Resources\TSpicture\1.jpg";
        public static string Agriculture = @"I:\Wcode\VStudio\c_1114\bin\Debug\Resources\TSpicture\2.jpg";
        public static string Sports = @"I:\Wcode\VStudio\c_1114\bin\Debug\Resources\TSpicture\3.jpg";
        public static string Grandstand = @"I:\Wcode\VStudio\c_1114\bin\Debug\Resources\TSpicture\4.jpg";
        public static string Deepsee = @"I:\Wcode\VStudio\c_1114\bin\Debug\Resources\TSpicture\5.jpg";
        public static string PolarExploration = @"I:\Wcode\VStudio\c_1114\bin\Debug\Resources\TSpicture\6.jpg";
        public static string Military = @"I:\Wcode\VStudio\c_1114\bin\Debug\Resources\TSpicture\7.jpg";

        public Form7()
        {
            InitializeComponent();
        }
        public void selectcase(string casename)
        {
        }
        private void button1_Click(object sender, EventArgs e)
        {
            switch (comboBox1.Text)
            {
                case "防灾救灾":
                    Form1 Disatershelter = new Form1();
                    Disatershelter.Text = "防灾减灾 - 临时结构设计助手";
                    Disatershelter.ShowDialog();
                    break;
                case "农业过程":
                    Form1 Agriculture = new Form1();
                    Agriculture.Text = "农业过程 - 临时结构设计助手"; 
                    Agriculture.ShowDialog();
                    break;
                case "体育文化":
                    Form1 Sports = new Form1();
                    Sports.ShowDialog();
                    break;
                case "舞台看台":
                    Form1 Grandstand = new Form1();
                    Grandstand.ShowDialog();
                    break;
                case "深海深地":
                    Form1 Deepsee = new Form1();
                    Deepsee.ShowDialog();
                    break;
                case "极地探测":
                    Form1 PolarExploration = new Form1();
                    PolarExploration.ShowDialog();
                    break;
                case "军工保障":
                    Form1 Military = new Form1();
                    Military.ShowDialog();
                    break;
                case "测试01":
                    Form8 Testform01 = new Form8();
                    Testform01.ShowDialog();
                    break;
            }



        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox1.Text)
            {
                case "防灾救灾":
                    pictureBox1.Load(Disatershelter);
                    break;
                case "农业过程":
                    pictureBox1.Load(Agriculture);
                    break;
                case "体育文化":
                    pictureBox1.Load(Sports);

                    break;
                case "舞台看台":
                    pictureBox1.Load(Grandstand);

                    break;
                case "深海深地":
                    pictureBox1.Load(Deepsee);

                    break;
                case "极地探测":
                    pictureBox1.Load(PolarExploration);
                    break;
                case "军工保障":
                    pictureBox1.Load(Military);

                    break;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            switch (comboBox1.Text)
            {
                case "防灾救灾":
                    Form1 Disatershelter = new Form1();
                    Disatershelter.Text = "防灾减灾 - 临时结构设计助手";
                    Disatershelter.ShowDialog();
                    break;
                case "农业过程":
                    Form1 Agriculture = new Form1();
                    Agriculture.Text = "农业过程 - 临时结构设计助手";
                    Agriculture.ShowDialog();
                    break;
                case "体育文化":
                    Form1 Sports = new Form1();
                    Sports.ShowDialog();
                    break;
                case "舞台看台":
                    Form1 Grandstand = new Form1();
                    Grandstand.ShowDialog();
                    break;
                case "深海深地":
                    Form1 Deepsee = new Form1();
                    Deepsee.ShowDialog();
                    break;
                case "极地探测":
                    Form1 PolarExploration = new Form1();
                    PolarExploration.ShowDialog();
                    break;
                case "军工保障":
                    Form1 Military = new Form1();
                    Military.ShowDialog();
                    break;
                case "测试01":
                    Form8 Testform01 = new Form8();
                    Testform01.Text ="墙体智能设计";
                    Testform01.ShowDialog();
                    break;
            }
        }
    }
}
