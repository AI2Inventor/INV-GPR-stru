using System;
using System.Windows.Forms;

namespace c_1114
{
    public partial class Form6 : Form
    {
        public static double mu, Pcr, I, b, t, r, l, E;

        private void Form6_Load(object sender, EventArgs e)
        {

        }

        public Form6()
        {
            InitializeComponent();
        }

        public void button1_Click(object sender, EventArgs e)
        {

            mu = 0;
            Pcr = 0;
            I = 0;
            r = Convert.ToDouble(textBox4.Text);
            t = Convert.ToDouble(textBox6.Text);
            b = Convert.ToDouble(textBox2.Text);
            l = Convert.ToDouble(textBox1.Text);
            E = 2.06 * Math.Pow(10, 3);




            switch (comboBox1.Text)
            {
                case "两端铰接":
                    mu = 1.0;
                    break;
                case "两端固定":
                    mu = 0.65;
                    break;
                case "上端铰接下端固定":
                    mu = 0.8;
                    break;
                case "上端自由下端固定":
                    mu = 2.1;
                    break;
            }
            I = Math.Pow(b, 4) / 12 - Math.Pow((b - 2 * t), 4) / 12;


            Pcr = (Math.Pow(Math.PI, Math.PI) * (E * I)) / (Math.Pow(mu * l, 2));
            // -1 / 12 * Math.Pow((b - 2 * t), 4)
            Console.WriteLine(Pcr / 1000);







        }



    }
}
