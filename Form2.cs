using Inventor;
using System;
using System.Windows.Forms;

namespace c_1114
{
    public partial class Form2 : Form
    {
        public static Form2 form2;
        public static int n;
        public Form2()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            Form1.beam_height = Convert.ToDouble(textBox3.Text) / 10;
            //上面这两个不是变量
            Double b = Convert.ToDouble(textBox2.Text) / 10;
            Double t1 = Convert.ToDouble(textBox4.Text) / 10;
            Double t2 = Convert.ToDouble(textBox5.Text) / 10;
            textBox1.Text = ((Form1.w_d - t2) / 2).ToString();
            n = Convert.ToInt32(textBox6.Text);
            Beam_section beam_Section = new Beam_section(b, t1, t2);
            Beam_section_Y beam_Section_Y = new Beam_section_Y(b, t1, t2);
            this.Close();
        }
        public class Beam_section
        {
            public Beam_section(double b, double t1, double t2)
            {
                PartDocument odoc = (PartDocument)Form1.invapp.Documents.Open(Form1.Models_path + "梁" + "\\" + "双拼C型钢.ipt", false);
                var oParams = odoc.ComponentDefinition.Parameters;
                var oUserParams = oParams.UserParameters;
                oUserParams["高"].Value = Form1.beam_height;
                oUserParams["卷边"].Value = b;
                oUserParams["厚度"].Value = t1;
                oUserParams["厚度2"].Value = t2;
                odoc.Update();
                odoc.Save2(true);
            }
        }
        public class Beam_section_Y
        {
            public Beam_section_Y(double b, double t1, double t2)
            {
                PartDocument odoc = (PartDocument)Form1.invapp.Documents.Open(Form1.Models_path + "梁" + "\\" + "双拼C型钢2.ipt", false);
                var oParams = odoc.ComponentDefinition.Parameters;
                var oUserParams = oParams.UserParameters;
                oUserParams["高"].Value = Form1.beam_height;

                oUserParams["卷边"].Value = b;
                oUserParams["厚度"].Value = t1;
                oUserParams["厚度2"].Value = t2;

                odoc.Update();
                odoc.Save2(true);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }
    }
}
