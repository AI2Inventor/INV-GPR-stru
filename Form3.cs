using Inventor;
using System;
using System.Windows.Forms;
namespace c_1114
{
    public partial class Form3 : Form
    {


        public Form3()
        {
            InitializeComponent();

        }
        private void button1_Click(object sender, EventArgs e)
        {
            double D = Convert.ToDouble(textBox1.Text) / 10;
            Form1.w_d = Convert.ToDouble(textBox2.Text) / 10;
            double S = Convert.ToDouble(textBox3.Text) / 10;
            double t1 = Convert.ToDouble(textBox4.Text) / 10;
            double t2 = Convert.ToDouble(textBox5.Text) / 10;
            double r = Convert.ToDouble(textBox6.Text) / 10;
            double h = Convert.ToDouble(textBox7.Text) / 10;
            //MessageBox.Show(Form1.w_d.ToString());
            column_connector column_Connector = new column_connector(D, S, t1, r, h);
            this.Close();
        }

        public class column_connector
        {

            public column_connector(double D, double S, double t1, double r, double h)
            {
                PartDocument odoc = (PartDocument)Form1.invapp.Documents.Open(@"I:\代码\VStudio\c_1114\模型文件\柱\柱端.ipt", false);
                var oParams = odoc.ComponentDefinition.Parameters;
                var oUserParams = oParams.UserParameters;
                oUserParams["厚度"].Value = t1;
                oUserParams["孔距"].Value = S;
                oUserParams["孔大小"].Value = D;
                oUserParams["长宽"].Value = Form1.w_d;
                oUserParams["高度"].Value = h;
                odoc.Update();
                odoc.Save2(true);

            }
        }
        //public class column_section
        //{
        //    public column_section(double r, double t2)
        //    {
        //        PartDocument odoc = (PartDocument)Form1.invapp.Documents.Open(Form1.Models_path + "柱" + "\\" + "RHS_section-01.ipt", false);
        //        var oParams = odoc.ComponentDefinition.Parameters;
        //        var oModelParams = oParams.ModelParameters;
        //        oModelParams["r"].Value =r;
        //        oModelParams["t"] .Value= t2;
        //        odoc.Update();
        //        odoc.Save2(true);
        //    }



        //}



        public class Test
        {
            public void Run(/*double r, double t2*/)
            {
                PartDocument odoc = (PartDocument)Form1.invapp.Documents.Open(Form1.Models_path + "柱" + "\\" + "RHS_section-01.ipt", true);
                var oParams = odoc.ComponentDefinition.Parameters;
                var oModelParams = oParams.ModelParameters;
                //oModelParams["r"].Value = r;
                //oModelParams["t"].Value = t2;
                odoc.Update();
                odoc.Save2(true);

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

            textBox1.Text = "26";
            textBox2.Text = "200";
            textBox3.Text = "38";
            textBox4.Text = "3";
            textBox5.Text = "3";
            textBox6.Text = "10";
            textBox7.Text = "50";




        }
    }
}
