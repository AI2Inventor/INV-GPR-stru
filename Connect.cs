using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using static System.Activator;

namespace c_1114
{




    internal class Connect
    {
        public Inventor.Application invapp;
        string inventor_app = "Inventor.Application";
        public static bool Startde = false;
        public object Conect()
        {
            try
            {
                invapp = (Inventor.Application)Marshal.GetActiveObject(inventor_app);
                MessageBox.Show("已连接到Inventor", "提示");
            }
            catch (Exception ex)
            {
                try
                {
                    MessageBox.Show("Inventor没有启动，启动inventor？");
                    Type invappType = System.Type.GetTypeFromProgID("Inventor.Application");
                    invapp = (Inventor.Application)CreateInstance(invappType);
                    invapp.Visible = true;
                    Startde = true;
                }
                catch (Exception ex2)// 改成ex2，名字不冲突 
                {
                    MessageBox.Show(ex2.ToString());
                    MessageBox.Show("无法连接Inventor");
                }
            }
            Console.WriteLine("sss");
            Console.WriteLine("success");
            return invapp;
        }
    }

}


