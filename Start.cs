using Inventor;
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using static System.Activator;

namespace c_1114
{
    internal class Start
    {

        


        public void Connect()
        {


            try
            {
                Form1.invapp = (Inventor.Application)Marshal.GetActiveObject("Inventor.Application");
                System.Threading.Thread.Sleep(1 * 2000);
                PartDocument opartdoc = (PartDocument)Form1.invapp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, Form1.part_templatepath, false);
                Form1.addin = Form1.invapp.ApplicationAddIns.ItemById[Form1.ClientID];
                Form1.iLogicAutomation = Form1.addin.Automation;
                //Form1.form1.richTextBox1.AppendText("已连接。\n");
            }
            catch (Exception ex)
            {
                try
                {
                    Form1.form1.richTextBox1.AppendText(ex.Message);
                    Type invappType = System.Type.GetTypeFromProgID("Inventor.Application");
                    Form1.invapp = (Inventor.Application)CreateInstance(invappType);
                    Form1.invapp.Visible = true;
                    Form1.Startde = true;
                }
                catch (Exception ex2)// 改成ex2，名字不冲突 
                {
                    MessageBox.Show(ex2.ToString());
                    MessageBox.Show("无法连接Inventor");

                }
            }


        }
    }
}




