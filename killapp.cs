using c_1114.Properties;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace c_1114
{
    internal class killapp
    {
        public void Killinventor()
        {
            Process proc = null;
            try
            {
                proc = new Process();
                proc.StartInfo.FileName = @"I:\Wcode\VStudio\c_1114\Resources\bat\killinventor.bat";
                proc.StartInfo.Arguments = string.Format("10");//this is argument
                proc.StartInfo.CreateNoWindow = false;
                proc.Start();
                proc.WaitForExit();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception Occurred :{0},{1}", ex.Message, ex.StackTrace.ToString());
            }
        }
    }
}
