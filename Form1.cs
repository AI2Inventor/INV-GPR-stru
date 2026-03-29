using Inventor;
using Microsoft.Office.Interop.Excel;
using System;
//using Autodesk.iLogic.Automation;
//using Autodesk.iLogic.Interfaces;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Numerics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using static System.Activator;
using Application = Microsoft.Office.Interop.Excel.Application;


namespace c_1114
{
    public partial class Form1 : Form  //为什么：Form
    {
        public static bool Startde = false;
        public static Inventor.Application invapp;
        public static dynamic iLogicAutomation = null;
        public static ApplicationAddIn addin = null;
        public static Form1 form1;  //注意别写错了 form......之前写了两个Form所以出错了
        public static double whole_length, whole_width, whole_height, column_num, x_distance, y_distance, z_distance;
        public static string save_path;
        public static int x_count, y_count, z_count;
        public static Vector3[,,] points = new Vector3[101, 101, 101];
        public static string part_templatepath = @"c:\users\public\documents\autodesk\inventor 2023\templates\zh-cn\metric\standard (mm).ipt";//可以设置为自动获取
        public static string assem_templatepath = @"c:\users\public\documents\autodesk\inventor 2023\templates\zh-cn\metric\standard (mm).iam";
        public static string Models_path = @"I:\Wcode\VStudio\c_1114\模型文件\";
        public static string rules_path = @"I:\Wcode\VStudio\c_1114\模型文件\rules\";
        public static string ClientID = "{3BDD8D79-2179-4B11-8A5A-257B1C0263AC}";

        //以下是最关键的参数：从窗体获取；后面加空间

        public static double w_d, floornum, interwallnum, exwallnum_x, exwallnum_y, beam_height;
        //w_d为cm单位


        public Form1()
        {
            InitializeComponent();
            CheckForIllegalCrossThreadCalls = false;
            form1 = this;
        }
        //public static bool startde = false;

        //public static Inventor.Application invapp;


        private void 连接ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            richTextBox1.AppendText("请指定目录...\n");
            Start start = new Start();
            start.Connect();
            // Connect connect = new Connect();
            richTextBox1.AppendText("请指定目录...\n");

        }
        private void Form1_Load(object sender, EventArgs e)

        {

            //Connect connect = new Connect();



            //MessageBox.Show("1.删除2.连接3.关闭打开的文档", "提示");
        }//窗口加载





        private void 删除文件_Click(object sender, EventArgs e)
        {
            DirectoryInfo dir = new DirectoryInfo(save_path);
            if (dir.Exists)
            {
                DirectoryInfo[] childs = dir.GetDirectories();
                foreach (DirectoryInfo child in childs)
                {
                    child.Delete(true);
                }
                dir.Delete(true);
                richTextBox1.AppendText("已删除\n");

            }
            else
            {
                richTextBox1.AppendText("目录不存在\n");
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            key_points key_Points = new key_points();

            AssemblyDocument assemDoc = (AssemblyDocument)invapp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, assem_templatepath, true);

            var oPositionMatrix = invapp.TransientGeometry.CreateMatrix(); // 创建位置矩阵
                                                                           // Console.WriteLine(oPositionMatrix.Translation.X)
                                                                           // Console.WriteLine(oPositionMatrix.Translation.Y)
                                                                           // Console.WriteLine(oPositionMatrix.Translation.Z)
                                                                           //ssemDoc.ComponentDefinition.Occurrences.Add(Models_path + "组合柱.iam", oPositionMatrix);

            // Dim oCylinder(5) As ComponentOccurrence '当前的部件

            ComponentOccurrence[] ocolumn = new ComponentOccurrence[10001], oxBeam = new ComponentOccurrence[10001],
                oyBeam = new ComponentOccurrence[10001],
                oFloor = new ComponentOccurrence[10010],
                oRoof = new ComponentOccurrence[10010], oInternalWalls = new ComponentOccurrence[10000],
                oExternalWalls = new ComponentOccurrence[10000];


            WorkPlane[] Column_YZ = new WorkPlane[6], Column_XZ = new WorkPlane[6], Column_XY = new WorkPlane[6], Beam_YZ = new WorkPlane[12],
                Beam_XZ = new WorkPlane[12], Beam_XY = new WorkPlane[12]; // column 和 beam 的 三个平面
                                                                          // 每个零件的planes

            // Dim oPartPlane1 As WorkPlane = ocolumn(5).Definition.WorkPlanes.Item(3)
            // Dim opartaxis1 As WorkAxis = oOcc1.Definition.workaxes.item(3)
            // 注意 这个definition 没有workplanes 定义
            var workplane_YZ = assemDoc.ComponentDefinition.WorkPlanes[1]; // 装配体的三个平面
            var workplane_XZ = assemDoc.ComponentDefinition.WorkPlanes[2];
            var workplane_XY = assemDoc.ComponentDefinition.WorkPlanes[3];

            var axis_x = invapp.TransientGeometry.CreateVector(1, 0, 0);
            var axis_y = invapp.TransientGeometry.CreateVector(0, 1, 0);
            var axis_z = invapp.TransientGeometry.CreateVector(0, 0, 1);
            var center_point = invapp.TransientGeometry.CreatePoint(0, 0, 0);
            // 把单独零件的平面设置成装配文档代理来使用单独零件里的平面。
            WorkPlaneProxy[] Column_YZ_Osm = new WorkPlaneProxy[6], Column_XZ_Osm = new WorkPlaneProxy[6], Column_XY_Osm = new WorkPlaneProxy[6];
            WorkPlaneProxy[] Beam_YZ_Osm = new WorkPlaneProxy[12], Beam_XZ_Osm = new WorkPlaneProxy[12], Beam_XY_Osm = new WorkPlaneProxy[12];
            // 转换到各自位置，变换操作
            Inventor.Vector[] oTrans = new Inventor.Vector[101], oTrans_x_beam = new Inventor.Vector[101], oTrans_y_beam = new Inventor.Vector[101],
                oTrans_L_joint = new Inventor.Vector[101], oTrans_x_Floor = new Inventor.Vector[1000],
                oTrans_y_Roof = new Inventor.Vector[100], oTrans_y_internalwalls = new Inventor.Vector[1000],
                oTrans_x_internalwalls = new Inventor.Vector[1000],
                oTrans_y_externalwalls = new Inventor.Vector[1000], oTrans_x_externalwalls = new Inventor.Vector[1000];

            int i = 0, j = 0, k = 0, m, m_1, m_2, m_3, m_4, m_5, m_6, m_7;
            double x, y, z;
            x = points[i, j, k].X;
            y = points[i, j, k].Y;
            z = points[i, j, k].Z;
            m = 0; // column
            m_1 = 0; // xbeam
            m_2 = 0; // ybeam
            m_3 = 0; // xfloor
            m_4 = 0;
            m_5 = 0;
            m_6 = 0;
            m_7 = 0;
            for (k = 0; k <= z_count - 1; k++)
            {
                int loopTo1 = y_count - 1;
                for (j = 0; j <= loopTo1; j++)
                {
                    var loopTo2 = x_count - 1;
                    for (i = 0; i <= loopTo2; i++)
                    {
                        oTrans[m] = invapp.TransientGeometry.CreateVector(points[i, j, k].X, points[i, j, k].Y, points[i, j, k].Z);
                        m = m + 1;
                    }
                }
            } // column
            for (k = 0; k <= z_count; k++)
            {
                var loopTo4 = y_count - 1;
                for (j = 0; j <= loopTo4; j++)
                {
                    var loopTo5 = x_count - 2;
                    for (i = 0; i <= loopTo5; i++)
                    {
                        oTrans_x_beam[m_1] = invapp.TransientGeometry.CreateVector(points[i, j, k].X, points[i, j, k].Y, points[i, j, k].Z);
                        m_1 = m_1 + 1;
                    }
                }
            } // xbeam


            for (k = 0; k <= z_count - 1; k++)
            {
                var loopTo7 = y_count - 2;
                for (j = 0; j <= loopTo7; j++)
                {
                    var loopTo8 = x_count - 1;
                    for (i = 0; i <= loopTo8; i++)
                    {
                        oTrans_y_beam[m_2] = invapp.TransientGeometry.CreateVector(points[i, j, k].X, points[i, j, k].Y, points[i, j, k].Z);
                        m_2 = m_2 + 1;
                    }
                }
            } // ybeam


            for (k = 0; k <= z_count - 1; k++)
            {

                for (j = 0; j <= y_count - 2; j++)
                {
                    for (i = 0; i <= x_count - 2; i++)
                    {
                        oTrans_x_Floor[m_3] = invapp.TransientGeometry.CreateVector(points[i, j, k].X, points[i, j, k].Y, points[i, j, k].Z);
                        m_3 = m_3 + 1;
                    }


                }
            }//xfloor




            k = z_count;
            for (i = 0; i <= x_count - 1; i++)
            {
                oTrans_y_Roof[m_4] = invapp.TransientGeometry.CreateVector(points[i, 0, k].X, 0, points[i, 0, k].Z);
                m_4 = m_4 + 1;
            }


            //YROOF


            for (k = 0; k <= z_count - 1; k++)
            {

                for (j = 0; j <= y_count - 2; j++)
                {
                    for (i = 1; i <= x_count - 2; i++)
                    {
                        oTrans_x_internalwalls[m_5] = invapp.TransientGeometry.CreateVector(points[i, j, k].X, points[i, j, k].Y, points[i, j, k].Z);
                        m_5 = m_5 + 1;
                    }

                }

            }//yinternal walls 

            for (k = 0; k <= z_count - 1; k++)
            {

                j = 0;

                for (i = 0; i < x_count - 1; i++)
                {
                    oTrans_x_externalwalls[m_6] = invapp.TransientGeometry.CreateVector(points[i, j, k].X, points[i, j, k].Y, points[i, j, k].Z);
                    m_6 = m_6 + 1;
                }
                j = y_count - 1;
                for (i = 0; i < x_count - 1; i++)
                {
                    oTrans_x_externalwalls[m_6] = invapp.TransientGeometry.CreateVector(points[i, j, k].X, points[i, j, k].Y, points[i, j, k].Z);
                    m_6 = m_6 + 1;
                }
            }//xexternal walls 

            for (k = 0; k <= z_count - 1; k++)
            {

                i = 0;

                for (j = 0; j < y_count - 1; j++)
                {
                    oTrans_y_externalwalls[m_7] = invapp.TransientGeometry.CreateVector(points[i, j, k].X, points[i, j, k].Y, points[i, j, k].Z);
                    m_7 = m_7 + 1;
                }
                i = x_count - 1;
                for (j = 0; j < y_count - 1; j++)
                {
                    oTrans_y_externalwalls[m_7] = invapp.TransientGeometry.CreateVector(points[i, j, k].X, points[i, j, k].Y, points[i, j, k].Z);
                    m_7 = m_7 + 1;
                }
            }//yexternal walls 
            // For k = 0 To z_count
            // For j = 0 To y_count - 2
            // For i = 0 To x_count - 1
            // oTrans_y_beam(m_2) = invapp.TransientGeometry.CreateVector(points(i, j, k).X, points(i, j, k).Y, points(i, j, k).Z)
            // m_2 = m_2 + 1
            // Next
            // Next
            // Next
            var loopTo9 = m - 1;
            for (i = 0; i <= loopTo9; i++) // 柱子
            {
                oPositionMatrix.SetTranslation(oTrans[i]);
                ocolumn[i] = assemDoc.ComponentDefinition.Occurrences.Add(Models_path + "柱" + "\\" + "组合柱.iam", oPositionMatrix);
                ocolumn[i].Name = "柱_" + (i + 1).ToString();
                ocolumn[i].Grounded = true;

            }
            richTextBox1.AppendText("已插入柱\n");
            var view = invapp.ActiveView;
            view.DisplayMode = DisplayModeEnum.kShadedWithEdgesRendering;

            view.Fit();
            for (i = 0; i <= m_1 - 1; i++) // x梁
            {
                var opositionMatrix_1 = invapp.TransientGeometry.CreateMatrix();
                var xMatrix_1 = invapp.TransientGeometry.CreateMatrix();
                xMatrix_1.SetTranslation(oTrans_x_beam[i]);
                opositionMatrix_1.PreMultiplyBy(xMatrix_1);
                oxBeam[i] = assemDoc.ComponentDefinition.Occurrences.Add(Models_path + "梁" + "\\" + "x_梁.iam", opositionMatrix_1);
                oxBeam[i].Name = "双肢梁X" + (i + 1).ToString();
                oxBeam[i].Grounded = true;

            }


            for (i = 0; i <= m_2 - 1; i++) // y梁
            {
                var opositionMatrix_1 = invapp.TransientGeometry.CreateMatrix();
                var xMatrix_1 = invapp.TransientGeometry.CreateMatrix();
                xMatrix_1.SetTranslation(oTrans_y_beam[i]);
                opositionMatrix_1.PreMultiplyBy(xMatrix_1);
                oyBeam[i] = assemDoc.ComponentDefinition.Occurrences.Add(Models_path + "梁" + "\\" + "Y_梁.iam", opositionMatrix_1);
                oyBeam[i].Name = "双肢梁Y" + (i + 1).ToString();
                oyBeam[i].Grounded = true;
            }
            richTextBox1.AppendText("已插入梁\n");
            k = 0;
            for (i = 0; i <= m_3 - 1; i++)
            {
                for (j = 0; j < floornum; j++)
                {
                    var opositionMatrix = invapp.TransientGeometry.CreateMatrix();
                    var xMatrix_1 = invapp.TransientGeometry.CreateMatrix();
                    var xMatrix_2 = invapp.TransientGeometry.CreateMatrix();
                    xMatrix_1.SetTranslation(oTrans_x_Floor[i]);
                    xMatrix_2.SetTranslation(invapp.TransientGeometry.CreateVector(j * ((x_distance - w_d) / floornum), 0, 0));
                    opositionMatrix.PreMultiplyBy(xMatrix_1);
                    opositionMatrix.PreMultiplyBy(xMatrix_2);
                    oFloor[k] = assemDoc.ComponentDefinition.Occurrences.Add(Models_path + "板" + "\\" + "组合楼板.iam", opositionMatrix);
                    oFloor[k].Name = "组合楼板" + (k + 1).ToString();
                    oFloor[k].Grounded = true;
                    k = k + 1;
                }
            }
            richTextBox1.AppendText("已插入组合楼板\n");
            k = 0;
            for (k = 0; k < 1; k++)
            {
                var opositionMatrix = invapp.TransientGeometry.CreateMatrix();
                var xMatrix_1 = invapp.TransientGeometry.CreateMatrix();
                xMatrix_1.SetTranslation(oTrans_y_Roof[0]);
                opositionMatrix.PreMultiplyBy(xMatrix_1);
                oRoof[0] = assemDoc.ComponentDefinition.Occurrences.Add(Models_path + "屋架" + "\\" + "全部屋架.iam", opositionMatrix);
                oRoof[0].Name = "屋架" + (k + 1).ToString();
                oRoof[0].Grounded = true;
                richTextBox1.AppendText("已插入屋架\n");
            }
            k = 0;

            if (Convert.ToInt32(TextBox2.Text) > 2)
            {
                for (i = 0; i <= m_5 - 1; i++)
                {
                    for (j = 0; j < interwallnum; j++)
                    {
                        var opositionMatrix = invapp.TransientGeometry.CreateMatrix();
                        var xMatrix_1 = invapp.TransientGeometry.CreateMatrix();
                        var xMatrix_2 = invapp.TransientGeometry.CreateMatrix();
                        xMatrix_1.SetTranslation(oTrans_x_internalwalls[i]);
                        xMatrix_2.SetTranslation(invapp.TransientGeometry.CreateVector(0, (y_distance - w_d) / interwallnum * j, 0));
                        opositionMatrix.PreMultiplyBy(xMatrix_1);
                        opositionMatrix.PreMultiplyBy(xMatrix_2);
                        oInternalWalls[k] = assemDoc.ComponentDefinition.Occurrences.Add(Models_path + "内墙" + "\\" + "组合内墙.iam", opositionMatrix);
                        oInternalWalls[k].Name = "内隔墙" + (k + 1).ToString();
                        oInternalWalls[k].Grounded = true;
                        k++;
                    }
                }
            }
            k = 0;
            for (i = 0; i <= m_6 - 1; i++)
            {
                for (j = 0; j < exwallnum_x; j++)
                {
                    var opositionMatrix = invapp.TransientGeometry.CreateMatrix();
                    var xMatrix_1 = invapp.TransientGeometry.CreateMatrix();
                    var xMatrix_2 = invapp.TransientGeometry.CreateMatrix();
                    xMatrix_1.SetTranslation(oTrans_x_externalwalls[i]);
                    xMatrix_2.SetTranslation(invapp.TransientGeometry.CreateVector((x_distance - w_d) / exwallnum_x * j, 0, 0));
                    opositionMatrix.PreMultiplyBy(xMatrix_1);
                    opositionMatrix.PreMultiplyBy(xMatrix_2);
                    oExternalWalls[k] = assemDoc.ComponentDefinition.Occurrences.Add(Models_path + "外墙" + "\\" + "标准外墙x.iam", opositionMatrix);
                    oExternalWalls[k].Name = "外墙X" + (k + 1).ToString();
                    oExternalWalls[k].Grounded = true;
                    k++;
                }
            }
            for (i = 0; i <= m_7 - 1; i++)
            {
                for (j = 0; j < exwallnum_y; j++)
                {
                    var opositionMatrix = invapp.TransientGeometry.CreateMatrix();
                    var xMatrix_1 = invapp.TransientGeometry.CreateMatrix();
                    var xMatrix_2 = invapp.TransientGeometry.CreateMatrix();
                    xMatrix_1.SetTranslation(oTrans_y_externalwalls[i]);

                    xMatrix_2.SetTranslation(invapp.TransientGeometry.CreateVector(0, (y_distance - w_d) / exwallnum_y * j, 0));
                    opositionMatrix.PreMultiplyBy(xMatrix_1);
                    opositionMatrix.PreMultiplyBy(xMatrix_2);
                    oExternalWalls[k] = assemDoc.ComponentDefinition.Occurrences.Add(Models_path + "外墙" + "\\" + "标准外墙y.iam", opositionMatrix);
                    oExternalWalls[k].Name = "外墙Y" + (k + 1).ToString();
                    oExternalWalls[k].Grounded = true;
                    k++;
                }
            }




            if (Convert.ToDouble(form1.TextBox5.Text) == 1)
            {
                var odoc = invapp.ActiveDocument;
                dynamic myRule = null;
                myRule = @"I:\Wcode\VStudio\c_1114\模型文件\Rule_3.iLogicVb";
                iLogicAutomation.RunExternalRule(odoc, myRule);
            }
            //// 把最上面的梁降低到与柱子平齐
            //var loopTo12 = (x_count - 1) * y_count * z_count - 1;
            //for (i = (x_count - 1) * y_count; i <= loopTo12; i++)
            //{
            //    oxBeam[i].Grounded = false;
            //    var matrix = invapp.TransientGeometry.CreateMatrix();
            //    matrix.SetTranslation(invapp.TransientGeometry.CreateVector(0d, 0d, -this.ComboBox3.Text / 20d));
            //    var opositionMatrix0 = invapp.TransientGeometry.CreateMatrix();
            //    opositionMatrix0.PreMultiplyBy(oxBeam[i].Transformation.Copy());
            //    opositionMatrix0.PreMultiplyBy(matrix);
            //    oxBeam[i].Transformation = opositionMatrix0;
            //    oxBeam[i].Grounded = true;
            //}

            //var loopTo13 = x_count * (y_count - 1) * z_count - 1;
            //for (i = x_count * (y_count - 1); i <= loopTo13; i++)
            //{
            //    oyBeam[i].Grounded = false;
            //    var matrix = invapp.TransientGeometry.CreateMatrix();
            //    matrix.SetTranslation(invapp.TransientGeometry.CreateVector(0d, 0d, -this.ComboBox3.Text / 20d));
            //    var opositionMatrix0 = invapp.TransientGeometry.CreateMatrix();
            //    opositionMatrix0.PreMultiplyBy(oyBeam[i].Transformation.Copy());
            //    opositionMatrix0.PreMultiplyBy(matrix);
            //    oyBeam[i].Transformation = opositionMatrix0;
            //    oyBeam[i].Grounded = true;
            //}
            //var loopTo14 = (x_count - 1) * y_count * (z_count + 1) - 1;
            //for (i = (x_count - 1) * y_count * z_count; i <= loopTo14; i++)
            //{
            //    oxBeam[i].Grounded = false;
            //    var matrix = invapp.TransientGeometry.CreateMatrix();
            //    matrix.SetTranslation(invapp.TransientGeometry.CreateVector(0d, 0d, -this.ComboBox3.Text / 10d));
            //    var opositionMatrix0 = invapp.TransientGeometry.CreateMatrix();
            //    opositionMatrix0.PreMultiplyBy(oxBeam[i].Transformation.Copy());
            //    opositionMatrix0.PreMultiplyBy(matrix);
            //    oxBeam[i].Transformation = opositionMatrix0;
            //    oxBeam[i].Grounded = true;
            //}
            //var loopTo15 = x_count * (y_count - 1) * (z_count + 1) - 1;
            //for (i = x_count * (y_count - 1) * z_count; i <= loopTo15; i++)
            //{
            //    oyBeam[i].Grounded = false;
            //    var matrix = invapp.TransientGeometry.CreateMatrix();
            //    matrix.SetTranslation(invapp.TransientGeometry.CreateVector(0d, 0d, -this.ComboBox3.Text / 10d));
            //    var opositionMatrix0 = invapp.TransientGeometry.CreateMatrix();
            //    opositionMatrix0.PreMultiplyBy(oyBeam[i].Transformation.Copy());
            //    opositionMatrix0.PreMultiplyBy(matrix);
            //    oyBeam[i].Transformation = opositionMatrix0;
            //    oyBeam[i].Grounded = true;
            //}
            //DialogResult result;
            //result = MessageBox.Show("是否插入节点 ？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            //if ((int)result == 1)
            //{

            //}
            //result = MessageBox.Show("是否插入门窗 ？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            //if ((int)result == 1)
            //{
            //    form1.Button19.PerformClick();



            //}
            //result = MessageBox.Show("是否插入龙骨 ？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            //if ((int)result == 1)
            //{
            //    this.Button18.PerformClick();
            //}
            //result = MessageBox.Show("是否插入屋架 ？", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            //if ((int)result == 1)
            //{
            //    this.Button21.PerformClick();
            //}




            // For i = 0 To 13
            // Beam_YZ(i) = oBeam(i).Definition.WorkPlanes.Item(1)
            // Beam_XZ(i) = oBeam(i).Definition.WorkPlanes.Item(2)
            // Beam_XY(i) = oBeam(i).Definition.WorkPlanes.Item(3)
            // oBeam(i).CreateGeometryProxy(Beam_YZ(i), Beam_YZ_Osm(i))
            // oBeam(i).CreateGeometryProxy(Beam_XZ(i), Beam_XZ_Osm(i))
            // oBeam(i).CreateGeometryProxy(Beam_XY(i), Beam_XY_Osm(i))
            // Next
            // For i = 0 To 5
            // assemDoc.ComponentDefinition.Constraints.AddFlushConstraint(Column_XY_Osm(i), workplane_XY, 0)

            // Next
            // For i = 0 To 11
            // assemDoc.ComponentDefinition.Constraints.AddFlushConstraint(Beam_XY_Osm(i), workplane_XY, 0)
            // Next
            // assemDoc.ComponentDefinition.Constraints.AddFlushConstraint(Column_YZ_Osm(0), workplane_YZ, 0)
            // assemDoc.ComponentDefinition.Constraints.AddFlushConstraint(Column_XZ_Osm(0), workplane_XZ, 0)

            // assemDoc.ComponentDefinition.Constraints.AddFlushConstraint(Column_YZ_Osm(1), workplane_YZ, offset_column_1 & " m")
            // assemDoc.ComponentDefinition.Constraints.AddFlushConstraint(Column_XZ_Osm(1), workplane_XZ, 0)

            // assemDoc.ComponentDefinition.Constraints.AddFlushConstraint(Column_YZ_Osm(2), workplane_YZ, offset_column_1 * 2 & " m")
            // assemDoc.ComponentDefinition.Constraints.AddFlushConstraint(Column_XZ_Osm(2), workplane_XZ, 0)

            // assemDoc.ComponentDefinition.Constraints.AddFlushConstraint(Column_YZ_Osm(3), workplane_YZ, 0)
            // assemDoc.ComponentDefinition.Constraints.AddFlushConstraint(Column_XZ_Osm(3), workplane_XZ, offset_column_2 & " m")

            // assemDoc.ComponentDefinition.Constraints.AddFlushConstraint(Column_YZ_Osm(4), workplane_YZ, offset_column_1 & " m")
            // assemDoc.ComponentDefinition.Constraints.AddFlushConstraint(Column_XZ_Osm(4), workplane_XZ, offset_column_2 & " m")

            // assemDoc.ComponentDefinition.Constraints.AddFlushConstraint(Column_YZ_Osm(5), workplane_YZ, offset_column_1 * 2 & " m")
            // assemDoc.ComponentDefinition.Constraints.AddFlushConstraint(Column_XZ_Osm(5), workplane_XZ, offset_column_2 & " m")


            view.DisplayMode = DisplayModeEnum.kShadedWithEdgesRendering;
            view.Fit();

            assemDoc.Update();

        }
        private void Button_1_Click(object sender, EventArgs e)
        {
        }
        private void 选择截面_Click(object sender, EventArgs e)
        {
        }
        private void Button1_Click(object sender, EventArgs e)
        {
            Form2 f = new Form2();
            f.ShowDialog();


        }
        private void 预览_Click(object sender, EventArgs e)
        {
            var draw_grids = new Draw_grids(Convert.ToDouble(TextBox2.Text), Convert.ToDouble(TextBox4.Text));
            richTextBox1.AppendText("共计" + Convert.ToDouble(TextBox2.Text) * Convert.ToDouble(TextBox4.Text) + "个轴网点\n");
            richTextBox1.AppendText("请选择截面...\n");
        }

        private void label3_Click(object sender, EventArgs e)
        {
            Draw_internalwalls draw_Internalwalls = new Draw_internalwalls(Convert.ToDouble(TextBox2.Text), Convert.ToDouble(TextBox4.Text));

        }
        private void 一键加载默认值1_Click(object sender, EventArgs e)
        {
            TextBox1.Text = "12";
            TextBox2.Text = "3";
            TextBox3.Text = "6";
            TextBox4.Text = "2";
            TextBox5.Text = "2";
            TextBox6.Text = "7";
        }
        private void 一键加载默认值2_Click(object sender, EventArgs e)
        {
            TextBox1.Text = "24";
            TextBox2.Text = "4";
            TextBox3.Text = "12";
            TextBox4.Text = "3";
            TextBox5.Text = "4";
            TextBox6.Text = "12";
        }
        private void 一键加载默认值3_Click(object sender, EventArgs e)
        {
            TextBox1.Text = "12";
            TextBox2.Text = "3";
            TextBox3.Text = "5";
            TextBox4.Text = "2";
            TextBox5.Text = "1";
            TextBox6.Text = "3.5";
        }
        private void 一键加载默认值4_Click(object sender, EventArgs e)
        {
            TextBox1.Text = "8";
            TextBox2.Text = "2";
            TextBox3.Text = "4";
            TextBox4.Text = "2";
            TextBox5.Text = "1";
            TextBox6.Text = "3.2";
        }



        private void Button16_Click_1(object sender, EventArgs e)
        {
            readbom readbom = new readbom();

        }



        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            //richTextBox1.AppendText("双击进入选择\n");
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                switch (e.Node.Text)
                {
                    case "截面尺寸":

                        richTextBox1.AppendText("请输入关键截面参数..\n");


                        break;
                    case "节点尺寸":
                        richTextBox1.AppendText("请输入节点尺寸参数..\n");


                        break;


                }




            }







        }
        private void treeView1_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {

            switch (e.Node.Name)
            {
                case "组合柱截面尺寸":
                case "组合柱节点尺寸":

                    column_connector column_Connector = new column_connector();
                    richTextBox1.AppendText("已保存\n");

                    break;

                case "双肢C型钢主梁截面尺寸":
                    Beam_sectiom beam_Sectiom = new Beam_sectiom();
                    richTextBox1.AppendText("已保存\n");
                    break;
                case "屋架形式":
                    Roof_apperance roof_Apperance = new Roof_apperance();

                    break;
                    //  case “双肢C型钢主梁节点尺寸“:



            }

        }


        private void 确定1_Click(object sender, EventArgs e)
        {
            floornum = 10;
            internalwalls internalwalls = new internalwalls(10);
            internalwalls internalwalls2 = new internalwalls(10);
        }
        private void 预览_1_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
            richTextBox1.AppendText("信息汇总\n");
            richTextBox1.AppendText("组合柱-截面尺寸：" + "150×150×3" + "\n");
            richTextBox1.AppendText("外墙尺寸" + "150×150×3" + "\n");

            double length;
            floornum = 10;
            interwallnum = 10;
            exwallnum_y = 8;
            exwallnum_x = 6;
            Column column = new Column(z_distance);
            BeamXX beam = new BeamXX();
            beam.BeamX(x_distance - w_d, 4);
            //BeamXX beam = new BeamXX(x_distance - w_d, 4);
            BeamY beamY = new BeamY(y_distance - w_d, 3);
            Roof roof = new Roof(y_distance * (y_count - 1), w_d, 15 * Math.PI / 180);
            RoofS roofS = new RoofS(12);
            Floor floor = new Floor(beam_height);
            internalwalls internalwalls2 = new internalwalls(interwallnum);
            externalwallx externalwallx = new externalwallx(exwallnum_x);
            externalwally externalwally = new externalwally(exwallnum_y);
            //  https://adndevblog.typepad.com/manufacturing/2013/09/run-ilogic-rule-from-external-application.html
        }
        private void 临时结构设计软件_Click(object sender, EventArgs e)//标签链接
        {
            Process.Start("https://www.4399.com/");
        }
        private void 指定目录_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("是否使用默认目录：桌面", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
            if (result == DialogResult.Yes)
            {
                save_path = @"C:\Users\LLL\Desktop\新建文件夹" + "\\";//两个斜杆代表
                                                                 //
                Directory.CreateDirectory(save_path);
                richTextBox1.AppendText("当前目录为：\n");
                richTextBox1.AppendText(save_path + "\n");
            }
            else
            {
                FolderBrowserDialog dilog = new FolderBrowserDialog();//folderbrowser 选择文件夹
                dilog.Description = "请选择一个文件夹";
                if (dilog.ShowDialog() == DialogResult.OK || dilog.ShowDialog() == DialogResult.Yes)
                {

                    save_path = dilog.SelectedPath;
                    save_path += "\\";
                    richTextBox1.AppendText("当前目录为：\n");
                    richTextBox1.AppendText(save_path + "\n");
                }
            }
        }//指定目录
        private void 导入dxf_Click(object sender, EventArgs e)
        {
        }//读入DXF
        //''''~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~////////分界线
        public class Floor //4参数
        {
            public Floor(double height)
            {
                AssemblyDocument odoc = (AssemblyDocument)invapp.Documents.Open(Models_path + "板" + "\\" + "组合楼板.iam", false);
                var oParams = odoc.ComponentDefinition.Parameters;


                var oUserParams = oParams.UserParameters;
                oUserParams["总跨度"].Value = y_distance - w_d;
                oUserParams["w_d"].Value = w_d;
                oUserParams["总宽度"].Value = (x_distance - w_d) / floornum;
                oUserParams["总高度"].Value = height;
                dynamic myRule = null;
                foreach (dynamic eachRule in iLogicAutomation.Rules(odoc))
                {
                    if (eachRule.Name == "Rule_2")
                    {
                        myRule = eachRule;
                        break;
                    }
                    iLogicAutomation.RunRule(odoc, "MyRule");
                }
                odoc.Update();
                odoc.Save2(true);
                form1.richTextBox1.AppendText("楼板已更新\n");
            }
        }

        public class column_connector
        {
            public column_connector()
            {
                Form3 f = new Form3();
                f.ShowDialog();
            }
        }
        public class Beam_sectiom
        {
            public Beam_sectiom()
            {
                Form2 f = new Form2();
                f.ShowDialog();
            }
        }
        public class beam_joint
        {
            public beam_joint()
            {
                Form4 f = new Form4();
                f.ShowDialog();
            }
        }
        public class Roof_apperance
        {
            public Roof_apperance()
            {
                Form5 f = new Form5();
                f.ShowDialog();
            }
        }
        public class Column//两个参数
        {
            public Column(double length)
            {
                AssemblyDocument odoc = (AssemblyDocument)invapp.Documents.Open(Models_path + "柱" + "\\" + "组合柱.iam", false);
                var oParams = odoc.ComponentDefinition.Parameters;
                var oUserParams = oParams.UserParameters;
                oUserParams["length"].Value = length;
                oUserParams["w_d"].Value = w_d;
                addin.Activate();
                dynamic myRule = null;
                foreach (dynamic eachRule in iLogicAutomation.Rules(odoc))
                {
                    if (eachRule.Name == "Rule_1")
                    {
                        myRule = eachRule;
                        //list the code of rule to the list box
                        //MessageBox.Show(myRule.Text);
                        break;
                    }
                    iLogicAutomation.RunRule(odoc, "MyRule");
                }
                odoc.Update();

                odoc.Save2(true);
                form1.richTextBox1.AppendText("柱已更新\n");
            }
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {




        }//生成墙体数据集-批量输出图片

        //private void button3_Click(object sender, EventArgs e)
        //{
        //    PartDocument odoc1 = (PartDocument)invapp.Documents.Open(@"I:\模型\参数化-生成数据集\墙体\wall_frame.ipt");

        //    AssemblyDocument odoc2 = (AssemblyDocument)invapp.Documents.Open(@"I:\模型\参数化-生成数据集\墙体\wall_frame.iam");

        //    var view = invapp.ActiveView;

        //    form1.richTextBox1.AppendText("pass");
        //    view.SetCurrentAsHome(true);

        //    var oParams = odoc1.ComponentDefinition.Parameters;
        //    var Ran = new Random();
        //    int wall_num = Convert.ToInt32(textBox15.Text);
        //    string savepath = @"C:\Users\LLL\Desktop\walls\";
        //    for (int i = 0; i < wall_num; i++)
        //    {

        //        for (int j = 17; j < 41; j++)
        //        {
        //            string parastring = "d" + j.ToString();
        //            form1.richTextBox1.AppendText(parastring);
        //            oParams.ModelParameters[parastring].Value = Ran.Next(0, 1000) / 10;
        //        }
        //        odoc1.Update();
        //        odoc2.Update();
        //        //view.SaveAsBitmap(savepath + "wallpic" + i + ".png", 64, 64);
        //        view.Camera.SaveAsBitmap(savepath + "wallpic" + i + ".png", 150, 150);
        //        // invapp.ActiveView.Camera.SaveAsBitmap(savepath + "wallpic" + i + ".png", 64, 64);

        //        odoc2.Update();
        //    }
        //}

        private void button4_Click(object sender, EventArgs e)
        {
            PartDocument odoc1 = (PartDocument)invapp.Documents.Open(@"I:\模型\参数化-生成数据集\整体尺寸\零件1.ipt");

            //AssemblyDocument odoc2 = (AssemblyDocument)invapp.Documents.Open(@"I:\模型\参数化-生成数据集\墙体\wall_frame.iam");



            //var view = invapp.ActiveView;
            //form1.richTextBox1.AppendText("pass");
            //view.SetCurrentAsHome(true);

            //var oParams = odoc1.ComponentDefinition.Parameters;
            //var Ran = new Random();
            //int pic_num = Convert.ToInt32(textBox15.Text);
            //string savepath = @"C:\Users\LLL\Desktop\buildings\";
            //for (int i = 0; i < pic_num; i++)
            //{
            //    oParams.UserParameters["length"].Value = Ran.Next(5000, 10000) / 10;
            //    oParams.UserParameters["width"].Value = Ran.Next(4000, 6000) / 10;
            //    oParams.UserParameters["height"].Value = Ran.Next(3000, 4200) / 10;
            //    oParams.UserParameters["theta"].Value = Ran.Next(0, 30) * Math.PI / 180;
            //    odoc1.Update();
            //    view.Fit();
            //    //view.SaveAsBitmap(savepath + "wallpic" + i + ".png", 64, 64);
            //    view.Camera.SaveAsBitmap(savepath + "buildingpic" + i + ".png", 150, 150);


            //    // invapp.ActiveView.Camera.SaveAsBitmap(savepath + "wallpic" + i + ".png", 64, 64);
            //}
        }

        public void WriteBatFile(string filePath, string fileContent)
        {

            FileStream fs1 = new FileStream(filePath, FileMode.Create, FileAccess.Write);//创建写入文件
            StreamWriter sw = new StreamWriter(fs1);
            sw.WriteLine(fileContent);//开始写入值
            sw.Close();
            fs1.Close();
        } //写入bat文件

        private void button5_Click(object sender, EventArgs e)
        {
            form1.richTextBox1.AppendText("正在写入批处理文件...\n");
            WriteBatFile(@"C:\Users\LLL\Desktop\buildings\1.bat", "python");
            form1.richTextBox1.AppendText("完成!\n");

            WriteBatFile(@"C:\Users\LLL\Desktop\buildings\1.bat", "python");

            WriteBatFile(@"C:\Users\LLL\Desktop\buildings\1.bat", "python");
            WriteBatFile(@"C:\Users\LLL\Desktop\buildings\1.bat", "python");

        }

        private void button6_Click(object sender, EventArgs e)
        {
            Form6 sectionclcu = new Form6();
            sectionclcu.ShowDialog();

        }



        private void Button11_Click(object sender, EventArgs e)
        {
            PictureBox1.Refresh();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();

        }
        private void 删除文件ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void 生成轴网_Click(object sender, EventArgs e)
        {

        }



        public class BeamXX
        {
            public void BeamX(double length, int num)
            {
                num = Form2.n;
                AssemblyDocument odoc = (AssemblyDocument)invapp.Documents.Open(Models_path + "梁" + "\\" + "x_梁.iam", false);
                foreach (ComponentOccurrence oCompOcc in odoc.ComponentDefinition.Occurrences)
                {
                    oCompOcc.Delete();
                }
                var oParams = odoc.ComponentDefinition.Parameters;
                var oUserParams = oParams.UserParameters;
                oUserParams["length"].Value = length;
                oUserParams["w_d"].Value = w_d;
                oUserParams["拼接段数"].Value = num;
                addin.Activate();
                dynamic myRule = null;
                myRule = "Rule_1";
                iLogicAutomation.RunRule(odoc, myRule);
                odoc.Update2();
                odoc.Save2(true);
                form1.richTextBox1.AppendText("x_梁已更新\n");
            }
            public void Beeem(int numss)
            {
                Console.WriteLine("s");
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            info info = new info();
            bom bom = new bom();
            readbom readbom = new readbom();
        }

        public class BeamY
        {
            public BeamY(double length, int num)
            {
                AssemblyDocument odoc = (AssemblyDocument)invapp.Documents.Open(Models_path + "梁" + "\\" + "Y_梁.iam", false);

                var oParams = odoc.ComponentDefinition.Parameters;
                var oUserParams = oParams.UserParameters;
                oUserParams["length"].Value = length;
                oUserParams["w_d"].Value = w_d;
                oUserParams["拼接段数"].Value = num;

                addin.Activate();
                foreach (ComponentOccurrence oCompOcc in odoc.ComponentDefinition.Occurrences)
                {
                    oCompOcc.Delete();
                }

                dynamic myRule = null;
                myRule = "Rule_1";

                iLogicAutomation.RunRule(odoc, myRule);

                odoc.Update2();
                odoc.Save2(true);
                form1.richTextBox1.AppendText("Y_梁已更新\n");


            }
        }
        public class Roof
        {
            public Roof(double length, double w_d, double theta)
            {
                AssemblyDocument odoc = (AssemblyDocument)invapp.Documents.Open(Models_path + "屋架" + "\\" + "屋架.iam", false);

                var oParams = odoc.ComponentDefinition.Parameters;
                var oUserParams = oParams.UserParameters;
                oUserParams["Y跨度"].Value = length;
                oUserParams["w_b"].Value = w_d;
                oUserParams["theta"].Value = theta;
                addin.Activate();

                dynamic myRule = null;
                foreach (dynamic eachRule in iLogicAutomation.Rules(odoc))
                {
                    if (eachRule.Name == "Rule_1")
                    {
                        myRule = eachRule;
                        //list the code of rule to the list box
                        //MessageBox.Show(myRule.Text);
                        break;
                    }
                    iLogicAutomation.RunRule(odoc, "MyRule");
                }
                odoc.Update2();
                odoc.Save2(true);
                form1.richTextBox1.AppendText("屋架已更新\n");
            }
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void 方钢管_Click(object sender, EventArgs e)
        {
            double size = Convert.ToDouble(comboBox1.Text);
            double thick = Convert.ToDouble(comboBox2.Text);
            int num = Convert.ToInt32(comboBox3.Text);
            double length = Convert.ToDouble(comboBox4.Text);
            RHS_section RHS = new RHS_section();
            RHS.mutiRHS(num, size, thick, 10, 1, length);

        }

        private void 关闭ToolStripMenuItem_Click(object sender, EventArgs e)
        {

            killapp killapp = new killapp();
            killapp.Killinventor();

        }

        private void button9_Click(object sender, EventArgs e)
        {

            richTextBox1.AppendText("s");
        }

        private void PictureBox1_Click(object sender, EventArgs e)
        {

        }

        public class RoofS//新建一个空白的生成
        {
            public RoofS(int purlin_num)
            {
                AssemblyDocument odoc = (AssemblyDocument)invapp.Documents.Add(DocumentTypeEnum.kAssemblyDocumentObject, assem_templatepath, true);

                var oParams = odoc.ComponentDefinition.Parameters;
                var oUserParams = oParams.UserParameters;
                oUserParams.AddByValue("x_distance", 4000, "mm");
                oUserParams.AddByValue("x_count", 5, "ul");
                oUserParams.AddByValue("purlin_num", 12, "ul");

                oUserParams["x_count"].Value = x_count;
                oUserParams["x_distance"].Value = x_distance;
                oUserParams["purlin_num"].Value = purlin_num;
                addin.Activate();

                dynamic MyRule = null;
                MyRule = @"I:\Wcode\VStudio\c_1114\模型文件\屋架\Rule_2.iLogicVb";

                iLogicAutomation.RunExternalRule(odoc, MyRule);

                odoc.Update();

                odoc.SaveAs(Models_path + "屋架" + "\\" + "全部屋架.iam", true);

                form1.richTextBox1.AppendText("全部屋架已更新\n");


            }

        }
        public class internalwalls
        {
            public internalwalls(double interwallnum)
            {
                AssemblyDocument odoc = (AssemblyDocument)invapp.Documents.Open(Models_path + "内墙" + "\\" + "组合内墙.iam", false);
                var oParams = odoc.ComponentDefinition.Parameters;
                var oUserParams = oParams.UserParameters;
                oUserParams["总跨度"].Value = z_distance - beam_height;
                oUserParams["总宽度"].Value = (y_distance - w_d) / interwallnum;
                oUserParams["总高度"].Value = w_d;
                oUserParams["数量"].Value = interwallnum;
                oUserParams["梁高"].Value = beam_height;
                addin.Activate();
                dynamic MyRule = null;
                MyRule = "Rule_1";
                iLogicAutomation.RunRule(odoc, MyRule);
                odoc.Update();
                odoc.Save2(true);
                form1.richTextBox1.AppendText("内墙已更新\n");
            }
        }

        public class externalwallx
        {

            public externalwallx(double exwallnum_x)
            {
                AssemblyDocument odoc = (AssemblyDocument)invapp.Documents.Open(Models_path + "外墙" + "\\" + "外墙x.iam", false);

                var oParams = odoc.ComponentDefinition.Parameters;
                var oUserParams = oParams.UserParameters;
                oUserParams["总跨度"].Value = (z_distance - beam_height) / 3;
                oUserParams["总宽度"].Value = (x_distance - w_d) / exwallnum_x;
                oUserParams["总高度"].Value = w_d;
                oUserParams["w_d"].Value = w_d;
                oUserParams["数量"].Value = exwallnum_x;
                addin.Activate();
                dynamic MyRule = null;

                MyRule = "Rule_2";
                iLogicAutomation.RunRule(odoc, MyRule);

                odoc.Update();

                odoc.Save2(true);
                form1.richTextBox1.AppendText("外墙X已更新\n");
                AssemblyDocument odoc1 = (AssemblyDocument)invapp.Documents.Open(Models_path + "外墙" + "\\" + "标准外墙x.iam", false);
                oParams = odoc1.ComponentDefinition.Parameters;
                oUserParams = oParams.UserParameters;
                oUserParams["w_d"].Value = w_d;
                oUserParams["梁高"].Value = beam_height;
                MyRule = "Rule_2";
                iLogicAutomation.RunRule(odoc1, MyRule);
                odoc1.Update();
                odoc1.Save2(true);
                form1.richTextBox1.AppendText("标准外墙X已更新\n");


            }
        }
        public class externalwally
        {

            public externalwally(double exwallnum_y)
            {
                AssemblyDocument odoc = (AssemblyDocument)invapp.Documents.Open(Models_path + "外墙" + "\\" + "外墙y.iam", false);

                var oParams = odoc.ComponentDefinition.Parameters;
                var oUserParams = oParams.UserParameters;
                oUserParams["总跨度"].Value = (z_distance - beam_height) / 3;
                oUserParams["总宽度"].Value = (y_distance - w_d) / exwallnum_y;
                oUserParams["总高度"].Value = w_d;
                oUserParams["w_d"].Value = w_d;
                oUserParams["数量"].Value = exwallnum_y;
                addin.Activate();
                dynamic MyRule = null;
                MyRule = "Rule_2";
                iLogicAutomation.RunRule(odoc, MyRule);
                odoc.Update();
                odoc.Save2(true);
                form1.richTextBox1.AppendText("外墙Y已更新\n");

                AssemblyDocument odoc1 = (AssemblyDocument)invapp.Documents.Open(Models_path + "外墙" + "\\" + "标准外墙y.iam", false);
                oParams = odoc1.ComponentDefinition.Parameters;
                oUserParams = oParams.UserParameters;
                oUserParams["w_d"].Value = w_d;
                oUserParams["梁高"].Value = beam_height;
                MyRule = "Rule_2";
                iLogicAutomation.RunRule(odoc1, MyRule);
                odoc1.Update();
                odoc1.Save2(true);
                form1.richTextBox1.AppendText("标准外墙Y已更新\n");
            }
        }




        public class RHS_section //5个参数
        {
            string part_templatepath = @"c:\users\public\documents\autodesk\inventor 2023\templates\zh-cn\metric\standard (mm).ipt";
            public void mutiRHS(int num, double b, double t, double r, int plane, double length)

            {
                b /= 10;
                t /= 10;
                r /= 10;
                // 1. 新建零件
                PartDocument opartdoc = (PartDocument)invapp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, part_templatepath, true);
                // 2. 新建2d草图   /  重命名
                var workplane = opartdoc.ComponentDefinition.WorkPlanes[plane];
                var osketch = opartdoc.ComponentDefinition.Sketches.Add(workplane);
                osketch.Name = "sketch_RHS";  // defaut is sketch1/2/3
                                              // 3.新建点(只是点)
                var point = new Point2d[11];
                point[0] = invapp.TransientGeometry.CreatePoint2d(-b / 2, -b / 2);
                point[1] = invapp.TransientGeometry.CreatePoint2d(b / 2, -b / 2);
                point[2] = invapp.TransientGeometry.CreatePoint2d(b / 2, b / 2);
                point[3] = invapp.TransientGeometry.CreatePoint2d(-b / 2, b / 2);
                var oCoord = new SketchPoint[11];
                oCoord[0] = osketch.SketchPoints.Add(point[0], false);
                oCoord[1] = osketch.SketchPoints.Add(point[1], false);
                oCoord[2] = osketch.SketchPoints.Add(point[2], false);
                oCoord[3] = osketch.SketchPoints.Add(point[3], false);
                var oLines = new SketchLine[4];
                oLines[0] = osketch.SketchLines.AddByTwoPoints(oCoord[0], oCoord[1]); // 线1
                oLines[1] = osketch.SketchLines.AddByTwoPoints(oCoord[1], oCoord[2]);
                oLines[2] = osketch.SketchLines.AddByTwoPoints(oCoord[2], oCoord[3]);
                oLines[3] = osketch.SketchLines.AddByTwoPoints(oCoord[3], oCoord[0]);
                //var sketchArc = new SketchArc[11];
                //sketchArc[0] = osketch.SketchArcs.AddByFillet((SketchEntity)oLines[0], (SketchEntity)oLines[1], r, oLines[0].StartSketchPoint.Geometry, oLines[1].EndSketchPoint.Geometry);
                //sketchArc[1] = osketch.SketchArcs.AddByFillet((SketchEntity)oLines[1], (SketchEntity)oLines[2], r, oLines[1].StartSketchPoint.Geometry, oLines[2].EndSketchPoint.Geometry);
                //sketchArc[2] = osketch.SketchArcs.AddByFillet((SketchEntity)oLines[2], (SketchEntity)oLines[3], r, oLines[2].StartSketchPoint.Geometry, oLines[3].EndSketchPoint.Geometry);
                //sketchArc[3] = osketch.SketchArcs.AddByFillet((SketchEntity)oLines[3], (SketchEntity)oLines[0], r, oLines[3].StartSketchPoint.Geometry, oLines[0].EndSketchPoint.Geometry);
                var oCollection = invapp.TransientObjects.CreateObjectCollection();
                oCollection.Add(oLines[0]);
                osketch.OffsetSketchEntitiesUsingDistance(oCollection, t, false, true);
                var oProfile = osketch.Profiles.AddForSolid();
                var oExtrudeDef = opartdoc.ComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProfile, (PartFeatureOperationEnum)20481);
                oExtrudeDef.SetDistanceExtent(length / 10, (PartFeatureExtentDirectionEnum)20993);
                var oExtrude = opartdoc.ComponentDefinition.Features.ExtrudeFeatures.Add(oExtrudeDef);
                string filename;
                b *= 10;
                t *= 10;

                for (int i = 0; i < num; i++)
                {
                    filename = string.Format("方钢管{0}-{1}-{2}-{3}.ipt", b, t, length, i);
                    opartdoc.SaveAs(save_path + filename, true);

                }
                string textname = string.Format("已生成{0}根方钢管\n", num);
                form1.richTextBox1.AppendText(textname);

            }



            public void singleRHS(double b, double t, double r, int plane, double length)
            {

                b /= 10;
                t /= 10;
                r /= 10;


                // 1. 新建零件

                PartDocument opartdoc = (PartDocument)invapp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, part_templatepath, true);
                // 2. 新建2d草图   /  重命名
                var workplane = opartdoc.ComponentDefinition.WorkPlanes[plane];
                var osketch = opartdoc.ComponentDefinition.Sketches.Add(workplane);
                osketch.Name = "sketch_RHS";  // defaut is sketch1/2/3
                                              // 3.新建点(只是点)
                var point = new Point2d[11];
                point[0] = invapp.TransientGeometry.CreatePoint2d(-b / 2, -b / 2);
                point[1] = invapp.TransientGeometry.CreatePoint2d(b / 2, -b / 2);
                point[2] = invapp.TransientGeometry.CreatePoint2d(b / 2, b / 2);
                point[3] = invapp.TransientGeometry.CreatePoint2d(-b / 2, b / 2);

                var oCoord = new SketchPoint[11];
                oCoord[0] = osketch.SketchPoints.Add(point[0], false);
                oCoord[1] = osketch.SketchPoints.Add(point[1], false);
                oCoord[2] = osketch.SketchPoints.Add(point[2], false);
                oCoord[3] = osketch.SketchPoints.Add(point[3], false);

                var oLines = new SketchLine[4];
                oLines[0] = osketch.SketchLines.AddByTwoPoints(oCoord[0], oCoord[1]); // 线1
                oLines[1] = osketch.SketchLines.AddByTwoPoints(oCoord[1], oCoord[2]);
                oLines[2] = osketch.SketchLines.AddByTwoPoints(oCoord[2], oCoord[3]);
                oLines[3] = osketch.SketchLines.AddByTwoPoints(oCoord[3], oCoord[0]);

                var sketchArc = new SketchArc[11];
                sketchArc[0] = osketch.SketchArcs.AddByFillet((SketchEntity)oLines[0], (SketchEntity)oLines[1], r, oLines[0].StartSketchPoint.Geometry, oLines[1].EndSketchPoint.Geometry);
                sketchArc[1] = osketch.SketchArcs.AddByFillet((SketchEntity)oLines[1], (SketchEntity)oLines[2], r, oLines[1].StartSketchPoint.Geometry, oLines[2].EndSketchPoint.Geometry);
                sketchArc[2] = osketch.SketchArcs.AddByFillet((SketchEntity)oLines[2], (SketchEntity)oLines[3], r, oLines[2].StartSketchPoint.Geometry, oLines[3].EndSketchPoint.Geometry);
                sketchArc[3] = osketch.SketchArcs.AddByFillet((SketchEntity)oLines[3], (SketchEntity)oLines[0], r, oLines[3].StartSketchPoint.Geometry, oLines[0].EndSketchPoint.Geometry);

                var oCollection = invapp.TransientObjects.CreateObjectCollection();

                oCollection.Add(oLines[0]);
                osketch.OffsetSketchEntitiesUsingDistance(oCollection, t, false, true);

                var oProfile = osketch.Profiles.AddForSolid();

                var oExtrudeDef = opartdoc.ComponentDefinition.Features.ExtrudeFeatures.CreateExtrudeDefinition(oProfile, (PartFeatureOperationEnum)20481);
                oExtrudeDef.SetDistanceExtent(length, (PartFeatureExtentDirectionEnum)20993);
                var oExtrude = opartdoc.ComponentDefinition.Features.ExtrudeFeatures.Add(oExtrudeDef);
                opartdoc.SaveAs(save_path + "RHS_section.ipt", true);
                form1.richTextBox1.AppendText("矩形截面已生成\n");














            }
        }           //5个参数



        public class H_section
        {
            public H_section()
            {












            }
        }//5个参数



        public class key_points
        {
            public key_points()
            {
                x_count = Convert.ToInt32(form1.TextBox2.Text);
                y_count = Convert.ToInt32(form1.TextBox4.Text);
                z_count = Convert.ToInt32(form1.TextBox5.Text);
                x_distance = Convert.ToDouble(form1.TextBox1.Text) / (double)(x_count - 1) * 100d;
                y_distance = Convert.ToDouble(form1.TextBox3.Text) / (double)(y_count - 1) * 100d;
                z_distance = Convert.ToDouble(form1.TextBox6.Text) / (double)z_count * 100d;
                // Dim points(x_count - 1, y_count - 1, z_count) As Vector3
                var points_count = x_count * y_count * z_count;
                // 把xyz坐标作为数组元素，值等于向量
                for (int i = 0, loopTo = z_count; i <= loopTo; i++)
                {
                    for (int j = 0, loopTo1 = y_count - 1; j <= loopTo1; j++)
                    {
                        for (int k = 0, loopTo2 = x_count - 1; k <= loopTo2; k++)
                            points[k, j, i].X = (float)(k * x_distance);
                    }
                }

                for (int i = 0, loopTo3 = z_count; i <= loopTo3; i++)
                {
                    for (int j = 0, loopTo4 = x_count - 1; j <= loopTo4; j++)
                    {
                        for (int k = 0, loopTo5 = y_count - 1; k <= loopTo5; k++)
                            points[j, k, i].Y = (float)(k * y_distance);
                    }
                }

                for (int i = 0, loopTo6 = x_count - 1; i <= loopTo6; i++)
                {
                    for (int j = 0, loopTo7 = y_count - 1; j <= loopTo7; j++)
                    {
                        for (int k = 0, loopTo8 = z_count; k <= loopTo8; k++)
                            points[i, j, k].Z = (float)(k * z_distance);
                    }
                }
                form1.richTextBox1.AppendText("当前时间：\n");
                form1.richTextBox1.AppendText(System.DateTime.Now.ToString("f") + "\n");
            }
        }





        public class Draw_grids //画平面轴网点.已完成，两个参数
        {
            public Draw_grids(double X, double Y)
            {
                x_count = Convert.ToInt32(form1.TextBox2.Text);
                y_count = Convert.ToInt32(form1.TextBox4.Text);
                z_count = Convert.ToInt32(form1.TextBox5.Text);
                x_distance = Convert.ToDouble(form1.TextBox1.Text) / (double)(x_count - 1) * 100d;
                y_distance = Convert.ToDouble(form1.TextBox3.Text) / (double)(y_count - 1) * 100d;
                z_distance = Convert.ToDouble(form1.TextBox6.Text) / (double)z_count * 100d;
                var grid_dimen = form1.PictureBox1.CreateGraphics();
                var color_black = System.Drawing.Color.FromName("black");
                var pen = new Pen(color_black);
                pen.Width = 2;
                Brush brush = new SolidBrush(color_black);
                int width = form1.PictureBox1.Size.Width - 40;
                int height = form1.PictureBox1.Size.Height - 40;
                var x_distance_draw = width / (X - 1);
                var y_distance_draw = height / (Y - 1);
                int x, y, n, m;
                n = 0;
                var loopTo = X;
                for (x = 1; x <= loopTo; x++)
                {
                    m = 0;
                    var loopTo1 = Y;
                    for (y = 1; y <= loopTo1; y++)
                    {
                        grid_dimen.DrawRectangle(pen, (float)(20 + n * x_distance_draw - 5), (float)(20 + m * y_distance_draw - 5), 10, 10);
                        m = m + 1;
                    }
                    n = n + 1;
                }
                Console.WriteLine(x_distance_draw);
            }
        }

        public class Draw_internalwalls //画平面轴网点.已完成，两个参数

        {
            public Draw_internalwalls(double X, double Y)
            {
                x_count = Convert.ToInt32(form1.TextBox2.Text);
                y_count = Convert.ToInt32(form1.TextBox4.Text);
                z_count = Convert.ToInt32(form1.TextBox5.Text);
                x_distance = Convert.ToDouble(form1.TextBox1.Text) / (double)(x_count - 1) * 100d;
                y_distance = Convert.ToDouble(form1.TextBox3.Text) / (double)(y_count - 1) * 100d;
                z_distance = Convert.ToDouble(form1.TextBox6.Text) / (double)z_count * 100d;
                var grid_dimen = form1.pictureBox2.CreateGraphics();
                var color_black = System.Drawing.Color.FromName("black");
                var pen = new Pen(color_black);
                pen.Width = 2;
                Brush brush = new SolidBrush(color_black);

                int width = form1.pictureBox2.Size.Width - 40;
                int height = form1.pictureBox2.Size.Height - 40;
                var x_distance_draw = width / (X - 1);
                var y_distance_draw = height / (Y - 1);
                int x, y, n, m;
                n = 0;
                var loopTo = X;
                for (x = 1; x <= loopTo; x++)
                {
                    m = 0;
                    var loopTo1 = Y;
                    for (y = 1; y <= loopTo1; y++)
                    {
                        grid_dimen.DrawRectangle(pen, (float)(20 + n * x_distance_draw - 5), (float)(20 + m * y_distance_draw - 5), 10, 10);
                        m = m + 1;
                    }
                    n = n + 1;
                }
                Console.WriteLine(x_distance_draw);

            }
        }







        public class Connect//连接通信
        {
            public Connect()
            {
                try
                {
                    invapp = (Inventor.Application)Marshal.GetActiveObject("Inventor.Application");

                    System.Threading.Thread.Sleep(1 * 2000);

                    PartDocument opartdoc = (PartDocument)invapp.Documents.Add(DocumentTypeEnum.kPartDocumentObject, part_templatepath, false);
                    addin = invapp.ApplicationAddIns.ItemById[ClientID];

                    iLogicAutomation = addin.Automation;
                    form1.richTextBox1.AppendText("已连接。\n");

                }
                catch (Exception ex)
                {
                    try
                    {
                        form1.richTextBox1.AppendText("正在启动.......\n");
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


            }
        }
        public class info
        {
            public info()
            {

                AssemblyDocument odoc = (AssemblyDocument)invapp.ActiveDocument;
                double X_length = Convert.ToDouble(form1.TextBox1.Text);
                double Y_length = Convert.ToDouble(form1.TextBox3.Text);
                double Z_length = Convert.ToDouble(form1.TextBox6.Text);
                double Z_num = Convert.ToDouble(form1.TextBox5.Text);
                var oMassProps = odoc.ComponentDefinition.MassProperties;

                oMassProps.Accuracy = MassPropertiesAccuracyEnum.k_Medium;
                var mass = Convert.ToString((int)oMassProps.Mass / 1000);
                var steelusuage = (int)oMassProps.Mass / (X_length * Y_length * Z_num) / 2;
                steelusuage = Math.Round(steelusuage, 2);
                form1.richTextBox1.AppendText("该建筑为" + X_length + "×" + Y_length + "×" + Z_length + "m" + "的" +
                    Z_num + "层框架结构\n");
                form1.richTextBox1.AppendText("上部结构总重量为:" + mass + "吨" + "\n");
                form1.richTextBox1.AppendText("用钢量为:" + steelusuage + "KG/平方米" + "\n");
            }
        }
        public class bom
        {
            public bom()
            {
                Application app = new Application();
                app.Workbooks.Close();
                app.Quit();
                AssemblyDocument odoc = (AssemblyDocument)invapp.ActiveDocument;

                BOM oBOM = odoc.ComponentDefinition.BOM;
                oBOM.StructuredViewFirstLevelOnly = false;
                oBOM.StructuredViewEnabled = true;

                BOMView assemblyBOMView = oBOM.BOMViews["结构化"];


                assemblyBOMView.Export(save_path + "bills.xlsx", FileFormatEnum.kMicrosoftExcelFormat);
                form1.richTextBox1.AppendText("材料表已创建" + "\n");


            }
        }
        public class readbom
        {
            public readbom()
            {

                Application app = new Application();
                Workbook workbook = app.Workbooks.Open(save_path + "bills.xlsx");
                //Workbook workbook = app.Workbooks.Open(save_path + "\\" + "bills1.xlsx");
                //读取工作表，索引由1开始。
                Worksheet worksheet = (Worksheet)workbook.Worksheets[1];
                form1.richTextBox1.AppendText("正在读取材料表..." + "\n");
                //保存原文件
                //workbook.Save();
                //保存为新的Excel文件
                //结尾记得关闭服务，不然会导致excel在后台开启

                Range xlRange = worksheet.UsedRange;
                //定义空数组
                string[] partname = new string[30];
                string[] partnum = new string[30];

                int totalnum = 0;//构件总数求和

                for (int i = 2; i < 22; i++)
                {
                    partname[i - 2] = ((Range)xlRange.Cells[i, 2]).Value2.ToString();//注意cell 第一个元素是列第二个是行

                    partnum[i - 2] = ((Range)xlRange.Cells[i, 6]).Value2.ToString();
                    totalnum = totalnum + Convert.ToInt32(partnum[i - 2]);
                    form1.richTextBox1.AppendText(partname[i - 2] + ":" + partnum[i - 2] + "\n");
                }

                //form1.richTextBox1.AppendText(partname[i - 2] + ":" + partnum[i - 2] + "\n");
                //form1.richTextBox1.AppendText(partname[1] + "\n");
                form1.richTextBox1.AppendText("所有构件数量:" + totalnum + "\n");
                form1.richTextBox1.AppendText("螺栓数量:" + totalnum * 4 + "\n");
                workbook.Close();
                app.Quit();
                invapp.Quit();
            }
        }
        // write a line



    }
}



