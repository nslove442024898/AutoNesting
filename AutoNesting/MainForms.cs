
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

using acadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using acadWin = Autodesk.AutoCAD.Windows;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;

namespace AutoNesting
{
    public partial class MainForms : Form
    {
        public List<Part> listParts = new List<Part>();
        public Document acDoc;

        public MainForms()
        {
            InitializeComponent();
        }

        private void MainForms_Load(object sender, EventArgs e)
        {
            for (int i = 0; i < listParts.Count; i++)
            {
                var index = this.dataGridView1.Rows.Add(listParts[i].partObjId.ToString(), listParts[i].partNumber, listParts[i].Length, listParts[i].Width, Math.Round(listParts[i].Area, 0));
            }
        }

        private void TextBox_Leave(object sender, EventArgs e)
        {
            var txtbox = sender as TextBox;
            if (!txtbox.Text.IsNumeric())
            {
                if (MessageBox.Show("你输入了一个非法字符！", "数字检查", MessageBoxButtons.RetryCancel, MessageBoxIcon.Exclamation) == DialogResult.Retry) txtbox.Focus(); else txtbox.Text = "";
            }
        }

        private void ComboBox1_MouseDown(object sender, MouseEventArgs e)
        {
            this.comboBox1.Items.Clear();
            string fileName = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\原材料设置文件.txt";
            if (File.Exists(fileName))
            {
                var rawMatls = File.ReadAllLines(fileName, Encoding.UTF8);
                foreach (var item in rawMatls)
                {
                    var size = item.Split(new char[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
                    this.comboBox1.Items.Add($"{size[0]}x{size[1]}");
                }
            }
            else MessageBox.Show("原材料配置文件丢失，请检查程序文件目录");
        }

        private void ComboBox2_MouseDown(object sender, MouseEventArgs e)
        {
            this.comboBox2.Items.Clear();
            string fileName = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\零件间隔参数设置.txt";
            if (File.Exists(fileName))
            {
                var partGap = File.ReadAllLines(fileName, Encoding.UTF8);
                foreach (var item in partGap)
                {
                    this.comboBox2.Items.Add($"{item}");
                }
            }
            else MessageBox.Show("零件间隔参数设置丢失，请检查程序文件目录");
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            if (this.comboBox1.Text != string.Empty && this.comboBox2.Text != string.Empty && this.comboBox3.Text != string.Empty)
            {
                List<RawMatl> lsitNCformats = new List<RawMatl>();
                var ed = acadApp.DocumentManager.MdiActiveDocument.Editor;
                var ppr = ed.GetPoint("拾取一个点摆放套料结果");
                if (ppr.Status != PromptStatus.OK) return;
                var inpnt = ppr.Value;
                string[] sizeofRaw = this.comboBox1.Text.Split(new char[] { 'x' }, StringSplitOptions.RemoveEmptyEntries);
                int rawCount = 0;
                using (DocumentLock m_DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
                {
                    while (this.listParts.Count > 0)
                    {
                        RawMatl r = new RawMatl(this.acDoc.Database, new Point3d(inpnt.X + 3 * double.Parse(sizeofRaw[0]) * rawCount, inpnt.Y, inpnt.Z), double.Parse(sizeofRaw[0]), double.Parse(sizeofRaw[1]));
                         r.NestedPart2Plate(this.listParts, double.Parse(this.comboBox2.Text), double.Parse(this.comboBox3.Text));
                        rawCount++;
                        lsitNCformats.Add(r);
                    }
                }
            }
            else MessageBox.Show("选择原材料和零件间隔参数！");
        }

        private void ComboBox3_MouseDown(object sender, MouseEventArgs e)
        {
            //零件距板边间隔
            this.comboBox3.Items.Clear();
            string fileName = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\Resources\零件距板边间隔.txt";
            if (File.Exists(fileName))
            {
                var partGap = File.ReadAllLines(fileName, Encoding.UTF8);
                foreach (var item in partGap)
                {
                    this.comboBox3.Items.Add($"{item}");
                }
            }
            else MessageBox.Show("零件间隔参数设置丢失，请检查程序文件目录");


        }


        private void DataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (!this.dataGridView1.SelectedRows[0].IsNewRow)
            {
                string oidString = this.dataGridView1.SelectedRows[0].Cells[0].Value.ToString().Replace('(', ' ').Replace(')', ' ').Trim();
                long intId = Convert.ToInt64(oidString);//这里的strId是一个纯数字字符串，将其转换成64的long类型，32的会报错
                IntPtr init = new IntPtr(intId);
                ObjectId oid = new ObjectId(init);
                using (DocumentLock m_DocumentLock = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.LockDocument())
                {
                    using (Transaction trans = acDoc.Database.TransactionManager.StartTransaction())
                    {
                        Polyline pline = oid.GetObject(OpenMode.ForWrite) as Polyline;
                        var thisdrawing = acDoc.GetAcadDocument() as Autodesk.AutoCAD.Interop.AcadDocument;
                        var pnts = pline.GeometricExtents;
                        thisdrawing.Application.ZoomWindow(new double[3] { pnts.MinPoint.X, pnts.MinPoint.Y, pnts.MinPoint.Z },
                            new double[3] { pnts.MaxPoint.X, pnts.MaxPoint.Y, pnts.MaxPoint.Z });
                        trans.Commit();
                        thisdrawing.Application.Update();
                    }
                }
            }
        }

        private void DataGridView1_CellMouseLeave(object sender, DataGridViewCellEventArgs e)
        {
            this.dataGridView1.SelectionChanged -= DataGridView1_SelectionChanged;
        }

        private void DataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            this.dataGridView1.SelectionChanged += DataGridView1_SelectionChanged;
        }

    }
}
