using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;

namespace AutoNesting
{
    public class FunctionCls01
    {
        public MainForms mfrm { get; set; }
        [CommandMethod("myNetDllUninstall")]
        public void 取消自启动()//
        {
            helper.UnRegister2HKCR();
        }
        [CommandMethod("myNetDllinstall")]
        public void 设置自启动()
        {
            helper.Register2HKCR();
        }
        [CommandMethod("myAutoNesting")]
        public void 零件自动套料()
        {
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acDb = acDoc.Database;
            Editor acEd = acDoc.Editor;
            //if (Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("NEXTFIBERWORLD").ToString() != "0") Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("NEXTFIBERWORLD", 0);
            if (acEd.CurrentUserCoordinateSystem.IsEqualTo(Matrix3d.Identity) == false) acEd.CurrentUserCoordinateSystem = Matrix3d.Identity;
            List<Part> listParts = new List<Part>();
            PromptSelectionOptions pso = new PromptSelectionOptions();
            pso.RejectObjectsFromNonCurrentSpace = true;
            pso.AllowDuplicates = false;
            pso.MessageForAdding = "选择零件" + Environment.NewLine;
            SelectionFilter sf = new SelectionFilter(new TypedValue[] { new TypedValue((int)DxfCode.Start, RXClass.GetClass(typeof(Polyline)).DxfName) });
            var psr = acEd.GetSelection(sf);
            if (psr.Status != PromptStatus.OK) return;
            int i = 1;
            //System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
            //sw.Start();
            foreach (SelectedObject dbobj in psr.Value)
            {
                using (Transaction trans = acDb.TransactionManager.StartTransaction())
                {
                    Polyline pline = trans.GetObject(dbobj.ObjectId, OpenMode.ForWrite) as Polyline;
                    Part p = new Part(pline, i, acEd);
                    DBText t = new DBText() { Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 171), TextString = "Part" + p.partNumber, Position = p.CenterPnt, Rotation = p.RotateAngle };
                    if (p.Length < p.Width) t.Height = p.Length / 5; else t.Height = p.Width / 5;
                    p.itemInsidePart.Add(Tools.AddToModelSpace(acDb, t));
                    listParts.Add(p);
                    trans.Commit();
                }
                i++;
            }
            var orderParts = listParts.OrderBy(c => c.Length).ThenBy(c => c.Width).ToList();
            if (orderParts.Count > 0)
            {
                if (this.mfrm == null)
                {
                    this.mfrm = new MainForms { listParts = orderParts, acDoc = acDoc };
                    this.mfrm.FormClosing += Mfrm_FormClosing;
                    Application.ShowModelessDialog(this.mfrm);
                }
            }
            else Application.ShowAlertDialog("未找到零件请重新选择零件！");
            //Autodesk.AutoCAD.ApplicationServices.Application.SetSystemVariable("NEXTFIBERWORLD", 1);
        }

        private void Mfrm_FormClosing(object sender, System.Windows.Forms.FormClosingEventArgs e)
        {
            this.mfrm = null;
        }
    }
    public class Part
    {
        public int partNumber { get; set; }
        public double Area { get; set; }
        public double Length { get; set; }
        public double Width { get; set; }
        public double RotateAngle { get; set; }
        public Point3d CenterPnt { get; set; }
        public ObjectId partObjId { get; set; }
        public List<ObjectId> itemInsidePart { get; set; }
        public Part(Polyline partBoundary, int num, Editor ed)
        {
            this.Area = partBoundary.Area;
            this.partNumber = num;
            this.itemInsidePart = new List<ObjectId>();
            this.partObjId = partBoundary.ObjectId;
            Point3dCollection pnts = new Point3dCollection();//存放多线的顶点
            List<double[]> pntsDis = new List<double[]>();//存放2个点的距离
            for (int i = 0; i < partBoundary.NumberOfVertices; i++)
            {
                //if (partBoundary.GetBulgeAt(i) > 1)//A bulge of 0 indicates a straight segment, and a bulge of 1 is a semicircle.
                //{
                var curPnt = partBoundary.GetPoint3dAt(i);
                if (i != partBoundary.NumberOfVertices - 1)
                {
                    var NextPnt = partBoundary.GetPoint3dAt(i + 1);
                    pntsDis.Add(new double[] { curPnt.DistanceTo(NextPnt), curPnt.X, curPnt.Y, NextPnt.X, NextPnt.Y });
                }
                pnts.Add(curPnt);
            }
            var psr = ed.SelectWindowPolygon(pnts);
            if (psr.Status == PromptStatus.OK) foreach (SelectedObject item in psr.Value) this.itemInsidePart.Add(item.ObjectId);
            var pntCenter = new Point3d(
                (partBoundary.GeometricExtents.MaxPoint.X + partBoundary.GeometricExtents.MinPoint.X) * 0.5,
                (partBoundary.GeometricExtents.MaxPoint.Y + partBoundary.GeometricExtents.MinPoint.Y) * 0.5,
                 (partBoundary.GeometricExtents.MaxPoint.Z + partBoundary.GeometricExtents.MinPoint.Z) * 0.5);
            double partLength = partBoundary.GeometricExtents.MaxPoint.X - partBoundary.GeometricExtents.MinPoint.X;
            double partWidth = partBoundary.GeometricExtents.MaxPoint.Y - partBoundary.GeometricExtents.MinPoint.Y;
            this.CenterPnt = pntCenter;
            List<double[]> maxDis = pntsDis.Where(c => c[0] == pntsDis.Max(t => t[0])).ToList();
            if (maxDis.Count(c => c[1] == c[3] || c[2] == c[4]) > 0)
            {
                this.Length = Math.Round(partLength, 0);
                this.Width = Math.Round(partWidth, 0);
                if (this.Length < this.Width)
                {
                    var temp = this.Length;
                    this.Length = this.Width;
                    this.Width = temp;
                    this.RotateAngle = Math.PI * 0.5;
                }
                else this.RotateAngle = 0;
            }
            else
            {
                var detlaX = Math.Abs(maxDis.First()[1] - maxDis.First()[3]);
                var detlaY = Math.Abs(maxDis.First()[2] - maxDis.First()[4]);
                this.RotateAngle = Math.Atan(detlaY / detlaX);
                using (Transaction trans = partBoundary.Database.TransactionManager.StartTransaction())
                {
                    EntTools.Rotate(partBoundary, pntCenter, Math.PI * 2 - this.RotateAngle);
                    this.Length = Math.Round(partBoundary.GeometricExtents.MaxPoint.X - partBoundary.GeometricExtents.MinPoint.X, 0);
                    this.Width = Math.Round(partBoundary.GeometricExtents.MaxPoint.Y - partBoundary.GeometricExtents.MinPoint.Y, 0);
                }
            }
        }
    }

    public class RawMatl
    {
        public double Length { get; set; }
        public double Width { get; set; }
        public Point3d LeftBtmPnt { get; set; }
        public Point3d RightTopPnt { get; set; }

        public ObjectId Oid { get; set; }
        public Database CurDb { get; set; }

        public List<Part> NestedParts { get; set; }

        public RawMatl(Database db, Point3d insPnt, double _l, double _w)
        {
            this.LeftBtmPnt = insPnt;
            this.Length = _l;
            this.Width = _w;
            this.CurDb = db;
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                Polyline pline = new Polyline();
                pline.CreateRectangle(new Point2d(insPnt.X, insPnt.Y), new Point2d(insPnt.X + _l, insPnt.Y + _w));
                pline.Layer = "0";
                pline.Color = Color.FromColorIndex(ColorMethod.ByAci, 201);
                pline.ConstantWidth = 5;
                this.Oid = Tools.AddToModelSpace(db, pline);
                this.RightTopPnt = new Point3d(this.LeftBtmPnt.X + this.Length, this.LeftBtmPnt.Y + this.Width, this.LeftBtmPnt.Z);
                trans.Commit();
            }
        }

        public void NestedPart2Plate(List<Part> allParts, double partsGap, double partPlateGap)
        {
            Point3d partIns = new Point3d(this.LeftBtmPnt.X + partPlateGap, this.LeftBtmPnt.Y + partPlateGap, this.LeftBtmPnt.Z);
            List<Part> nestedPart = new List<Part>();
            double lengthLeft = this.Length;
            double widthLeft = this.Width;
            bool firstParts = true;
            NextPartLocation nextPartLoc = NextPartLocation.putRight;//当前零件如何摆放
            Part NextParts = null; Part curpart = null;
            do
            {
                if (NextParts == null) { curpart = allParts[allParts.Count - 1]; } else curpart = NextParts;
                using (Transaction trans = this.CurDb.TransactionManager.StartTransaction())
                {
                    Polyline curpart_Shape = trans.GetObject(curpart.partObjId, OpenMode.ForWrite) as Polyline;
                    DBObjectCollection dbobjs = new DBObjectCollection();
                    if (curpart.itemInsidePart.Count > 0) foreach (ObjectId item in curpart.itemInsidePart) dbobjs.Add(trans.GetObject(item, OpenMode.ForWrite));
                    Point3d curPartLeftbtm = Point3d.Origin;
                    #region//旋转套入当前零件
                    if (curpart.RotateAngle != 0 || curpart.RotateAngle != 0.5 * Math.PI)//不是水平的零件
                    {
                        curpart_Shape.Rotate(curpart.CenterPnt, Math.PI * 2 - curpart.RotateAngle);//把零件转平
                        double l = curpart_Shape.GeometricExtents.MaxPoint.X - curpart_Shape.GeometricExtents.MinPoint.X;
                        double w = curpart_Shape.GeometricExtents.MaxPoint.Y - curpart_Shape.GeometricExtents.MinPoint.Y;
                        Point3d newCen = new Point3d(0.5 * (curpart_Shape.GeometricExtents.MaxPoint.X + curpart_Shape.GeometricExtents.MinPoint.X),
                            0.5 * (curpart_Shape.GeometricExtents.MaxPoint.Y + curpart_Shape.GeometricExtents.MinPoint.Y),
                            0.5 * (curpart_Shape.GeometricExtents.MaxPoint.Z + curpart_Shape.GeometricExtents.MinPoint.Z));
                        curPartLeftbtm = new Point3d(newCen.X - 0.5 * l, newCen.Y - 0.5 * w, newCen.Z);
                        if (dbobjs.Count > 0)
                        {
                            foreach (DBObject item in dbobjs) (item as Entity).Rotate(curpart.CenterPnt, Math.PI * 2 - curpart.RotateAngle);
                            foreach (DBObject item in dbobjs) (item as Entity).Move(partIns, curPartLeftbtm);
                        }
                    }
                    else
                    {
                        if (curpart.RotateAngle == Math.PI * 0.5)//竖直零件
                        {
                            curpart_Shape.Rotate(curpart.CenterPnt, Math.PI * 1.5);//把零件转平
                            curPartLeftbtm = new Point3d(curpart.CenterPnt.X - 0.5 * curpart.Length, curpart.CenterPnt.Y - 0.5 * curpart.Width, curpart.CenterPnt.Z);
                            if (dbobjs.Count > 0)
                            {
                                foreach (DBObject item in dbobjs) (item as Entity).Rotate(curpart.CenterPnt, Math.PI * 1.5);
                                foreach (DBObject item in dbobjs) (item as Entity).Move(partIns, curPartLeftbtm);
                            }
                        }
                        else//水平的零件
                        {
                            curPartLeftbtm = new Point3d(curpart.CenterPnt.X - 0.5 * curpart.Length, curpart.CenterPnt.Y - 0.5 * curpart.Width, curpart.CenterPnt.Z);
                            if (dbobjs.Count > 0) foreach (DBObject item in dbobjs) (item as Entity).Move(partIns, curPartLeftbtm);
                        }
                    }
                    curpart_Shape.Move(partIns, curPartLeftbtm);
                    trans.Commit();
                    #endregion
                    if (nextPartLoc == NextPartLocation.putUp)
                    {
                        lengthLeft = this.Length - partsGap - curpart.Length;//剩余长度
                        widthLeft = widthLeft - partsGap - curpart.Width;//剩余宽度
                    }
                    else
                    {
                        if (firstParts)
                        {
                            lengthLeft = lengthLeft - partsGap - curpart.Length - partPlateGap;
                            widthLeft = widthLeft - partsGap - curpart.Width - partPlateGap;
                            firstParts = false;
                        }
                        else lengthLeft = lengthLeft - partsGap - curpart.Length - partPlateGap;

                    }//剩余长度

                    if (allParts.Count > 0)
                    {
                        var fitLengthParts = allParts.Where(c => c.Length <= lengthLeft && c.Width <= curpart.Width && c.partObjId != curpart.partObjId).ToList();
                        var fitWidthParts = allParts.Where(c => c.Width <= widthLeft && c.partObjId != curpart.partObjId).ToList();
                        if (fitLengthParts.Count == 0 && fitWidthParts.Count == 0)//长度宽度都不合适的零件
                        {
                            NextParts = null;
                        }
                        else
                        {
                            if (fitLengthParts.Count > 0 && fitWidthParts.Count > 0)
                            {
                                NextParts = fitLengthParts.OrderByDescending(c => c.Width).ToList()[0];//
                                partIns = new Point3d(partIns.X + partsGap + curpart.Length, partIns.Y, partIns.Z);//
                                nextPartLoc = NextPartLocation.putRight;
                            }
                            else
                            {
                                if (fitWidthParts.Count > 0)
                                {
                                    NextParts = fitWidthParts.OrderByDescending(c => c.Length).ToList()[0];//
                                    partIns = new Point3d(this.LeftBtmPnt.X + partPlateGap, this.LeftBtmPnt.Y + partPlateGap + (this.Width - widthLeft) + partsGap, partIns.Z);//
                                    nextPartLoc = NextPartLocation.putUp;
                                }
                                else
                                {
                                    NextParts = fitLengthParts.OrderByDescending(c => c.Width).ToList()[0];//
                                    partIns = new Point3d(partIns.X + partsGap + curpart.Length, partIns.Y, partIns.Z);//
                                    nextPartLoc = NextPartLocation.putRight;
                                }
                            }

                        }
                    }
                    else NextParts = null;
                }
                allParts.Remove(curpart);//移除义套入的零件
                nestedPart.Add(curpart);
            } while (NextParts != null);
        }
    }
    public enum NextPartLocation
    {
        putRight,
        putUp
    }
}
