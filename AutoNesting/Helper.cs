using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Interop;
using Autodesk.AutoCAD.Runtime;
using Microsoft.Win32;

namespace AutoNesting
{
    public static class helper
    {/// <summary>
     ///提取所有的命令
     /// </summary>
     /// <param name="dllFiles">dll的路径</param>
     /// <returns></returns>
        public static List<myAcCmd> GetDllCmds()
        {
            List<myAcCmd> cmds = new List<myAcCmd>();
            #region 提取所以的命令
            Assembly ass = Assembly.GetExecutingAssembly();
            var clsCollection = ass.GetTypes().Where(t => t.IsClass && t.IsPublic).ToList();
            if (clsCollection.Count > 0)
            {
                foreach (var cls in clsCollection)
                {
                    var methods = cls.GetMethods().Where(m => m.IsPublic && m.GetCustomAttributes(true).Length > 0).ToList();
                    if (methods.Count > 0)
                    {
                        foreach (MethodInfo mi in methods)
                        {
                            var atts = mi.GetCustomAttributes(true).Where(c => c is CommandMethodAttribute).ToList();
                            if (atts.Count == 1)
                            {
                                myAcCmd cmd = new myAcCmd(  (atts[0] as CommandMethodAttribute).GlobalName, mi.Name);
                                cmds.Add(cmd);
                            }
                        }
                    }
                }
            }
            #endregion
            return cmds;
            //
        }
        public static void AddCmdtoMenuBar(List<myAcCmd> cmds,string MenuName)
        {
            var gcadApp = Application.AcadApplication as AcadApplication;
            AcadMenuGroup mg = null;
            for (int i = 0; i < gcadApp.MenuGroups.Count; i++) if (gcadApp.MenuGroups.Item(i).Name == "ACAD") mg = gcadApp.MenuGroups.Item(i);
            for (int i = 0; i < mg.Menus.Count; i++) if (mg.Menus.Item(i).Name == MenuName) mg.Menus.Item(i).RemoveFromMenuBar();
            AcadPopupMenu popMenu = mg.Menus.Add(MenuName);
            for (int i = 0; i < cmds.Count; i++)
            {
                var dllPopMenu = popMenu.AddMenuItem(popMenu.Count + 1, cmds[i].Label, cmds[i].Macro+" ");
            }
            popMenu.InsertInMenuBar(mg.Menus.Count + 1);
        }
        public static void DeleteCmdMenu(string MenuName)
        {
            var gcadApp = Application.AcadApplication as AcadApplication;
            AcadMenuGroup mg = null;
            for (int i = 0; i < gcadApp.MenuGroups.Count; i++) if (gcadApp.MenuGroups.Item(i).Name == "ACAD") mg = gcadApp.MenuGroups.Item(i);
            for (int i = 0; i < mg.Menus.Count; i++) if (mg.Menus.Item(i).Name == MenuName) mg.Menus.Item(i).RemoveFromMenuBar();
        }

        /// <summary>
        /// 将菜单加载到AutoCAD
        /// </summary>
        public static void Register2HKCR()
        {
            string hkcrKey = HostApplicationServices.Current.UserRegistryProductRootKey;
            var assName = Assembly.GetExecutingAssembly().CodeBase;
            var apps_Acad = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(Path.Combine(hkcrKey, "Applications"));
            if (apps_Acad.GetSubKeyNames().Count(c => c == Path.GetFileNameWithoutExtension(assName)) == 0)
            {
                var myNetLoader = apps_Acad.CreateSubKey(Path.GetFileNameWithoutExtension(assName), RegistryKeyPermissionCheck.Default);
                myNetLoader.SetValue("DESCRIPTION", "加载自定义dll文件", Microsoft.Win32.RegistryValueKind.String);
                myNetLoader.SetValue("LOADCTRLS", 2, Microsoft.Win32.RegistryValueKind.DWord);
                myNetLoader.SetValue("LOADER", assName, Microsoft.Win32.RegistryValueKind.String);
                myNetLoader.SetValue("MANAGED", 1, Microsoft.Win32.RegistryValueKind.DWord);
                Application.ShowAlertDialog(Path.GetFileNameWithoutExtension(assName) + "自动加载设置成功，重启AutoCAD 生效！");
            } else Application.ShowAlertDialog(Path.GetFileNameWithoutExtension(assName) + "程序已经设置为自启动，无需重复设置！");
        }
        public static void UnRegister2HKCR()
        {
            string hkcrKey = HostApplicationServices.Current.UserRegistryProductRootKey;
            var assName = Assembly.GetExecutingAssembly().CodeBase;
            var apps_Acad = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(Path.Combine(hkcrKey, "Applications"));
            if (apps_Acad.GetSubKeyNames().Count(c => c == Path.GetFileNameWithoutExtension(assName)) != 0)
            {
                var myNetLoader = apps_Acad.CreateSubKey(Path.GetFileNameWithoutExtension(assName), RegistryKeyPermissionCheck.Default);
                apps_Acad.DeleteSubKeyTree(Path.GetFileNameWithoutExtension(assName));
                Application.ShowAlertDialog(Path.GetFileNameWithoutExtension(assName) + "程序卸载完成，重启AutoCAD 生效！");
            } else Application.ShowAlertDialog(Path.GetFileNameWithoutExtension(assName) + "程序没有设置为自启动，无需重复删除！");
        }
        public static bool CheckFileReadOnly(string fileName)
        {
            bool inUse = true;
            FileStream fs = null;
            try
            {
                fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.None);
                inUse = false;
            }
            catch { }
            return inUse;//true表示正在使用,false没有使用  
        }
    }

    /// <summary>
    /// 图层操作类
    /// </summary>
    public static class LayerTools
    {
        /// <summary>
        /// 创建新图层
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="layerName">图层名</param>
        /// <returns>返回新建图层的ObjectId</returns>
        public static ObjectId AddLayer(this Database db, string layerName)
        {
            //打开层表
            LayerTable lt = (LayerTable)db.LayerTableId.GetObject(OpenMode.ForRead);
            if (!lt.Has(layerName))//如果不存在名为layerName的图层，则新建一个图层
            {
                //定义一个新的层表记录
                LayerTableRecord ltr = new LayerTableRecord();
                ltr.Name = layerName;//设置图层名
                lt.UpgradeOpen();//切换层表的状态为写以添加新的图层
                //将层表记录的信息添加到层表中
                lt.Add(ltr);
                //把层表记录添加到事务处理中
                db.TransactionManager.AddNewlyCreatedDBObject(ltr, true);
                lt.DowngradeOpen();//为了安全，将层表的状态切换为读
            }
            return lt[layerName];//返回新添加的层表记录的ObjectId
        }

        /// <summary>
        /// 设置图层的颜色
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="layerName">图层名</param>
        /// <param name="colorIndex">颜色索引</param>
        /// <returns>如果成功设置图层颜色，则返回true，否则返回false</returns>
        public static bool SetLayerColor(this Database db, string layerName, short colorIndex)
        {
            //打开层表
            LayerTable lt = (LayerTable)db.LayerTableId.GetObject(OpenMode.ForRead);
            //如果不存在名为layerName的图层，则返回
            if (!lt.Has(layerName)) return false;
            ObjectId layerId = lt[layerName];//获取名为layerName的层表记录的Id
            //以写的方式打开名为layerName的层表记录
            LayerTableRecord ltr = (LayerTableRecord)layerId.GetObject(OpenMode.ForWrite);
            //设置图层的颜色
            ltr.Color = Color.FromColorIndex(ColorMethod.ByAci, colorIndex);
            ltr.DowngradeOpen();//为了安全，将图层的状态切换为读
            return true;//设置图层颜色成功
        }

        /// <summary>
        /// 将指定的图层设置为当前层
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="layerName">图层名</param>
        /// <returns>如果设置成功，则返回ture</returns>
        public static bool SetCurrentLayer(this Database db, string layerName)
        {
            //打开层表
            LayerTable lt = (LayerTable)db.LayerTableId.GetObject(OpenMode.ForRead);
            //如果不存在名为layerName的图层，则返回
            if (!lt.Has(layerName)) return false;
            //获取名为layerName的层表记录的Id
            ObjectId layerId = lt[layerName];
            //如果指定的图层为当前层，则返回
            if (db.Clayer == layerId) return false;
            db.Clayer = layerId;//指定当前层
            return true;//指定当前图层成功
        }

        /// <summary>
        /// 获取当前图形中所有的图层
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <returns>返回所有的层表记录</returns>
        public static List<LayerTableRecord> GetAllLayers(this Database db)
        {
            //打开层表
            LayerTable lt = (LayerTable)db.LayerTableId.GetObject(OpenMode.ForRead);
            //用于返回层表记录的列表
            List<LayerTableRecord> ltrs = new List<LayerTableRecord>();
            foreach (ObjectId id in lt)//遍历层表
            {
                //打开层表记录
                LayerTableRecord ltr = (LayerTableRecord)id.GetObject(OpenMode.ForRead);
                ltrs.Add(ltr);//添加到返回列表中
            }
            return ltrs;//返回所有的层表记录
        }

        /// <summary>
        /// 删除指定名称的图层
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="layerName">图层名</param>
        /// <returns>如果删除成功，则返回true，否则返回false</returns>
        public static bool DeleteLayer(this Database db, string layerName)
        {
            //打开层表
            LayerTable lt = (LayerTable)db.LayerTableId.GetObject(OpenMode.ForRead);
            //如果层名为0或Defpoints，则返回（这两个图层不能删除）
            if (layerName == "0" || layerName == "Defpoints") return false;
            //如果不存在名为layerName的图层，则返回
            if (!lt.Has(layerName)) return false;
            ObjectId layerId = lt[layerName];//获取名为layerName的层表记录的Id
            //如果要删除的图层为当前层，则返回（不能删除当前层）
            if (layerId == db.Clayer) return false;
            //打开名为layerName的层表记录
            LayerTableRecord ltr = (LayerTableRecord)layerId.GetObject(OpenMode.ForRead);
            //如果要删除的图层包含对象或依赖外部参照，则返回（不能删除这些层）
            lt.GenerateUsageData();
            if (ltr.IsUsed) return false;
            ltr.UpgradeOpen();//切换层表记录为写的状态
            ltr.Erase(true);//删除名为layerName的图层
            return true;//删除图层成功
        }

        /// <summary>
        /// 获取所有图层的ObjectId
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <returns>返回所有图层的ObjectId</returns>
        public static List<ObjectId> GetAllLayerObjectIds(this Database db)
        {
            //打开层表
            LayerTable lt = (LayerTable)db.LayerTableId.GetObject(OpenMode.ForRead);
            //用于返回层表记录ObjectId的列表
            List<ObjectId> ltrs = new List<ObjectId>();
            //遍历层表
            foreach (ObjectId id in lt)
            {
                ltrs.Add(id);//添加到返回列表中
            }
            return ltrs;//返回所有的层表记录的ObjectId
        }
    }
    public static partial class GeTools
    {
        /// <summary>
        /// 将弧度值转换为角度值
        /// </summary>
        /// <param name="angle">弧度</param>
        /// <returns>返回角度值</returns>
        public static double RadianToDegree(this double angle)
        {
            return angle * (180.0 / Math.PI);
        }

        /// <summary>
        /// 将角度值转换为弧度值
        /// </summary>
        /// <param name="angle">角度</param>
        /// <returns>返回弧度值</returns>
        public static double DegreeToRadian(this double angle)
        {
            return angle * (Math.PI / 180.0);
        }

        /// <summary>
        /// 获取与给定点指定角度和距离的点
        /// </summary>
        /// <param name="point">给定点</param>
        /// <param name="angle">角度</param>
        /// <param name="dist">距离</param>
        /// <returns>返回与给定点指定角度和距离的点</returns>
        public static Point3d PolarPoint(this Point3d point, double angle, double dist)
        {
            return new Point3d(point.X + dist * Math.Cos(angle), point.Y + dist * Math.Sin(angle), point.Z);
        }

        /// <summary>
        /// 获取两个点之间的中点
        /// </summary>
        /// <param name="pt1">第一点</param>
        /// <param name="pt2">第二点</param>
        /// <returns>返回两个点之间的中点</returns>
        public static Point3d MidPoint(Point3d pt1, Point3d pt2)
        {
            Point3d midPoint = new Point3d((pt1.X + pt2.X) / 2.0,
                                        (pt1.Y + pt2.Y) / 2.0,
                                        (pt1.Z + pt2.Z) / 2.0);
            return midPoint;
        }

        /// <summary>
        /// 计算从第一点到第二点所确定的矢量与X轴正方向的夹角
        /// </summary>
        /// <param name="pt1">第一点</param>
        /// <param name="pt2">第二点</param>
        /// <returns>返回两点所确定的矢量与X轴正方向的夹角</returns>
        public static double AngleFromXAxis(this Point3d pt1, Point3d pt2)
        {
            //构建一个从第一点到第二点所确定的矢量
            Vector2d vector = new Vector2d(pt1.X - pt2.X, pt1.Y - pt2.Y);
            //返回该矢量和X轴正半轴的角度（弧度）
            return vector.Angle;
        }
    }

    /// <summary>
    /// 辅助操作类
    /// </summary>
    public static partial class Tools
    {
        /// <summary>
        /// 判断字符串是否为数字
        /// </summary>
        /// <param name="value">字符串</param>
        /// <returns>如果字符串为数字，返回true，否则返回false</returns>
        public static bool IsNumeric(this string value)
        {
            return Regex.IsMatch(value, @"^[+-]?\d*[.]?\d*$");
        }

        /// <summary>
        /// 判断字符串是否为整数
        /// </summary>
        /// <param name="value">字符串</param>
        /// <returns>如果字符串为整数，返回true，否则返回false</returns>
        public static bool IsInt(this string value)
        {
            return Regex.IsMatch(value, @"^[+-]?\d*$");
        }

        /// <summary>
        /// 获取当前.NET程序所在的目录
        /// </summary>
        /// <returns>返回当前.NET程序所在的目录</returns>
        public static string GetCurrentPath()
        {
            var moudle = System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0];
            return System.IO.Path.GetDirectoryName(moudle.FullyQualifiedName);
        }

        /// <summary>
        /// 判断字符串是否为空或空白
        /// </summary>
        /// <param name="value">字符串</param>
        /// <returns>如果字符串为空或空白，返回true，否则返回false</returns>
        public static bool IsNullOrWhiteSpace(this string value)
        {
            if (value == null) return false;
            return string.IsNullOrEmpty(value.Trim());
        }

        /// <summary>
        /// 获取模型空间的ObjectId
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <returns>返回模型空间的ObjectId</returns>
        public static ObjectId GetModelSpaceId(this Database db)
        {
            return SymbolUtilityServices.GetBlockModelSpaceId(db);
        }

        /// <summary>
        /// 获取图纸空间的ObjectId
        /// </summary>
        /// <param name="db"></param>
        /// <returns>返回图纸空间的ObjectId</returns>
        public static ObjectId GetPaperSpaceId(this Database db)
        {
            return SymbolUtilityServices.GetBlockPaperSpaceId(db);
        }

        /// <summary>
        /// 将实体添加到模型空间
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="ent">要添加的实体</param>
        /// <returns>返回添加到模型空间中的实体ObjectId</returns>
        public static ObjectId AddToModelSpace(this Database db, Entity ent)
        {
            ObjectId entId;//用于返回添加到模型空间中的实体ObjectId
            //定义一个指向当前数据库的事务处理，以添加直线
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                //以读方式打开块表
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                //以写方式打开模型空间块表记录.
                BlockTableRecord btr = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);
                entId = btr.AppendEntity(ent);//将图形对象的信息添加到块表记录中
                trans.AddNewlyCreatedDBObject(ent, true);//把对象添加到事务处理中
                trans.Commit();//提交事务处理
            }
            return entId; //返回实体的ObjectId
        }

        /// <summary>
        /// 将实体添加到模型空间
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="ents">要添加的多个实体</param>
        /// <returns>返回添加到模型空间中的实体ObjectId集合</returns>
        public static ObjectIdCollection AddToModelSpace(this Database db, params Entity[] ents)
        {
            ObjectIdCollection ids = new ObjectIdCollection();
            var trans = db.TransactionManager;
            BlockTableRecord btr = (BlockTableRecord)trans.GetObject(SymbolUtilityServices.GetBlockModelSpaceId(db), OpenMode.ForWrite);
            foreach (var ent in ents)
            {
                ids.Add(btr.AppendEntity(ent));
                trans.AddNewlyCreatedDBObject(ent, true);
            }
            btr.DowngradeOpen();
            return ids;
        }

        /// <summary>
        /// 将实体添加到图纸空间
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="ent">要添加的实体</param>
        /// <returns>返回添加到图纸空间中的实体ObjectId</returns>
        public static ObjectId AddToPaperSpace(this Database db, Entity ent)
        {
            var trans = db.TransactionManager;
            BlockTableRecord btr = (BlockTableRecord)trans.GetObject(SymbolUtilityServices.GetBlockPaperSpaceId(db), OpenMode.ForWrite);
            ObjectId id = btr.AppendEntity(ent);
            trans.AddNewlyCreatedDBObject(ent, true);
            btr.DowngradeOpen();
            return id;
        }

        /// <summary>
        /// 将实体添加到图纸空间
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="ents">要添加的多个实体</param>
        /// <returns>返回添加到图纸空间中的实体ObjectId集合</returns>
        public static ObjectIdCollection AddToPaperSpace(this Database db, params Entity[] ents)
        {
            ObjectIdCollection ids = new ObjectIdCollection();
            var trans = db.TransactionManager;
            BlockTableRecord btr = (BlockTableRecord)trans.GetObject(SymbolUtilityServices.GetBlockPaperSpaceId(db), OpenMode.ForWrite);
            foreach (var ent in ents)
            {
                ids.Add(btr.AppendEntity(ent));
                trans.AddNewlyCreatedDBObject(ent, true);
            }
            btr.DowngradeOpen();
            return ids;
        }

        /// <summary>
        /// 将实体添加到当前空间
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="ent">要添加的实体</param>
        /// <returns>返回添加到当前空间中的实体ObjectId</returns>
        public static ObjectId AddToCurrentSpace(this Database db, Entity ent)
        {
            var trans = db.TransactionManager;
            BlockTableRecord btr = (BlockTableRecord)trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
            ObjectId id = btr.AppendEntity(ent);
            trans.AddNewlyCreatedDBObject(ent, true);
            btr.DowngradeOpen();
            return id;
        }

        /// <summary>
        /// 将实体添加到当前空间
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="ents">要添加的多个实体</param>
        /// <returns>返回添加到当前空间中的实体ObjectId集合</returns>
        public static ObjectIdCollection AddToCurrentSpace(this Database db, params Entity[] ents)
        {
            ObjectIdCollection ids = new ObjectIdCollection();
            var trans = db.TransactionManager;
            BlockTableRecord btr = (BlockTableRecord)trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
            foreach (var ent in ents)
            {
                ids.Add(btr.AppendEntity(ent));
                trans.AddNewlyCreatedDBObject(ent, true);
            }
            btr.DowngradeOpen();
            return ids;
        }

        /// <summary>
        /// 将字符串形式的句柄转化为ObjectId
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <param name="handleString">句柄字符串</param>
        /// <returns>返回实体的ObjectId</returns>
        public static ObjectId HandleToObjectId(this Database db, string handleString)
        {
            Handle handle = new Handle(Convert.ToInt64(handleString, 16));
            ObjectId id = db.GetObjectId(false, handle, 0);
            return id;
        }

        /// <summary>
        /// 亮显实体
        /// </summary>
        /// <param name="ids">要亮显的实体的Id集合</param>
        public static void HighlightEntities(this ObjectIdCollection ids)
        {
            if (ids.Count == 0) return;
            var trans = ids[0].Database.TransactionManager;
            foreach (ObjectId id in ids)
            {
                Entity ent = trans.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent != null)
                {
                    ent.Highlight();
                }
            }
        }

        /// <summary>
        /// 亮显选择集中的实体
        /// </summary>
        /// <param name="selectionSet">选择集</param>
        public static void HighlightEntities(this SelectionSet selectionSet)
        {
            if (selectionSet == null) return;
            ObjectIdCollection ids = new ObjectIdCollection(selectionSet.GetObjectIds());
            ids.HighlightEntities();
        }

        /// <summary>
        /// 取消亮显实体
        /// </summary>
        /// <param name="ids">实体的Id集合</param>
        public static void UnHighlightEntities(this ObjectIdCollection ids)
        {
            if (ids.Count == 0) return;
            var trans = ids[0].Database.TransactionManager;
            foreach (ObjectId id in ids)
            {
                Entity ent = trans.GetObject(id, OpenMode.ForRead) as Entity;
                if (ent != null)
                {
                    ent.Unhighlight();
                }
            }
        }

        /// <summary>
        /// 将字符串格式的点转换为Point3d格式
        /// </summary>
        /// <param name="stringPoint">字符串格式的点</param>
        /// <returns>返回对应的Point3d</returns>
        public static Point3d StringToPoint3d(this string stringPoint)
        {
            string[] strPoint = stringPoint.Trim().Split(new char[] { '(', ',', ')' }, StringSplitOptions.RemoveEmptyEntries);
            double x = Convert.ToDouble(strPoint[0]);
            double y = Convert.ToDouble(strPoint[1]);
            double z = Convert.ToDouble(strPoint[2]);
            return new Point3d(x, y, z);
        }

        /// <summary>
        /// 获取数据库对应的文档对象
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <returns>返回数据库对应的文档对象</returns>
        public static Document GetDocument(this Database db)
        {
            return Application.DocumentManager.GetDocument(db);
        }

        /// <summary>
        /// 根据数据库获取命令行对象
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <returns>返回命令行对象</returns>
        public static Editor GetEditor(this Database db)
        {
            return Application.DocumentManager.GetDocument(db).Editor;
        }

        /// <summary>
        /// 在命令行输出信息
        /// </summary>
        /// <param name="ed">命令行对象</param>
        /// <param name="message">要输出的信息</param>
        public static void WriteMessage(this Editor ed, object message)
        {
            ed.WriteMessage(message.ToString());
        }

        /// <summary>
        /// 在命令行输出信息，信息显示在新行上
        /// </summary>
        /// <param name="ed">命令行对象</param>
        /// <param name="message">要输出的信息</param>
        public static void WriteMessageWithReturn(this Editor ed, object message)
        {
            ed.WriteMessage("\n" + message.ToString());
        }
    }
    public static class EntTools
    {
        /// <summary>
        /// 移动实体
        /// </summary>
        /// <param name="id">实体的ObjectId</param>
        /// <param name="sourcePt">移动的源点</param>
        /// <param name="targetPt">移动的目标点</param>
        public static void Move(this ObjectId id, Point3d sourcePt, Point3d targetPt)
        {
            //构建用于移动实体的矩阵
            Vector3d vector = targetPt.GetVectorTo(sourcePt);
            Matrix3d mt = Matrix3d.Displacement(vector);
            //以写的方式打开id表示的实体对象
            Entity ent = (Entity)id.GetObject(OpenMode.ForWrite);
            ent.TransformBy(mt);//对实体实施移动
            ent.DowngradeOpen();//为防止错误，切换实体为读的状态
        }

        /// <summary>
        /// 移动实体
        /// </summary>
        /// <param name="ent">实体</param>
        /// <param name="sourcePt">移动的源点</param>
        /// <param name="targetPt">移动的目标点</param>
        public static void Move(this Entity ent, Point3d sourcePt, Point3d targetPt)
        {
            if (ent.IsNewObject) // 如果是还未被添加到数据库中的新实体
            {
                // 构建用于移动实体的矩阵
                Vector3d vector = targetPt.GetVectorTo(sourcePt);
                Matrix3d mt = Matrix3d.Displacement(vector);
                ent.TransformBy(mt);//对实体实施移动
            }
            else // 如果是已经添加到数据库中的实体
            {
                ent.ObjectId.Move(sourcePt, targetPt);
            }
        }

        /// <summary>
        /// 复制实体
        /// </summary>
        /// <param name="id">实体的ObjectId</param>
        /// <param name="sourcePt">复制的源点</param>
        /// <param name="targetPt">复制的目标点</param>
        /// <returns>返回复制实体的ObjectId</returns>
        public static ObjectId Copy(this ObjectId id, Point3d sourcePt, Point3d targetPt)
        {
            //构建用于复制实体的矩阵
            Vector3d vector = targetPt.GetVectorTo(sourcePt);
            Matrix3d mt = Matrix3d.Displacement(vector);
            //获取id表示的实体对象
            Entity ent = (Entity)id.GetObject(OpenMode.ForRead);
            //获取实体的拷贝
            Entity entCopy = ent.GetTransformedCopy(mt);
            //将复制的实体对象添加到模型空间
            ObjectId copyId = id.Database.AddToModelSpace(entCopy);
            return copyId; //返回复制实体的ObjectId
        }

        /// <summary>
        /// 复制实体
        /// </summary>
        /// <param name="ent">实体</param>
        /// <param name="sourcePt">复制的源点</param>
        /// <param name="targetPt">复制的目标点</param>
        /// <returns>返回复制实体的ObjectId</returns>
        public static ObjectId Copy(this Entity ent, Point3d sourcePt, Point3d targetPt)
        {
            ObjectId copyId;
            if (ent.IsNewObject) // 如果是还未被添加到数据库中的新实体
            {
                //构建用于复制实体的矩阵
                Vector3d vector = targetPt.GetVectorTo(sourcePt);
                Matrix3d mt = Matrix3d.Displacement(vector);
                //获取实体的拷贝
                Entity entCopy = ent.GetTransformedCopy(mt);
                //将复制的实体对象添加到模型空间
                copyId = ent.Database.AddToModelSpace(entCopy);
            }
            else
            {
                copyId = ent.ObjectId.Copy(sourcePt, targetPt);
            }
            return copyId; //返回复制实体的ObjectId
        }

        /// <summary>
        /// 旋转实体
        /// </summary>
        /// <param name="id">实体的ObjectId</param>
        /// <param name="basePt">旋转基点</param>
        /// <param name="angle">旋转角度</param>
        public static void Rotate(this ObjectId id, Point3d basePt, double angle)
        {
            Matrix3d mt = Matrix3d.Rotation(angle, Vector3d.ZAxis, basePt);
            Entity ent = (Entity)id.GetObject(OpenMode.ForWrite);
            ent.TransformBy(mt);
            ent.DowngradeOpen();
        }

        /// <summary>
        /// 旋转实体
        /// </summary>
        /// <param name="ent">实体</param>
        /// <param name="basePt">旋转基点</param>
        /// <param name="angle">旋转角度</param>
        public static void Rotate(this Entity ent, Point3d basePt, double angle)
        {
            if (ent.IsNewObject)
            {
                Matrix3d mt = Matrix3d.Rotation(angle, Vector3d.ZAxis, basePt);
                ent.TransformBy(mt);
            }
            else
            {
                ent.ObjectId.Rotate(basePt, angle);
            }
        }

        /// <summary>
        /// 缩放实体
        /// </summary>
        /// <param name="id">实体的ObjectId</param>
        /// <param name="basePt">缩放基点</param>
        /// <param name="scaleFactor">缩放比例</param>
        public static void Scale(this ObjectId id, Point3d basePt, double scaleFactor)
        {
            Matrix3d mt = Matrix3d.Scaling(scaleFactor, basePt);
            Entity ent = (Entity)id.GetObject(OpenMode.ForWrite);
            ent.TransformBy(mt);
            ent.DowngradeOpen();
        }

        /// <summary>
        /// 缩放实体
        /// </summary>
        /// <param name="ent">实体</param>
        /// <param name="basePt">缩放基点</param>
        /// <param name="scaleFactor">缩放比例</param>
        public static void Scale(this Entity ent, Point3d basePt, double scaleFactor)
        {
            if (ent.IsNewObject)
            {
                Matrix3d mt = Matrix3d.Scaling(scaleFactor, basePt);
                ent.TransformBy(mt);
            }
            else
            {
                ent.ObjectId.Scale(basePt, scaleFactor);
            }
        }

        /// <summary>
        /// 镜像实体
        /// </summary>
        /// <param name="id">实体的ObjectId</param>
        /// <param name="mirrorPt1">镜像轴的第一点</param>
        /// <param name="mirrorPt2">镜像轴的第二点</param>
        /// <param name="eraseSourceObject">是否删除源对象</param>
        /// <returns>返回镜像实体的ObjectId</returns>
        public static ObjectId Mirror(this ObjectId id, Point3d mirrorPt1, Point3d mirrorPt2, bool eraseSourceObject)
        {
            Line3d miLine = new Line3d(mirrorPt1, mirrorPt2);//镜像线
            Matrix3d mt = Matrix3d.Mirroring(miLine);//镜像矩阵
            ObjectId mirrorId = id;
            Entity ent = (Entity)id.GetObject(OpenMode.ForWrite);
            //如果删除源对象，则直接对源对象实行镜像变换
            if (eraseSourceObject == true)
                ent.TransformBy(mt);
            else //如果不删除源对象，则镜像复制源对象
            {
                Entity entCopy = ent.GetTransformedCopy(mt);
                mirrorId = id.Database.AddToModelSpace(entCopy);
            }
            return mirrorId;
        }

        /// <summary>
        /// 镜像实体
        /// </summary>
        /// <param name="ent">实体</param>
        /// <param name="mirrorPt1">镜像轴的第一点</param>
        /// <param name="mirrorPt2">镜像轴的第二点</param>
        /// <param name="eraseSourceObject">是否删除源对象</param>
        /// <returns>返回镜像实体的ObjectId</returns>
        public static ObjectId Mirror(this Entity ent, Point3d mirrorPt1, Point3d mirrorPt2, bool eraseSourceObject)
        {
            Line3d miLine = new Line3d(mirrorPt1, mirrorPt2);//镜像线
            Matrix3d mt = Matrix3d.Mirroring(miLine);//镜像矩阵
            ObjectId mirrorId = ObjectId.Null;
            if (ent.IsNewObject)
            {
                //如果删除源对象，则直接对源对象实行镜像变换
                if (eraseSourceObject == true)
                    ent.TransformBy(mt);
                else //如果不删除源对象，则镜像复制源对象
                {
                    Entity entCopy = ent.GetTransformedCopy(mt);
                    mirrorId = ent.Database.AddToModelSpace(entCopy);
                }
            }
            else
            {
                mirrorId = ent.ObjectId.Mirror(mirrorPt1, mirrorPt2, eraseSourceObject);
            }
            return mirrorId;
        }

        /// <summary>
        /// 偏移实体
        /// </summary>
        /// <param name="id">实体的ObjectId</param>
        /// <param name="dis">偏移距离</param>
        /// <returns>返回偏移后的实体Id集合</returns>
        public static ObjectIdCollection Offset(this ObjectId id, double dis)
        {
            ObjectIdCollection ids = new ObjectIdCollection();
            Curve cur = id.GetObject(OpenMode.ForWrite) as Curve;
            if (cur != null)
            {
                try
                {
                    //获取偏移的对象集合
                    DBObjectCollection offsetCurves = cur.GetOffsetCurves(dis);
                    //将对象集合类型转换为实体类的数组，以方便加入实体的操作
                    Entity[] offsetEnts = new Entity[offsetCurves.Count];
                    offsetCurves.CopyTo(offsetEnts, 0);
                    //将偏移的对象加入到数据库
                    ids = id.Database.AddToModelSpace(offsetEnts);
                }
                catch
                {
                    Application.ShowAlertDialog("无法偏移！");
                }
            }
            else
                Application.ShowAlertDialog("无法偏移！");
            return ids;//返回偏移后的实体Id集合
        }

        /// <summary>
        /// 偏移实体
        /// </summary>
        /// <param name="ent">实体</param>
        /// <param name="dis">偏移距离</param>
        /// <returns>返回偏移后的实体集合</returns>
        public static DBObjectCollection Offset(this Entity ent, double dis)
        {
            DBObjectCollection offsetCurves = new DBObjectCollection();
            Curve cur = ent as Curve;
            if (cur != null)
            {
                try
                {
                    offsetCurves = cur.GetOffsetCurves(dis);
                    Entity[] offsetEnts = new Entity[offsetCurves.Count];
                    offsetCurves.CopyTo(offsetEnts, 0);
                }
                catch
                {
                    Application.ShowAlertDialog("无法偏移！");
                }
            }
            else
                Application.ShowAlertDialog("无法偏移！");
            return offsetCurves;
        }

        /// <summary>
        /// 矩形阵列实体
        /// </summary>
        /// <param name="id">实体的ObjectId</param>
        /// <param name="numRows">矩形阵列的行数,该值必须为正数</param>
        /// <param name="numCols">矩形阵列的列数,该值必须为正数</param>
        /// <param name="disRows">行间的距离</param>
        /// <param name="disCols">列间的距离</param>
        /// <returns>返回阵列后的实体集合的ObjectId</returns>
        public static ObjectIdCollection ArrayRectang(this ObjectId id, int numRows, int numCols, double disRows, double disCols)
        {
            // 用于返回阵列后的实体集合的ObjectId
            ObjectIdCollection ids = new ObjectIdCollection();
            // 以读的方式打开实体
            Entity ent = (Entity)id.GetObject(OpenMode.ForRead);
            for (int m = 0; m < numRows; m++)
            {
                for (int n = 0; n < numCols; n++)
                {
                    // 获取平移矩阵
                    Matrix3d mt = Matrix3d.Displacement(new Vector3d(n * disCols, m * disRows, 0));
                    Entity entCopy = ent.GetTransformedCopy(mt);// 复制实体
                    // 将复制的实体添加到模型空间
                    ObjectId entCopyId = id.Database.AddToModelSpace(entCopy);
                    ids.Add(entCopyId);// 将复制实体的ObjectId添加到集合中
                }
            }
            ent.UpgradeOpen();// 切换实体为写的状态
            ent.Erase();// 删除实体
            return ids;// 返回阵列后的实体集合的ObjectId
        }

        /// <summary>
        /// 环形阵列实体
        /// </summary>
        /// <param name="id">实体的ObjectId</param>
        /// <param name="cenPt">环形阵列的中心点</param>
        /// <param name="numObj">在环形阵列中所要创建的对象数量</param>
        /// <param name="angle">以弧度表示的填充角度，正值表示逆时针方向旋转，负值表示顺时针方向旋转，如果角度为0则出错</param>
        /// <returns>返回阵列后的实体集合的ObjectId</returns>
        public static ObjectIdCollection ArrayPolar(this ObjectId id, Point3d cenPt, int numObj, double angle)
        {
            ObjectIdCollection ids = new ObjectIdCollection();
            Entity ent = (Entity)id.GetObject(OpenMode.ForRead);
            for (int i = 0; i < numObj - 1; i++)
            {
                Matrix3d mt = Matrix3d.Rotation(angle * (i + 1) / numObj, Vector3d.ZAxis, cenPt);
                Entity entCopy = ent.GetTransformedCopy(mt);
                ObjectId entCopyId = id.Database.AddToModelSpace(entCopy);
                ids.Add(entCopyId);
            }
            return ids;
        }

        /// <summary>
        /// 删除实体
        /// </summary>
        /// <param name="id">实体的ObjectId</param>
        public static void Erase(this ObjectId id)
        {
            DBObject ent = id.GetObject(OpenMode.ForWrite);
            ent.Erase();
        }

        /// <summary>
        /// 计算图形数据库模型空间中所有实体的包围框
        /// </summary>
        /// <param name="db">数据库对象</param>
        /// <returns>返回模型空间中所有实体的包围框</returns>
        public static Extents3d GetAllEntsExtent(this Database db)
        {
            Extents3d totalExt = new Extents3d();
            using (Transaction trans = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btRcd = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                foreach (ObjectId entId in btRcd)
                {
                    Entity ent = trans.GetObject(entId, OpenMode.ForRead) as Entity;
                    totalExt.AddExtents(ent.GeometricExtents);
                }
            }
            return totalExt;
        }
    }

    public static class PolylineTools
    {
        /// <summary>
        /// 通过三维点集合创建多段线
        /// </summary>
        /// <param name="pline">多段线对象</param>
        /// <param name="pts">多段线的顶点</param>
        public static void CreatePolyline(this Polyline pline, Point3dCollection pts)
        {
            for (int i = 0; i < pts.Count; i++)
            {
                //添加多段线的顶点
                pline.AddVertexAt(i, new Point2d(pts[i].X, pts[i].Y), 0, 0, 0);
            }
        }

        /// <summary>
        /// 通过二维点集合创建多段线
        /// </summary>
        /// <param name="pline">多段线对象</param>
        /// <param name="pts">多段线的顶点</param>
        public static void CreatePolyline(this Polyline pline, Point2dCollection pts)
        {
            for (int i = 0; i < pts.Count; i++)
            {
                //添加多段线的顶点
                pline.AddVertexAt(i, pts[i], 0, 0, 0);
            }
        }

        /// <summary>
        /// 通过不固定的点创建多段线
        /// </summary>
        /// <param name="pline">多段线对象</param>
        /// <param name="pts">多段线的顶点</param>
        public static void CreatePolyline(this Polyline pline, params Point2d[] pts)
        {
            pline.CreatePolyline(new Point2dCollection(pts));
        }

        /// <summary>
        /// 创建矩形
        /// </summary>
        /// <param name="pline">多段线对象</param>
        /// <param name="pt1">矩形的角点</param>
        /// <param name="pt2">矩形的角点</param>
        public static void CreateRectangle(this Polyline pline, Point2d pt1, Point2d pt2)
        {
            //设置矩形的4个顶点
            double minX = Math.Min(pt1.X, pt2.X);
            double maxX = Math.Max(pt1.X, pt2.X);
            double minY = Math.Min(pt1.Y, pt2.Y);
            double maxY = Math.Max(pt1.Y, pt2.Y);
            Point2dCollection pts = new Point2dCollection();
            pts.Add(new Point2d(minX, minY));
            pts.Add(new Point2d(minX, maxY));
            pts.Add(new Point2d(maxX, maxY));
            pts.Add(new Point2d(maxX, minY));
            pline.CreatePolyline(pts);
            pline.Closed = true;//闭合多段线以形成矩形
        }

        /// <summary>
        /// 创建多边形
        /// </summary>
        /// <param name="pline">多段线对象</param>
        /// <param name="centerPoint">多边形中心点</param>
        /// <param name="number">边数</param>
        /// <param name="radius">外接圆半径</param>
        public static void CreatePolygon(this Polyline pline, Point2d centerPoint, int number, double radius)
        {
            Point2dCollection pts = new Point2dCollection(number);
            double angle = 2 * Math.PI / number;//计算每条边对应的角度
            //计算多边形的顶点
            for (int i = 0; i < number; i++)
            {
                Point2d pt = new Point2d(centerPoint.X + radius * Math.Cos(i * angle), centerPoint.Y + radius * Math.Sin(i * angle));
                pts.Add(pt);
            }
            pline.CreatePolyline(pts);
            pline.Closed = true;//闭合多段线以形成多边形
        }

        /// <summary>
        /// 创建多段线形式的圆
        /// </summary>
        /// <param name="pline">多段线对象</param>
        /// <param name="centerPoint">圆心</param>
        /// <param name="radius">半径</param>
        public static void CreatePolyCircle(this Polyline pline, Point2d centerPoint, double radius)
        {
            //计算多段线的顶点
            Point2d pt1 = new Point2d(centerPoint.X + radius, centerPoint.Y);
            Point2d pt2 = new Point2d(centerPoint.X - radius, centerPoint.Y);
            Point2d pt3 = new Point2d(centerPoint.X + radius, centerPoint.Y);
            Point2dCollection pts = new Point2dCollection();
            //添加多段线的顶点
            pline.AddVertexAt(0, pt1, 1, 0, 0);
            pline.AddVertexAt(1, pt2, 1, 0, 0);
            pline.AddVertexAt(2, pt3, 1, 0, 0);
            pline.Closed = true;//闭合曲线以形成圆
        }

        /// <summary>
        /// 创建多段线形式的圆弧
        /// </summary>
        /// <param name="pline">多段线对象</param>
        /// <param name="centerPoint">圆弧的圆心</param>
        /// <param name="radius">圆弧的半径</param>
        /// <param name="startAngle">起始角度</param>
        /// <param name="endAngle">终止角度</param>
        public static void CreatePolyArc(this Polyline pline, Point2d centerPoint, double radius, double startAngle, double endAngle)
        {
            //计算多段线的顶点
            Point2d pt1 = new Point2d(centerPoint.X + radius * Math.Cos(startAngle),
                                    centerPoint.Y + radius * Math.Sin(startAngle));
            Point2d pt2 = new Point2d(centerPoint.X + radius * Math.Cos(endAngle),
                                    centerPoint.Y + radius * Math.Sin(endAngle));
            //添加多段线的顶点
            pline.AddVertexAt(0, pt1, Math.Tan((endAngle - startAngle) / 4), 0, 0);
            pline.AddVertexAt(1, pt2, 0, 0, 0);
        }
    }
    public class myAcCmd
    {
        public string Macro { get; set; }
        public string Label { get; set; }
        public myAcCmd(string _cmd, string cmdName)
        {
            this.Macro = _cmd;
            this.Label = cmdName;
        }
    }
}
