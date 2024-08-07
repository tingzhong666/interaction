using Bentley.DgnPlatformNET;
using Bentley.DgnPlatformNET.Elements;
using Bentley.GeometryNET;
using Bentley.MstnPlatformNET;
using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace interaction
{
    class Commands
    {
        public static void OutputSuccess(string unparsed)
        {
            MessageBox.Show("Success");
        }
        public static void OutInform(string unparsed)
        {
            //MessageBox.Show("OutInform");
            OutInformTool.Install();
        }
    }

    class OutInformTool : DgnElementSetTool
    {
        static DgnModel dgnModel { get; set; }
        public OutInformTool()
        {
        }

        public static void Install()
        {
            var obj = new OutInformTool();
            obj.InstallTool();
        }

        protected override void OnRestartTool()
        {
            Install();
        }


        public override StatusInt OnElementModify(Element element)
        {
            return StatusInt.Error;
        }

        protected override void OnPostInstall()
        {
            //MessageBox.Show("OnPostInstall");
            NotificationManager.OutputPrompt("请选择元素");

            dgnModel = dgnModel ?? Session.Instance.GetActiveDgnModel();
        }

        protected override StatusInt ProcessAgenda(DgnButtonEvent ev)
        {
            //MessageBox.Show("ProcessAgenda");
            var outPath = "D:/work/temp/elementInform.xls";

            var outData = new List<ElementInf>();

            // 遍历获取信息
            for (uint i = 0; i < ElementAgenda.GetCount(); i++)
            {
                var e = (DisplayableElement)ElementAgenda.GetEntry(i);
                e.GetTransformOrigin(out DPoint3d point);
                outData.Add(new ElementInf(e.ElementId, point, e.LevelId, e.ElementType, e.TypeName, e.Description));
            }
            // 输出表格
            //var str = "";
            //foreach (var v in outData)
            //{
            //    str += "元素id: " + v.ElementId_ + " 坐标点: " + v.Point_.X + ", " + v.Point_.Y + ", " + v.Point_.Z + " 图层: " + v.LevelId_ + " 元素类型: " + v.MSElementType_ + " " + v.TypeName + " 元素名: " + v.Description + "\n";
            //}
            //MessageBox.Show(str);

            var workbook = new HSSFWorkbook(); // 一个表格对象
            var sheet = workbook.CreateSheet("表"); // 工作表

            // 行
            var title = sheet.CreateRow(0);
            title.CreateCell(0).SetCellValue("序号"); // 单元格写入
            title.CreateCell(1).SetCellValue("元素ID");
            title.CreateCell(2).SetCellValue("坐标点值");
            title.CreateCell(3).SetCellValue("图层名称");
            title.CreateCell(4).SetCellValue("元素类型");

            var flc = dgnModel.GetFileLevelCache();
            Type LevelHandleType;
            for (int i = 0; i < outData.Count; i++)
            {
                var row = sheet.CreateRow(i + 1);
                row.CreateCell(0).SetCellValue(i + 1);
                row.CreateCell(1).SetCellValue(outData.ElementAt(i).ElementId_.ToString());
                row.CreateCell(2).SetCellValue(outData.ElementAt(i).Point_.X + ", " + outData.ElementAt(i).Point_.Y + ", " + outData.ElementAt(i).Point_.Z);
                var level = flc.GetLevelByCode((int)outData.ElementAt(i).LevelId_);
                row.CreateCell(3).SetCellValue(level.DisplayName == "" ? "缺省" : level.DisplayName);
                row.CreateCell(4).SetCellValue(outData.ElementAt(i).Description);
            }

            //保持文件
            if (!Directory.Exists(Path.GetDirectoryName(outPath)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(outPath));
            }
            try
            {
                var fileData = new FileStream(outPath, FileMode.Create);
                workbook.Write(fileData);
                fileData.Close();
            }
            catch (Exception e)
            {
                if (e.Message.IndexOf("正由另一进程使用") != -1)
                {
                    MessageBox.Show(outPath + " 此文件已被其他程序打开，请关闭后重新操作");
                    return StatusInt.Error;
                }
                throw e;
            }


            MessageBox.Show("输出到 " + outPath);

            Process.Start(outPath);

            return StatusInt.Success;
        }
    }

    class ElementInf
    {
        public ElementId ElementId_ { get; set; } // 元素ID
        public DPoint3d Point_ { get; set; } // 坐标
        public LevelId LevelId_ { get; set; }    // 图层
        public MSElementType MSElementType_ { get; set; }
        public string TypeName { get; set; }
        public string Description { get; set; }

        public ElementInf(ElementId elementId_, DPoint3d point_, LevelId levelId_, MSElementType mSElementType_, string typeName, string description)
        {
            ElementId_ = elementId_;
            Point_ = point_;
            LevelId_ = levelId_;
            MSElementType_ = mSElementType_;
            TypeName = typeName;
            Description = description;
        }
    }
}
