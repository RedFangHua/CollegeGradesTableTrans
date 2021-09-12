using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;


namespace OriginGradeTrans
{
    class Translator
    {
        static string mainObjFlags = "公共必修;学科基础;专业必修;专业核心";

        string originFileName;  //源文件
        string targetFileName;  //新文件

        IWorkbook originBook;
        IWorkbook targetBook;
        ISheet originSheet;
        ISheet targetSheet;

        MainObjectRect objRect;
        List<GradeMeta> metaList;

        public delegate void DebugInfoOut(string InfoString);
        DebugInfoOut debugOut = new DebugInfoOut(EmptyOut);

        static void EmptyOut(string InfoString) { }

        public void SetDebugOut(DebugInfoOut Func)
        {
            debugOut = Func;
        }

        //转换开始方法
        public bool GradesTableTranslate(string originGradeTableFile)
        {
            originFileName = originGradeTableFile;
            targetFileName = originGradeTableFile.Replace(".xlsx", "_trans.xlsx");

            originBook = new XSSFWorkbook(originFileName);
            targetBook = new XSSFWorkbook();
            targetSheet = targetBook.CreateSheet("Sheet1");

            InitObjRect();
            GenerateMetaList();
            GenerateTargetFormatSheet();
            WriteTransedSheet();

            originBook.Close();
            return false;
        }

        //初始化主科记录矩阵
        void InitObjRect()
        {
            metaList = new List<GradeMeta>();
            objRect = new MainObjectRect();
            originSheet = originBook.GetSheetAt(0); //开第一张表

            int StartRow = 2;   //从第三行开始有有效数据

            debugOut("成绩单采样中");

            IRow Row_One = originSheet.GetRow(2);
            String ObjName_Current;
            String ObjType_Current;
            
            while (Row_One != null)
            {
                ObjName_Current = Row_One.GetCell(6).ToString();
                ObjType_Current = Row_One.GetCell(7).ToString();

                if(mainObjFlags.IndexOf(ObjType_Current) != -1)
                {
                    objRect.AppendMainObj(ObjName_Current);

                    //debug
                    debugOut(string.Format("主科: {0} 所在列: {1}", ObjName_Current, objRect.GetObjectColumn(ObjName_Current)));
                }

                StartRow++;
                Row_One = originSheet.GetRow(StartRow);
            }
        }

        //生成每个人的成绩单的表
        void GenerateMetaList()
        {
            int StartRow = 2;
            int ID = 1;

            IRow Row_Any = originSheet.GetRow(2);
            String Name_Current = Row_Any.GetCell(2).ToString();
            String ObjName_Current;
            String ObjGrade_Current;
            GradeMeta Meta_Current = new GradeMeta(ID, Row_Any.GetCell(1).ToString(), Name_Current);

            while (Row_Any != null)
            {
                ObjName_Current = Row_Any.GetCell(6).ToString();
                ObjGrade_Current = Row_Any.GetCell(17).ToString();
                String NameTemp = Row_Any.GetCell(2).ToString();

                if (NameTemp != Name_Current)   //名字与上一个不同，说明到了下一个人
                {
                    //debug
                    debugOut(Name_Current + ":");
                    string gradeLine = "";
                    foreach (ObjectInfo objectInfo in objRect.Objects)
                    {
                        GradeItem gradeItem = Meta_Current.GetGrade(objectInfo.ObjectName);
                        if (gradeItem != null)
                        {
                            gradeLine += gradeItem.Value.ToString();
                        }
                        else
                        {
                            gradeLine += "  ";
                        }

                        gradeLine += " ";
                    }
                    debugOut(gradeLine);

                    metaList.Add(Meta_Current);
                    Name_Current = NameTemp;
                    Meta_Current = new GradeMeta(ID + 1, Row_Any.GetCell(1).ToString(), NameTemp);
                    ID++;
                }

                Meta_Current.SetMaxGrade(ObjName_Current, int.Parse(ObjGrade_Current));

                StartRow++;
                Row_Any = originSheet.GetRow(StartRow);
            }
            metaList.Add(Meta_Current);

            debugOut(string.Format("已处理完{0}个人", metaList.Count));
        }

        //创建新表中的特定格式
        void GenerateTargetFormatSheet()
        {
            for(int i = 0; i < metaList.Count + 2; i++)
            {
                targetSheet.CreateRow(i);
            }
            IRow Line2 = targetSheet.GetRow(1);
            ICell[] LineCell = new ICell[3 + objRect.Objects.Count];
            for(int i = 0; i < 3 + objRect.Objects.Count; i++)
            {
                LineCell[i] = Line2.CreateCell(i);
            }
            LineCell[0].SetCellValue("序号");
            LineCell[1].SetCellValue("学号");
            LineCell[2].SetCellValue("姓名");
            for(int i = 3; i < 3 + objRect.Objects.Count; i++)
            {
                LineCell[i].SetCellValue(objRect.Objects[i - 3].ObjectName);
            }
        }

        //将成绩单表写入新表
        void WriteTransedSheet()
        {
            int StartLine = 2;
            for(int i = 0; i < metaList.Count; i++) //每个人
            {
                IRow CLine = targetSheet.GetRow(i + StartLine);
                ICell[] CCell = new ICell[3 + objRect.Objects.Count];

                bool ExpectedKClass = false;        //K班例外情况
                if (metaList[i].GetGrade("大学英语（三）") != null)
                    ExpectedKClass = true;

                for(int j = 3; j < objRect.Objects.Count + 3; j++)  //每个成绩
                {
                    CCell[j] = CLine.CreateCell(j);
                    GradeItem gradeItem = metaList[i].GetGrade(objRect.Objects[j - 3].ObjectName);
                    if (gradeItem != null)
                        CCell[j].SetCellValue(gradeItem.Value);
                    else
                        CCell[j].SetCellValue(0);
                }

                if (ExpectedKClass)         //K班特例
                {
                    GradeItem it = metaList[i].GetGrade("大学英语（二）");
                    if (it != null)
                        CCell[3 + objRect.GetObjectColumn("大学英语（一）")].SetCellValue(it.Value);
                    else
                        CCell[3 + objRect.GetObjectColumn("大学英语（一）")].SetCellValue(0);

                    it = metaList[i].GetGrade("大学英语（三）");
                    if (it != null)
                        CCell[3+objRect.GetObjectColumn("大学英语（二）")].SetCellValue(it.Value);
                    else
                        CCell[3 + objRect.GetObjectColumn("大学英语（二）")].SetCellValue(0);
                }

                for(int j = 0; j < 3; j++)  //序号、学号、姓名
                {
                    CCell[j] = CLine.CreateCell(j);
                }
                CCell[0].SetCellValue(metaList[i].ID);
                CCell[1].SetCellValue(metaList[i].StdID);
                CCell[2].SetCellValue(metaList[i].Name);
            }

            FileStream targetFile = new FileStream(targetFileName, FileMode.Create);
            targetBook.Write(targetFile);
            targetFile.Close();
            targetBook.Close();

            debugOut("转换已完成，文件为：");
            debugOut(targetFileName);
        }

        public Translator()
        {
            objRect = new MainObjectRect();
        }
    }

    //单个科目在表中的名称和所处列位置
    class ObjectInfo        
    {
        public string ObjectName {  get; set; }
        public int ObjectColumn {  get; set; }
    }

    //科目列表矩阵，记录每个科目在表中的列位置
    class MainObjectRect    
    {
        public List<ObjectInfo> Objects = new List<ObjectInfo>();
        int MainObjCount = 0;   //仅记录从0累加的列数，偏移应按实际应用加上

        //检查科目名是否是主科科目
        public bool CheckIsMainObj(string ObjectName)
        {
            foreach(ObjectInfo obj_T in Objects)
            {
                if (obj_T.ObjectName == ObjectName)
                    return true;
            }
            return false;
        }

        //向主科记录表中添加新的主科
        public void AppendMainObj(string ObjectName)
        {
            if (ObjectName == "大学英语（三）")
                return;
            if (!CheckIsMainObj(ObjectName))    //查重
            {
                var Obj_T = new ObjectInfo()
                { ObjectName = ObjectName, ObjectColumn = MainObjCount };
                Objects.Add(Obj_T);
                MainObjCount++;
            }
        }

        //获取科目列数
        public int GetObjectColumn(string ObjectName)
        {
            foreach(ObjectInfo obj_T in Objects)
            {
                if (obj_T.ObjectName == ObjectName)
                    return obj_T.ObjectColumn;
            }
            return -1;
        }
    }
}
