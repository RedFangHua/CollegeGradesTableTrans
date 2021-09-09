using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OriginGradeTrans
{
    class GradeMeta
    {
        public int ID { get; set; }

        public string StdID {  get; set; }
        public string Name {  get; set; }

        List<GradeItem> Grades = new List<GradeItem>();

        //添加、设置最大成绩
        public void SetMaxGrade(String ObjName, int ObjGrade)
        {
            for(int i = 0; i < Grades.Count; i++)
            {
                if(Grades[i].Name == ObjName)
                {
                    if (Grades[i].Value < ObjGrade)
                    {
                        Grades[i].Value = ObjGrade;
                        return;
                    }
                }
            }

            //若该科成绩没有记录则直接添加
            GradeItem NewGrade = new GradeItem() { Name = ObjName, Value = ObjGrade};
            Grades.Add(NewGrade);
        }

        //按科目名获取单科成绩对象
        public GradeItem GetGrade(string ObjName)
        {
            foreach(GradeItem Grade in Grades)
            {
                if (Grade.Name == ObjName)
                    return Grade;
            }
            return null;
        }


        public GradeMeta(int sID, string sStdID, string sName)
        {
            ID = sID;
            StdID = sStdID; 
            Name = sName;
        }
    }
}
