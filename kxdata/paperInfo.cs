using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace kxrealtime.kxdata
{
    // 弃用
    class paperInfo
    {
    }
    // 弃用
    class paperItemInfo
    {
        public string t;
        public paperItemContent c;
    }
    // 弃用
    class paperItemContent
    {
    
    }
    // 弃用
    class paperItemContentAnalysis
    {

    }
    // 试卷list答案结构
    public class simplePaper
    {
        public string type;
        public List<simplePaperItem> data;
    }
    // 试卷答案结构
    public class simplePaperItem
    {
        public string type;
        public string title;
        public float score;
        public List<simpleAnswerItem> answers;
        public List<string> options;
    }
    // 填空答案结构
    public class simpleAnswerItem
    {
        public string text;
        public float score;
    }
    // 填空答案结构
    public class simpleFillAnswer
    {
        public string answer;
        public float score;
    }

}
