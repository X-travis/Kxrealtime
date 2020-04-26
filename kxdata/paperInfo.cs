using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace kxrealtime.kxdata
{
    class paperInfo
    {
    }

    class paperItemInfo
    {
        public string t;
        public paperItemContent c;
    }

    class paperItemContent
    {
    
    }

    class paperItemContentAnalysis
    {

    }

    public class simplePaper
    {
        public string type;
        public List<simplePaperItem> data;
    }

    public class simplePaperItem
    {
        public string type;
        public string title;
        public float score;
        public List<simpleAnswerItem> answers;
        public List<string> options;
    }

    public class simpleAnswerItem
    {
        public string text;
        public float score;
    }

    public class simpleFillAnswer
    {
        public string answer;
        public float score;
    }

}
