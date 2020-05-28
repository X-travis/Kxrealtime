using System.Collections;

namespace kxrealtime
{


    public static class AnswerStore
    {

        public static Hashtable answerArrHansh = new Hashtable();

        // 记录答案
        public static void setAnswer(string key, object value)
        {
            if (answerArrHansh == null)
            {
                answerArrHansh = new Hashtable();
            }
            if (answerArrHansh.Contains(key))
            {
                answerArrHansh[key] = value;
            }
            else
            {
                answerArrHansh.Add(key, value);
            }
        }

        // 读取答案
        public static object getAnswer(string key)
        {
            if (answerArrHansh.Contains(key))
            {
                return answerArrHansh[key];
            }
            else
            {
                return null;
            }
        }

    }

    // 弃用
    public class answerFormat
    {
        public string key { get; set; }
        public string[] answer { get; set; }
    }
}
