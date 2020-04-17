using System.Collections;

namespace kxrealtime
{


    public static class AnswerStore
    {

        public static Hashtable answerArrHansh = new Hashtable();

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

    public class answerFormat
    {
        public string key { get; set; }
        public string[] answer { get; set; }
    }
}
