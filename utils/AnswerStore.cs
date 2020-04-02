using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace kxrealtime
{

    
    public static class AnswerStore
    {

        public static Hashtable answerArrHansh = new Hashtable();

        public static void setAnswer(string key, object value)
        {
            if(answerArrHansh == null)
            {
                answerArrHansh = new Hashtable();
            }
            if(answerArrHansh.Contains(key))
            {
                answerArrHansh[key] = value;
            } else
            {
                answerArrHansh.Add(key, value);
            }
        }

        public static object getAnswer(string key)
        {
            if(answerArrHansh.Contains(key))
            {
                return answerArrHansh[key];
            } else
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
