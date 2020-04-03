using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace kxrealtime.kxdata
{
    public class slideExamInfo
    {
        public string slideName { get; set; }
        public string paperId { get; set; }
        public string testId { get; set; }
        public Int64 startTimeStamp { get; set; }

        public Int64 duringTime { get; set; }

        public bool noTime { get; set; }

        public string paperTitle { get; set; }

    }
}
