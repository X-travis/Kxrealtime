using System;

namespace kxrealtime.kxdata
{
    // 幻灯片考试信息数据结构
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
