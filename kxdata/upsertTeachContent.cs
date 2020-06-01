using System;
using System.Collections.Generic;

namespace kxrealtime.kxdata
{
    // 记录授课数据结构
    public class upsertTeachContentItem
    {
        public Int64 tid;
        public string snapshot;
    }
    // teach_content_list 数据结构
    public class upsertTeachContent
    {
        public List<upsertTeachContentItem> teach_content_list;
    }
}
