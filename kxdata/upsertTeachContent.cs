using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace kxrealtime.kxdata
{
    public class upsertTeachContentItem
    {
        public Int64 tid;
        public string snapshot;
    }

    public class upsertTeachContent
    {
        public List<upsertTeachContentItem> teach_content_list;
    }
}
