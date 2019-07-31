using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using clsdatabaseinfo;

namespace clsCommon
{

    public class maichangmingchengComparer : IEqualityComparer<clsribaodatasoureinfo>
    {
        public bool Equals(clsribaodatasoureinfo x, clsribaodatasoureinfo y)
        {
            if (x == null)
                return y == null;
            return x.mingcheng == y.mingcheng;
        }

        public int GetHashCode(clsribaodatasoureinfo obj)
        {
            if (obj == null || obj.mingcheng == null)
                return 0;
            return obj.mingcheng.GetHashCode();
        }
    }
}
