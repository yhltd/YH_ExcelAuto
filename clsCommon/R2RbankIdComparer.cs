using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using clsdatabaseinfo;

namespace clsCommon
{

    public class R2RbankIdComparer : IEqualityComparer<clsR2Rbankchargeinfo>
    {
        public bool Equals(clsR2Rbankchargeinfo x, clsR2Rbankchargeinfo y)
        {
            if (x == null)
                return y == null;
            return x.yinhang == y.yinhang;
        }

        public int GetHashCode(clsR2Rbankchargeinfo obj)
        {
            if (obj == null)
                return 0;
            return obj.yinhang.GetHashCode();
        }
    }
}
