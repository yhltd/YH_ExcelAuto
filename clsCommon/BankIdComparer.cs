using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using clsdatabaseinfo;

namespace clsCommon
{

    public class BankIdComparer : IEqualityComparer<clsWangyininfo>
    {
        public bool Equals(clsWangyininfo x, clsWangyininfo y)
        {
            if (x == null)
                return y == null;
            return x.Maichangdaima == y.Maichangdaima;
        }

        public int GetHashCode(clsWangyininfo obj)
        {
            if (obj == null)
                return 0;
            return obj.Maichangdaima.GetHashCode();
        }
    }
}
