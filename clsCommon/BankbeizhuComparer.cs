using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using clsdatabaseinfo;

namespace clsCommon
{

    public class BankbeizhuComparer : IEqualityComparer<clsWangyininfo>
    {
        public bool Equals(clsWangyininfo x, clsWangyininfo y)
        {
            if (x == null)
                return y == null;
            return x.beizhu == y.beizhu;
        }

        public int GetHashCode(clsWangyininfo obj)
        {
            if (obj == null)
                return 0;
            return obj.beizhu.GetHashCode();
        }
    }
}
