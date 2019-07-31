using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using clsdatabaseinfo;
 
namespace clsCommon
{

    public class WangyinYinhangIdComparer : IEqualityComparer<clsWangyininfo>
    {
        public bool Equals(clsWangyininfo x, clsWangyininfo y)
        {
            if (x == null)
                return y == null;
            return x.yinhang == y.yinhang;
        }

        public int GetHashCode(clsWangyininfo obj)
        {
            if (obj == null)
                return 0;
            return obj.yinhang.GetHashCode();
        }
    }
}
