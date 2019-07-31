using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using clsdatabaseinfo;

namespace clsCommon
{

    public class ProductmingchengComparer : IEqualityComparer<clsribaodatasoureinfo>
    {
        public bool Equals(clsribaodatasoureinfo x, clsribaodatasoureinfo y)
        {
            if (x == null)
                return y == null;
            return x.diqudaima == y.diqudaima;
        }

        public int GetHashCode(clsribaodatasoureinfo obj)
        {
            if (obj == null || obj.diqudaima == null)
                return 0;
            return obj.diqudaima.GetHashCode();
        }
    }
}
