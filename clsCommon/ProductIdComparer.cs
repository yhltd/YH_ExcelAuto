using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using clsdatabaseinfo;

namespace clsCommon
{

    public class ProductIdComparer : IEqualityComparer<clsribaodatasoureinfo>
    {
        public bool Equals(clsribaodatasoureinfo x, clsribaodatasoureinfo y)
        {
            if (x == null)
                return y == null;
            return x.Maichangdaima == y.Maichangdaima;
        }

        public int GetHashCode(clsribaodatasoureinfo obj)
        {
            if (obj == null)
                return 0;
            return obj.Maichangdaima.GetHashCode();
        }
    }
}
