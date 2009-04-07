using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Google.GData.Extensions;

namespace contactsdump
{
    class AddressComparer: IEqualityComparer<PostalAddress>
    {
        #region IEqualityComparer<PostalAddress> Members

        public bool Equals(PostalAddress x, PostalAddress y)
        {
            return x.Value == y.Value;
        }

        public int GetHashCode(PostalAddress obj)
        {
            return obj.Value.GetHashCode();
        }

        #endregion
    }
}
