using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Google.GData.Extensions;

namespace contactsdump
{
    class PhoneComparer: IEqualityComparer<PhoneNumber>
    {
        #region IEqualityComparer<PhoneNumber> Members

        public bool Equals(PhoneNumber x, PhoneNumber y)
        {
            return x.Value == y.Value;
        }

        public int GetHashCode(PhoneNumber obj)
        {
            return obj.Value.GetHashCode();
        }

        #endregion
    }
}
