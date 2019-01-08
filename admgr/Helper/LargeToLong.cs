using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ADmgr.Helper
{
    public static class LargeToLong
    {
        #region 变large整型为long整形
        public static long ConvertLargeIntegerToLong(object largeInteger)
        {
            Type type = largeInteger.GetType();

            int highPart = (int)type.InvokeMember("HighPart", BindingFlags.GetProperty, null, largeInteger, null);
            int lowPart = (int)type.InvokeMember("LowPart", BindingFlags.GetProperty | BindingFlags.Public, null, largeInteger, null);

            return (((long)highPart) << 32) + (uint)lowPart;
        }
        #endregion
    }
}
