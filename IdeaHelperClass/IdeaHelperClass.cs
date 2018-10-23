using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace IdeaHelperClass
{
    public class IdeaHelperClass
    {
      public static void DisposeCom(object obj)
      {
         if (obj == null)
         {
            return;
         }
         try
         {
            Marshal.ReleaseComObject(obj);
         }
         catch (Exception)
         {
            obj = null;
         }
      }

   }
}
