using Interop.IdeaLib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace IdeaHelperClass
{
    public class IdeaHelperClass
    {
      public static string GetProjectDirectory()
      {

         IdeaClient client = null;
         string dir;
         try
         {
            client = new IdeaClient();
            try
            {
               string temp = client.WorkingDirectory;
               temp = temp.Remove(temp.Length - 1);
               string temp2 = Path.GetFileName(temp);

               dir = client.ManagedProject;

               if (dir == temp2)
               {
                  return client.WorkingDirectory; ;
               }
               else
               {
                  return Directory.GetParent(temp).FullName + "\\";
               }


            }
            catch (Exception)
            {
               try
               {
                  dir = client.ExternalProject;
                  return dir;
               }
               catch (Exception ex)
               {
                  MessageBox.Show("could not get current projet" + Environment.NewLine + ex.Message);
                  Environment.Exit(1);
                  throw ex;
               }
            }
         }
         catch (Exception ex)
         {
            MessageBox.Show(ex.Message);
            Environment.Exit(1);
            throw ex;
         }
         finally
         {
            DisposeCom(client);
         }
      }

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
