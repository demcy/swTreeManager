using System.Runtime.InteropServices;

namespace SW
{
    public class SwTools
    {
        public bool swConnect()
        {
            try
            {
                object swApp = Marshal.GetActiveObject("SldWorks.Application");
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}