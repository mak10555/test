using System.Runtime.InteropServices;

namespace TSP.Export
{
    public static class ServiceMethods
    {
        public static void ReleaseComObject(object comObject)
        {
            if (comObject != null)
                Marshal.ReleaseComObject(comObject);
        }
    }
}
