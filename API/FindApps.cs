using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace MSOfficeManager.API
{
    internal static class FindApps
    {
        private struct RunningObject
        {
            public string name;
            public object o;
        }

        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

        private static List<RunningObject> GetRunningObjects()
        {
            List<RunningObject> res = new List<RunningObject>();
            IBindCtx bc;
            CreateBindCtx(0, out bc);
            IRunningObjectTable runningObjectTable;
            bc.GetRunningObjectTable(out runningObjectTable);
            IEnumMoniker monikerEnumerator;
            runningObjectTable.EnumRunning(out monikerEnumerator);
            monikerEnumerator.Reset();

            IMoniker[] monikers = new IMoniker[1];
            IntPtr numFetched = IntPtr.Zero;
            while (monikerEnumerator.Next(1, monikers, numFetched) == 0)
            {
                RunningObject running;
                monikers[0].GetDisplayName(bc, null, out running.name);
                runningObjectTable.GetObject(monikers[0], out running.o);
                res.Add(running);
            }
            return res;
        }

        public static object GetObjects(Type type)
        {
            IList ret = Activator.CreateInstance(typeof(List<>).MakeGenericType(type)) as IList;
            foreach (RunningObject running in GetRunningObjects())
                if (running.o?.GetType() == type)
                    ret.Add(running.o);
            return ret;
        }
    }
}
