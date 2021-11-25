using System.Diagnostics.Tracing;
using System.IO;

namespace Helper
{
    [EventSource(Name = "OfficeSuiteEventProvider")]
    class Logging : EventSource
    {

        public static string GetEventName()
        {
            string sysPath = Directory.GetCurrentDirectory();
            string eventProviderName = Path.GetFileNameWithoutExtension(sysPath);
            return eventProviderName;
        }


        public void StartBenchmark()
        {
            WriteEvent(1, "Start of Program");
        }


        public void EndBenchmark()
        {
            WriteEvent(2, "Stop of microbenchmark test");
        }

        public static Logging Log = new Logging();
    }
}
