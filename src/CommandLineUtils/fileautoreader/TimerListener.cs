using System.Threading;
using TWUtilities40;

namespace FileAutoReader
{
    class TimerListener : ITimerExpiryListener
    {
        public void TimerExpired(ref TimerExpiredEventData ev)
        {
            Program.writeSingleLineToConsole("");
        }
    }
}
