using System;
using System.Management.Automation;
using System.Threading;

namespace PSWordCloud
{
    internal class ProcessingState : IDisposable
    {
        internal PSObject? Data { get; private set; }

        internal EventWaitHandle WaitHandle { get; }

        internal ProcessingState(PSObject data)
        {
            Data = data;
            WaitHandle = new EventWaitHandle(initialState: false, EventResetMode.ManualReset);
        }

        public void Dispose()
        {
            Data = null;
            WaitHandle.Dispose();
        }
    }
}
