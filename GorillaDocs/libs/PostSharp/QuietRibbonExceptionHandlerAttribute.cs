using PostSharp.Aspects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;

namespace GorillaDocs.libs.PostSharp
{
    [Serializable]
    public sealed class QuietRibbonExceptionHandlerAttribute : OnMethodBoundaryAspect
    {
        public override void OnException(MethodExecutionArgs args)
        {
            if (args.Exception is COMException || args.Exception is InvalidOperationException)
                Message.LogWarning(args.Exception);
            else
                Message.LogError(args.Exception);
            args.FlowBehavior = FlowBehavior.Return;
        }
    }
}
