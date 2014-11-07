using PostSharp.Aspects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;

namespace GorillaDocs.libs.PostSharp
{
    [Serializable]
    public sealed class LoudRibbonExceptionHandlerAttribute : OnMethodBoundaryAspect
    {
        public override void OnException(MethodExecutionArgs args)
        {
            Message.ShowError(args.Exception, Assembly.GetCallingAssembly());
            args.FlowBehavior = FlowBehavior.Return;
        }
    }
}
