using System;
using System.Collections.Generic;
using System.Linq;
using PostSharp.Aspects;

namespace GorillaDocs.libs.PostSharp
{
    [Serializable]
    public sealed class LogAttribute : OnMethodBoundaryAspect
    {
        [System.Diagnostics.DebuggerStepThrough]
        public override void OnEntry(MethodExecutionArgs args)
        {
            if (Message.IsDebugEnabled())
                Message.LogDebug(args.Method, string.Format("{0} Enter", GetParameters(args)));
        }

        [System.Diagnostics.DebuggerStepThrough]
        public override void OnExit(MethodExecutionArgs args)
        {
            if (Message.IsDebugEnabled())
                Message.LogDebug(args.Method, "Returns: " + args.ReturnValue);
        }

        [System.Diagnostics.DebuggerStepThrough]
        public override void OnException(MethodExecutionArgs args)
        {
            if (Message.IsWarnEnabled())
                Message.LogWarning(args.Method, args.Exception);
        }

        [System.Diagnostics.DebuggerStepThrough]
        string GetParameters(MethodExecutionArgs args)
        {
            string parameters = string.Empty;
            for (var i = 0; i < args.Arguments.Count; i++)
            {
                string name = args.Method.GetParameters()[i].Name;
                object arg = args.Arguments[i];
                Type type = arg.NullableGetType();
                string value = Convert.ToString(arg);
                if (Convert.ToString(type) == value)
                    parameters += string.Format(", {0} [{1}]", name, type);
                else
                    parameters += string.Format(", {0} [{1}] {2}", name, type, value);
            }
            return string.Format("({0})", parameters.TrimStart(new char[] { ',', ' ' }));
        }
    }
}
