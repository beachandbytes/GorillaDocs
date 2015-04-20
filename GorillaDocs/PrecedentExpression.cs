using System;
using System.Collections.Generic;
using System.Linq;
using E = System.Linq.Expressions;
using System.Linq.Dynamic;
using System.Reflection;

namespace GorillaDocs
{
    public class PrecedentExpression
    {
        public static bool Resolve<T>(string Expression, T Data, string VariableNameUsedInExpression)
        {
            var p = E.Expression.Parameter(typeof(T), VariableNameUsedInExpression);
            var e = DynamicExpression.ParseLambda(new[] { p }, null, Expression);
            return (bool)e.Compile().DynamicInvoke(Data);
        }
    }
}
