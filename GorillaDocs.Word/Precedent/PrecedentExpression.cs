using System;
using System.Collections.Generic;
using System.Linq;
using E = System.Linq.Expressions;
using System.Linq.Dynamic;

namespace GorillaDocs.Word.Precedent
{
    public class PrecedentExpression1
    {
        public static bool Resolve<T>(string Expression, T Data, string VariableNameUsedInExpression = "obj")
        {
            var p = E.Expression.Parameter(typeof(T), VariableNameUsedInExpression);
            var e = DynamicExpression.ParseLambda(new[] { p }, null, Expression);
            return (bool)e.Compile().DynamicInvoke(Data);
        }
    }
}
