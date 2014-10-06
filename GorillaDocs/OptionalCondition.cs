using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace GorillaDocs
{
    public class OptionalCondition
    {
        //TODO: Implement AND & NOT

        readonly Stack<BooleanExpression> expressions = new Stack<BooleanExpression>();

        public OptionalCondition(string value, XDocument data)
        {
            var args = value.Split(new string[] { "or", "OR", "Or", "oR", "||" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string arg in args)
                expressions.Push(new BooleanExpression(arg, data));
        }

        public bool? Evaluate()
        {
            while (expressions.Count > 0)
            {
                var result = expressions.Pop().Evaluate();
                if (result == true)
                    return true;
                else if (result == null)
                    return null;
                // Keep evaluating
            }
            return null;
        }
    }
}
