using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace GorillaDocs
{
    [Obsolete]
    public class BooleanExpression
    {
        enum OperatorType { NoOperator, Equals, NotEquals, GreaterThan, LessThan, GreaterThanOrEqual, LessThanOrEqual }
        readonly OperatorType operatorType;
        readonly string operatorValue;
        string LHS;
        string RHS;
        public BooleanExpression(string value, XDocument data)
        {
            if (value.Contains("<="))
            {
                operatorType = OperatorType.LessThanOrEqual;
                operatorValue = "<=";
            }
            else if (value.Contains(">="))
            {
                operatorType = OperatorType.GreaterThanOrEqual;
                operatorValue = ">=";
            }
            else if (value.Contains("<>"))
            {
                operatorType = OperatorType.NotEquals;
                operatorValue = "<>";
            }
            else if (value.Contains("!="))
            {
                operatorType = OperatorType.NotEquals;
                operatorValue = "!=";
            }
            else if (value.Contains("<"))
            {
                operatorType = OperatorType.LessThan;
                operatorValue = "<";
            }
            else if (value.Contains(">"))
            {
                operatorType = OperatorType.GreaterThan;
                operatorValue = ">";
            }
            else if (value.Contains("="))
            {
                operatorType = OperatorType.Equals;
                operatorValue = "=";
            }
            else
            {
                operatorType = OperatorType.NoOperator;
            }

            ParseArguments(value, operatorValue, data);
        }

        void ParseArguments(string expression, string operatorValue, XDocument data)
        {
            var args = expression.Split(new string[] { operatorValue }, StringSplitOptions.RemoveEmptyEntries);
            ExpandArguments(args, data);
            LHS = args[0].Trim().Trim('\"');
            if (operatorType != OperatorType.NoOperator)
                RHS = args[1].Trim().Trim('\"');
        }

        static void ExpandArguments(string[] args, XDocument data)
        {
            for (int i = 0; i < args.Length; i++)
            {
                args[i] = args[i].Trim();
                var element = data.Descendants().SingleOrDefault(e => e.Name.LocalName == args[i]);
                if (element != null)
                    args[i] = element.Value;
            }
        }

        public bool? Evaluate()
        {
            if (string.IsNullOrEmpty(LHS) || string.IsNullOrEmpty(RHS))
                return null;

            switch (operatorType)
            {
                case OperatorType.NoOperator:
                    return bool.Parse(LHS);
                case OperatorType.Equals:
                    return LHS == RHS;
                case OperatorType.NotEquals:
                    return LHS != RHS;
                case OperatorType.GreaterThan:
                    return float.Parse(LHS) > float.Parse(RHS);
                case OperatorType.LessThan:
                    return float.Parse(LHS) < float.Parse(RHS);
                case OperatorType.GreaterThanOrEqual:
                    return float.Parse(LHS) >= float.Parse(RHS);
                case OperatorType.LessThanOrEqual:
                    return float.Parse(LHS) <= float.Parse(RHS);
            }
            throw new InvalidOperationException("Unable to evaluate expression.");
        }
    }
}
