using System.Collections.Generic;

namespace EasyDox
{
    class Properties : Dictionary<string, string>
    {
        public Properties() {}
        public Properties(Dictionary<string, string> dict) : base(dict) { }

        public string Eval(IExpression exp)
        {
            return new ExpressionEvaluator(this, exp).Result;
        }

        public void FindMissingProperties(IExpression exp, List<string> missingProperties)
        {
            exp.Accept(new MissingPropertiesFinder(this, missingProperties));
        }
    }
}