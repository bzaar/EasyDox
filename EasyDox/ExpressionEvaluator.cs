using System.Linq;

namespace EasyDox
{
    class ExpressionEvaluator : IExpressionVisitor
    {
        private readonly Properties properties;

        string result;

        public ExpressionEvaluator (Properties properties, IExpression expression)
        {
            this.properties = properties;

            expression.Accept (this);
        }

        public string Result
        {
            get { return result; }
        }

        void IExpressionVisitor.Visit(Literal literal)
        {
            result = literal.Value;
        }

        void IExpressionVisitor.Visit(Field field)
        {
            result = properties [field.Name];
        }

        void IExpressionVisitor.Visit(Function function)
        {
            result = function.Definition.Eval (function.Args.Select (a => new ExpressionEvaluator (properties, a).result).ToArray());
        }
    }
}