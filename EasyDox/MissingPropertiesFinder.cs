using System.Collections.Generic;

namespace EasyDox
{
    class MissingPropertiesFinder : IExpressionVisitor
    {
        private readonly Properties properties;

        private readonly List<string> missingProperties;

        public MissingPropertiesFinder(Properties properties, List<string> missingProperties)
        {
            this.properties = properties;
            this.missingProperties = missingProperties;
        }

        void IExpressionVisitor.Visit(Literal literal) { }

        void IExpressionVisitor.Visit(Field field)
        {
            if (!properties.ContainsKey(field.Name) && !missingProperties.Contains(field.Name))
            {
                missingProperties.Add(field.Name);
            }
        }

        void IExpressionVisitor.Visit(Function function)
        {
            foreach (var arg in function.Args)
            {
                arg.Accept(this);
            }
        }
    }
}