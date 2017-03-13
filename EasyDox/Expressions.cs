using System;
using System.Collections.Generic;

namespace EasyDox
{
    interface IExpression
    {
        void Accept (IExpressionVisitor visitor);
    }

    interface IExpressionVisitor
    {
        void Visit (Literal literal);
        void Visit (Field field);
        void Visit (Function function);
    }

    /// <summary>
    /// Represents a literal string in an expression.
    /// </summary>
    class Literal : IExpression
    {
        public Literal (string s)
        {
            Value = s;
        }

        void IExpression.Accept(IExpressionVisitor visitor)
        {
            visitor.Visit (this);
        }

        public string Value { get; }
    }

    /// <summary>
    /// Represents a field reference in an expression.
    /// </summary>
    class Field : IExpression
    {
        public Field(string name)
        {
            Name = name;
        }

        void IExpression.Accept(IExpressionVisitor visitor)
        {
            visitor.Visit (this);
        }

        /// <summary>
        /// Gets the field name.
        /// </summary>
        public string Name { get; }
    }

    class Function : IExpression
    {
        public Function (Delegate def, IEnumerable<IExpression> args)
        {
            Definition = def;
            Args = args;
        }

        public Delegate Definition { get; }

        public IEnumerable <IExpression> Args { get; }

        void IExpression.Accept(IExpressionVisitor visitor)
        {
            visitor.Visit (this);
        }
    }
}