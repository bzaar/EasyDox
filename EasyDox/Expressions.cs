using System;
using System.Collections.Generic;
using System.Linq;

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
        public Function (FunctionDefinition def, IEnumerable<IExpression> args)
        {
            Definition = def;
            Args = args;
        }

        public FunctionDefinition Definition { get; }

        public IEnumerable <IExpression> Args { get; }

        void IExpression.Accept(IExpressionVisitor visitor)
        {
            visitor.Visit (this);
        }
    }

    public class FunctionDefinition
    {
        readonly IFuncN func;

        public FunctionDefinition (Func<string,string> func)
        {
            this.func = new Func1 (func);
        }

        public FunctionDefinition (Func<string,string,string> func)
        {
            this.func = new Func2 (func);
        }

        public int ArgCount => func.ArgCount;

        public string Eval (string [] args)
        {
            return func.Eval (args);
        }
    }

    interface IFuncN
    {
        string Eval (string [] args);
        int ArgCount {get;}
    }

    class Func1 : IFuncN
    {
        private readonly Func <string, string> func;

        public Func1 (Func<string, string> func)
        {
            this.func = func;
        }

        string IFuncN.Eval(string [] args)
        {
            return func(args.Single());
        }

        int IFuncN.ArgCount => 1;
    }

    class Func2 : IFuncN
    {
        private readonly Func <string, string, string> func;

        public Func2 (Func<string, string, string> func)
        {
            this.func = func;
        }

        string IFuncN.Eval(string [] args)
        {
            return func(args[0], args[1]);
        }

        int IFuncN.ArgCount => 2;
    }
}