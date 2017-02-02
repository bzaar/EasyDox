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
        private readonly string s;

        public Literal (string s)
        {
            this.s = s;
        }

        void IExpression.Accept(IExpressionVisitor visitor)
        {
            visitor.Visit (this);
        }

        public string Value { get {return s;} }
    }

    /// <summary>
    /// Represents a field reference in an expression.
    /// </summary>
    class Field : IExpression
    {
        private readonly string name;

        public Field(string name)
        {
            this.name = name;
        }

        void IExpression.Accept(IExpressionVisitor visitor)
        {
            visitor.Visit (this);
        }

        /// <summary>
        /// Gets the field name.
        /// </summary>
        public string Name { get {return name;} }
    }

    class Function : IExpression
    {
        private readonly FunctionDefinition def;
        private readonly IEnumerable <IExpression> args;

        public Function (FunctionDefinition def, IEnumerable<IExpression> args)
        {
            this.def = def;
            this.args = args;
        }

        public FunctionDefinition Definition
        {
            get { return def; }
        }

        public IEnumerable <IExpression> Args
        {
            get { return args; }
        }

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

        public int ArgCount
        {
            get {return func.ArgCount;}
        }

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

        int IFuncN.ArgCount
        {
            get { return 1; }
        }
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

        int IFuncN.ArgCount
        {
            get { return 2; }
        }
    }
}