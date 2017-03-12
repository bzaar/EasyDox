using System;

namespace EasyDox
{
    public class Func2 : IFuncN
    {
        readonly Func <string, string, string> func;

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