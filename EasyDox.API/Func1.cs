using System;
using System.Linq;

namespace EasyDox
{
    public class Func1 : IFuncN
    {
        readonly Func <string, string> func;

        public Func1 (Func<string, string> func)
        {
            this.func = func;
        }

        string IFuncN.Eval(string [] args)
        {
            return func(args[0]);
        }

        int IFuncN.ArgCount => 1;
    }
}