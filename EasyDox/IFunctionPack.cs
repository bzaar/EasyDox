using System.Collections.Generic;

namespace EasyDox
{
    public interface IFunctionPack
    {
        Dictionary <string, FunctionDefinition> Functions { get; }
    }
}