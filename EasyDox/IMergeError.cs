namespace EasyDox
{
    public interface IMergeError
    {
        string Accept(IMergeErrorVisitor visitor);
    }
}