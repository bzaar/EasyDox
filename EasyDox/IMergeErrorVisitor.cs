namespace EasyDox
{
    public interface IMergeErrorVisitor
    {
        string InvalidExpression(string expr);
        string MissingField(string fieldName);
    }
}