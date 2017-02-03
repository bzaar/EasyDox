namespace EasyDox
{
    public class ErrorToRussianString : IMergeErrorVisitor
    {
        string IMergeErrorVisitor.InvalidExpression(string expr)
        {
            return "Ошибка в выражении: " + expr + ".";
        }

        string IMergeErrorVisitor.MissingField(string fieldName)
        {
            return "Не заполнено поле " + fieldName + ".";
        }
    }
}