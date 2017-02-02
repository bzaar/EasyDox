namespace EasyDox
{
    public class ErrorToRussianString : Docx.IMergeErrorVisitor
    {
        string Docx.IMergeErrorVisitor.InvalidExpression(string expr)
        {
            return "Ошибка в выражении: " + expr + ".";
        }

        string Docx.IMergeErrorVisitor.MissingField(string fieldName)
        {
            return "Не заполнено поле " + fieldName + ".";
        }
    }
}