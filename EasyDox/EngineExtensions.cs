using System.Collections.Generic;
using System.IO;
using EasyDox;

public static class EngineExtensions
{
    /// <summary>
    /// Merges the field values into the template specified by <paramref name="templatePath"/> and saves the output to <paramref name="outputPath"/>.
    /// </summary>
    /// <param name="engine"></param>
    /// <param name="fieldValues">A dictionary of field values keyed by field name.</param>
    /// <param name="templatePath">Path to template docx.</param>
    /// <param name="outputPath">Path to output docx.</param>
    /// <returns></returns>
    public static IEnumerable <IMergeError> Merge (this Engine engine, string templatePath, Dictionary <string, string> fieldValues, string outputPath)
    {
        File.Copy (templatePath, outputPath, true);

        return Docx.MergeInplace(engine, outputPath, fieldValues);
    }

    public static IEnumerable<IMergeError> MergeXL(this Engine engine, string templatePath, Dictionary<string, string> fieldValues, string outputPath)
    {
        File.Copy(templatePath, outputPath, true);
        return Xlsx.MergeInplace(engine, outputPath, fieldValues);
    }
}