using System;
using System.Collections.Generic;

namespace EasyDox
{
    public static class Xlsx
    {
        /// <summary>
        /// Merges <paramref name="fieldValues"/> into the xlsx template specified 
        /// by <paramref name="xlsxPath"/> and replaces the original file.
        /// </summary>
        /// <param name="engine">Expression evaluation engine.</param>
        /// <param name="xlsxPath">Template and output path.</param>
        /// <param name="fieldValues">A dictionary of field values keyed by field name.</param>
        /// <returns></returns>
        public static IEnumerable<IMergeError> MergeInplace(
            Engine engine, 
            string xlsxPath,
            Dictionary<string, string> fieldValues)
        {
            throw new NotImplementedException();
        }
    }
}
