using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.XPath;

namespace EasyDox
{
    public static class Docx
    {
        /// <summary>
        /// Merges <paramref name="fieldValues"/> into the docx template specified by <paramref name="docxPath"/> and replaces the original file.
        /// </summary>
        /// <param name="engine">Expression evaluation engine.</param>
        /// <param name="docxPath">Template and output path.</param>
        /// <param name="fieldValues">A dictionary of field values keyed by field name.</param>
        /// <returns></returns>
        public static IEnumerable<IMergeError> MergeInplace(Engine engine, string docxPath,
            Dictionary<string, string> fieldValues)
        {
            using (var pkg = Package.Open(docxPath, FileMode.Open, FileAccess.ReadWrite))
            {
                // Specify the URI of the part to be read
                PackagePart part = pkg.GetPart(new Uri("/word/document.xml", UriKind.Relative));

                // Get the document part from the package.
                // Load the XML in the part into an XmlDocument instance.
                var xdoc = new XmlDocument();

                using (var partStream = part.GetStream(FileMode.Open, FileAccess.Read))
                {
                    xdoc.Load(partStream);
                }

                var fields = ReplaceMergeFieldsAndReturnMissingFieldNames(xdoc, fieldValues, engine);

                using (var partStream = part.GetStream(FileMode.Open, FileAccess.Write))
                using (var partWrt = new StreamWriter(partStream))
                {
                    xdoc.Save(partWrt);
                }

                return fields;
            }
        }

        class MergeError : IMergeError
        {
            private readonly Func<IMergeErrorVisitor, string> callback;

            public MergeError(Func<IMergeErrorVisitor, string> callback)
            {
                this.callback = callback;
            }

            public string Accept(IMergeErrorVisitor visitor)
            {
                return callback(visitor);
            }
        }

        static IEnumerable<IMergeError> ReplaceMergeFieldsAndReturnMissingFieldNames(XmlDocument xdoc,
            Dictionary<string, string> dict, Engine engine)
        {
            var fields = GetFields(xdoc);

            var errors = new List<IMergeError>();

            var properties = new Properties(dict);

            foreach (var field in fields)
            {
                var matches = regex.Match(field.InstrText);

                if (!matches.Success) continue;

                var fieldName = matches.Groups["name"].Value;

                var format = matches.Groups["format"].Captures;

                var exp = engine.Parse(fieldName);

                if (exp == null)
                {
                    errors.Add(new MergeError(v => v.InvalidExpression(fieldName)));
                }
                else
                {
                    var missingProperties = new List<string>();

                    properties.FindMissingProperties(exp, missingProperties);

                    errors.AddRange(missingProperties.Select(p => new MergeError(v => v.MissingField(p))));

                    if (missingProperties.Count == 0) // otherwise Eval will throw
                    {
                        field.Value = ApplyFormat(properties.Eval(exp), format);
                    }
                }
            }

            return errors;
        }

        private static string ApplyFormat(string s, CaptureCollection format)
        {
            foreach (Capture capture in format)
            {
                switch (capture.Value.ToUpperInvariant())
                {
                    case "FIRSTCAP":
                        s = s.Substring(0, 1).ToUpperInvariant() + s.Substring(1);
                        break;
                }
            }

            return s;
        }

        internal static IEnumerable<IField> GetFields(XmlDocument xdoc)
        {
            var nsManager = new XmlNamespaceManager(new NameTable());
            nsManager.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");

            XPathNavigator xDocNavigator = xdoc.CreateNavigator();

            var nodes = xDocNavigator.Select("//w:r[w:fldChar/@w:fldCharType='begin']",
                nsManager);

            foreach (XPathNavigator navigator in nodes)
            {
                yield return new ComplexField(navigator, nsManager);
            }

            var simpleNodes = xDocNavigator.Select("//w:fldSimple[starts-with(@w:instr,' MERGEFIELD')]",
                nsManager);

            foreach (XPathNavigator navigator in simpleNodes)
            {
                yield return new SimpleField(navigator, nsManager);
            }
        }

        internal static readonly Regex regex = new Regex(

            @"^[\s]*MERGEFIELD[\s]+(?<name>[^\s""]+)|([""](?<name>[^""]+)[""])
            ([\s]*\\\*[\s]*(?<format>\w+))*[\s]*?",

            RegexOptions.Compiled
            | RegexOptions.CultureInvariant
            | RegexOptions.ExplicitCapture
            | RegexOptions.IgnoreCase
            | RegexOptions.IgnorePatternWhitespace
            | RegexOptions.Singleline);

        internal interface IField
        {
            string InstrText { get; }
            string Value { get; set; }
        }

        internal class SimpleField : IField
        {
            private readonly XPathNavigator node;
            private readonly XmlNamespaceManager namespaceManager;

            public SimpleField(XPathNavigator node, XmlNamespaceManager namespaceManager)
            {
                this.node = node;
                this.namespaceManager = namespaceManager;
            }

            string IField.InstrText => node.GetAttribute("instr", namespaceManager.LookupNamespace("w"));

            string IField.Value
            {
                get => ValueNode.Value;
                set => ValueNode.SetValue(value);
            }

            private XPathNavigator ValueNode => node.Select("w:r/w:t", namespaceManager).Cast<XPathNavigator>().Single();
        }

        internal class ComplexField : IField
        {
            private readonly XPathNavigator begin;
            private readonly XmlNamespaceManager nsManager;

            public ComplexField(XPathNavigator begin, XmlNamespaceManager nsManager)
            {
                this.begin = begin;
                this.nsManager = nsManager;
            }

            string IField.InstrText
            {
                get
                {
                    // TODO: change this to look for w:fldCharType="separate" and get rid of the hard coded 5.
                    var nodes = begin.Select("following-sibling::w:r[position() <= 5]/w:instrText", nsManager);

                    return string.Join("", nodes.Cast<XPathNavigator>().Select(n => n.Value));
                }
            }

            string IField.Value
            {
                get => throw new NotImplementedException();
                set
                {
                    var nodes = begin.Select(
                        "following-sibling::w:r[w:fldChar/@w:fldCharType='separate']/following-sibling::w:r/w:t",
                        nsManager);

                    var mode = nodes.Cast<XPathNavigator>().First();

                    mode.SetValue(value);
                }
            }
        }
    }
}