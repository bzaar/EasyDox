using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Xml;
using System.Xml.XPath;
using System.Text.RegularExpressions;
using System.Linq;

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
            using (var pkg = Package.Open(xlsxPath, FileMode.Open, FileAccess.ReadWrite))
            {
                // Specify the URI of the part to be read
                PackagePart part = pkg.GetPart(new Uri ("/xl/sharedStrings.xml", UriKind.Relative));

                // Get the document part from the package.
                // Load the XML in the part into an XmlDocument instance.
                var xdoc = new XmlDocument();
                xdoc.Load(part.GetStream(FileMode.Open, FileAccess.Read));

                var fields = ReplaceMergeFieldsAndReturnMissingFieldNames(xdoc, fieldValues, engine);

                using (var partWrt = new StreamWriter(part.GetStream(FileMode.Open, FileAccess.Write)))
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

        public static IEnumerable<IMergeError> ReplaceMergeFieldsAndReturnMissingFieldNames(XmlDocument xdoc, Dictionary<string, string> dict, Engine engine)
        {
            var fields = GetFields(xdoc);

            var errors = new List<IMergeError>();

            var properties = new Properties(dict);

            foreach (var field in fields)
            {
                var value = field.Text;
                var matches = regex.Match(field.Text);

                if (matches.Success)
                {
                    var fieldNames = matches.Groups["name"].Captures;
                    var fieldTemplates = matches.Groups["template"].Captures;

                    for(var fieldIdx = 0; fieldIdx < fieldNames.Count; ++fieldIdx )
                    {
                        var fieldName = fieldNames[fieldIdx].ToString();
                        var fieldTemlate = fieldTemplates[fieldIdx].ToString();

                        var exp = engine.Parse(fieldName.ToString());

                        if (exp == null)
                        {
                            errors.Add(new MergeError(v => v.InvalidExpression(fieldName.ToString())));
                        }
                        else
                        {
                            var missingProperties = new List<string>();

                            properties.FindMissingProperties(exp, missingProperties);

                            errors.AddRange(missingProperties.Select(p => new MergeError(v => v.MissingField(p))));

                            if (missingProperties.Count == 0) // otherwise Eval will throw
                            {
                                value = value.Replace(fieldTemlate, properties.Eval(exp));
                            }
                        }
                    }

                    field.Value = value;
                }
            }

            return errors;
        }

        internal static IEnumerable<IField> GetFields(XmlDocument xdoc)
        {
            var nsManager = new XmlNamespaceManager(new NameTable());
            nsManager.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            XPathNavigator xDocNavigator = xdoc.CreateNavigator();

            var nodes = xDocNavigator.Select("//s:sst/s:si/s:t", nsManager);

            foreach (XPathNavigator navigator in nodes)
            {
                yield return new SimpleField(navigator, nsManager);
            }
        }

        internal static readonly Regex regex = new Regex(

            @"^(.*?(?<template>\[\[[\s]*(((?<name>[^\s""]+?)|([""](?<name>[^""]+?)[""]))[\s]*)\]\]))+",

            RegexOptions.Compiled
            | RegexOptions.CultureInvariant
            | RegexOptions.ExplicitCapture
            | RegexOptions.IgnoreCase
            | RegexOptions.IgnorePatternWhitespace
            | RegexOptions.Singleline);

        internal interface IField
        {
            string Text { get; }
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

            string IField.Text
            {
                // ReSharper disable AssignNullToNotNullAttribute
                get { return node.InnerXml; }
                // ReSharper restore AssignNullToNotNullAttribute
            }

            string IField.Value
            {
                get
                {
                    return ValueNode.Value;
                }
                set
                {
                    ValueNode.SetValue(value);
                }
            }

            private XPathNavigator ValueNode
            {
                //get { return node.Select("w:r/w:t", namespaceManager).Cast<XPathNavigator>().Single(); }
                get { return node; }

            }
        }

    }
}
