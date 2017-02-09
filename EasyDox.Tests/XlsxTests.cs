using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EasyDox.Tests
{
    [TestClass]
    public class XlsxTests
    {
        [TestMethod]
        public void XlsxRegexMethod1()
        {
            var match = Xlsx.regex.Match("[[Должность]]");
            Assert.IsTrue(match.Success);
            Assert.AreEqual("Должность", match.Groups["name"].Value);
            Assert.AreEqual("[[Должность]]", match.Groups["template"].Value);
        }

        [TestMethod]
        public void XlsxRegexMethod2()
        {
            var match = Xlsx.regex.Match(" [[ Должность ]] ");
            Assert.IsTrue(match.Success);
            Assert.AreEqual("Должность", match.Groups["name"].Value);
            Assert.AreEqual("[[ Должность ]]", match.Groups["template"].Value);
        }

        [TestMethod]
        public void XlsxRegexMethod3()
        {
            var match = Xlsx.regex.Match("ООО \"Самая лучшая компания\", [[Адрес]] ");
            
            Assert.IsTrue(match.Success);
            Assert.AreEqual("Адрес", match.Groups["name"].Value);
            Assert.AreEqual("[[Адрес]]", match.Groups["template"].Value);
        }

        [TestMethod]
        public void XlsxRegexMethod4()
        {
            var match = Xlsx.regex.Match("Счет на оплату № [[№ счета]] от [[Дата счета]]");

            Assert.IsTrue(match.Success);
            Assert.AreEqual(2, match.Groups["name"].Captures.Count);

            Assert.AreEqual("№ счета", match.Groups["name"].Captures[0].Value);
            Assert.AreEqual("Дата счета", match.Groups["name"].Captures[1].Value);

            Assert.AreEqual(2, match.Groups["template"].Captures.Count);
            Assert.AreEqual("[[№ счета]]", match.Groups["template"].Captures[0].Value);
            Assert.AreEqual("[[Дата счета]]", match.Groups["template"].Captures[1].Value);
        }

        [TestMethod]
        public void XlsxRegexMethod6()
        {
            var match = Xlsx.regex.Match("[[Доверенность]][[Фамилия]][[Отчество]]");
            Assert.IsTrue(match.Success);

            Assert.AreEqual(3, match.Groups["name"].Captures.Count);
            Assert.AreEqual("Доверенность", match.Groups["name"].Captures[0].Value);
            Assert.AreEqual("Фамилия", match.Groups["name"].Captures[1].Value);
            Assert.AreEqual("Отчество", match.Groups["name"].Captures[2].Value);

            Assert.AreEqual(3, match.Groups["template"].Captures.Count);
            Assert.AreEqual("[[Доверенность]]", match.Groups["template"].Captures[0].Value);
            Assert.AreEqual("[[Фамилия]]", match.Groups["template"].Captures[1].Value);
            Assert.AreEqual("[[Отчество]]", match.Groups["template"].Captures[2].Value);

        }

        [TestMethod]
        public void XlsxRegexMethod7()
        {
            var match = Xlsx.regex.Match(" На основании доверенности № [[Доверенность]] в лице [[ Фамилия ]], проживающий по адресу:[[ Адрес]] ");
            Assert.IsTrue(match.Success);

            Assert.AreEqual(3, match.Groups["name"].Captures.Count);
            Assert.AreEqual("Доверенность", match.Groups["name"].Captures[0].Value);
            Assert.AreEqual("Фамилия", match.Groups["name"].Captures[1].Value);
            Assert.AreEqual("Адрес", match.Groups["name"].Captures[2].Value);

            Assert.AreEqual(3, match.Groups["template"].Captures.Count);
            Assert.AreEqual("[[Доверенность]]", match.Groups["template"].Captures[0].Value);
            Assert.AreEqual("[[ Фамилия ]]", match.Groups["template"].Captures[1].Value);
            Assert.AreEqual("[[ Адрес]]", match.Groups["template"].Captures[2].Value);

        }

        [TestMethod]
        [DeploymentItem("XlsxSharedStrings.xml")]
        public void XlsxParseSharedStrings()
        {
            var xdoc = new XmlDocument();
            xdoc.Load("XlsxSharedStrings.xml");
            var fields = Xlsx.GetFields(xdoc).ToArray();

            Assert.AreEqual(4, fields.Length);

            Assert.AreEqual("[[Родительный (Должность представителя Лицензиата) ]]", fields[0].Text);
            Assert.AreEqual("[[ Родительный (ФИО представителя Лицензиата) ]] ", fields[1].Text);
            Assert.AreEqual("[[Полное наименование компании Лицензиата]]", fields[2].Text);
            Assert.AreEqual("  На основании доверенности № [[Доверенность]] в лице [[ Фамилия ]], проживающий по адресу:[[ Адрес]] ", fields[3].Text);
        }

        [TestMethod]
        [DeploymentItem("XlsxSharedStrings2.xml")]
        public void XlsxSubstitureSharedStrings()
        {
            var xdoc = new XmlDocument();
            xdoc.Load("XlsxSharedStrings2.xml");

            var stream = new MemoryStream();
            xdoc.Save(stream);

            var xdoc2 = new XmlDocument();
            stream.Position = 0;
            xdoc2.Load(stream);

            var dict = new Dictionary<string, string>()
            {
                {"Доверенность", "123-456/АГ"},
                {"Адрес",  "Петропавловск-Камчатский"},
                {"Фамилия", "Иванов И.П."},
            };

            var engine = new Engine();
            Xlsx.ReplaceMergeFieldsAndReturnMissingFieldNames(xdoc2, dict, engine);

            var fields = Xlsx.GetFields(xdoc2);

            Assert.AreEqual("  На основании доверенности № 123-456/АГ в лице Иванов И.П., проживающий по адресу: Петропавловск-Камчатский ", fields.Single().Value);
        }

        [TestMethod, DeploymentItem("Invoice.xlsx")]
        public void SmokeTest()
        {
            Console.WriteLine(Environment.CurrentDirectory);

            var dict = new Dictionary<string, string>
            {
                {"№ счета", "123"},
                {"Лицензиат",  "ООО \"Тюльпан\""},
                {"Дата счета", "8 марта 2017"},
                {"Продукт", "Право использования библиотеки \"Морфер\""},
                {"Цена", "1500"},
            };

            var engine = new Engine();

            var errors = engine.MergeXL("Invoice.xlsx", dict, "invoice1.xlsx");

            Assert.IsFalse(errors.Any());
        }
    }
}
