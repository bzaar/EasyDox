using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EasyDox.Tests
{
    [TestClass]
    public class DocxTests
    {
        [TestMethod]
        public void TestMethod1()
        {
            var match = Docx.regex.Match ("  MERGEFIELD Должность ");

            Assert.IsTrue (match.Success);

            Assert.AreEqual ("Должность", match.Groups ["name"].Value);
        }

        [TestMethod]
        public void TestMethod2()
        {
            var match = Docx.regex.Match ("  MERGEFIELD \"Должность представителя Лицензиата\" ");

            Assert.IsTrue (match.Success);

            Assert.AreEqual ("Должность представителя Лицензиата", match.Groups ["name"].Value);
        }

        [TestMethod]
        public void TestMethod3()
        {
            var match = Docx.regex.Match ("  MERGEFIELD \"Должность представителя Лицензиата\" \\* MERGEFORMAT");

            Assert.IsTrue (match.Success);

            Assert.AreEqual ("Должность представителя Лицензиата", match.Groups ["name"].Value);

            Assert.AreEqual ("MERGEFORMAT", match.Groups ["format"].Value);
        }

        [TestMethod]
        public void TestMethod4()
        {
            var match = Docx.regex.Match ("  MERGEFIELD \"Должность представителя Лицензиата\" \\* MERGEFORMAT \\* FirstCap");

            Assert.IsTrue (match.Success);

            Assert.AreEqual ("Должность представителя Лицензиата", match.Groups ["name"].Value);

            Assert.AreEqual ("MERGEFORMAT", match.Groups ["format"].Captures [0].Value);
            Assert.AreEqual ("FirstCap", match.Groups ["format"].Captures [1].Value);
        }


        [TestMethod]
        public void TestMethod5()
        {
            var match = Docx.regex.Match ("  MERGEFIELD  \"Родительный (Должность представителя Лицензиата)\"  \\* MERGEFORMAT ");

            Assert.IsTrue (match.Success);

            Assert.AreEqual ("Родительный (Должность представителя Лицензиата)", match.Groups ["name"].Value);

            Assert.AreEqual ("MERGEFORMAT", match.Groups ["format"].Captures [0].Value);
        }
        
        [TestMethod]
        public void MatchQuote ()
        {
            Assert.AreEqual ("asdf", Regex.Match ("\"asdf\"", @"^""(?<name>[^""]+)""$").Groups ["name"].Value);
        }

        [TestMethod]
        public void MatchNonSpace ()
        {
            Assert.AreEqual ("asdf x", Regex.Match ("\"asdf x\"", @"^(?<name>[^\s""]+)|(""(?<name>[^""]+)"")$").Groups ["name"].Value);
        }

        [TestMethod]
        [DeploymentItem ("DocxComplexField.xml")]
        public void TestMethod6 ()
        {
            var xdoc = new XmlDocument ();
            xdoc.Load ("DocxComplexField.xml");
            var fields = Docx.GetFields (xdoc);
            Assert.IsTrue (fields.Select (f => f.InstrText).SequenceEqual (new [] {
                                                                                      " MERGEFIELD  \"Родительный (Должность представителя Лицензиата)\"  \\* MERGEFORMAT ",
                                                                                      " MERGEFIELD  \"Родительный (ФИО представителя Лицензиата)\"  \\* MERGEFORMAT ",
                                                                                      " MERGEFIELD  \"Полное наименование компании Лицензиата\"  \\* MERGEFORMAT "}));
        }

        [TestMethod]
        [DeploymentItem ("DocxComplexField2.xml")]
        public void TestMethod8 ()
        {
            var xdoc = new XmlDocument ();
            xdoc.Load ("DocxComplexField2.xml");
            var fields = Docx.GetFields (xdoc);
            var instrText = fields.Single().InstrText;

            Assert.AreEqual(" MERGEFIELD  Продукт  \\* MERGEFORMAT ", instrText);
        }

        [TestMethod]
        [DeploymentItem ("DocxSimpleField.xml")]
        public void SimpleField ()
        {
            var xdoc = new XmlDocument ();
            xdoc.Load ("DocxSimpleField.xml");
            var fields = Docx.GetFields (xdoc);
            var instrText = fields.Single().InstrText;

            Assert.AreEqual(" MERGEFIELD  \"№ паспорта Лицензиара\"  \\* MERGEFORMAT ", instrText);
        }

        [TestMethod]
        [DeploymentItem ("DocxSimpleField.xml")]
        public void SimpleFieldSetValue ()
        {
            var xdoc = new XmlDocument ();
            xdoc.Load ("DocxSimpleField.xml");
            var fields = Docx.GetFields (xdoc);
            var field = fields.Single();
            field.Value = "new value";

            Assert.AreEqual("new value", field.Value);
        }

        [TestMethod]
        [DeploymentItem ("Dogovor.xml")]
        public void Dogovor ()
        {
            var xdoc = new XmlDocument ();
            xdoc.Load ("Dogovor.xml");
            var fields = Docx.GetFields (xdoc);
            const string instrText = " MERGEFIELD  \"Цена (цифрами и прописью)\"  \\* MERGEFORMAT ";
            var field = fields.Single(f => f.InstrText == instrText);
            field.Value = "123456";

            var buffer = new MemoryStream();
            xdoc.Save(buffer);
            var xdoc2 = new XmlDocument ();
            buffer.Position = 0;
            xdoc2.Load(buffer);

            var fields2 = Docx.GetFields (xdoc);
            var field2 = fields2.Single(f => f.InstrText == instrText);
            
            Assert.AreEqual("123456", field2.Value);
        }

        [TestMethod]
        [DeploymentItem ("DocxComplexField.xml")]
        public void TestMethod7 ()
        {
            var xdoc = new XmlDocument ();
            xdoc.Load ("DocxComplexField.xml");
            var fields = Docx.GetFields (xdoc);
            fields.First ().Value = "Иванов И.И.";
        }
    }
}