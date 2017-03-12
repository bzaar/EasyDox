using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EasyDox.Tests
{
    [TestClass]
    public class EngineEvalTests
    {
        [TestMethod]
        public void Literal()
        {
            Assert.AreEqual("действующий", new Engine().Eval(" \"действующий\"", new Properties()));
        }

        [TestMethod]
        public void Property()
        {
            var properties = new Properties {{"Подписант","Иванов В.П."}};

            Assert.AreEqual("Иванов В.П.", new Engine().Eval("Подписант", properties));
        }

        [TestMethod]
        public void PropertyWithInnerSpace()
        {
            var properties = new Properties {{"ФИО Покупателя","Иванов В.П."}};

            Assert.AreEqual("Иванов В.П.", new Engine().Eval("ФИО Покупателя", properties));
        }

        [TestMethod]
        public void PropertyWithTrailingSpace()
        {
            var properties = new Properties {{"ФИО Покупателя","Иванов В.П."}};

            Assert.AreEqual("Иванов В.П.", new Engine().Eval("ФИО Покупателя ", properties));
        }

        [TestMethod]
        public void FunctionOfProperty()
        {
            var properties = new Properties {{"ФИО Покупателя","Иванов В.П."}};

            var functions = new Dictionary <string, IFuncN> 
            {
                {"родительный", new Func1 (s => s == "Иванов В.П." ? "Иванова В.П." : "")}
            };

            Assert.AreEqual("Иванова В.П.", new Engine(functions).Eval("ФИО Покупателя (родительный)", properties));
        }

        [TestMethod]
        public void FunctionOfTwoProperties()
        {
            var properties = new Properties {{"Подписант","Иванова В.П."}};

            var functions = new Dictionary <string, IFuncN> 
            {
                {"род как у", new Func2 ((s,r) => s == "действующий" && r == "Иванова В.П." ? "действующая" : "")}
            };

            Assert.AreEqual("действующая", new Engine(functions).Eval("\"действующий\" (род как у Подписант)", properties));
        }

        [TestMethod]
        public void FunctionOfTwoPropertiesCalledWithOneArgument()
        {
            var properties = new Properties();

            var functions = new Dictionary <string, IFuncN> 
            {
                {"род как у", new Func2((s,r) => "")}
            };

            Assert.AreEqual(null, new Engine(functions).Eval("\"действующий\" (род как у )", properties));
        }

        [TestMethod]
        public void InvalidFunctionName()
        {
            var properties = new Properties {{"Подписант","Иванова В.П."}};

            var functions = new Dictionary <string, IFuncN> 
            {
                {"род как у", new Func2((s,r) => "")}
            };

            Assert.AreEqual(null, new Engine(functions).Eval("\"действующий\" (род Подписант)", properties));
        }

        [TestMethod]
        public void FunctionOfFunctionOfTwoProperties()
        {
            var properties = new Properties {{"Подписант","Иванова В.П."}};

            var functions = new Dictionary <string, IFuncN> 
            {
                {"род как у", new Func2((s,r) => s == "действующий" && r == "Иванова В.П." ? "действующая" : "")},
                {"родительный", new Func1(s => s == "действующая" ? "действующей" : "")}
            };

            Assert.AreEqual("действующей", new Engine(functions).Eval("\"действующий\" (род как у Подписант) (родительный)", properties));
        }

        [TestMethod]
        public void FunctionOfTwoLiterals()
        {
            var properties = new Properties();

            var functions = new Dictionary <string, IFuncN> 
            {
                {"род как у", new Func2((s,r) => s == "беременный" && r == "женщина" ? "беременная" : "")},
            };

            Assert.AreEqual("беременная", new Engine(functions).Eval("\"беременный\" (род как у \"женщина\")", properties));
        }
    }
}