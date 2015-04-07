using System;
using System.Text;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Configuration;
using DocumentFormat.OpenXml.Packaging;
using OpenXml.Word.Extensions;
using System.IO;

namespace OpenXml.Word.Tests
{
    /// <summary>
    /// Summary description for DocxTemplateTeste01
    /// </summary>
    [TestClass]
    public class DocxTemplateTeste01
    {
        public DocxTemplateTeste01()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        private TestContext testContextInstance;

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }

        #region Additional test attributes
        //
        // You can use the following additional attributes as you write your tests:
        //
        // Use ClassInitialize to run code before running the first test in the class
        // [ClassInitialize()]
        // public static void MyClassInitialize(TestContext testContext) { }
        //
        // Use ClassCleanup to run code after all tests in a class have run
        // [ClassCleanup()]
        // public static void MyClassCleanup() { }
        //
        // Use TestInitialize to run code before running each test 
        // [TestInitialize()]
        // public void MyTestInitialize() { }
        //
        // Use TestCleanup to run code after each test has run
        // [TestCleanup()]
        // public void MyTestCleanup() { }
        //
        #endregion

        [TestMethod]
        public void TestMethod1()
        {
            //
            // TODO: Add test logic here
            //
        }

        [TestMethod]
        public void TestarGeracaoDocx()
        {
            Dictionary<String, String> substituicoes =
                new Dictionary<String, String>();
            substituicoes["#NOME_CLIENTE#"] =
                "João da Silva";
            //substituicoes["#ENDERECO_CLIENTE#"] =
            //  "Avenida Paulista, 950 - São Paulo - SP";
            substituicoes.Add("#ENDERECO_CLIENTE#", "Avenida Paulista, 950 - São Paulo - SP");
            substituicoes["#NOME_ASSINATURA#"] =
                "Pedro Oliveira";

            //string caminhoTemplate =
            //  ConfigurationManager.AppSettings.Get("caminhoTemplate");
            string caminhoTemplate =
                ConfigurationManager.AppSettings["CaminhoArquivoTemplate"];
            string caminhoArquivoDestino =
                ConfigurationManager.AppSettings["DiretorioArquivoTeste"] +
                "Teste_" + DateTime.Now.ToString("dd-MM-yyyy_HH'h'mm'min'ss's'") + ".docx";

            DocxTemplate.CriarNovoDocumento(
                caminhoTemplate,
                caminhoArquivoDestino,
                substituicoes);

            Assert.IsTrue(File.Exists(caminhoArquivoDestino));
        }
    }
}
