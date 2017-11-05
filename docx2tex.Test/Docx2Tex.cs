using System;
using System.Text;
using System.Collections.Generic;

using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using System.Diagnostics;

namespace docx2tex.Test
{
    /// <summary>
    /// Summary description for Docx2Tex
    /// </summary>
    [TestClass]
    public class Docx2Tex
    {
        public Docx2Tex()
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
        public void PreparedDocTest()
        {
            Do("test", "example");
        }

        [TestMethod]
        public void PreparedCSCSTest()
        {
            Do("testcscs", "regression-test-final2");
        }

        [TestMethod]
        public void PreparedTUGTest()
        {
            Do("testtug", "docx2tex");
        }

        [TestMethod]
        public void PreparedMACSTest()
        {
            Do("testmacs", "macs_secdistr_draft");
        }

        [TestMethod]
        public void PreparedEquationsest()
        {
            Do("testeq", "eqtest");
        }

        private void Do(string dir, string fileName)
        {
            string solutionDir = new DirectoryInfo(testContextInstance.TestDir).Parent.Parent.FullName;
            string prog = Path.Combine(solutionDir, "out\\docx2tex.exe");
            string testProjDir = Path.Combine(solutionDir, "docx2tex.Test");
            string baseInputDir = Path.Combine(testProjDir, "Input");
            string baseOutputDir = Path.Combine(testProjDir, "Output");
            string baseExpectedDir = Path.Combine(testProjDir, "Expected");

            string inputDir = Path.Combine(baseInputDir, dir);
            string outputDir = Path.Combine(baseOutputDir, dir);
            string expectedDir = Path.Combine(baseExpectedDir, dir);

            string outputMediaDir = Path.Combine(outputDir, "media");
            string expectedMediaDir = Path.Combine(expectedDir, "media");

            string inputFile = Path.Combine(inputDir, fileName + ".docx");
            string outputFile = Path.Combine(outputDir, fileName + ".tex");
            string expectedFile = Path.Combine(expectedDir, fileName + ".tex");

            if (Directory.Exists(outputMediaDir))
            {
                Directory.Delete(outputMediaDir, true);
            }
            if (File.Exists(outputFile))
            {
                File.Delete(outputFile);
            }

            ProcessStartInfo psi = new ProcessStartInfo();
            psi.FileName = prog;
            psi.Arguments = inputFile + " " + outputFile;

            Process proc = Process.Start(psi);
            proc.WaitForExit();

            Assert.AreEqual(File.ReadAllText(expectedFile), File.ReadAllText(outputFile));

            FileInfo[] outFiles = new DirectoryInfo(outputMediaDir).GetFiles();
            FileInfo[] expFiles = new DirectoryInfo(expectedMediaDir).GetFiles();

            Assert.AreEqual(expFiles.Length, outFiles.Length);
        }
    }
}
