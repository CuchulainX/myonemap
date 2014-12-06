using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Mindjet.MindManager.Interop;
using myonemap.onenote;
using onemapApp = Mindjet.MindManager.Interop.Application;

namespace myonemap.Tests
{
    /// <summary>
    /// Summary description for UnitTest1
    /// </summary>
    [TestClass]
    public class UnitTest1
    {
        private onemapApp mm = new onemapApp();

        public UnitTest1()
        {

        }

        private int mmRunning
        {
            get { return mm.hWnd; }
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
            var temp =
                @"onenote:///H:\mybase\onenote\notebooks\asp.net\asp.net\testsection1.one#tttooo&section-id={0BEEFA4B-1CE9-4EA5-9483-E48211243F16}&page-id={166B3538-91ED-4CC5-AFC5-6F9B0EBCCA61}&end&mm-guid={4f3850ca-29b5-4b2b-9aca-947c31384868}@onenote:///H:\mybase\onenote\notebooks\asp.net\asp.net\testsection1.one#Oaa&section-id={0BEEFA4B-1CE9-4EA5-9483-E48211243F16}&page-id={DFC69630-51C3-49C9-8E33-02653A8CCF93}&end&mm-guid={c9e1e5bc-e2c1-480a-bce2-ab53973ca93e}@onenote:///H:\mybase\onenote\notebooks\asp.net\asp.net\testsection1.one#Obb&section-id={0BEEFA4B-1CE9-4EA5-9483-E48211243F16}&page-id={B1B3514C-5695-4416-B738-915681D71189}&end&mm-guid={69871266-e520-44ff-a0ce-86ffff1524bc}@onenote:///H:\mybase\onenote\notebooks\asp.net\asp.net\testsection1.one#Occ23&section-id={0BEEFA4B-1CE9-4EA5-9483-E48211243F16}&page-id={2B8003F0-E949-42CE-8CAC-769557087002}&end&mm-guid={c66720fc-26cf-4293-898c-1dbd38a54ca6}
onenote:///H:\mybase\onenote\notebooks\asp.net\asp.net\testsection1.one#Odd&section-id={0BEEFA4B-1CE9-4EA5-9483-E48211243F16}&page-id={51C1EAFE-168E-4B03-8629-2A792A1FA99C}&end&mm-guid={b6db3e91-b974-4dce-9648-0b51e1a6f995}@onenote:///H:\mybase\onenote\notebooks\asp.net\asp.net\testsection1.one#section-id={0BEEFA4B-1CE9-4EA5-9483-E48211243F16}&page-id={B6DF9CF1-6F07-4AEB-98C6-D909D60FAE83}&end&mm-guid={79c8db69-4e7d-4fb1-921b-aa0a6be0f479}@onenote:///H:\mybase\onenote\notebooks\asp.net\asp.net\testsection1.one#section-id={0BEEFA4B-1CE9-4EA5-9483-E48211243F16}&page-id={3B6715AF-B557-4FB4-9E17-CB0ED1160F68}&end&mm-guid={9406af74-8ed0-45f7-8846-9ec2802d602b}@onenote:///H:\mybase\onenote\notebooks\asp.net\asp.net\testsection1.one#Fdsfdsfds&section-id={0BEEFA4B-1CE9-4EA5-9483-E48211243F16}&page-id={2200CC75-9986-48CE-80FE-8FE4F276AAF1}&end&mm-guid={72571467-a8ac-40dd-8f43-84052857676a}@onenote:///H:\mybase\onenote\notebooks\asp.net\asp.net\testsection1.one#Main%20Topic&section-id={0BEEFA4B-1CE9-4EA5-9483-E48211243F16}&page-id={A2C520A7-346B-43D5-B907-CB2B38F7DF3D}&end&mm-guid={15467fe1-6ef5-4e67-8369-e92958cdce29}";
            string[] items = temp.Split('@');
            foreach (var it in items)
            {
                Assert.IsTrue((OnenoteUtils.ValidateOnenoteLink(it) != "success"));
            }
        }

        [TestMethod]
        public void MyMethod()
        {
            var a = mmRunning;
        }
    }
}
