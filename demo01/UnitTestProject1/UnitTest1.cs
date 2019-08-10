using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using demo01;

namespace unitTest01
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            Assert.AreEqual("Hello SURYA", Program.createMessage());
                
        }
    }
}
