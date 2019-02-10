using System;
using Xunit;
using PSWordCloud;

namespace Cmdlet.Tests
{
    public class WCUtilsTests
    {
        [Fact]
        public void Test_ToRadians()
        {
            //Assert.Equal(KnownResult, TestResult);
            
            Assert.Equal(0,((float)0).ToRadians());
        }
    }
}
