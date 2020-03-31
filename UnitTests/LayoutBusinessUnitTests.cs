using AzulBusiness;
using Microsoft.VisualStudio.TestTools.UnitTesting;


namespace UnitTests
{

    [TestClass]
    public class LayoutBusinessUnitTests
    {
        private string _SYSTEM_SKYSPEED_MOCK = "ss";
        private string _SYSTEM_ANAC_MOCK = "ac";
        private string _SYSTEM_REDEMET_MOCK = "rm";
        private string _SYSTEM_FLIGHTUTILITIES_MOCK = "fo";
        private string _SYSTEM_MIDLEWARE_MOCK = "md";
        private string _PNR1_MOCK = "PNR1";
        private string _PNR2_MOCK = "PNR2";
        private string _PNR3_MOCK = "PNR3";


        [TestMethod]
        public void GetTableTitle_ReturnTableTitle_Sucess()
        {
            //Tests correct result of dictionary tabletitles from layout

            //Arrange
            string _flow = "bd";
            string _expectedResult = "ESCLARECIMENTO_SKYSPEED_BREAKDOWN_PNR1";

            //Arrange
            LayoutBusiness layoutBusiness = new LayoutBusiness();
            var result = layoutBusiness.GetTableTitle(_SYSTEM_SKYSPEED_MOCK, _flow, _PNR1_MOCK);


            //Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(_expectedResult, result);
        }


        [TestMethod]
        public void GetTableTitle_ReturnTableTitleMidlewareSystem_Sucess()
        {

            //Tests correct result of dictionary tabletitles from layout with middleware system data 

            //Arrange
            string _flow = "default";
            string _expectedResult = "ESCLARECIMENTO_TUDOAZUL_MIDDLEWARE_PNR2";

            //Arrange
            LayoutBusiness layoutBusiness = new LayoutBusiness();
            var result = layoutBusiness.GetTableTitle(_SYSTEM_MIDLEWARE_MOCK, _flow, _PNR2_MOCK);


            //Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(_expectedResult, result);
        }

        [TestMethod]
        public void GetTableTitle_ReturnTableTitleRedemetSystem_Sucess()
        {

            //Tests correct result of dictionary tabletitles from layout with redemet system data 
            //Arrange
            string _flow = "default";
            string _expectedResult = "ESCLARECIMENTO_REDEMET_PNR3";

            //Arrange
            LayoutBusiness layoutBusiness = new LayoutBusiness();
            var result = layoutBusiness.GetTableTitle(_SYSTEM_REDEMET_MOCK, _flow, _PNR3_MOCK);


            //Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(_expectedResult, result);
        }

        [TestMethod]
        public void GetTableTitle_ReturnTableTitleFlightUtilitiesSystem_Sucess()
        {

            //Tests correct result of dictionary tabletitles from layout with flightUtilities system data 
            //Arrange
            string _flow = "default";
            string _expectedResult = "ESCLARECIMENTO_FLIGHT_UTILITIES_PNR1";

            //Arrange
            LayoutBusiness layoutBusiness = new LayoutBusiness();
            var result = layoutBusiness.GetTableTitle(_SYSTEM_FLIGHTUTILITIES_MOCK, _flow, _PNR1_MOCK);


            //Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(_expectedResult, result);
        }


        [TestMethod]
        public void GetTableTitle_ReturnTableTitleAnacSystem_Sucess()
        {

            //Tests correct result of dictionary tabletitles from layout with anac system data 
            //Arrange
            string _flow = "default";
            string _expectedResult = "ESCLARECIMENTO_ANAC_PNR1";

            //Arrange
            LayoutBusiness layoutBusiness = new LayoutBusiness();
            var result = layoutBusiness.GetTableTitle(_SYSTEM_ANAC_MOCK, _flow, _PNR1_MOCK);


            //Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(_expectedResult, result);
        }

        [TestMethod]
        public void GetTableTitle_ReturnTableTitle_fail()
        {
            //Tests invalids results when wrong parameters are passed 

            //Arrange
            string _flow = "ps";
            string _expectedResult = "ESCLARECIMENTO_SKYSPEED_BREAKDOWN_PNR1";

            //Arrange
            LayoutBusiness layoutBusiness = new LayoutBusiness();
            var result = layoutBusiness.GetTableTitle(_SYSTEM_SKYSPEED_MOCK, _flow, _PNR1_MOCK);

            //Assert
            Assert.IsNotNull(result);
            Assert.AreNotEqual(_expectedResult, result);
        }



        [TestMethod]
        public void GetTableTitle_ReturnTableTitleWithEmptyParameters_fail()
        {
            //Tests method when empty strings are passed as parameters 

            //Arrange
            string _flow = string.Empty;
            string _expectedResult = "Table Not Found";

            //Arrange
            LayoutBusiness layoutBusiness = new LayoutBusiness();
            var result = layoutBusiness.GetTableTitle(_SYSTEM_SKYSPEED_MOCK, _flow, _PNR1_MOCK);

            //Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(_expectedResult, result);
        }


        [TestMethod]
        public void GetTableTitle_WithWrongParameters_fail()
        {
            //Tests method when wrong parameters ares passed

            //Arrange
            string _systemMock = "system no even exists";
            string _flowMock = "no flow!! total caos instead";
            string _pnrxMock = "PNRX";
            string _expectedResult = "Table Not Found";

            //Arrange
            LayoutBusiness layoutBusiness = new LayoutBusiness();
            var result = layoutBusiness.GetTableTitle(_systemMock, _flowMock, _pnrxMock);


            //Assert
            Assert.IsNotNull(result);
            Assert.AreEqual(_expectedResult, result);

        }
    }
}
