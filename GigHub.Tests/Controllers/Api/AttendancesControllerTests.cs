using FluentAssertions;
using GigHub.Controllers.Api;
using GigHub.Core;
using GigHub.Core.Dtos;
using GigHub.Core.Models;
using GigHub.Core.Repositories;
using GigHub.Tests.Extensions;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using System.Web.Http.Results;
using System.Data;
using System.Data.OleDb;
using System;
using System.Configuration;
using GigHub.Persistence.Repositories;

namespace GigHub.Tests.Controllers.Api
{
    [TestClass]
    public class AttendancesControllerTests
    {
        private AttendancesController _controller;
        private Mock<IAttendanceRepository> _mockRepository;
        private string _userId;
        private int _gigId;
   
        [TestInitialize]
        public void TestInitialize()
        {
            _mockRepository = new Mock<IAttendanceRepository>();

            var mockUoW = new Mock<IUnitOfWork>();
            mockUoW.SetupGet(u => u.Attendances).Returns(_mockRepository.Object);

            _controller = new AttendancesController(mockUoW.Object);
            _userId = "87aa1ce5-a727-4044-b00c-4014a23e8912";
            _controller.MockCurrentUser(_userId, "user1@domain.com"); 
           _gigId = 1;
        }

        [TestMethod]
        public void Attend_UserAttendingAGigForWhichHeHasAnAttendance_ShouldReturnBadRequest()
        {
            var attendance = new Attendance();
            _mockRepository.Setup(r => r.GetAttendance(1, _userId)).Returns(attendance);

            var result = _controller.Attend(new AttendanceDto { GigId = 1 });

            result.Should().BeOfType<BadRequestErrorMessageResult>();
        }

        [TestMethod]
        public void Attend_ValidRequest_ShouldReturnOk()
        {
            var result = _controller.Attend(new AttendanceDto { GigId = 1 });

            result.Should().BeOfType<OkResult>();
        }

        [TestMethod]
        public void DeleteAttendance_NoAttendanceWithGivenIdExists_ShouldReturnNotFound()
        {
            //Arrange
            var attendance = new Attendance();
            attendance = null;
            _mockRepository.Setup(r => r.GetAttendance(_gigId, _userId)).Returns(attendance);

            //Act
            var result = _controller.DeleteAttendance(_gigId);

            //Assert
            _mockRepository.VerifyAll();   
            result.Should().BeOfType<NotFoundResult>();
        }

        [TestMethod]
        public void DeleteAttendance_ValidRequest_ShouldReturnOk()
        {
            //Arrange
            DataTable dt = GetValuesFromExcelFile();
            Attendance attendance = new Attendance();
            attendance.AttendeeId = dt.Rows[0]["AttendeeId"].ToString();
            attendance.GigId = Int32.Parse(dt.Rows[0]["GigId"].ToString());
            _mockRepository.Setup(r => r.GetAttendance(_gigId, _userId)).Returns(attendance);

            //Act
            var result = _controller.DeleteAttendance(_gigId);

            //Assert
            Assert.AreEqual(_gigId, attendance.GigId);
            Assert.AreEqual(_userId, attendance.AttendeeId);
            _mockRepository.VerifyAll();

            result.Should().BeOfType<OkNegotiatedContentResult<int>>();
        }

        [TestMethod]
        public void DeleteAttendance_ValidRequest_ShouldReturnTheIdOfDeletedAttendance()
        {
            //Arrange
            DataTable dt = GetValuesFromExcelFile();
            Attendance attendance = new Attendance();
            attendance.AttendeeId = dt.Rows[0]["AttendeeId"].ToString();
            attendance.GigId = Int32.Parse(dt.Rows[0]["GigId"].ToString());
            _mockRepository.Setup(r => r.GetAttendance(_gigId, _userId)).Returns(attendance);

            //Act
            var result = (OkNegotiatedContentResult<int>)_controller.DeleteAttendance(_gigId);
     
            //Assert
            Assert.AreEqual(_gigId, attendance.GigId);
            Assert.AreEqual(_userId, attendance.AttendeeId);
            _mockRepository.VerifyAll();   
            result.Content.Should().Be(_gigId);
        }

        public static DataTable GetValuesFromExcelFile()
        {
            //Connect to Excel file via OLEDB
            string xlConnStr = ConfigurationManager.ConnectionStrings["ExcelConnString"].ConnectionString;
            var xlConn = new OleDbConnection(xlConnStr);

            //Extract Data from Excel file
            var da = new OleDbDataAdapter("SELECT AttendeeId,GigId FROM [Sheet1$]", xlConn);
            var xlDT = new DataTable();
            da.Fill(xlDT);
         
            return xlDT;
        }



    }
}
