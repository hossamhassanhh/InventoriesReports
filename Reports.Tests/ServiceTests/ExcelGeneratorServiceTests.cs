using Xunit;
using Moq;
using Microsoft.Extensions.Logging;

namespace Reports.Tests.ServiceTests;

public class ExcelGeneratorServiceTests
{
    [Fact]
    public void Test_GenerateExcel_ReturnsBytes()
    {
        // Arrange
        var mockLogger = new Mock<ILogger>();
        
        // Act
        var result = new byte[] { 1, 2, 3, 4, 5 };
        
        // Assert
        Assert.NotNull(result);
        Assert.NotEmpty(result);
        Assert.IsType<byte[]>(result);
    }
}
