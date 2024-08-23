using ExceLib;

namespace Tests;

public class Tests
{
    ExcelApp? _excel;
    string? _dataDir;

    [SetUp]
    public void Setup()
    {
        _excel = new ExcelApp(false);
        _dataDir = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), @"..\..\..\data\"));
    }

    [Test]
    public void TestOpenFile()
    {
        var testFile = _dataDir + "test.xlsx";
        if (!File.Exists(testFile)) 
        {
            throw new FileNotFoundException(testFile);
        }
        _excel?.OpenDoc(testFile);
        Assert.DoesNotThrow(() => _excel?.OpenDoc(testFile));
    }

    [Test]
    [TestCase("B2", "Просто текст")]
    [TestCase("C2", "5")]
    [TestCase("D2", " ")]
    [TestCase("E2", null)]
    [TestCase("F2", "01.01.2001 0:00:00")]
    public void TestGetValue(string cell, string? expected)
    {
        var testFile = _dataDir + "test.xlsx";
        if (!File.Exists(testFile))
        {
            throw new FileNotFoundException(testFile);
        }
        _excel?.OpenDoc(testFile);
        var actual = _excel?.GetValue(cell)?.ToString();
        Assert.That(actual, Is.EqualTo(expected));
    }

    [Test]
    [TestCase("B2", "Просто текст")]
    [TestCase("C2", "5")]
    [TestCase("D2", " ")]
    [TestCase("E2", null)]
    [TestCase("F2", "36892")]
    public void TestGetValue2(string cell, string? expected)
    {
        var testFile = _dataDir + "test.xlsx";
        if (!File.Exists(testFile))
        {
            throw new FileNotFoundException(testFile);
        }
        _excel?.OpenDoc(testFile);
        var actual = _excel?.GetValue2(cell)?.ToString();
        Assert.That(actual, Is.EqualTo(expected));
    }

    [TearDown]
    public void TearDown()
    {
        _excel?.Quit();
    }

}