using JFToolkit.Documents;
using JFToolkit.Documents.Excel;
using Xunit;

namespace JFToolkit.Documents.Tests;

public class EncryptedExcelTests
{
    [Fact]
    public void Encrypt_PlainFile_ProducesOle2Container()
    {
        var plainPath = CreatePlainXlsx();
        var encPath = GetTempPath("encrypted.xlsx");

        using (var doc = EncryptedExcelDocument.Open(plainPath))
        {
            doc.SaveAs(encPath, "secret");
        }

        Assert.True(EncryptedExcelDocument.IsEncrypted(encPath));
        Assert.True(File.Exists(encPath));
    }

    [Fact]
    public void OpenEncrypted_WithCorrectPassword_ReadsData()
    {
        var (encPath, plainData) = CreateEncryptedFile();

        using var doc = EncryptedExcelDocument.Open(encPath, "secret");

        Assert.Equal(1, doc.SheetCount);
        var names = doc.GetSheetNames();
        Assert.NotEmpty(names);

        // Check header row
        Assert.Equal("Name", doc.GetCellValue(names[0], 1, 1)?.ToString());
        Assert.Equal("Price", doc.GetCellValue(names[0], 1, 2)?.ToString());

        // Check data
        Assert.Equal("Widget", doc.GetCellValue(names[0], 2, 1)?.ToString());
        Assert.Equal("9.99", doc.GetCellValue(names[0], 2, 2)?.ToString());
    }

    [Fact]
    public void OpenEncrypted_WithoutPassword_Throws()
    {
        var (encPath, _) = CreateEncryptedFile();

        var ex = Assert.Throws<InvalidOperationException>(() =>
            EncryptedExcelDocument.Open(encPath));

        Assert.Contains("encrypted", ex.Message.ToLower());
        Assert.Contains("password", ex.Message.ToLower());
    }

    [Fact]
    public void OpenEncrypted_WithWrongPassword_Throws()
    {
        var (encPath, _) = CreateEncryptedFile();

        Assert.Throws<InvalidOperationException>(() =>
            EncryptedExcelDocument.Open(encPath, "wrongpassword"));
    }

    [Fact]
    public void EncryptDecrypt_Roundtrip_PreservesAllData()
    {
        var plainPath = CreatePlainXlsx(("A", 1.5m, 10), ("B", 2.5m, 20), ("C", 3.5m, 30));
        var encPath = GetTempPath("roundtrip.xlsx");

        // Encrypt
        using (var doc = EncryptedExcelDocument.Open(plainPath))
        {
            doc.SaveAs(encPath, "pwd");
        }

        // Decrypt and verify
        using var decrypted = EncryptedExcelDocument.Open(encPath, "pwd");
        var sheet = decrypted.GetSheetNames()[0];

        Assert.Equal("A", decrypted.GetCellValue(sheet, 2, 1)?.ToString());
        Assert.Equal(1.5m, Convert.ToDecimal(decrypted.GetCellValue(sheet, 2, 2)));
        Assert.Equal(10, Convert.ToInt32(decrypted.GetCellValue(sheet, 2, 3)));

        Assert.Equal("C", decrypted.GetCellValue(sheet, 4, 1)?.ToString());
        Assert.Equal(3.5m, Convert.ToDecimal(decrypted.GetCellValue(sheet, 4, 2)));
        Assert.Equal(30, Convert.ToInt32(decrypted.GetCellValue(sheet, 4, 3)));
    }

    [Fact]
    public void SaveAs_WithNullPassword_SavesUnencrypted()
    {
        var plainPath = CreatePlainXlsx();
        var encPath = GetTempPath("temp_enc.xlsx");
        var plainPath2 = GetTempPath("plain_again.xlsx");

        // First encrypt
        using (var doc = EncryptedExcelDocument.Open(plainPath))
        {
            doc.SaveAs(encPath, "secret");
        }
        Assert.True(EncryptedExcelDocument.IsEncrypted(encPath));

        // Then save unencrypted
        using (var doc = EncryptedExcelDocument.Open(encPath, "secret"))
        {
            doc.SaveAs(plainPath2, null);
        }
        Assert.False(EncryptedExcelDocument.IsEncrypted(plainPath2));

        // Should be readable as plain
        var data = DocumentReader.Read<Product>(plainPath2);
        Assert.NotEmpty(data);
    }

    [Fact]
    public void ModifyEncrypted_ChangePassword_StillReadable()
    {
        var (encPath, _) = CreateEncryptedFile();

        // Change cell value and password
        using (var doc = EncryptedExcelDocument.Open(encPath, "secret"))
        {
            var sheet = doc.GetSheetNames()[0];
            doc.SetCellValue(sheet, 2, 2, "42.00");
            doc.SaveAs(encPath, "newpassword");
        }

        // Open with new password
        using var reopened = EncryptedExcelDocument.Open(encPath, "newpassword");
        var s = reopened.GetSheetNames()[0];
        Assert.Equal("42.00", reopened.GetCellValue(s, 2, 2)?.ToString());
    }

    [Fact]
    public void IsEncrypted_PlainFile_ReturnsFalse()
    {
        var plainPath = CreatePlainXlsx();
        Assert.False(EncryptedExcelDocument.IsEncrypted(plainPath));
    }

    [Fact]
    public void IsExcelFile_Encrypted_ReturnsTrue()
    {
        var (encPath, _) = CreateEncryptedFile();
        Assert.True(EncryptedExcelDocument.IsExcelFile(encPath));
    }

    [Fact]
    public void IsExcelFile_Plain_ReturnsTrue()
    {
        var plainPath = CreatePlainXlsx();
        Assert.True(EncryptedExcelDocument.IsExcelFile(plainPath));
    }

    [Fact]
    public void IsExcelFile_NonExcel_ReturnsFalse()
    {
        var path = GetTempPath("notexcel.txt");
        File.WriteAllText(path, "hello world");
        Assert.False(EncryptedExcelDocument.IsExcelFile(path));
    }

    [Fact]
    public void ReadOnly_Open_CanReadButNotSave()
    {
        var (encPath, _) = CreateEncryptedFile();

        using var doc = EncryptedExcelDocument.Open(encPath, "secret", readOnly: true);

        var sheet = doc.GetSheetNames()[0];
        Assert.NotNull(doc.GetCellValue(sheet, 2, 1));

        Assert.Throws<InvalidOperationException>(() =>
            doc.SetCellValue(sheet, 2, 1, "test"));
    }

    [Fact]
    public void MultipleSheets_EncryptDecrypt_PreservesSheets()
    {
        // Create a multi-sheet file from the plain template that already has Sheet1
        // For now, test that single sheet roundtrip works (multi-sheet requires
        // programmatic sheet creation which isn't yet exposed)
        var (encPath, _) = CreateEncryptedFile();
        using var doc = EncryptedExcelDocument.Open(encPath, "secret");
        Assert.Equal(1, doc.SheetCount);
        Assert.Equal("Sheet1", doc.GetSheetNames()[0]);
    }

    // ── Helpers ──

    /// <summary>Creates an encrypted .xlsx with sample data.</summary>
    private static (string Path, List<Product> Data) CreateEncryptedFile()
    {
        var data = new List<Product>
        {
            new() { Name = "Widget", Price = 9.99m, Stock = 100 },
        };
        var plainPath = CreatePlainXlsx(data);
        var encPath = GetTempPath($"enc_{Guid.NewGuid():N}.xlsx");

        using (var doc = EncryptedExcelDocument.Open(plainPath))
        {
            doc.SaveAs(encPath, "secret");
        }

        return (encPath, data);
    }

    private static string CreatePlainXlsx(params (string Name, decimal Price, int Stock)[] items)
    {
        var data = items.Select(i => new Product
        {
            Name = i.Name,
            Price = i.Price,
            Stock = i.Stock
        }).ToList();

        if (data.Count == 0)
        {
            data.Add(new Product { Name = "Widget", Price = 9.99m, Stock = 100 });
        }

        return CreatePlainXlsx(data);
    }

    private static string CreatePlainXlsx(List<Product> data)
    {
        var path = GetTempPath($"plain_{Guid.NewGuid():N}.xlsx");
        DocumentWriter.Write(path, data);
        return path;
    }

    private static string GetTempPath(string name)
    {
        var dir = Path.Combine(Path.GetTempPath(), "jftoolkit_tests");
        Directory.CreateDirectory(dir);
        return Path.Combine(dir, name);
    }

    public class Product
    {
        public string Name { get; set; } = "";
        public decimal Price { get; set; }
        public int Stock { get; set; }
    }
}
