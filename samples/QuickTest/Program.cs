using VibeCoderToolkit.Documents;
using VibeCoderToolkit.Documents.Excel;

var testDir = Path.Combine(Path.GetTempPath(), "vibecoder_verify");
Directory.CreateDirectory(testDir);

int passed = 0, failed = 0;

void Check(bool cond, string msg)
{
    if (cond) { passed++; Console.WriteLine($"  ✅ {msg}"); }
    else { failed++; Console.WriteLine($"  ❌ {msg}"); }
}

Console.WriteLine("🔬 VibeCoderToolkit.Documents — verification test\n");

// ──── Data ────
var products = new List<Product>
{
    new() { Name = "Keyboard", Price = 899.50m, Stock = 12 },
    new() { Name = "Mouse", Price = 349m, Stock = 45 },
    new() { Name = "Monitor", Price = 2499m, Stock = 3 },
};

// ──── CSV ────
Console.WriteLine("📄 CSV:");
var csvPath = Path.Combine(testDir, "test.csv");
DocumentWriter.Write(csvPath, products);
var csvRead = DocumentReader.Read<Product>(csvPath);
Check(csvRead.Count == 3, "Write and read 3 rows");
Check(csvRead[0].Name == "Keyboard", "First row — name");
Check(csvRead[1].Price == 349m, "Second row — price");
Check(csvRead[2].Stock == 3, "Third row — stock");

// ──── JSON ────
Console.WriteLine("📄 JSON:");
var jsonPath = Path.Combine(testDir, "test.json");
DocumentWriter.Write(jsonPath, products);
var jsonRead = DocumentReader.Read<Product>(jsonPath);
Check(jsonRead.Count == 3, "Write and read 3 rows");
Check(jsonRead[0].Name == "Keyboard", "First row — name");

// ──── Excel (unencrypted) ────
Console.WriteLine("📄 Excel (unencrypted):");
var xlsxPath = Path.Combine(testDir, "test.xlsx");
DocumentWriter.Write(xlsxPath, products);
var xlsxRead = DocumentReader.Read<Product>(xlsxPath);
Check(xlsxRead.Count == 3, "Write and read 3 rows");
Check(xlsxRead[0].Name == "Keyboard", "First row — name");
Check(xlsxRead[2].Price == 2499m, "Third row — price");

// ──── Excel (encrypted with password) ────
Console.WriteLine("📄 Excel (encrypted, password: \"secret\"):");
var encPath = Path.Combine(testDir, "encrypted.xlsx");

// Write encrypted: open unencrypted → save with password
using (var doc = EncryptedExcelDocument.Open(xlsxPath))
{
    doc.SaveAs(encPath, "secret");
}
Check(EncryptedExcelDocument.IsEncrypted(encPath), "File is encrypted (OLE2 format)");

// Try to read WITHOUT password → should throw
try
{
    EncryptedExcelDocument.Open(encPath);
    Check(false, "Opened encrypted file without password — should have thrown");
}
catch (InvalidOperationException)
{
    Check(true, "Open without password → InvalidOperationException (as expected)");
}

// Try to read with WRONG password → should throw
try
{
    EncryptedExcelDocument.Open(encPath, "wrongpassword");
    Check(false, "Opened with wrong password — should have thrown");
}
catch (Exception)
{
    Check(true, "Open with wrong password → error (as expected)");
}

// Read with CORRECT password → should work
using (var encDoc = EncryptedExcelDocument.Open(encPath, "secret"))
{
    var names = encDoc.GetSheetNames();
    Check(names.Length >= 1, $"Has {names.Length} sheet(s)");

    var val = encDoc.GetCellValue(names[0], 2, 1)?.ToString();
    Check(val == "Keyboard", $"Row 2, col 1 = '{val}' (expected 'Keyboard')");

    var price = encDoc.GetCellValue(names[0], 3, 2);
    Check(Convert.ToDecimal(price) == 349m, $"Row 3, col 2 = {price} (expected 349)");

    // Modify a cell and resave with a new password
    encDoc.SetCellValue(names[0], 2, 3, 999);
    encDoc.SaveAs(encPath, "newpassword");
}
Check(EncryptedExcelDocument.IsEncrypted(encPath), "File still encrypted after modification");

// Reopen with new password
using (var reopened = EncryptedExcelDocument.Open(encPath, "newpassword"))
{
    var sheet = reopened.GetSheetNames()[0];
    var modified = reopened.GetCellValue(sheet, 2, 3);
    Check(Convert.ToInt32(modified) == 999, $"Modified cell = {modified} (expected 999)");

    var unchanged = reopened.GetCellValue(sheet, 1, 2)?.ToString();
    Check(unchanged == "Price", $"Unchanged header = '{unchanged}' (expected 'Price')");
}

Console.WriteLine($"📂 Unencrypted: {xlsxPath}");
Console.WriteLine($"📂 Encrypted:  {encPath}");

// ──── Cleanup (keep Excel files for inspection) ────
try { File.Delete(csvPath); File.Delete(jsonPath); } catch { }

Console.WriteLine($"\n━━━ {passed}/{passed + failed} tests passed ━━━");
if (failed > 0)
{
    Console.WriteLine("❌ Some tests failed!");
    Environment.Exit(1);
}
else
{
    Console.WriteLine("🎉 All tests pass!");
}

public class Product
{
    public string Name { get; set; } = "";
    public decimal Price { get; set; }
    public int Stock { get; set; }
}
