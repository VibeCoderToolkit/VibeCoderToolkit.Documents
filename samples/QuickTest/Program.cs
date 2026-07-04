using JFToolkit.Documents;
using JFToolkit.Documents.Excel;

var testDir = Path.Combine(Path.GetTempPath(), "jftoolkit_verify");
Directory.CreateDirectory(testDir);

int passed = 0, failed = 0;

void Check(bool cond, string msg)
{
    if (cond) { passed++; Console.WriteLine($"  ✅ {msg}"); }
    else { failed++; Console.WriteLine($"  ❌ {msg}"); }
}

Console.WriteLine("🔬 JFToolkit.Documents — verifiseringstest\n");

// ──── Data ────
var products = new List<Product>
{
    new() { Name = "Tastatur", Pris = 899.50m, Antall = 12 },
    new() { Name = "Mus", Pris = 349m, Antall = 45 },
    new() { Name = "Skjerm", Pris = 2499m, Antall = 3 },
};

// ──── CSV ────
Console.WriteLine("📄 CSV:");
var csvPath = Path.Combine(testDir, "test.csv");
DocumentWriter.Write(csvPath, products);
var csvRead = DocumentReader.Read<Product>(csvPath);
Check(csvRead.Count == 3, "Skriv og les 3 rader");
Check(csvRead[0].Name == "Tastatur", "Første rad — navn");
Check(csvRead[1].Pris == 349m, "Andre rad — pris");
Check(csvRead[2].Antall == 3, "Tredje rad — antall");

// ──── JSON ────
Console.WriteLine("📄 JSON:");
var jsonPath = Path.Combine(testDir, "test.json");
DocumentWriter.Write(jsonPath, products);
var jsonRead = DocumentReader.Read<Product>(jsonPath);
Check(jsonRead.Count == 3, "Skriv og les 3 rader");
Check(jsonRead[0].Name == "Tastatur", "Første rad — navn");

// ──── Excel (ukryptert) ────
Console.WriteLine("📄 Excel (ukryptert):");
var xlsxPath = Path.Combine(testDir, "test.xlsx");
DocumentWriter.Write(xlsxPath, products);
var xlsxRead = DocumentReader.Read<Product>(xlsxPath);
Check(xlsxRead.Count == 3, "Skriv og les 3 rader");
Check(xlsxRead[0].Name == "Tastatur", "Første rad — navn");
Check(xlsxRead[2].Pris == 2499m, "Tredje rad — pris");

// ──── Excel (kryptert med passord) ────
Console.WriteLine("📄 Excel (kryptert, passord: \"hemmelig\"):");
var encPath = Path.Combine(testDir, "kryptert.xlsx");

// Skriv kryptert: opne ukryptert → lagre med passord
using (var doc = EncryptedExcelDocument.Open(xlsxPath))
{
    doc.SaveAs(encPath, "hemmelig");
}
Check(EncryptedExcelDocument.IsEncrypted(encPath), "Fila er kryptert (OLE2-format)");

// Prøv å lese UTAN passord → skal kaste feil
try
{
    EncryptedExcelDocument.Open(encPath);
    Check(false, "Opna kryptert fil utan passord — skulle kasta feil");
}
catch (InvalidOperationException)
{
    Check(true, "Opna utan passord → InvalidOperationException (som forventa)");
}

// Prøv å lese med FEIL passord → skal kaste feil
try
{
    EncryptedExcelDocument.Open(encPath, "feilpassord");
    Check(false, "Opna med feil passord — skulle kasta feil");
}
catch (Exception)
{
    Check(true, "Opna med feil passord → feil (som forventa)");
}

// Les med RIKTIG passord → skal fungere
using (var encDoc = EncryptedExcelDocument.Open(encPath, "hemmelig"))
{
    var names = encDoc.GetSheetNames();
    Check(names.Length >= 1, $"Har {names.Length} ark");

    var val = encDoc.GetCellValue(names[0], 2, 1)?.ToString();
    Check(val == "Tastatur", $"Rad 2, col 1 = '{val}' (forventa 'Tastatur')");

    var pris = encDoc.GetCellValue(names[0], 3, 2);
    Check(Convert.ToDecimal(pris) == 349m, $"Rad 3, col 2 = {pris} (forventa 349)");

    // Endre ei celle og lagre på nytt med nytt passord
    encDoc.SetCellValue(names[0], 2, 3, 999);
    encDoc.SaveAs(encPath, "nyttpassord");
}
Check(EncryptedExcelDocument.IsEncrypted(encPath), "Fila framleis kryptert etter endring");

// Les tilbake med nytt passord
using (var reopened = EncryptedExcelDocument.Open(encPath, "nyttpassord"))
{
    var sheet = reopened.GetSheetNames()[0];
    var endret = reopened.GetCellValue(sheet, 2, 3);
    Check(Convert.ToInt32(endret) == 999, $"Endra celle = {endret} (forventa 999)");

    var uendra = reopened.GetCellValue(sheet, 1, 2)?.ToString();
    Check(uendra == "Pris", $"Uendra overskrift = '{uendra}' (forventa 'Pris')");
}

Console.WriteLine($"📂 Ukryptert: {xlsxPath}");
Console.WriteLine($"📂 Kryptert:  {encPath}");

// ──── Cleanup (behald Excel-filene for inspeksjon) ────
try { File.Delete(csvPath); File.Delete(jsonPath); } catch { }

Console.WriteLine($"\n━━━ {passed}/{passed + failed} testar bestått ━━━");
if (failed > 0)
{
    Console.WriteLine("❌ Nokre testar feila!");
    Environment.Exit(1);
}
else
{
    Console.WriteLine("🎉 Alt fungerer!");
}

public class Product
{
    public string Name { get; set; } = "";
    public decimal Pris { get; set; }
    public int Antall { get; set; }
}
