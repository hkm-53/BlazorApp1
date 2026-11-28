using System.Globalization;
using BlazorApp1.Models;
using ClosedXML.Excel;

namespace BlazorApp1.Services
{
    public class ExcelService
    {
        private readonly CultureInfo _atCulture = new("de-AT");

        public ExcelSteuerungDaten LeseSteuerung(XLWorkbook workbook)
        {
            var steuerung = new ExcelSteuerungDaten();

            // 1. SYS_Param auslesen (globale & blattbezogene Parameter)
            if (workbook.TryGetWorksheet("SYS_Param", out var paramSheet))
            {
                int letzteZeile = paramSheet.LastRowUsed()?.RowNumber() ?? 1;
                string modus = "Global";

                for (int r = 1; r <= letzteZeile; r++)
                {
                    string c1 = paramSheet.Cell(r, 1).GetString().Trim();
                    string c2 = paramSheet.Cell(r, 2).GetString().Trim();
                    string c3 = paramSheet.Cell(r, 3).GetString().Trim();

                    if (c1.Equals("Parameter", StringComparison.OrdinalIgnoreCase)) { modus = "Global"; continue; }
                    if (c1.Equals("Blatt", StringComparison.OrdinalIgnoreCase)) { modus = "Blatt"; continue; }
                    if (string.IsNullOrEmpty(c1) || string.IsNullOrEmpty(c2)) continue;

                    if (modus == "Global")
                        steuerung.GlobalParameter[c1] = c2;
                    else
                        steuerung.BlattParameter[$"{c1}|{c2}"] = c3;
                }
            }

            // 2. SYS_Felder auslesen
            if (workbook.TryGetWorksheet("SYS_Felder", out var felderSheet))
            {
                int letzteZeile = felderSheet.LastRowUsed()?.RowNumber() ?? 1;
                for (int r = 2; r <= letzteZeile; r++)
                {
                    string blatt = felderSheet.Cell(r, 1).GetString().Trim();
                    string feldname = felderSheet.Cell(r, 2).GetString().Trim();
                    string typ = felderSheet.Cell(r, 3).GetString().Trim();
                    string bezeichnung = felderSheet.Cell(r, 4).GetString().Trim();

                    if (string.IsNullOrEmpty(blatt) || string.IsNullOrEmpty(feldname) || blatt.StartsWith("xxx") || blatt.StartsWith("Info"))
                        continue;

                    steuerung.FeldDefinitionen.Add(new FeldDefinition
                    {
                        Blatt = blatt,
                        Feldname = feldname,
                        Typ = string.IsNullOrEmpty(typ) ? "Text" : typ,
                        Bezeichnung = bezeichnung
                    });
                }
            }

            // 3. SYS_Kombifelder auslesen
            if (workbook.TryGetWorksheet("SYS_Kombifelder", out var kombiSheet))
            {
                int letzteZeile = kombiSheet.LastRowUsed()?.RowNumber() ?? 1;
                for (int r = 2; r <= letzteZeile; r++)
                {
                    string blatt = kombiSheet.Cell(r, 1).GetString().Trim();
                    string feldname = kombiSheet.Cell(r, 2).GetString().Trim();
                    string wert = kombiSheet.Cell(r, 3).GetString().Trim();
                    string bezeichnung = kombiSheet.Cell(r, 4).GetString().Trim();

                    if (string.IsNullOrEmpty(blatt) || string.IsNullOrEmpty(feldname) || string.IsNullOrEmpty(wert))
                        continue;

                    steuerung.Kombifelder.Add(new KombiEintrag
                    {
                        Blatt = blatt,
                        Feldname = feldname,
                        Wert = wert,
                        Bezeichnung = bezeichnung
                    });
                }
            }

            return steuerung;
        }

        public IXLWorksheet GetArbeitsblatt(XLWorkbook workbook, string bevorzugterName = "Erfassung")
        {
            if (workbook.TryGetWorksheet(bevorzugterName, out var sheet))
                return sheet;

            return workbook.Worksheets.First(w => !w.Name.StartsWith("SYS_", StringComparison.OrdinalIgnoreCase));
        }

        public List<ZeilenInfo> LeseZeilenListe(IXLWorksheet sheet, ExcelSteuerungDaten steuerung, out int standardZeile)
        {
            var liste = new List<ZeilenInfo>();
            int ersterDs = steuerung.GetErsterDs(sheet.Name);
            int letzterDs = steuerung.GetLetzterDs(sheet.Name);
            if (letzterDs <= 0) letzterDs = sheet.LastRowUsed()?.RowNumber() ?? ersterDs;

            string heuteText = DateTime.Now.ToString("dd.MM.yyyy");
            string heuteIso = DateTime.Now.ToString("yyyy-MM-dd");
            standardZeile = ersterDs;

            for (int row = ersterDs; row <= letzterDs; row++)
            {
                var zelle1 = sheet.Cell(row, 1);
                string datumText = zelle1.DataType == XLDataType.DateTime
                    ? zelle1.GetDateTime().ToString("dd.MM.yyyy")
                    : zelle1.GetString();

                // Dummy-Zeilen wie 31.12.1899 am Monatsende überspringen
                if (datumText.Contains("1899")) continue;

                string zeitText = sheet.Cell(row, 2).GetFormattedString();
                string anzeige = $"{datumText} {zeitText}".Trim();
                if (string.IsNullOrEmpty(anzeige)) anzeige = "(Leer)";

                liste.Add(new ZeilenInfo { ZeilenNummer = row, AnzeigeText = anzeige });

                if (datumText.Contains(heuteText) || datumText.Contains(heuteIso))
                    standardZeile = row;
            }

            return liste;
        }

        public List<FeldItem> LeseZeilenFelder(IXLWorksheet sheet, int zeile, ExcelSteuerungDaten steuerung)
        {
            var felder = new List<FeldItem>();
            int letzteSpalte = sheet.Row(1).LastCellUsed()?.Address.ColumnNumber ?? 0;

            for (int col = 1; col <= letzteSpalte; col++)
            {
                string spaltenName = sheet.Cell(1, col).GetString().Trim();
                if (string.IsNullOrEmpty(spaltenName)) continue;

                var def = steuerung.FeldDefinitionen.FirstOrDefault(f =>
                    f.Blatt.Equals(sheet.Name, StringComparison.OrdinalIgnoreCase) &&
                    f.Feldname.Equals(spaltenName, StringComparison.OrdinalIgnoreCase));

                string typ = def?.Typ ?? "Text";
                if (typ.Equals("Ausblenden", StringComparison.OrdinalIgnoreCase)) continue;

                var zelle = sheet.Cell(zeile, col);
                bool istFormel = zelle.HasFormula || typ.StartsWith("Formel", StringComparison.OrdinalIgnoreCase);

                // Typ automatisch erkennen, falls in SYS_Felder nichts steht
                if (def == null)
                {
                    if (spaltenName.Equals("Datum", StringComparison.OrdinalIgnoreCase)) typ = "DatumAusw";
                    else if (spaltenName.Equals("Zeit", StringComparison.OrdinalIgnoreCase)) typ = "ZeitAusw";
                    else if (zelle.DataType == XLDataType.Number) typ = "Zahl";
                }

                bool istProzent = zelle.Style.NumberFormat.Format.Contains("%");
                string wert;

                if (istFormel)
                {
                    wert = zelle.GetFormattedString();
                }
                else if (zelle.DataType == XLDataType.DateTime && typ == "DatumAusw")
                {
                    wert = zelle.GetDateTime().ToString("yyyy-MM-dd");
                }
                else if (typ == "ZeitAusw" && (zelle.DataType == XLDataType.DateTime || zelle.DataType == XLDataType.TimeSpan))
                {
                    wert = zelle.GetFormattedString();
                }
                else if (istProzent && zelle.DataType == XLDataType.Number)
                {
                    wert = Math.Round(zelle.GetDouble() * 100, 1).ToString(_atCulture);
                }
                else if (zelle.DataType == XLDataType.Number)
                {
                    wert = zelle.GetDouble().ToString(_atCulture);
                }
                else
                {
                    wert = zelle.GetString();
                }

                var kombis = steuerung.Kombifelder
                    .Where(k => k.Blatt.Equals(sheet.Name, StringComparison.OrdinalIgnoreCase) &&
                                k.Feldname.Equals(spaltenName, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (kombis.Count > 0) typ = "Kombifeld";

                felder.Add(new FeldItem
                {
                    SpaltenNummer = col,
                    SpaltenName = spaltenName,
                    AnzeigeLabel = string.IsNullOrWhiteSpace(def?.Bezeichnung) ? spaltenName : def.Bezeichnung,
                    Wert = wert,
                    Typ = typ,
                    IstFormel = istFormel,
                    IstProzent = istProzent,
                    KombiOptionen = kombis
                });
            }

            return felder;
        }

        public byte[] SchreibeZeile(byte[] originalBytes, string blattName, int zeile, List<FeldItem> felder, ExcelSteuerungDaten steuerung)
        {
            using var ms = new MemoryStream(originalBytes);
            using var workbook = new XLWorkbook(ms);
            var sheet = GetArbeitsblatt(workbook, blattName);
            int ersterDs = steuerung.GetErsterDs(sheet.Name);

            foreach (var feld in felder)
            {
                var zelle = sheet.Cell(zeile, feld.SpaltenNummer);

                // Formatierung aus der ersten Datenzeile (Erster_Ds) übernehmen
                if (zeile != ersterDs)
                {
                    zelle.Style = sheet.Cell(ersterDs, feld.SpaltenNummer).Style;
                }

                // Formelzellen niemals überschreiben!
                if (feld.IstFormel || zelle.HasFormula)
                {
                    continue;
                }

