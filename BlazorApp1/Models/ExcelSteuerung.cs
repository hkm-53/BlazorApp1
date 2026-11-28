namespace BlazorApp1.Models
{
    public class KombiEintrag
    {
        public string Blatt { get; set; } = "";
        public string Feldname { get; set; } = "";
        public string Wert { get; set; } = "";
        public string Bezeichnung { get; set; } = "";
        public string AnzeigeText => string.IsNullOrWhiteSpace(Bezeichnung)
            ? Wert
            : $"{Wert} – {Bezeichnung}";
    }

    public class FeldDefinition
    {
        public string Blatt { get; set; } = "";
        public string Feldname { get; set; } = "";
        public string Typ { get; set; } = "Text";
        public string Bezeichnung { get; set; } = "";
    }

    public class FeldItem
    {
        public int SpaltenNummer { get; set; }
        public string SpaltenName { get; set; } = "";
        public string AnzeigeLabel { get; set; } = "";
        public string Wert { get; set; } = "";
        public string Typ { get; set; } = "Text";
        public bool IstFormel { get; set; } = false;
        public bool IstProzent { get; set; } = false;
        public List<KombiEintrag> KombiOptionen { get; set; } = new();
    }

    public class ZeilenInfo
    {
        public int ZeilenNummer { get; set; }
        public string AnzeigeText { get; set; } = "";
    }

    public class ExcelSteuerungDaten
    {
        public Dictionary<string, string> GlobalParameter { get; set; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, string> BlattParameter { get; set; } = new(StringComparer.OrdinalIgnoreCase);
        public List<KombiEintrag> Kombifelder { get; set; } = new();
        public List<FeldDefinition> FeldDefinitionen { get; set; } = new();

        public int GetErsterDs(string blatt) =>
            BlattParameter.TryGetValue($"{blatt}|Erster_Ds", out var v) && int.TryParse(v, out int z) ? z : 3;

        public int GetLetzterDs(string blatt) =>
            BlattParameter.TryGetValue($"{blatt}|Letzter_Ds", out var v) && int.TryParse(v, out int z) ? z : 0;

        public int GetSummenzeile(string blatt) =>
            BlattParameter.TryGetValue($"{blatt}|Summenzeile", out var v) && int.TryParse(v, out int z) ? z : 0;
    }
}
