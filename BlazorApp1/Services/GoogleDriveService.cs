using System.Net.Http.Headers;
using System.Text.Json;

namespace BlazorApp1.Services
{
    public class GoogleDriveService
    {
        private readonly HttpClient _http;
        private const string ClientNummer = "532480087224-7cv3n70c8jechg9gnn57asvc0hesdvm9";

        public GoogleDriveService(HttpClient http)
        {
            _http = http;
        }

        public string GeneriereGoogleAuthUrl()
        {
            string volleClientId = ClientNummer + ".apps.googleusercontent.com";
            string authBasis = "https:" + "//accounts." + "google.com" + "/o/oauth2/v2/auth";
            string zielSeite = "https:" + "//hkm-53." + "github.io" + "/BlazorApp1/";
            string bereich = "https:" + "//www." + "googleapis.com" + "/auth/drive";

            return authBasis +
                   "?client_id=" + Uri.EscapeDataString(volleClientId) +
                   "&redirect_uri=" + Uri.EscapeDataString(zielSeite) +
                   "&response_type=token" +
                   "&scope=" + Uri.EscapeDataString(bereich);
        }


        public async Task<(string? DateiId, string DateiName, byte[]? Bytes, string? Fehler)> LadeExakteDateiAsync(string token, string exakterDateiname)
        {
            try
            {
                string query = Uri.EscapeDataString($"name = '{exakterDateiname}' and trashed = false");
                string suchUrl = $"https://www.googleapis.com{query}&fields=files(id,name)";

                var suchReq = new HttpRequestMessage(HttpMethod.Get, suchUrl);
                suchReq.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                var suchRes = await _http.SendAsync(suchReq);
                string json = await suchRes.Content.ReadAsStringAsync();

                if (!suchRes.IsSuccessStatusCode)
                    return (null, exakterDateiname, null, $"Suchfehler ({suchRes.StatusCode}): {json}");

                using var doc = JsonDocument.Parse(json);
                var files = doc.RootElement.GetProperty("files");
                if (files.GetArrayLength() == 0)
                    return (null, exakterDateiname, null, $"Datei '{exakterDateiname}' wurde auf Google Drive nicht gefunden.");

                string? fileId = files.GetProperty("id").GetString();
                string fileName = files.GetProperty("name").GetString() ?? exakterDateiname;

                string dlUrl = $"https://www.googleapis.com{fileId}?alt=media";
                var dlReq = new HttpRequestMessage(HttpMethod.Get, dlUrl);
                dlReq.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                var dlRes = await _http.SendAsync(dlReq);

                if (!dlRes.IsSuccessStatusCode)
                    return (fileId, fileName, null, $"Download-Fehler: {dlRes.StatusCode}");

                byte[] bytes = await dlRes.Content.ReadAsByteArrayAsync();
                return (fileId, fileName, bytes, null);
            }
            catch (Exception ex)
            {
                return (null, exakterDateiname, null, ex.Message);
            }
        }

        public async Task<(bool Erfolg, string? Fehler)> UeberschreibeDateiAsync(string token, string fileId, byte[] daten)
        {
            try
            {
                string uploadUrl = $"https://www.googleapis.com{fileId}?uploadType=media";
                var upReq = new HttpRequestMessage(HttpMethod.Patch, uploadUrl);
                upReq.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token);
                upReq.Content = new ByteArrayContent(daten);
                upReq.Content.Headers.ContentType = new MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

                var upRes = await _http.SendAsync(upReq);
                return upRes.IsSuccessStatusCode
                    ? (true, null)
                    : (false, $"Upload-Fehler ({upRes.StatusCode})");
            }
            catch (Exception ex)
            {
                return (false, ex.Message);
            }
        }

    }
}