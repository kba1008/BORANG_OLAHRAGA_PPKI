using System;
using System.Collections.Generic;
using System.IO;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Drive.v3; // Pastikan anda memasang NuGet package: Google.Apis.Drive.v3

namespace WebScannerBackend
{
    public class GoogleWorkspaceManager
    {
        // ==============================================================
        // TETAPAN GOOGLE API (SILA MASUKKAN ID ANDA DI SINI)
        // ==============================================================
        static string spreadsheetId = "MASUKKAN_SPREADSHEET_ID_ANDA_DI_SINI";
        static string folderId = "MASUKKAN_FOLDER_ID_GOOGLE_DRIVE_DI_SINI"; 
        
        static string sheetPengguna = "Pengguna"; // Sheet untuk Login Guru
        static string sheetSijil = "SijilPelajar"; // Sheet untuk simpan rekod sijil
        
        private static GoogleCredential GetCredential()
        {
            // Perlukan kebenaran untuk akses Sheets dan Drive
            string[] scopes = { SheetsService.Scope.Spreadsheets, DriveService.Scope.DriveFile };
            GoogleCredential credential;
            
            using (var stream = new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(scopes);
            }
            return credential;
        }

        // ==============================================================
        // FUNGSI 1: MUAT NAIK SIJIL (PDF) KE GOOGLE DRIVE
        // ==============================================================
        public static string MuatNaikSijilKeDrive(string namaFail, byte[] failPdf)
        {
            try
            {
                var service = new DriveService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = GetCredential(),
                    ApplicationName = "WebScanner PWA"
                });

                // Tetapkan meta data fail (Nama dan Folder tujuan)
                var fileMetadata = new Google.Apis.Drive.v3.Data.File()
                {
                    Name = namaFail,
                    Parents = new List<string> { folderId }
                };

                // Muat naik fail menggunakan memory stream
                string fileId = "";
                using (var stream = new MemoryStream(failPdf))
                {
                    var request = service.Files.Create(fileMetadata, stream, "application/pdf");
                    request.Fields = "id, webViewLink";
                    request.Upload();
                    
                    var file = request.ResponseBody;
                    fileId = file.WebViewLink; // Ambil link untuk diletakkan di Google Sheet
                }
                
                return fileId;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Ralat Drive: " + ex.Message);
                return null;
            }
        }

        // ==============================================================
        // FUNGSI 2: SIMPAN REKOD SIJIL KE GOOGLE SHEET
        // ==============================================================
        public static string SimpanRekodSijil(string namaPelajar, string namaSijil, string namaGuru, string pdfLink)
        {
            try
            {
                var service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = GetCredential(),
                    ApplicationName = "WebScanner PWA"
                });

                var range = $"{sheetSijil}!A:E"; // Kolum: Tarikh, Pelajar, Sijil, Guru, Link Drive
                var valueRange = new ValueRange();
                
                var objectList = new List<object>() { 
                    DateTime.Now.ToString("dd-MM-yyyy HH:mm"), 
                    namaPelajar, 
                    namaSijil, 
                    namaGuru, 
                    pdfLink 
                };
                
                valueRange.Values = new List<IList<object>> { objectList };

                var appendRequest = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range);
                appendRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
                
                appendRequest.Execute();
                return "Rekod sijil berjaya disimpan dalam sistem.";
            }
            catch (Exception ex)
            {
                return "Ralat Sheet: " + ex.Message;
            }
        }
    }
}
