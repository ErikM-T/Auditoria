using System;
using System.Diagnostics;
using System.Net.Http;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace AuditoriaEquipos
{
    public class AuditExcelExporter
    {
        public async Task<bool> SubirReporteADriveAsync(string rutaExcelLocal, string cargo, string tipoArchivo, string urlWebAppScript)
        {
            try
            {
                byte[] fileBytes = File.ReadAllBytes(rutaExcelLocal);
                string base64String = Convert.ToBase64String(fileBytes);
                string nombreArchivo = Path.GetFileName(rutaExcelLocal);

                var payload = new
                {
                    cargo = cargo,
                    tipoArchivo = tipoArchivo,
                    nombreArchivo = nombreArchivo,
                    contenidoBase64 = base64String
                };

                string jsonBody = System.Text.Json.JsonSerializer.Serialize(payload);

                using (HttpClient client = new HttpClient())
                {
                    var content = new StringContent(jsonBody, Encoding.UTF8, "application/json");
                    HttpResponseMessage response = await client.PostAsync(urlWebAppScript, content);

                    string responseText = await response.Content.ReadAsStringAsync();
                    return response.IsSuccessStatusCode && responseText.Contains("success");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error al subir a Drive: {ex.Message}");
                return false;
            }
        }

        public string ExportarReporteAuditoria(
            string proceso,
            string ejecutadoPor,
            string cargo,
            string responsable,
            string usuarioDominio,
            string nombreEquipo,
            string versionSo,
            string tipoEquipo,
            string isAdmin,
            string usbState,
            string cdState,
            string btState,
            string wifiState,
            string navState,
            string outlookState,
            string gmailState,
            string gmailServices,
            string cloudState,
            string cloudServiceUsed,
            string histoState,
            string avName,
            string avVer,
            string avUpdate,
            string appsState,
            string vpnRdpState,
            int deskFiles,
            int recycleCount,
            int picsCount,
            int musicCount,
            int videoCount,
            int docsCount,
            string bitlockerState,
            string ntpState,
            string aipWordState,
            string aipExcelState,
            string aipOutlookState,
            string passState,
            TextBox txtResultados,
            string rutaDestino = "")
        {
            try
            {
                if (txtResultados != null)
                {
                    txtResultados.AppendText("\r\n[Exportación] Preparando entorno para generación de Excel...\r\n");
                    Application.DoEvents();
                }

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                string rutaExe = Process.GetCurrentProcess().MainModule?.FileName ?? AppContext.BaseDirectory;
                string directorioEjecutable = Path.GetDirectoryName(rutaExe) ?? AppContext.BaseDirectory;
                string carpetaResultados = Path.Combine(directorioEjecutable, "Resultados");
                Directory.CreateDirectory(carpetaResultados);

                string nombrePersona = string.IsNullOrWhiteSpace(responsable) ? "Usuario" : responsable;
                foreach (char c in Path.GetInvalidFileNameChars())
                {
                    nombrePersona = nombrePersona.Replace(c, '_');
                }
                nombrePersona = nombrePersona.Trim();

                // Se conserva tal cual la nomenclatura original
                string nombreArchivoFinal = $"GE-FO-21 INSPECCION DE SEGURIDAD DE EQUIPOS {nombrePersona}.xlsx";
                string rutaSalida = Path.Combine(carpetaResultados, nombreArchivoFinal);

                string nombreRecurso = "AuditoriaEquipos.Resultados.GE-FO-21 INSPECCION DE SEGURIDAD DE EQUIPOS .xlsx";

                using (MemoryStream streamPlantilla = ManejadorPlantillas.ObtenerStreamPlantilla(nombreRecurso))
                {
                    File.WriteAllBytes(rutaSalida, streamPlantilla.ToArray());

                    using (var excelPkg = new ExcelPackage(new FileInfo(rutaSalida)))
                    {
                        var ws = excelPkg.Workbook.Worksheets[0];

                    ws.Cells["G8"].Value = proceso?.ToUpper();
                    ws.Cells["G7"].Value = DateTime.Now.ToString("dd/MM/yyyy");
                    ws.Cells["B8"].Value = ejecutadoPor?.ToUpper();
                    ws.Cells["B7"].Value = cargo?.ToUpper();

                    ws.Cells["B11"].Value = responsable?.ToUpper();
                    ws.Cells["B12"].Value = usuarioDominio?.ToUpper();
                    ws.Cells["B13"].Value = nombreEquipo?.ToUpper();
                    ws.Cells["B14"].Value = versionSo?.ToUpper();

                    if (tipoEquipo == "Portátil")
                        ws.Cells["C15"].Value = "X (Portátil)";
                    else
                        ws.Cells["E15"].Value = "X (Escritorio)";

                    void SetCheckMarkDirect(ExcelWorksheet wSheet, int row, string valStr)
                    {
                        if (valStr == "SI") wSheet.Cells[$"C{row}"].Value = "X";
                        else if (valStr == "NO") wSheet.Cells[$"D{row}"].Value = "X";
                        else wSheet.Cells[$"E{row}"].Value = "X";
                    }

                    void SetFileCheckMarkDirect(ExcelWorksheet wSheet, int row, int count)
                    {
                        if (count > 0) wSheet.Cells[$"C{row}"].Value = "X";
                        else wSheet.Cells[$"D{row}"].Value = "X";
                    }

                    SetCheckMarkDirect(ws, 19, isAdmin);
                    SetCheckMarkDirect(ws, 20, usbState);
                    SetCheckMarkDirect(ws, 21, cdState);
                    SetCheckMarkDirect(ws, 22, btState);
                    SetCheckMarkDirect(ws, 23, wifiState);
                    SetCheckMarkDirect(ws, 24, navState);
                    SetCheckMarkDirect(ws, 25, outlookState);

                    SetCheckMarkDirect(ws, 26, gmailState);
                    ws.Cells["F26"].Value = gmailServices?.ToUpper();

                    SetCheckMarkDirect(ws, 27, cloudState);
                    ws.Cells["F27"].Value = cloudServiceUsed?.ToUpper();
                    SetCheckMarkDirect(ws, 28, histoState);

                    ws.Cells["C29"].Value = "X";
                    ws.Cells["F29"].Value = $"{avName?.ToUpper()} {avVer?.ToUpper()}";
                    ws.Cells["C30"].Value = avUpdate?.ToUpper();

                    SetCheckMarkDirect(ws, 32, appsState);
                    SetCheckMarkDirect(ws, 33, vpnRdpState);

                    SetFileCheckMarkDirect(ws, 34, deskFiles);
                    ws.Cells["F34"].Value = $"{deskFiles} archivos";

                    SetFileCheckMarkDirect(ws, 35, recycleCount);
                    ws.Cells["F35"].Value = $"{recycleCount} archivos";

                    SetFileCheckMarkDirect(ws, 36, picsCount);
                    ws.Cells["F36"].Value = $"{picsCount} archivos";

                    SetFileCheckMarkDirect(ws, 37, musicCount);
                    ws.Cells["F37"].Value = $"{musicCount} archivos";

                    SetFileCheckMarkDirect(ws, 38, videoCount);
                    ws.Cells["F38"].Value = $"{videoCount} archivos";

                    SetFileCheckMarkDirect(ws, 39, docsCount);
                    ws.Cells["F39"].Value = $"{docsCount} archivos";

                    SetCheckMarkDirect(ws, 40, bitlockerState);
                    SetCheckMarkDirect(ws, 41, ntpState);
                    ws.Cells["F41"].Value = "NTP";

                    SetCheckMarkDirect(ws, 42, aipWordState);
                    SetCheckMarkDirect(ws, 43, aipExcelState);
                    SetCheckMarkDirect(ws, 44, aipOutlookState);
                    SetCheckMarkDirect(ws, 45, passState);

                    FileInfo fileInfo = new FileInfo(rutaSalida);
                    excelPkg.SaveAs(fileInfo);
                    }
                }

                if (txtResultados != null)
                {
                    txtResultados.AppendText("\r\n[OK] ¡Archivo Excel guardado en la carpeta de resultados con éxito!\r\n");
                }

                return rutaSalida; // Deuelve la ruta física del archivo guardado
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error guardando el archivo Excel localmente: {ex.Message}",
                    "Error de Exportación", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return string.Empty;
            }
        }
    }
}
