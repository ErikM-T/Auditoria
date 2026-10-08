using System;
using System.IO;
using System.Reflection;
using System.Drawing;
using System.Diagnostics;
using System.Text;
using System.Windows.Forms;
using System.ServiceProcess;
using System.Net.NetworkInformation;
using System.Threading.Tasks;
using System.Collections.Generic;
using ClosedXML.Excel;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using System.Linq;

namespace AuditoriaEquipos
{

    public static class ManejadorPlantillas
    {
        /// <summary>
        /// Obtiene la plantilla embebida como un MemoryStream para trabajar directamente en memoria.
        /// </summary>
        public static MemoryStream ObtenerStreamPlantilla(string nombreRecurso)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            
            using Stream? stream = assembly.GetManifestResourceStream(nombreRecurso);
            if (stream == null)
            {
                throw new FileNotFoundException($"No se encontró el recurso embebido: {nombreRecurso}");
            }

            MemoryStream ms = new MemoryStream();
            stream.CopyTo(ms);
            ms.Position = 0; // Reiniciar posición del stream
            return ms;
        }

        /// <summary>
        /// Extrae la plantilla del .exe y la guarda físicamente en un archivo local.
        /// </summary>
        public static string ExtraerPlantillaAArchivo(string nombreRecurso, string rutaDestino)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();

            using Stream? stream = assembly.GetManifestResourceStream(nombreRecurso);
            if (stream == null)
            {
                throw new FileNotFoundException($"No se encontró el recurso embebido: {nombreRecurso}");
            }

            using FileStream fileStream = new FileStream(rutaDestino, FileMode.Create, FileAccess.Write);
            stream.CopyTo(fileStream);

            return rutaDestino;
        }
    }
    public partial class Form1 : Form
    {
        private TextBox txtProceso = null!;
        private TextBox txtEjecutadoPor = null!;
        private TextBox txtFecha = null!;
        private ComboBox cmbCargo = null!;

        private TextBox txtRespEquipo = null!;
        private TextBox txtUsuarioDominio = null!;
        private TextBox txtNombreEquipo = null!;
        private TextBox txtVersionSO = null!;
        private RadioButton rbPortatil = null!, rbEscritorio = null!;

        private CheckBox chkAdminSi = null!, chkAdminNo = null!, chkAdminNa = null!;
        private TextBox txtAdminDetalle = null!;

        private CheckBox chkUsbSi = null!, chkUsbNo = null!, chkUsbNa = null!;
        private CheckBox chkCdSi = null!, chkCdNo = null!, chkCdNa = null!;
        private CheckBox chkBtSi = null!, chkBtNo = null!, chkBtNa = null!;
        private CheckBox chkWifiSi = null!, chkWifiNo = null!, chkWifiNa = null!;

        private CheckBox chkWebSi = null!, chkWebNo = null!, chkWebNa = null!;
        private CheckBox chkOutlookSi = null!, chkOutlookNo = null!, chkOutlookNa = null!;
        private CheckBox chkGmailSi = null!, chkGmailNo = null!, chkGmailNa = null!;
        private CheckBox chkCloudSi = null!, chkCloudNo = null!, chkCloudNa = null!;
        private CheckBox chkHistorialSi = null!, chkHistorialNo = null!, chkHistorialNa = null!;
        private TextBox txtCloudDetalle = null!;
        private TextBox txtGmailDetalle = null!;

        private CheckBox chkAvInstSi = null!, chkAvInstNo = null!, chkAvInstNa = null!;
        private TextBox txtAvInstDetalle = null!;
        private TextBox txtAvFechaUpd = null!;
        private TextBox txtAvAgente = null!;

        private CheckBox chkAppSi = null!, chkAppNo = null!, chkAppNa = null!;
        private CheckBox chkVpnSi = null!, chkVpnNo = null!, chkVpnNa = null!;

        private CheckBox chkEscritorioSi = null!, chkEscritorioNo = null!, chkEscritorioNa = null!;
        private TextBox txtEscritorioDetalle = null!;

        private CheckBox chkPapeleraSi = null!, chkPapeleraNo = null!, chkPapeleraNa = null!;
        private TextBox txtPapeleraDetalle = null!;

        private CheckBox chkImgSi = null!, chkImgNo = null!, chkImgNa = null!;
        private TextBox txtImgDetalle = null!;
        private CheckBox chkMusSi = null!, chkMusNo = null!, chkMusNa = null!;
        private TextBox txtMusDetalle = null!;
        private CheckBox chkVidSi = null!, chkVidNo = null!, chkVidNa = null!;
        private TextBox txtVidDetalle = null!;

        private CheckBox chkDocSi = null!, chkDocNo = null!, chkDocNa = null!;
        private TextBox txtDocDetalle = null!;

        private CheckBox chkCifradoSi = null!, chkCifradoNo = null!, chkCifradoNa = null!;
        private CheckBox chkNtpSi = null!, chkNtpNo = null!, chkNtpNa = null!;
        private TextBox txtNtpDetalle = null!;

        private CheckBox chkAipWordSi = null!, chkAipWordNo = null!, chkAipWordNa = null!;
        private CheckBox chkAipExcelSi = null!, chkAipExcelNo = null!, chkAipExcelNa = null!;
        private CheckBox chkAipOutlookSi = null!, chkAipOutlookNo = null!, chkAipOutlookNa = null!;
        private CheckBox chkPassSi = null!, chkPassNo = null!, chkPassNa = null!;

        private Button btnEjecutar = null!;
        private Button btnPlantillaExcel = null!;
        private Button btnCerrar = null!;
        private TextBox txtResultados = null!;

        // Estructuras auxiliares para la validación de software por área
        private class SoftwarePermitido
        {
            public string Nombre { get; set; } = "";
            public string VersionEsperada { get; set; } = "";
            public string Area { get; set; } = "";
        }

        private class ProgramaInstalado
        {
            public string Nombre { get; set; } = "";
            public string Version { get; set; } = "";
        }

        public Form1()
        {
            InitializeComponentCustom();
            chkWebSi.Checked = false;
            chkWebNo.Checked = false;
            chkWebNa.Checked = false;
            chkHistorialNo.Checked = false;
        }

        private bool DetectarUsbHabilitado()
        {
            try
            {
                using (RegistryKey? key = Registry.LocalMachine.OpenSubKey(@"SYSTEM\CurrentControlSet\Services\USBSTOR"))
                {
                    if (key != null)
                    {
                        object? startVal = key.GetValue("Start");
                        if (startVal != null && Convert.ToInt32(startVal) == 4)
                        {
                            return false;
                        }
                    }
                }

                bool tieneHubsUsb = System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable(); 
                using (var searcher = new System.Management.ManagementObjectSearcher(@"SELECT * FROM Win32_USBController"))
                {
                    using (var collection = searcher.Get())
                    {
                        if (collection.Count > 0)
                        {
                            return true;
                        }
                    }
                }
            }
            catch { }

            // Si no detecta bloqueo ni fallos en los controladores, asume habilitado por defecto
            return true;
        }
        private void InitializeComponentCustom()
        {
            this.Text = "Formato Oficial de Auditoría de Equipos - Turrisystem (GA-FO-02)";
            this.MaximumSize = SystemInformation.PrimaryMonitorMaximizedWindowSize;
            this.WindowState = FormWindowState.Maximized;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(245, 247, 250);

            Panel panel = new Panel();
            panel.Dock = DockStyle.Fill;
            panel.AutoScroll = true;
            this.Controls.Add(panel);

            int y = 20;

            CrearFilaEncabezado(panel, ref txtProceso, ref txtEjecutadoPor, ref txtFecha, ref cmbCargo, ref y);
            y += 10;

            CrearSeccionTitulo(panel, "DETALLE DEL EQUIPO Y SISTEMA", ref y);
            CrearFilaTexto(panel, "Nombre Responsable Equipo:", ref txtRespEquipo, ref y);
            CrearFilaTexto(panel, "Usuario Dominio:", ref txtUsuarioDominio, ref y);
            CrearFilaTexto(panel, "Nombre Equipo (Host):", ref txtNombreEquipo, ref y);
            CrearFilaTexto(panel, "Versión S.O.:", ref txtVersionSO, ref y);

            Label lblTipo = CrearLabel(panel, "Tipo de Equipo:", 15, y);
            rbPortatil = new RadioButton() { Text = "Portátil", Location = new Point(220, y), Size = new Size(75, 23), Checked = true };
            rbEscritorio = new RadioButton() { Text = "Escritorio", Location = new Point(310, y), Size = new Size(85, 23) };
            panel.Controls.Add(rbPortatil);
            panel.Controls.Add(rbEscritorio);
            y += 38;

            CrearCabeceraTabla(panel, ref y);

            CrearFilaPregunta(panel, "¿Es Usuario Administrador?", out chkAdminSi, out chkAdminNo, out chkAdminNa, out txtAdminDetalle, ref y);
            CrearFilaPreguntaSub(panel, "¿Tiene Habilitados los siguientes periféricos?", "USB", out chkUsbSi, out chkUsbNo, out chkUsbNa, ref y);
            CrearFilaPreguntaSub(panel, "", "CD/DVD", out chkCdSi, out chkCdNo, out chkCdNa, ref y);
            CrearFilaPreguntaSub(panel, "", "Bluetooth", out chkBtSi, out chkBtNo, out chkBtNa, ref y);
            CrearFilaPreguntaSub(panel, "", "WiFi", out chkWifiSi, out chkWifiNo, out chkWifiNa, ref y);

            CrearFilaPreguntaSub(panel, "¿El equipo cuenta con alguno de los siguientes accesos web?", "Navegación web", out chkWebSi, out chkWebNo, out chkWebNa, ref y);
            CrearFilaPreguntaSub(panel, "", "Outlook.live", out chkOutlookSi, out chkOutlookNo, out chkOutlookNa, ref y);
            CrearFilaPreguntaSub(panel, "", "Servicios Gmail", out chkGmailSi, out chkGmailNo, out chkGmailNa, ref y, out txtGmailDetalle);
            CrearFilaPreguntaSub(panel, "", "Servicios de almacenamiento", out chkCloudSi, out chkCloudNo, out chkCloudNa, ref y, out txtCloudDetalle);
            CrearFilaPreguntaSub(panel, "", "Historial de navegación", out chkHistorialSi, out chkHistorialNo, out chkHistorialNa, ref y);

            CrearFilaPreguntaSub(panel, "¿Cuenta con Antivirus?", "Instalado", out chkAvInstSi, out chkAvInstNo, out chkAvInstNa, ref y, out txtAvInstDetalle);
            CrearFilaPreguntaSubTextoLibre(panel, "", "Última comprobación de actualización de seguridad", ref txtAvFechaUpd, ref y);
            CrearFilaPreguntaSubTextoLibre(panel, "", "Última comunicación agente-servidor", ref txtAvAgente, ref y);

            CrearFilaPregunta(panel, "¿Cuenta con Aplicaciones no Autorizadas?", out chkAppSi, out chkAppNo, out chkAppNa, out _, ref y);
            CrearFilaPregunta(panel, "¿Cuenta con acceso a VPN o acceso remoto RDP?", out chkVpnSi, out chkVpnNo, out chkVpnNa, out _, ref y);

            CrearFilaPreguntaConDetalle(panel, "¿Cuenta con información en el escritorio? (Cuántos archivos)", out chkEscritorioSi, out chkEscritorioNo, out chkEscritorioNa, out txtEscritorioDetalle, ref y);
            CrearFilaPreguntaConDetalle(panel, "¿Cuenta con información en la Papelera de Reciclaje? (Cuántos archivos)", out chkPapeleraSi, out chkPapeleraNo, out chkPapeleraNa, out txtPapeleraDetalle, ref y);

            CrearFilaPreguntaSub(panel, "¿Cuenta con los siguientes tipos de archivos en el equipo?", "Imágenes", out chkImgSi, out chkImgNo, out chkImgNa, ref y, out txtImgDetalle);
            CrearFilaPreguntaSub(panel, "", "Música", out chkMusSi, out chkMusNo, out chkMusNa, ref y, out txtMusDetalle);
            CrearFilaPreguntaSub(panel, "", "Videos", out chkVidSi, out chkVidNo, out chkVidNa, ref y, out txtVidDetalle);
            CrearFilaPreguntaConDetalle(panel, "¿Cuenta con información en la carpeta Mis Documentos? (Cuántos archivos)", out chkDocSi, out chkDocNo, out chkDocNa, out txtDocDetalle, ref y);

            CrearFilaPregunta(panel, "¿El disco duro se encuentra cifrado? (Solo equipo portátil)", out chkCifradoSi, out chkCifradoNo, out chkCifradoNa, out _, ref y);
            CrearFilaPreguntaSubTextoLibreConCheck(panel, "¿La hora del equipo se encuentra sincronizada con una fuente NTP, cuál? (w32tm /query /status)", "NTP", out chkNtpSi, out chkNtpNo, out chkNtpNa, out txtNtpDetalle, ref y);
            CrearFilaPreguntaSub(panel, "¿Se encuentra habilitado Azure Information Protect (AIP)?", "Word", out chkAipWordSi, out chkAipWordNo, out chkAipWordNa, ref y);
            CrearFilaPreguntaSub(panel, "", "Excel", out chkAipExcelSi, out chkAipExcelNo, out chkAipExcelNa, ref y);
            CrearFilaPreguntaSub(panel, "", "Outlook", out chkAipOutlookSi, out chkAipOutlookNo, out chkAipOutlookNa, ref y);
            CrearFilaPregunta(panel, "¿Se visualizan contraseñas en el equipo o archivos con contraseñas? (Por ej. pósito)", out chkPassSi, out chkPassNo, out chkPassNa, out _, ref y);

            y += 20;

            btnEjecutar = new Button() { Text = "Ejecutar Consultas Automáticas", Location = new Point(15, y), Size = new Size(240, 40) };
            btnEjecutar.BackColor = Color.FromArgb(16, 185, 129);
            btnEjecutar.ForeColor = Color.White;
            btnEjecutar.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            btnEjecutar.Click += new EventHandler(BtnEjecutar_Click);
            panel.Controls.Add(btnEjecutar);

            btnPlantillaExcel = new Button() { Text = "Llenar Plantilla Excel", Location = new Point(270, y), Size = new Size(220, 40) };
            btnPlantillaExcel.BackColor = Color.FromArgb(59, 130, 246);
            btnPlantillaExcel.ForeColor = Color.White;
            btnPlantillaExcel.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            btnPlantillaExcel.Click += new EventHandler(BtnLlenarPlantillaExcel_Click);
            panel.Controls.Add(btnPlantillaExcel);

            btnCerrar = new Button() { Text = "Cerrar", Location = new Point(815, y), Size = new Size(200, 40) };
            btnCerrar.BackColor = Color.FromArgb(239, 68, 68);
            btnCerrar.ForeColor = Color.White;
            btnCerrar.Font = new Font("Segoe UI", 9, FontStyle.Bold);
            btnCerrar.Click += (sender, e) => { this.Close(); };
            panel.Controls.Add(btnCerrar);
            y += 60;

            txtResultados = new TextBox() { Multiline = true, ReadOnly = true, Location = new Point(15, y), Size = new Size(1000, 100), Visible = false };
            panel.Controls.Add(txtResultados);
        }

        private void CrearFilaEncabezado(Panel p, ref TextBox pProc, ref TextBox pEjec, ref TextBox pFecha, ref ComboBox pCargo, ref int y)
        {
            CrearLabel(p, "Proceso / Área:", 10, y);
            pCargo = new ComboBox() { Location = new Point(200, y), Size = new Size(275, 25), DropDownStyle = ComboBoxStyle.DropDown };
            pCargo.Items.Add("SOPORTE");
            pCargo.Items.Add("DESARROLLO");
            pCargo.Items.Add("ADMINISTRATIVOS");
            pCargo.SelectedIndex = 0;
            p.Controls.Add(pCargo);

            CrearLabel(p, "Fecha:", 500, y);
            pFecha = new TextBox() { Location = new Point(685, y), Size = new Size(150, 23), Text = DateTime.Now.ToString("yyyy-MM-dd"), ReadOnly = true };
            p.Controls.Add(pFecha);
            y += 32;

            CrearLabel(p, "Ejecutado por:", 15, y);
            string userFullName = "ERIK ALONSO MACIAS HOYOS";
            pEjec = new TextBox() { Location = new Point(200, y), Size = new Size(275, 23), Text = userFullName.ToUpper() };
            p.Controls.Add(pEjec);

            CrearLabel(p, "Cargo:", 500, y);
            pProc = new TextBox() { Location = new Point(685, y), Size = new Size(275, 23), Text = "AUXILIAR DE SOPORTE TECNICO" };
            p.Controls.Add(pProc);
            y += 42;
        }

        private void CrearSeccionTitulo(Panel p, string texto, ref int y)
        {
            Label lbl = new Label() { Text = texto, Location = new Point(15, y), Size = new Size(1005, 24) };
            lbl.Font = new Font("Segoe UI", 9.5F, FontStyle.Bold);
            lbl.BackColor = Color.FromArgb(209, 213, 219);
            lbl.TextAlign = ContentAlignment.MiddleCenter;
            p.Controls.Add(lbl);
            y += 30;
        }

        private Label CrearLabel(Panel p, string text, int x, int y)
        {
            Label lbl = new Label() { Text = text, Location = new Point(x, y), Size = new Size(185, 22) };
            lbl.Font = new Font("Segoe UI", 8.25F, FontStyle.Bold);
            p.Controls.Add(lbl);
            return lbl;
        }

        private void CrearFilaTexto(Panel p, string labelText, ref TextBox txt, ref int y)
        {
            CrearLabel(p, labelText, 15, y);
            txt = new TextBox() { Location = new Point(220, y), Size = new Size(580, 23) };
            p.Controls.Add(txt);
            y += 28;
        }

        private void CrearCabeceraTabla(Panel p, ref int y)
        {
            Panel header = new Panel() { Location = new Point(15, y), Size = new Size(1005, 28), BackColor = Color.FromArgb(229, 231, 235) };
            Label l1 = new Label() { Text = "ÍTEM / PREGUNTA", Location = new Point(5, 5), Size = new Size(425, 18), Font = new Font("Segoe UI", 8F, FontStyle.Bold) };
            Label l2 = new Label() { Text = "SI", Location = new Point(445, 5), Size = new Size(30, 18), Font = new Font("Segoe UI", 8F, FontStyle.Bold), TextAlign = ContentAlignment.MiddleCenter };
            Label l3 = new Label() { Text = "NO", Location = new Point(495, 5), Size = new Size(30, 18), Font = new Font("Segoe UI", 8F, FontStyle.Bold), TextAlign = ContentAlignment.MiddleCenter };
            Label l4 = new Label() { Text = "N/A", Location = new Point(545, 5), Size = new Size(30, 18), Font = new Font("Segoe UI", 8F, FontStyle.Bold), TextAlign = ContentAlignment.MiddleCenter };
            Label l5 = new Label() { Text = "DETALLES", Location = new Point(605, 5), Size = new Size(375, 18), Font = new Font("Segoe UI", 8F, FontStyle.Bold) };
            header.Controls.Add(l1); header.Controls.Add(l2); header.Controls.Add(l3); header.Controls.Add(l4); header.Controls.Add(l5);
            p.Controls.Add(header);
            y += 32;
        }

        private void ConfigurarExclusividad(CheckBox chkSi, CheckBox chkNo, CheckBox chkNa)
        {
            chkSi.CheckedChanged += (s, e) => {
                if (chkSi.Checked) { chkNo.Checked = false; chkNa.Checked = false; }
            };
            chkNo.CheckedChanged += (s, e) => {
                if (chkNo.Checked) { chkSi.Checked = false; chkNa.Checked = false; }
            };
            chkNa.CheckedChanged += (s, e) => {
                if (chkNa.Checked) { chkSi.Checked = false; chkNo.Checked = false; }
            };
        }

        private void CrearFilaPregunta(Panel p, string pregunta, out CheckBox chkSi, out CheckBox chkNo, out CheckBox chkNa, out TextBox txtDetalle, ref int y)
        {
            Label lbl = new Label() { Text = pregunta, Location = new Point(15, y), Size = new Size(430, 22) };
            lbl.Font = new Font("Segoe UI", 8.25F);
            p.Controls.Add(lbl);

            chkSi = new CheckBox() { Location = new Point(468, y + 2), Size = new Size(18, 18) };
            chkNo = new CheckBox() { Location = new Point(518, y + 2), Size = new Size(18, 18) };
            chkNa = new CheckBox() { Location = new Point(568, y + 2), Size = new Size(18, 18) };
            ConfigurarExclusividad(chkSi, chkNo, chkNa);
            p.Controls.Add(chkSi); p.Controls.Add(chkNo); p.Controls.Add(chkNa);

            txtDetalle = new TextBox() { Location = new Point(615, y), Size = new Size(405, 23) };
            p.Controls.Add(txtDetalle);
            y += 28;
        }

        private void CrearFilaPreguntaConDetalle(Panel p, string pregunta, out CheckBox chkSi, out CheckBox chkNo, out CheckBox chkNa, out TextBox txtDetalle, ref int y)
        {
            CrearFilaPregunta(p, pregunta, out chkSi, out chkNo, out chkNa, out txtDetalle, ref y);
        }

        private void CrearFilaPreguntaSub(Panel p, string categoria, string subItem, out CheckBox chkSi, out CheckBox chkNo, out CheckBox chkNa, ref int y, out TextBox txtDetalle)
        {
            if (!string.IsNullOrEmpty(categoria))
            {
                Label lblCat = new Label() { Text = categoria, Location = new Point(15, y), Size = new Size(240, 22) };
                lblCat.Font = new Font("Segoe UI", 8.25F);
                p.Controls.Add(lblCat);
            }

            Label lblSub = new Label() { Text = subItem, Location = new Point(265, y), Size = new Size(180, 22) };
            lblSub.Font = new Font("Segoe UI", 8.25F, FontStyle.Italic);
            p.Controls.Add(lblSub);

            chkSi = new CheckBox() { Location = new Point(468, y + 2), Size = new Size(18, 18) };
            chkNo = new CheckBox() { Location = new Point(518, y + 2), Size = new Size(18, 18) };
            chkNa = new CheckBox() { Location = new Point(568, y + 2), Size = new Size(18, 18) };
            ConfigurarExclusividad(chkSi, chkNo, chkNa);
            p.Controls.Add(chkSi); p.Controls.Add(chkNo); p.Controls.Add(chkNa);

            txtDetalle = new TextBox() { Location = new Point(615, y), Size = new Size(405, 23) };
            p.Controls.Add(txtDetalle);
            y += 28;
        }

        private void CrearFilaPreguntaSub(Panel p, string categoria, string subItem, out CheckBox chkSi, out CheckBox chkNo, out CheckBox chkNa, ref int y)
        {
            TextBox dummy;
            CrearFilaPreguntaSub(p, categoria, subItem, out chkSi, out chkNo, out chkNa, ref y, out dummy);
        }

        private void CrearFilaPreguntaSubTextoLibre(Panel p, string categoria, string subItem, ref TextBox txtDetalle, ref int y)
        {
            if (!string.IsNullOrEmpty(categoria))
            {
                Label lblCat = new Label() { Text = categoria, Location = new Point(15, y), Size = new Size(240, 22) };
                lblCat.Font = new Font("Segoe UI", 8.25F);
                p.Controls.Add(lblCat);
            }

            Label lblSub = new Label() { Text = subItem, Location = new Point(265, y), Size = new Size(345, 22) };
            lblSub.Font = new Font("Segoe UI", 8.25F, FontStyle.Italic);
            p.Controls.Add(lblSub);

            txtDetalle = new TextBox() { Location = new Point(615, y), Size = new Size(405, 23) };
            p.Controls.Add(txtDetalle);
            y += 28;
        }

        private void CrearFilaPreguntaSubTextoLibreConCheck(Panel p, string categoria, string subItem, out CheckBox chkSi, out CheckBox chkNo, out CheckBox chkNa, out TextBox txtDetalle, ref int y)
        {
            Label lblCat = new Label() { Text = categoria, Location = new Point(15, y), Size = new Size(245, 22) };
            lblCat.Font = new Font("Segoe UI", 8.25F);
            p.Controls.Add(lblCat);

            Label lblSub = new Label() { Text = subItem, Location = new Point(265, y), Size = new Size(180, 22) };
            lblSub.Font = new Font("Segoe UI", 8.25F, FontStyle.Italic);
            p.Controls.Add(lblSub);

            chkSi = new CheckBox() { Location = new Point(468, y + 2), Size = new Size(18, 18) };
            chkNo = new CheckBox() { Location = new Point(518, y + 2), Size = new Size(18, 18) };
            chkNa = new CheckBox() { Location = new Point(568, y + 2), Size = new Size(18, 18) };
            ConfigurarExclusividad(chkSi, chkNo, chkNa);
            p.Controls.Add(chkSi); p.Controls.Add(chkNo); p.Controls.Add(chkNa);

            txtDetalle = new TextBox() { Location = new Point(615, y), Size = new Size(405, 23) };
            p.Controls.Add(txtDetalle);
            y += 28;
        }

        private async void BtnLlenarPlantillaExcel_Click(object? sender, EventArgs e)
        {
            static bool GrupoValidado(CheckBox? si, CheckBox? no, CheckBox? na)
            {
                return (si != null && si.Checked) || (no != null && no.Checked) || (na != null && na.Checked);
            }

            if (!GrupoValidado(chkAdminSi, chkAdminNo, chkAdminNa) ||
                !GrupoValidado(chkUsbSi, chkUsbNo, chkUsbNa) ||
                !GrupoValidado(chkCdSi, chkCdNo, chkCdNa) ||
                !GrupoValidado(chkBtSi, chkBtNo, chkBtNa) ||
                !GrupoValidado(chkWifiSi, chkWifiNo, chkWifiNa) ||
                !GrupoValidado(chkWebSi, chkWebNo, chkWebNa) ||
                !GrupoValidado(chkOutlookSi, chkOutlookNo, chkOutlookNa) ||
                !GrupoValidado(chkGmailSi, chkGmailNo, chkGmailNa) ||
                !GrupoValidado(chkCloudSi, chkCloudNo, chkCloudNa) ||
                !GrupoValidado(chkHistorialSi, chkHistorialNo, chkHistorialNa) ||
                !GrupoValidado(chkAvInstSi, chkAvInstNo, chkAvInstNa) ||
                !GrupoValidado(chkAppSi, chkAppNo, chkAppNa) ||
                !GrupoValidado(chkVpnSi, chkVpnNo, chkVpnNa) ||
                !GrupoValidado(chkEscritorioSi, chkEscritorioNo, chkEscritorioNa) ||
                !GrupoValidado(chkPapeleraSi, chkPapeleraNo, chkPapeleraNa) ||
                !GrupoValidado(chkImgSi, chkImgNo, chkImgNa) ||
                !GrupoValidado(chkMusSi, chkMusNo, chkMusNa) ||
                !GrupoValidado(chkVidSi, chkVidNo, chkVidNa) ||
                !GrupoValidado(chkDocSi, chkDocNo, chkDocNa) ||
                !GrupoValidado(chkCifradoSi, chkCifradoNo, chkCifradoNa) ||
                !GrupoValidado(chkNtpSi, chkNtpNo, chkNtpNa) ||
                !GrupoValidado(chkAipWordSi, chkAipWordNo, chkAipWordNa) ||
                !GrupoValidado(chkAipExcelSi, chkAipExcelNo, chkAipExcelNa) ||
                !GrupoValidado(chkAipOutlookSi, chkAipOutlookNo, chkAipOutlookNa) ||
                !GrupoValidado(chkPassSi, chkPassNo, chkPassNa))
            {
                MessageBox.Show("Por favor, asegúrate de responder todas las preguntas seleccionando una opción (SI, NO o N/A) antes de continuar.", 
                    "Campos Incompletos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(txtEscritorioDetalle.Text) || 
                string.IsNullOrWhiteSpace(txtPapeleraDetalle.Text) || 
                string.IsNullOrWhiteSpace(txtImgDetalle.Text) || 
                string.IsNullOrWhiteSpace(txtMusDetalle.Text) || 
                string.IsNullOrWhiteSpace(txtVidDetalle.Text) || 
                string.IsNullOrWhiteSpace(txtDocDetalle.Text))
            {
                MessageBox.Show("Por favor, completa todas las cantidades de archivos. No pueden estar vacías.", 
                    "Campo Requerido", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            static string Estado(CheckBox si, CheckBox no, CheckBox na)
            {
                if (si != null && si.Checked) return "SI";
                if (no != null && no.Checked) return "NO";
                return "NA";
            }

            static int ParseCount(TextBox? txt)
            {
                if (txt == null || string.IsNullOrWhiteSpace(txt.Text)) return 0;
                string textoLimpio = System.Text.RegularExpressions.Regex.Match(txt.Text, @"\d+").Value;
                if (int.TryParse(textoLimpio, out var valor)) return valor;
                return int.TryParse(txt.Text, out var valor2) ? valor2 : 0;
            }

            var exporter = new AuditExcelExporter();
            string directorioEjecutable = AppDomain.CurrentDomain.BaseDirectory;
            string carpetaResultados = Path.Combine(directorioEjecutable, "Resultados");

            // 1. Exportar localmente usando el nombre exacto que ya tienes configurado
            string rutaExcelGenerado = exporter.ExportarReporteAuditoria(
                txtProceso.Text,
                txtEjecutadoPor.Text.ToUpper(),
                cmbCargo.Text.ToUpper(),
                txtRespEquipo.Text.ToUpper(),
                txtUsuarioDominio.Text.ToUpper(),
                txtNombreEquipo.Text.ToUpper(),
                txtVersionSO.Text.ToUpper(),
                rbPortatil.Checked ? "Portátil" : "Escritorio",
                Estado(chkAdminSi, chkAdminNo, chkAdminNa),
                Estado(chkUsbSi, chkUsbNo, chkUsbNa),
                Estado(chkCdSi, chkCdNo, chkCdNa),
                Estado(chkBtSi, chkBtNo, chkBtNa),
                Estado(chkWifiSi, chkWifiNo, chkWifiNa),
                Estado(chkWebSi, chkWebNo, chkWebNa),
                Estado(chkOutlookSi, chkOutlookNo, chkOutlookNa),
                Estado(chkGmailSi, chkGmailNo, chkGmailNa),
                txtGmailDetalle.Text,
                Estado(chkCloudSi, chkCloudNo, chkCloudNa),
                txtCloudDetalle.Text,
                Estado(chkHistorialSi, chkHistorialNo, chkHistorialNa),
                txtAvAgente.Text,
                txtAvInstDetalle.Text,
                txtAvFechaUpd.Text,
                Estado(chkAppSi, chkAppNo, chkAppNa),
                Estado(chkVpnSi, chkVpnNo, chkVpnNa),
                ParseCount(txtEscritorioDetalle),
                ParseCount(txtPapeleraDetalle),
                ParseCount(txtImgDetalle),
                ParseCount(txtMusDetalle),
                ParseCount(txtVidDetalle),
                ParseCount(txtDocDetalle),
                Estado(chkCifradoSi, chkCifradoNo, chkCifradoNa),
                Estado(chkNtpSi, chkNtpNo, chkNtpNa),
                Estado(chkAipWordSi, chkAipWordNo, chkAipWordNa),
                Estado(chkAipExcelSi, chkAipExcelNo, chkAipExcelNa),
                Estado(chkAipOutlookSi, chkAipOutlookNo, chkAipOutlookNa),
                Estado(chkPassSi, chkPassNo, chkPassNa),
                txtResultados,
                carpetaResultados);

            // 2. Si se generó el archivo local correctamente, enviarlo a Google Drive
            if (!string.IsNullOrEmpty(rutaExcelGenerado) && File.Exists(rutaExcelGenerado))
            {
                string urlScript = "https://script.google.com/macros/s/AKfycbxM2VqT9AY2FEUKxqGfU4v_pSKNj37ZfeKslxDnQx3jpMVdZK1vafTDs2M9D0fNw904/exec";

                bool subidoConExito = await exporter.SubirReporteADriveAsync(
                    rutaExcelLocal: rutaExcelGenerado,
                    cargo: cmbCargo.Text.Trim(),                  // Crea la subcarpeta del Cargo (ej: "SOPORTE")
                    tipoArchivo: "GE-FO-21_INSPECCION_DE_SEGURIDAD", // Subcarpeta tipo de archivo
                    urlWebAppScript: urlScript
                );

                if (subidoConExito)
                {
                    MessageBox.Show($"Reporte generado y subido exitosamente a Google Drive.\r\n\r\nRuta local:\r\n{rutaExcelGenerado}", 
                        "Éxito Completo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show($"Reporte guardado localmente, pero no se pudo subir a Google Drive.\r\n\r\nRuta local:\r\n{rutaExcelGenerado}", 
                        "Advertencia de Red", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private async void BtnEjecutar_Click(object? sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtRespEquipo.Text))
            {
                MessageBox.Show("Por favor, ingresa el 'Nombre Responsable Equipo' antes de ejecutar la consulta de software.", 
                    "Campo Requerido", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtRespEquipo.Focus();
                return;
            }

            CancellationTokenSource cts = new CancellationTokenSource();

            string escritorio = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string documentos = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string imagenes = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
            string musica = Environment.GetFolderPath(Environment.SpecialFolder.MyMusic);
            string videos = Environment.GetFolderPath(Environment.SpecialFolder.MyVideos);

            int countEscritorio = 0;
            int countImg = 0;
            int countMus = 0;
            int countVid = 0;
            int countPapelera = 0;
            string detalleDocumentosFinal = "";                

            Form progresoForm = new Form()
            {
                Text = "Ejecutando Auditoría...",
                Size = new Size(430, 160),
                StartPosition = FormStartPosition.CenterScreen,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false,
                MinimizeBox = false,
                ControlBox = false
            };

            Label lblProgreso = new Label()
            {
                Text = "Iniciando análisis del sistema...",
                Location = new Point(20, 20),
                Size = new Size(375, 23),
                Font = new Font("Segoe UI", 8.5F)
            };

            ProgressBar progressBar = new ProgressBar()
            {
                Location = new Point(20, 50),
                Size = new Size(375, 23),
                Minimum = 0,
                Maximum = 100,
                Value = 10,
                Style = ProgressBarStyle.Continuous
            };

            Button btnCancelar = new Button()
            {
                Text = "Cancelar",
                Location = new Point(160, 95),
                Size = new Size(100, 25),
                BackColor = Color.FromArgb(239, 68, 68),
                ForeColor = Color.White
            };

            btnCancelar.Click += (s, ev) => 
            { 
                cts.Cancel(); 
                btnCancelar.Enabled = false;
                btnCancelar.Text = "Cancelando...";
            };

            progresoForm.Controls.Add(btnCancelar);
            progresoForm.Controls.Add(lblProgreso);
            progresoForm.Controls.Add(progressBar);

            bool esAdminReal = false;
            List<string> serviciosDisponibles = new List<string>();
            bool antivirusActivo = false;
            string avNombreCompleto = "";
            string origenNtpResultado = "";
            string usuarioLogueadoResultado = "";

            List<string> reporteComparativoSoftware = new List<string>();

            this.Enabled = false;

            var backgroundTask = Task.Run(async () =>
            {
                try
                {
                    // 1. Conteo de Archivos
                    cts.Token.ThrowIfCancellationRequested();
                    countEscritorio = ContarArchivosSeguro(escritorio);
                    countImg = ContarArchivosSeguro(imagenes);
                    countMus = ContarArchivosSeguro(musica);
                    countVid = ContarArchivosSeguro(videos);
                    
                    ContarArchivosExcluyendoUSBYSumarOtroDisco(documentos, out detalleDocumentosFinal, cts.Token);

                    // Actualizar UI progreso
                    progresoForm.Invoke(new Action(() =>
                    {
                        lblProgreso.Text = "Verificando red, usuario y privilegios...";
                        progressBar.Value = 25;
                    }));

                    try
                    {
                        string domEnv = Environment.UserDomainName;
                        string usrEnv = Environment.UserName;
                        usuarioLogueadoResultado = $"{domEnv.ToUpper()}\\{usrEnv.ToUpper()}";
                    }
                    catch
                    {
                        usuarioLogueadoResultado = Environment.UserName.ToUpper();
                    }

                    try
                    {
                        ProcessStartInfo psi = new ProcessStartInfo("net", "localgroup Administradores")
                        {
                            RedirectStandardOutput = true,
                            UseShellExecute = false,
                            CreateNoWindow = true
                        };
                        Process? p = Process.Start(psi);
                        if (p != null)
                        {
                            string resultado = p.StandardOutput.ReadToEnd();
                            p.WaitForExit();
                            esAdminReal = resultado.Contains(Environment.UserName);
                        }
                    }
                    catch { }

                    progresoForm.Invoke(new Action(() =>
                    {
                        lblProgreso.Text = "Comparando software con plantilla de Excel...";
                        progressBar.Value = 40;
                    }));

                    // Lectura y comparación con la hoja de Excel
                    try
                    {
                        reporteComparativoSoftware.Clear();
                        string areaSeleccionada = "";
                        progresoForm.Invoke(new Action(() => { areaSeleccionada = cmbCargo.Text; }));

                        List<SoftwarePermitido> appsPermitidas = ObtenerAplicacionesPermitidas(areaSeleccionada);
                        List<ProgramaInstalado> appsReales = ObtenerProgramasInstalados();

                        reporteComparativoSoftware.Add($"=== REPORTE DE SOFTWARE INSTALADO - ÁREA: {areaSeleccionada} ===");
                        reporteComparativoSoftware.Add($"Total programas instalados detectados: {appsReales.Count}\n");
                        
                        foreach (var appInstalada in appsReales)
                        {
                            string LimpiarNombre(string texto)
                            {
                                string sinVersiones = Regex.Replace(texto, @"\b\d+(\.\d+)+\b", "");
                                return Regex.Replace(sinVersiones, @"\s+", " ").Trim();
                            }
                            string nombreInstaladoLimpio = LimpiarNombre(appInstalada.Nombre);

                            var tokensInstalado = nombreInstaladoLimpio.Split(new[] { ' ', '-', '_', '.', '(', ')' }, StringSplitOptions.RemoveEmptyEntries)
                                                                        .Where(w => w.Length > 2)
                                                                        .ToList();

                            var encontrada = appsPermitidas.FirstOrDefault(p =>
                            {
                                string nombreExcelLimpio = LimpiarNombre(p.Nombre);

                                if(nombreInstaladoLimpio.Equals(nombreExcelLimpio, StringComparison.OrdinalIgnoreCase))
                                    return true;

                                var tokensExcel = nombreExcelLimpio.Split(new[] { ' ', '-', '_', '.', '(', ')' }, StringSplitOptions.RemoveEmptyEntries)
                                           .Where(w => w.Length > 2)
                                           .ToList();
        
                                if (tokensExcel.Count == 0 || tokensInstalado.Count == 0) return false;

                                int coincidencias = tokensExcel.Count(tExcel => tokensInstalado.Contains(tExcel, StringComparer.OrdinalIgnoreCase));
                                double porcentaje = (double)coincidencias / tokensExcel.Count;
                                
                                return porcentaje >= 0.75;
                            });
                                
                            if (encontrada != null)
                            {
                                reporteComparativoSoftware.Add($"[AUTORIZADO] {appInstalada.Nombre} + {appInstalada.Version}");
                            }
                            else
                            {
                                reporteComparativoSoftware.Add($"[NO ENLISTADO] {appInstalada.Nombre} + {appInstalada.Version}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        reporteComparativoSoftware.Add($"Error al procesar el listado de software: {ex.Message}");
                    }

                    progresoForm.Invoke(new Action(() =>
                    {
                        lblProgreso.Text = "Comprobando servicios web y antivirus...";
                        progressBar.Value = 60;
                    }));

                    var serviciosAProbar = new (string Nombre, string Url)[]
                    {
                        ("GMAIL", "https://mail.google.com"),
                        ("MEET", "https://meet.google.com"),
                        ("DRIVE", "https://drive.google.com"),
                        ("YOUTUBE", "https://www.youtube.com"),
                        ("OUTLOOK", "https://outlook.live.com"),
                        ("WETRANSFER", "https://wetransfer.com"),
                        ("DROPBOX", "https://www.dropbox.com"),
                        ("BOX", "https://www.box.com"),
                        ("MEGA", "https://mega.io/es")
                    };

                    foreach (var servicio in serviciosAProbar)
                    {
                        cts.Token.ThrowIfCancellationRequested();
                        if (PuedeAccederAUrl(servicio.Url, TimeSpan.FromSeconds(4)))
                        {
                            serviciosDisponibles.Add(servicio.Nombre);
                        }
                    }

                    try
                    {
                        List<string> avsEncontrados = new List<string>();
                        string esetVersion = ObtenerVersionEsetDesdeRegistro();

                        foreach (var serv in ServiceController.GetServices())
                        {
                            string sNameLower = serv.ServiceName.ToLower();
                            string dNameLower = serv.DisplayName.ToLower();
                            if ((sNameLower.Contains("eset") || dNameLower.Contains("eset") || 
                                 sNameLower.Contains("windefend") || dNameLower.Contains("windows defender") || 
                                 sNameLower.Contains("avast") || dNameLower.Contains("avast") || 
                                 sNameLower.Contains("mcafee") || dNameLower.Contains("mcafee") || 
                                 sNameLower.Contains("norton") || dNameLower.Contains("norton") || 
                                 sNameLower.Contains("kaspersky") || dNameLower.Contains("kaspersky")) 
                                && serv.Status == ServiceControllerStatus.Running)
                            {
                                antivirusActivo = true;
                                string nombreAv = serv.DisplayName.ToUpper();
                                if (nombreAv.Contains("ESET") && !string.IsNullOrEmpty(esetVersion))
                                {
                                    avsEncontrados.Add($"{nombreAv} (VER: {esetVersion})");
                                }
                                else
                                {
                                    Match matchVersion = System.Text.RegularExpressions.Regex.Match(nombreAv, @"v?(\d+(\.\d+)+)");
                                    if (matchVersion.Success)
                                    {
                                        string versionEncontrada = matchVersion.Value;
                                        string nombreLimpio = nombreAv.Replace(versionEncontrada, "").Trim();
                                        avsEncontrados.Add($"{nombreLimpio} (VER: {versionEncontrada})");
                                    }
                                    else
                                    {
                                        avsEncontrados.Add(nombreAv);
                                    }
                                }
                            }
                        }

                        if (avsEncontrados.Count > 0)
                        {
                            avNombreCompleto = string.Join(", ", avsEncontrados);
                        }
                    }
                    catch { }

                    try
                    {
                        ProcessStartInfo psiNtp = new ProcessStartInfo("w32tm", "/query /status")
                        {
                            RedirectStandardOutput = true,
                            UseShellExecute = false,
                            CreateNoWindow = true
                        };
                        Process? pNtp = Process.Start(psiNtp);
                        if (pNtp != null)
                        {
                            string resNtp = pNtp.StandardOutput.ReadToEnd();
                            pNtp.WaitForExit();

                            foreach (string linea in resNtp.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
                            {
                                if (linea.Contains("Origen") || linea.Contains("Source"))
                                {
                                    origenNtpResultado = "NTP";
                                    break;
                                }
                            }
                        }
                        if (string.IsNullOrEmpty(origenNtpResultado)) origenNtpResultado = "NTP";
                    }
                    catch { origenNtpResultado = "NTP"; }

                    progresoForm.Invoke(new Action(() =>
                    {
                        lblProgreso.Text = "Analizando archivos y documentos...";
                        progressBar.Value = 85;
                    }));

                    try
                    {
                        int totalPapeleraUsuario = 0;

                        string sidUsuarioActual = System.Security.Principal.WindowsIdentity.GetCurrent().User?.Value ?? "";
                        if (!string.IsNullOrEmpty(sidUsuarioActual))
                        {
                            foreach (DriveInfo drive in DriveInfo.GetDrives())
                            {
                                if (drive.IsReady && drive.DriveType == DriveType.Fixed)
                                {
                                    string papeleraPath = Path.Combine(drive.RootDirectory.FullName, "$Recycle.Bin", sidUsuarioActual);
                                    if (Directory.Exists(papeleraPath))
                                    {
                                        totalPapeleraUsuario += ContarArchivosPapeleraSeguro(papeleraPath);
                                    }
                                }
                            }
                        }
                        countPapelera = totalPapeleraUsuario;
                    }
                    catch 
                    { 
                        countPapelera = 0; 
                    }

                    progresoForm.Invoke(new Action(() => progressBar.Value = 100));
                }
                catch (OperationCanceledException)
                {
                    // Cancelación controlada
                }
                finally
                {
                    // Asegurar que la ventana de progreso se cierre siempre desde el hilo de UI
                    if (!progresoForm.IsDisposed)
                    {
                        progresoForm.Invoke(new Action(() => {
                            if (!progresoForm.IsDisposed) progresoForm.Close();
                        }));
                    }
                }
            }, cts.Token);

            // Mostrar el dialogo de progreso bloqueando interacciones ajenas hasta que culmine o se cancele
            progresoForm.ShowDialog(this);

            // Rehabilitar ventana principal
            this.Enabled = true;

            // Si el usuario canceló, abortamos la asignación de campos para evitar datos vacíos o parciales
            if (cts.IsCancellationRequested)
            {
                return;
            }

            try
            {
                txtUsuarioDominio.Text = usuarioLogueadoResultado;
                txtNombreEquipo.Text = Environment.MachineName;
                string nombreWin = "WINDOWS";
                try
                {
                    using (Microsoft.Win32.RegistryKey? key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion"))
                    {
                        if (key != null)
                        {
                            int buildNumber = 0;
                            string buildStr = key.GetValue("CurrentBuild") as string ?? "";
                            int.TryParse(buildStr, out buildNumber);

                            string tipoOS = (buildNumber >= 22000) ? "WINDOWS 11" : "MICROSOFT WINDOWS 10";
                            string productNameReg = key.GetValue("ProductName") as string ?? "";

                            string sufijoEdicion = "PRO";
                            if (productNameReg.ToUpper().Contains("HOME")) sufijoEdicion = "HOME";
                            else if (productNameReg.ToUpper().Contains("ENTERPRISE")) sufijoEdicion = "ENTERPRISE";
                            else if (productNameReg.ToUpper().Contains("EDUCATION")) sufijoEdicion = "EDUCATION";

                            nombreWin = $"{tipoOS} {sufijoEdicion}";
                            if (!string.IsNullOrEmpty(buildStr)) nombreWin += $" (BUILD {buildStr})";
                        }
                    }
                }
                catch { }
                txtVersionSO.Text = nombreWin;

                chkAdminSi.Checked = esAdminReal;
                chkAdminNo.Checked = !esAdminReal;

                bool tieneUsbHabilitado = DetectarUsbHabilitado();
                chkUsbSi.Checked = tieneUsbHabilitado;
                chkUsbNo.Checked = !tieneUsbHabilitado;

                bool tieneCd = DriveInfo.GetDrives().Any(d => d.DriveType == DriveType.CDRom);
                chkCdSi.Checked = tieneCd;
                chkCdNo.Checked = !tieneCd;

                bool tieneBluetooth = false;
                try { tieneBluetooth = ServiceController.GetServices().Any(s => s.ServiceName.ToLower().Contains("bthserv")); } catch { }
                chkBtSi.Checked = tieneBluetooth;
                chkBtNo.Checked = !tieneBluetooth;

                bool tieneWifi = false;
                try { tieneWifi = NetworkInterface.GetAllNetworkInterfaces().Any(n => n.NetworkInterfaceType == NetworkInterfaceType.Wireless80211); } catch { }
                chkWifiSi.Checked = tieneWifi;
                chkWifiNo.Checked = !tieneWifi;

                if (serviciosDisponibles.Any(s => s == "GMAIL" || s == "MEET" || s == "DRIVE" || s == "YOUTUBE"))
                {
                    chkGmailSi.Checked = true;
                    txtGmailDetalle.Text = string.Join(", ", serviciosDisponibles.Where(s => s == "GMAIL" || s == "MEET" || s == "DRIVE" || s == "YOUTUBE"));
                }
                else
                {
                    chkGmailNo.Checked = true;
                    txtGmailDetalle.Text = "Ninguno";
                }

                chkOutlookSi.Checked = serviciosDisponibles.Contains("OUTLOOK");
                chkOutlookNo.Checked = !serviciosDisponibles.Contains("OUTLOOK");

                if (serviciosDisponibles.Any(s => s == "DRIVE" || s == "WETRANSFER" || s == "DROPBOX" || s == "BOX" || s == "MEGA"))
                {
                    chkCloudSi.Checked = true;
                    txtCloudDetalle.Text = string.Join(", ", serviciosDisponibles.Where(s => s == "DRIVE" || s == "WETRANSFER" || s == "DROPBOX" || s == "BOX" || s == "MEGA"));
                }
                else
                {
                    chkCloudNo.Checked = true;
                    txtCloudDetalle.Text = "Sin nube detectada";
                }

                if (antivirusActivo)
                {
                    chkAvInstSi.Checked = true;
                    txtAvInstDetalle.Text = avNombreCompleto;
                }
                else
                {
                    chkAvInstNo.Checked = true;
                    txtAvInstDetalle.Text = "No detectado";
                }
                txtAvFechaUpd.Text = DateTime.Now.ToString("yyyy-MM-dd");
                chkAppNo.Checked = true;

                bool vpnDetectada = false;
                try { vpnDetectada = ServiceController.GetServices().Any(s => s.ServiceName.ToLower().Contains("forticlient") || s.ServiceName.ToLower().Contains("globalprotect") || s.ServiceName.ToLower().Contains("openvpn")); } catch { }
                chkVpnSi.Checked = vpnDetectada;
                chkVpnNo.Checked = !vpnDetectada;

                chkEscritorioSi.Checked = true;
                txtEscritorioDetalle.Text = $"{countEscritorio} archivos";

                chkDocSi.Checked = true;
                txtDocDetalle.Text = detalleDocumentosFinal;

                chkImgSi.Checked = countImg > 0; chkImgNo.Checked = countImg == 0; txtImgDetalle.Text = $"{countImg} archivos";
                chkMusSi.Checked = countMus > 0; chkMusNo.Checked = countMus == 0; txtMusDetalle.Text = $"{countMus} archivos";
                chkVidSi.Checked = countVid > 0; chkVidNo.Checked = countVid == 0; txtVidDetalle.Text = $"{countVid} archivos";

                chkPapeleraSi.Checked = true;
                txtPapeleraDetalle.Text = $"{countPapelera} archivos";
                chkCifradoNo.Checked = true;
                chkNtpSi.Checked = true;
                txtNtpDetalle.Text = origenNtpResultado;

                // Mostrar la pestaña emergente con el resultado comparativo del software
                MostrarVentanaEmergenteSoftware(reporteComparativoSoftware);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ocurrió un error al procesar los resultados: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private int ContarArchivosPapeleraSeguro(string rutaPapelera)
        {
            int contador = 0;
            try
            {
                var opciones = new EnumerationOptions
                {
                    IgnoreInaccessible = true,
                    RecurseSubdirectories = true,
                    AttributesToSkip = FileAttributes.None 
                };

                contador = Directory.EnumerateFiles(rutaPapelera, "$R*", opciones)
                    .Where(f => !Path.GetFileName(f).Equals("desktop.ini", StringComparison.OrdinalIgnoreCase))
                    .Count();
            }
            catch { }
            return contador;
        }

        // Métodos de lectura y visualización de software por área
        private List<SoftwarePermitido> ObtenerAplicacionesPermitidas(string areaSeleccionada)
        {
            List<SoftwarePermitido> permitidas = new List<SoftwarePermitido>();
            // Asegúrate de que el nombre del archivo de Excel con el listado esté en la carpeta del .exe
            string nombreRecurso = "AuditoriaEquipos.Resultados.GA-FO-18 CATALOGO DE SOFTWARE UTILITARIO.xlsx"; 

            try
            {
                using (MemoryStream ms = ManejadorPlantillas.ObtenerStreamPlantilla(nombreRecurso))
                {
                    using (var workbook = new XLWorkbook(ms))
                    {
                        var hoja = workbook.Worksheets.FirstOrDefault(w => w.Name.Trim().Equals(areaSeleccionada.Trim(), StringComparison.OrdinalIgnoreCase));
                
                        // Si no encuentra una hoja con el nombre exacto, toma la primera por defecto
                        if (hoja == null && workbook.Worksheets.Count > 0)
                        {
                            hoja = workbook.Worksheet(1);
                        }

                        if (hoja != null)
                        {
                            var filas = hoja.RowsUsed().Skip(1); // Omitir encabezado

                            foreach (var fila in filas)
                            {
                                // Columna 2: NOMBRE DEL SOFTWARE O APLICATIVO
                                string nombreApp = fila.Cell(2).GetString().Trim().ToUpper();
                                // Columna 3: VERSIÓN
                                string versionEsp = fila.Cell(3).GetString().Trim().ToUpper();

                                if (!string.IsNullOrEmpty(nombreApp))
                                {
                                    permitidas.Add(new SoftwarePermitido
                                    {
                                        Nombre = nombreApp,
                                        VersionEsperada = versionEsp,
                                        Area = areaSeleccionada
                                    });
                                }
                            }
                        }      
                    }
                }    
            }
            catch(Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error al cargar la plantilla desde recursos: {ex.Message}");
            }

            return permitidas;
        }

        private List<ProgramaInstalado> ObtenerProgramasInstalados()
        {
            var lista = new List<ProgramaInstalado>();
            string[] rutasReg = {
                @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                @"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            };

            foreach (string ruta in rutasReg)
            {
                try
                {
                    using (RegistryKey? key = Registry.LocalMachine.OpenSubKey(ruta))
                    {
                        if (key != null)
                        {
                            foreach (string subKeyName in key.GetSubKeyNames())
                            {
                                using (RegistryKey? subKey = key.OpenSubKey(subKeyName))
                                {
                                    if (subKey != null)
                                    {
                                        string name = subKey.GetValue("DisplayName") as string ?? "";
                                        string ver = subKey.GetValue("DisplayVersion") as string ?? "";
                                        if (string.IsNullOrWhiteSpace(name)) continue;

                                        // 1. Omitir componentes ocultos del sistema
                                        int systemComponent = 0;
                                        try { systemComponent = Convert.ToInt32(subKey.GetValue("SystemComponent") ?? 0); } catch { }
                                        if (systemComponent == 1) continue;

                                        // 2. Omitir subcomponentes o paquetes dependientes de otros
                                        string parentKey = subKey.GetValue("ParentKeyName") as string ?? "";
                                        if (!string.IsNullOrEmpty(parentKey)) continue;

                                        string nameUpper = name.ToUpper();

                                        // 3. Filtrar drivers, actualizaciones de Windows y parches KB
                                        bool esDriverOActualizacion = nameUpper.Contains("DRIVER") || 
                                                                    nameUpper.Contains("SECURITY UPDATE") || 
                                                                    nameUpper.Contains("UPDATE FOR WINDOWS") || 
                                                                    nameUpper.Contains("HOTFIX") ||
                                                                    Regex.IsMatch(nameUpper, @"\bKB\d{5,}\b");

                                        if (!esDriverOActualizacion)
                                        {
                                            // Evitar duplicados exactos si se registran en ambas rutas del registro
                                            if (!lista.Any(p => p.Nombre == nameUpper))
                                            {
                                                lista.Add(new ProgramaInstalado { Nombre = nameUpper, Version = ver.Trim().ToUpper() });
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch { }
            }
            return lista;
        }

        private void MostrarVentanaEmergenteSoftware(List<string> lineasReporte)
        {
            Form popForm = new Form()
            {
                Text = "Resultado de Validación de Software vs. Plantilla Excel",
                Size = new Size(650, 450),
                StartPosition = FormStartPosition.CenterParent,
                FormBorderStyle = FormBorderStyle.Sizable
            };

            Panel panelInferior = new Panel()
            {
                Height = 55,
                Dock = DockStyle.Bottom,
                BackColor = Color.FromArgb(240, 242, 245)
            };

            Button btnExportar = new Button()
            {
                Text = "Exportar a .txt",
                Size = new Size(150, 35),
                Location = new Point(15, 10),
                BackColor = Color.FromArgb(59, 130, 246),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };

            btnExportar.Click += async (s, e) =>
            {
                string directorioEjecutable = AppDomain.CurrentDomain.BaseDirectory;
                string carpetaResultados = Path.Combine(directorioEjecutable, "Resultados");
                Directory.CreateDirectory(carpetaResultados);

                string responsable = txtRespEquipo.Text.Trim();
                string nombrePersona = string.IsNullOrWhiteSpace(responsable) ? "Usuario" : responsable.ToUpper();
                foreach (char c in Path.GetInvalidFileNameChars())
                {
                    nombrePersona = nombrePersona.Replace(c, '_');
                }

                // Se conserva la nomenclatura exacta de tu archivo .txt
                string nombreArchivoTxt = $"GA-FO-18 CATALOGO DE SOFTWARE UTILITARIO - {nombrePersona}.txt";
                string rutaTxtLocal = Path.Combine(carpetaResultados, nombreArchivoTxt);

                try
                {
                    // Guardar las aplicaciones seleccionadas en el .txt
                    using (StreamWriter sw = new StreamWriter(rutaTxtLocal, false, Encoding.UTF8))
                    {
                        sw.WriteLine($"REPORTE DE SOFTWARE INSTALADO - {nombrePersona.ToUpper()}");
                        sw.WriteLine(new string('-', 50));

                        foreach (var linea in lineasReporte)
                        {
                            sw.WriteLine($"- {linea}");
                        }
                    }

                    MessageBox.Show($"Archivo TXT guardado correctamente en:\n{rutaTxtLocal}", 
                        "Éxito Local", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // 2. Subir el archivo .txt a Google Drive
                    var exporter = new AuditExcelExporter();
                    string urlScript = "https://script.google.com/macros/s/AKfycbxM2VqT9AY2FEUKxqGfU4v_pSKNj37ZfeKslxDnQx3jpMVdZK1vafTDs2M9D0fNw904/exec";

                    bool subido = await exporter.SubirReporteADriveAsync(
                        rutaExcelLocal: rutaTxtLocal,
                        cargo: cmbCargo.Text.Trim(),
                        tipoArchivo: "GA-FO-18 CATALOGO DE SOFTWARE UTILITARIO",
                        urlWebAppScript: urlScript
                    );

                    if (subido)
                    {
                        MessageBox.Show("¡El archivo TXT de software utilitario también se subió exitosamente a Google Drive!", 
                            "Éxito en Drive", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("El archivo TXT se guardó localmente, pero hubo un detalle al subirlo a Google Drive.", 
                            "Advertencia de Red", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    popForm.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error al procesar el archivo TXT: {ex.Message}", 
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            };

            panelInferior.Controls.Add(btnExportar);

            RichTextBox txtBox = new RichTextBox()
            {
                Multiline = true,
                ReadOnly = true,
                Dock = DockStyle.Fill,
                Font = new Font("Arial", 9.5F),
            };

            popForm.Controls.Add(txtBox);
            popForm.Controls.Add(panelInferior);

            foreach (string linea in lineasReporte)
            {
                if (linea.Contains("[AUTORIZADO]"))
                {
                    txtBox.SelectionColor = Color.FromArgb(16, 185, 129);
                }
                else if (linea.Contains("[NO ENLISTADO]"))
                {
                    txtBox.SelectionColor = Color.FromArgb(239, 68, 68);
                }
                else if (linea.Contains("=== REPORTE"))
                {
                    txtBox.SelectionColor = Color.FromArgb(59, 130, 246);
                }
                else
                {
                    txtBox.SelectionColor = Color.Black;
                }

                txtBox.AppendText(linea + Environment.NewLine);
            }

            popForm.ShowDialog(this);
        }

        private string ObtenerVersionEsetDesdeRegistro()
        {
            string version = "";
            string[] rutasReg = new string[]
            {
                @"SOFTWARE\ESET\ESET Security\CurrentVersion\Info",
                @"SOFTWARE\ESET\ESET Endpoint Security\CurrentVersion\Info",
                @"SOFTWARE\WOW6432Node\ESET\ESET Security\CurrentVersion\Info",
                @"SOFTWARE\WOW6432Node\ESET\ESET Endpoint Security\CurrentVersion\Info"
            };

            foreach (string ruta in rutasReg)
            {
                try
                {
                    using (RegistryKey? key = Registry.LocalMachine.OpenSubKey(ruta))
                    {
                        if (key != null)
                        {
                            object? verObj = key.GetValue("VersionString") ?? key.GetValue("ProductVersion") ?? key.GetValue("Version");
                            if (verObj != null)
                            {
                                version = verObj.ToString() ?? "";
                                break;
                            }
                        }
                    }
                }
                catch { }
            }
            return version ?? "";
        }

        private bool PuedeAccederAUrl(string url, TimeSpan timeout)
        {
            if (string.IsNullOrWhiteSpace(url)) return false;
            try
            {
                using (var client = new System.Net.Http.HttpClient())
                {
                    client.Timeout = timeout;
                    using (var response = client.GetAsync(url).GetAwaiter().GetResult())
                    {
                        return response.StatusCode < System.Net.HttpStatusCode.InternalServerError;
                    }
                }
            }
            catch { return false; }
        }

        private int ContarArchivosSeguro(string rutaDirectorio)
        {
            
            if (string.IsNullOrEmpty(rutaDirectorio) || !Directory.Exists(rutaDirectorio)) return 0;
            int contador = 0;
            try
            {
                var opciones = new EnumerationOptions
                {
                    IgnoreInaccessible = true,
                    RecurseSubdirectories = true,
                };
                contador = Directory.EnumerateFiles(rutaDirectorio, "*", opciones).Count();
            }
            catch { }
            return contador;
        }

        private int ContarArchivosExcluyendoUSBYSumarOtroDisco(string rutaDocumentosOriginal, out string detalleResultado, CancellationToken token)
        {
            int totalArchivos = 0;
            var discosProcesados = new List<string>();
            string raizDocumentos = string.IsNullOrEmpty(rutaDocumentosOriginal) ? "" : (Path.GetPathRoot(Path.GetFullPath(rutaDocumentosOriginal)) ?? "");

            if (!string.IsNullOrEmpty(rutaDocumentosOriginal) && Directory.Exists(rutaDocumentosOriginal))
            {
                ContarArchivosRecursivoSeguro(rutaDocumentosOriginal, ref totalArchivos, token);
            }

            if (!string.IsNullOrEmpty(raizDocumentos))
            {
                discosProcesados.Add($"{raizDocumentos.TrimEnd('\\')} ({totalArchivos} archivos)");
            }

            foreach (DriveInfo drive in DriveInfo.GetDrives())
            {
                token.ThrowIfCancellationRequested();
                if (drive.IsReady && drive.DriveType == DriveType.Fixed)
                {
                    string raizDisco = drive.RootDirectory.FullName;
                    if (!string.Equals(raizDisco, raizDocumentos, StringComparison.OrdinalIgnoreCase))
                    {
                        int archivosOtroDisco = 0;
                        ContarArchivosRecursivoSeguro(raizDisco, ref archivosOtroDisco, token);
                        totalArchivos += archivosOtroDisco;
                        discosProcesados.Add($"{raizDisco.TrimEnd('\\')} ({archivosOtroDisco} archivos)");
                    }
                }
            }

            detalleResultado = $"{totalArchivos} archivos ({string.Join(" + ", discosProcesados)})";
            return totalArchivos;
        }

        private void ContarArchivosRecursivoSeguro(string directorio, ref int contadorArchivos, CancellationToken token)
        {
            if (string.IsNullOrEmpty(directorio) || !Directory.Exists(directorio)) return;
            try
            {
                token.ThrowIfCancellationRequested();

                try
                {
                    string[] archivos = Directory.GetFiles(directorio);
                    contadorArchivos += archivos.Length;
                }
                catch { } 

                string[] subDirs = Array.Empty<string>();
                try
                {
                    subDirs = Directory.GetDirectories(directorio);
                }
                catch { }

                foreach (string subDir in subDirs)
                {
                    token.ThrowIfCancellationRequested();
                    try
                    {
                        string nombreDir = Path.GetFileName(subDir);
                        if (nombreDir.Equals("$Recycle.Bin", StringComparison.OrdinalIgnoreCase) ||
                            nombreDir.Equals("System Volume Information", StringComparison.OrdinalIgnoreCase)) continue;

                        ContarArchivosRecursivoSeguro(subDir, ref contadorArchivos, token);
                    }
                    catch { }
                }
            }
            catch { }
        }
    }
    
}