using System;
using System.Windows.Forms;
using System.Diagnostics;

namespace LocalizacionColombia
{
    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {             
            try
            {//Conexion con SAP
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Conexion oConnection = new Conexion();
                oConnection.SetApplication();
                oConnection.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                try
                {
                    string versionNueva = string.Empty;
                    oConnection.SBO_Application.SetStatusBarMessage("Validacion de instalacion addon Localizacion Colombia", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                    System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
                    FileVersionInfo fvi = FileVersionInfo.GetVersionInfo(assembly.Location);
                    versionNueva = fvi.FileVersion;

                    //Instalacion
                    string versionActual = oConnection.GetVersionAddonBD();
                    long verNueva = VersionNumberCompareString(versionNueva);
                    if (verNueva == -1)
                    {
                        //La versión que trae el archivo no tiene una version comparable se cancela la instalación
                        string mensajeError = string.Format("Version addon {0} no coincide con una numeracion del addon valida", versionNueva);
                        oConnection.SBO_Application.SetStatusBarMessage(mensajeError, SAPbouiCOM.BoMessageTime.bmt_Long, true);
                        MessageBox.Show(mensajeError);
                    }
                    long verActual = VersionNumberCompareString(versionActual);
                    if (string.IsNullOrEmpty(versionActual) || verNueva > verActual)
                    {
                        //Se debe instalar el addon
                        oConnection.SBO_Application.SetStatusBarMessage("Instalacion addon Localizacion Colombia", SAPbouiCOM.BoMessageTime.bmt_Long, false);
                        oConnection.CargaCamposUsuarioDBSAP(versionNueva);
                        //Creacion de tablas, campos, informes, codigos de transaccion, categorias de consultas, consultas, 
                        oConnection.añadirComponentes();
                    }                 
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show(ex.Message);
                }
                Business oNegocio = null;
                oNegocio = new Business(oConnection.oCompany, oConnection.SBO_Application);
                oConnection.SBO_Application.SetStatusBarMessage("El addon de Localizacion Colombia se incio correctamente", SAPbouiCOM.BoMessageTime.bmt_Short, false);
                Application.Run();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        /// <summary>
        /// Compara la version del addon con versiones previamente instalada
        /// </summary>
        /// <param name="versionNumber">
        ///     String que contiene la version a comparar
        /// </param>
        /// <param name="MaxWidth1">
        ///     Maximo tamaño de la version
        /// </param>
        /// <returns>
        ///     Un valor entero que representa el resultado de la comparacion de las versiones del addon
        /// </returns>
        private static long VersionNumberCompareString(string versionNumber, int MaxWidth1 = 3)
        {
            try
            {
                string result = null;
                int puntos = versionNumber.Split('.').Length;
                var integerValues = versionNumber.Split('.');
                for (int i = 0; i < puntos; i++)
                {
                    result += integerValues[i].PadLeft(MaxWidth1, '0'); ;
                }
                return long.Parse(result);
            }
            catch (Exception)
            {
                return -1;
            }
        }

        /// <summary>
        /// Metodo que sirve para la identificacion de eventos dentro de SAP
        /// </summary>
        /// <param name="EventType">
        ///     Objeto que contiene informacion relevante del evento
        /// </param>
        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }
    }
}
