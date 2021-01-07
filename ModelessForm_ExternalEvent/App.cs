#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Office.Interop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Windows.Media.Imaging;

#endregion

namespace ModelessForm_ExternalEvent
{
    /// <summary>
    /// Implementa l'interfaccia del componente aggiuntivo Revit IExternalApplication
    /// </summary>
    /// 
    public class App : IExternalApplication
    {
        // Instanza della classe 
        internal static App thisApp = null;
        
        // Instanza della finestra di dialogo non modale
        private ModelessForm m_MyForm;

        #region IExternalApplication Members
        /// <summary>
        /// Implementa l'evento OnStartup
        /// </summary>
        /// <param name="application"></param>
        /// <returns></returns>
        public Result OnStartup(UIControlledApplication application)
        {
            m_MyForm = null;   // nessun dialogo ancora necessario; il comando lo porterà

            thisApp = this;  // accesso statico a questa istanza dell'applicazione

            // Metodo per aggiungere un Tab e un RibbonPanel
            RibbonPanel ribbonPanel = RibbonPanel(application);
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;

            // Creazione del Button da inserire nel Tab
            if(ribbonPanel.AddItem(
                new PushButtonData("ModelessForm_ExternalEvent", "App", thisAssemblyPath, "ModelessForm_ExternalEvent.Command"))
                is PushButton button)
            {
                // ToolTip mostrato
                button.ToolTip = "App base";
                // Icona del Button
                Uri uriImage = new Uri("C:\\DatiLDB\\Progetti_REVIT\\13.Winform\\ModelessForm_ExternalEvent - 2021\\ModelessForm_ExternalEvent\\Resources\\revit_small.png");
                BitmapImage image = new BitmapImage(uriImage);
                button.LargeImage = image;
            };
            return Result.Succeeded;
        }

        /// <summary>
        /// Implementa l'evento OnShutDown
        /// </summary>
        /// <param name="application"></param>
        /// <returns></returns>

        public Result OnShutdown(UIControlledApplication application)
        {
            if (m_MyForm != null && m_MyForm.Visible)
            {
                m_MyForm.Close();
            }

            return Result.Succeeded;
        }

        /// <summary>
        ///   Questo metodo crea e mostra una finestra di dialogo non modale, a meno che non esista già.
        /// </summary>
        /// <remarks>
        ///   Il comando esterno lo richiama su richiesta dell'utente finale
        /// </remarks>
        /// 
        public void ShowForm(UIApplication uiapp)
        {
            // If we do not have a dialog yet, create and show it
            if (m_MyForm == null || m_MyForm.IsDisposed)
            {
                
                // Un nuovo gestore per gestire l'invio delle richieste tramite la finestra di dialogo
                RequestHandler handler = new RequestHandler();

                // Evento esterno per la finestra di dialogo da utilizzare (per inviare richieste)
                ExternalEvent exEvent = ExternalEvent.Create(handler);

                // Diamo gli oggetti alla nuova finestra di dialogo. 
                // La finestra di dialogo diventa il proprietario responsabile della loro disposizione, alla fine.
                m_MyForm = new ModelessForm(exEvent, handler);
                m_MyForm.Show();
                m_MyForm.TopMost = true;
            }
        }

        /// <summary>
        ///   Restituisce la finestra di dialogo.
        /// </summary>
        /// 
        public ModelessForm RetriveForm()
        {
            return m_MyForm;            
        }

        /// <summary>
        ///   Rende la finestra di dialogo secondaria
        /// </summary>
        /// 
        public void DontShowFormTop()
        {
            m_MyForm.TopMost = false;
        }

        /// <summary>
        ///   Rende la finestra di dialogo primaria
        /// </summary>
        /// 
        public void ShowFormTop()
        {
            m_MyForm.TopMost = true;
        }

        /// <summary>
        ///   Riattivare la finestra di dialogo dal suo stato di attesa.
        /// </summary>
        /// 
        public void WakeFormUp()
        {
            if (m_MyForm != null)
            {
                m_MyForm.WakeUp();                
            }
        }
        #endregion

        #region Ribbon Panel

        public RibbonPanel RibbonPanel(UIControlledApplication uiapp)
        {
            // Nome del Tab
            string tab = "Bold";

            // Dichiara e inizializza un RibbonPanel vuoto
            RibbonPanel ribbonPanel = null;

            // Prova a creare un Ribbon Tab
            try
            {
                uiapp.CreateRibbonTab(tab);
            }
            catch (Exception ex)
            {
                Utils.HandleError(ex);
            }

            // Prova a creare un RibbonPanel
            try
            {
                RibbonPanel panel = uiapp.CreateRibbonPanel(tab, "Index");
            }
            catch (Exception ex)
            {
                Utils.HandleError(ex);
            }

            // Verifica se il tab del Panel esiste gia'
            List<RibbonPanel> panels = uiapp.GetRibbonPanels(tab);
            foreach (RibbonPanel p in panels.Where(x => x.Name == "Index"))
            {
                ribbonPanel = p;
            }

            return ribbonPanel;
        }

        #endregion


    }
}
