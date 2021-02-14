using Autodesk.Revit.UI;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

namespace ModelessForm_ExternalEvent
{
    /// <summary>
    /// La classe della nostra finestra di dialogo non modale.
    /// </summary>
    /// <remarks>
    /// Oltre ad altri metodi, ha un metodo per ogni pulsante di comando. 
    /// In ognuno di questi metodi non viene fatto nient'altro che il sollevamento
    /// di un evento esterno con una richiesta specifica impostata nel gestore delle richieste.
    /// </remarks>
    /// 
    public partial class ModelessForm : Form
    {
        // In questo esempio, la finestra di dialogo possiede il gestore e gli oggetti evento, 
        // ma non è un requisito. Potrebbero anche essere proprietà statiche dell'applicazione.

        #region Private data members
        private RequestHandler m_Handler;
        private ExternalEvent m_ExEvent;
        
        // Dichiarazione della stringa dell'elemento scelto nella ComboBox
        private string _elementSelectedComboBox = "";

        // Dichiarazione dell'indice  dell'elemento scelto nella ComboBox
        private int _indexSelectedComboBox = -1;
        #endregion

        #region Class public property
        /// <summary>
        /// Proprietà pubblica per accedere al valore della richiesta corrente
        /// </summary>
        public RequestHandler RequestHandler
        {
            get { return m_Handler; }
        }

        /// <summary>
        /// Proprietà pubblica per accedere al valore della richiesta corrente
        /// </summary>
        public string ElementSelectedComboBox
        {
            get { return _elementSelectedComboBox; }
        }

        /// <summary>
        /// Proprietà pubblica per accedere al valore della richiesta corrente
        /// </summary>
        public int IndexSelectedComboBox
        {
            get { return _indexSelectedComboBox; }
        }

        #endregion

        #region Class public method
        /// <summary>
        ///   Costruttore della finestra di dialogo
        /// </summary>
        /// 
        public ModelessForm(ExternalEvent exEvent, RequestHandler handler)
        {
            InitializeComponent();
            m_Handler = handler;
            m_ExEvent = exEvent;

        }
        #endregion

        /// <summary>
        /// Modulo gestore eventi chiuso
        /// </summary>
        /// <param name="e"></param>
        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            // possediamo sia l'evento che il gestore
            // dovremmo eliminarlo prima di chiudere la finestra

            m_ExEvent.Dispose();
            m_ExEvent = null;
            m_Handler = null;

            // non dimenticare di chiamare la classe base
            base.OnFormClosed(e);
        }

        /// <summary>
        ///   Attivatore / disattivatore del controllo
        /// </summary>
        ///
        private void EnableCommands(bool status)
        {
            foreach (Control ctrl in this.Controls)
            {
                ctrl.Enabled = status;
            }
            if (!status)
            {
                this.exitButton.Enabled = true;
            }
        }

        /// <summary>
        ///   Un metodo di supporto privato per effettuare una richiesta 
        ///   e allo stesso tempo mettere la finestra di dialogo in sospensione.
        /// </summary>
        /// <remarks>
        ///   Ci si aspetta che il processo che esegue la richiesta 
        ///   (l'helper Idling in questo caso particolare) 
        ///   riattivi anche la finestra di dialogo dopo aver terminato l'esecuzione.
        /// </remarks>
        ///
        private void MakeRequest(RequestId request)
        {
            App.thisApp.DontShowFormTop();
            m_Handler.Request.Make(request);
            m_ExEvent.Raise();
            DozeOff();            
        }


        /// <summary>
        ///   DozeOff -> disabilita tutti i controlli (tranne il pulsante Esci)
        /// </summary>
        /// 
        private void DozeOff()
        {
            EnableCommands(false);
        }

        /// <summary>
        ///   WakeUp -> abilita tutti i controlli
        /// </summary>
        /// 
        public void WakeUp()
        {
            EnableCommands(true);
        }

        /// <summary>
        ///   Metodo di interazione con la finestra di dialogo
        /// </summary>
        /// 
        private void ModelessForm_Load(object sender, EventArgs e)
        {
            // Operazioni iniziali
            MakeRequest(RequestId.Initial);
        }

        /// <summary>
        ///   Metodo per riempimento della ComboBox
        /// </summary>
        ///
        public void SetComboBox()
        {
            List<string> listParameters = m_Handler.ValuesForComboBox;
            foreach (string item in listParameters)
            {
                comboBox1.Items.Add(item);
            }            
        }

        /// <summary>
        ///   Metodo per catturare l'elemento scelto nella ComboBox
        /// </summary>
        ///
        private void showSelectedButton_SelectedIndexChanged(object sender, EventArgs e)
        {
            //listBox1.DataSource = null;
            MakeRequest(RequestId.ChangeComboBox);
        }


        /// <summary>
        ///   Metodo per catturare l'elemento scelto nella ComboBox
        /// </summary>
        ///
        public string GetComboBox()
        {
            _elementSelectedComboBox = comboBox1.SelectedItem as string;
            _indexSelectedComboBox = comboBox1.SelectedIndex + 1;
            comboBox1.Text = _elementSelectedComboBox;
            return _elementSelectedComboBox;
        }



        /// <summary>
        ///   Metodo che riempie il LISTBOX
        /// </summary>
        ///
        public void SetListBox()
        {
            listBox1.DataSource = m_Handler.ListParameters;
        }

        /// <summary>
        ///   Exit - chiude la finestra di dialogo
        /// </summary>
        /// 
        private void exitButton_Click_1(object sender, EventArgs e)
        {
            Close();
        }
    }  // class
}
