//
// (C) Copyright 2003-2017 by Autodesk, Inc.
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted,
// provided that the above copyright notice appears in all copies and
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE. AUTODESK, INC.
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
//
// Use, duplication, or disclosure by the U.S. Government is subject to
// restrictions set forth in FAR 52.227-19 (Commercial Computer
// Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
// (Rights in Technical Data and Computer Software), as applicable.
//
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace ModelessForm_ExternalEvent
{
    /// <summary>
    ///   Una classe con metodi per eseguire le richieste effettuate dall'utente della finestra di dialogo.
    /// </summary>
    /// 
    public class RequestHandler : IExternalEventHandler  // Un'istanza di una classe che implementa questa interfaccia verrà registrata prima con Revit e ogni volta che viene generato l'evento esterno corrispondente, verrà richiamato il metodo Execute di questa interfaccia.
    {
        #region Private data members
        // Il valore dell'ultima richiesta effettuata dal modulo non modale
        private Request m_request = new Request();

        // Dichiara un'istanza di questa classe
        private ModelessForm _modelessForm;

        // Path del file.txt
        private string _pathTxt = "";

        // Path del file.txt
        private string _pathTxt2 = "";

        // Valori della ComboBox
        private List<string> _titles;

        // Valori della ListBox
        private List<string[]> _parameters;

        // Valore stringa scelto dalla ComboBox
        private string _selectemItemComboBox;

        // Valori dei parametri da passare alla ListBOx
        private ArrayList _listParameters;
        #endregion

        #region Class public property
        /// <summary>
        /// Proprietà pubblica per accedere al valore della richiesta corrente
        /// </summary>
        public Request Request
        {
            get { return m_request; }
        }

        /// <summary>
        /// Proprietà pubblica per accedere al valore della richiesta corrente
        /// </summary>
        public List<string> ValuesForComboBox
        {
            get { return _titles; }
        }

        /// <summary>
        /// Proprietà pubblica per accedere al valore della richiesta corrente
        /// </summary>
        public List<string[]> ValuesForListBox
        {
            get { return _parameters; }
        }

        /// <summary>
        /// Proprietà pubblica per accedere al valore della richiesta corrente
        /// </summary>
        public ArrayList ListParameters
        {
            get { return _listParameters; }
        }
        #endregion

        #region Class public method
        /// <summary>
        /// Costruttore di default di RequestHandler
        /// </summary>
        public RequestHandler()
        {
            // Costruisce i membri dei dati per le proprietà
            _pathTxt = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) 
                + @"\Esperimenti_Revit\ParameterGroups.txt";
            _pathTxt2 = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                 + @"\Esperimenti_Revit\ParameterGroups2.txt";
            _titles = new List<string>();
            _parameters = new List<string[]>();
            _listParameters = new ArrayList();
        }
        #endregion

        /// <summary>
        ///   Un metodo per identificare questo gestore di eventi esterno
        /// </summary>
        public String GetName()
        {
            return "R2014 External Event Sample";
        }

        /// <summary>
        ///   Il metodo principale del gestore di eventi.
        /// </summary>
        /// <remarks>
        ///   Viene chiamato da Revit dopo che è stato generato l'evento esterno corrispondente 
        ///   (dal modulo non modale) e Revit ha raggiunto il momento in cui potrebbe 
        ///   chiamare il gestore dell'evento (cioè questo oggetto)
        /// </remarks>
        /// 
        public void Execute(UIApplication uiapp)
        {
            try
            {
                switch (Request.Take())
                {
                    case RequestId.None:
                        {
                            return;  // no request at this time -> we can leave immediately
                        }
                    case RequestId.Initial:
                        {
                            // Cattura tutti i parametri 
                            GetParameters(uiapp);
                            // Cattura i valori
                            ValueForComboBox();
                            // Riempie la ComboBox
                            _modelessForm = App.thisApp.RetriveForm();
                            _modelessForm.SetComboBox();
                            break;
                        }
                    case RequestId.ChangeComboBox:
                        {
                            // Svuoto la list di parametri precedente
                            _listParameters.Clear();
                            // Ottiene l'elemento contenuto nella ComboBox
                            _modelessForm = App.thisApp.RetriveForm();
                            _selectemItemComboBox = _modelessForm.GetComboBox();
                            // Metodo che ottine gli elementi da inserire nella ListBox
                            _listParameters = ListForListBox(_selectemItemComboBox);
                            // Metodo che riempie la ListBox
                            _modelessForm.SetListBox();
                            break;
                        }
                    default:
                        {
                            // Una sorta di avviso qui dovrebbe informarci di una richiesta imprevista
                            break;
                        }
                }
            }
            finally
            {
                App.thisApp.WakeFormUp();
                App.thisApp.ShowFormTop();
            }

            return;
        }

        /// <summary>
        ///   Metodo richiamato nello switch
        /// </summary>
        /// <remarks>
        /// </remarks>
        /// <param name="uiapp">L'oggetto Applicazione di Revit</param>m>
        /// 
        private void GetParameters(UIApplication uiapp)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            List<Element> allElements = ElementFinder.FindAllElements(doc);
            Dictionary<BuiltInParameterGroup, List<BuiltInParameter>> dict =
                new Dictionary<BuiltInParameterGroup, List<BuiltInParameter>>();

            foreach (Element e in allElements)
            {
                foreach (Parameter p in e.Parameters)
                {
                    if (p.IsShared)
                        continue;
                    if(p.Definition != null)
                    {
                        if (!dict.ContainsKey(p.Definition.ParameterGroup))
                        {
                            dict.Add(p.Definition.ParameterGroup, new List<BuiltInParameter>());
                        }

                        BuiltInParameter biParam = (p.Definition as InternalDefinition).BuiltInParameter;
                        if (!dict[p.Definition.ParameterGroup].Contains(biParam))
                        {
                            dict[p.Definition.ParameterGroup].Add(biParam);
                        }
                    }
                }
            }

            using (StreamWriter sw = new StreamWriter(_pathTxt))
            {
                int count = 1;
                foreach (KeyValuePair<BuiltInParameterGroup, List<BuiltInParameter>> kvp in dict)
                {
                    sw.WriteLine(count + ". " + kvp.Key);
                    foreach (BuiltInParameter v in kvp.Value)
                    {
                        sw.WriteLine(string.Format("        {0}", v));
                    }
                    count++;
                }
            }
        }

        /// <summary>
        ///   Metodo per riempire la COMBOBOX
        /// </summary>
        /// <remarks>
        /// </remarks>
        /// 
        private void ValueForComboBox()
        {
            // Leggi il file .txt
            string[] lines = System.IO.File.ReadAllLines(_pathTxt);

            int count = 0;
            foreach (var item in lines)
            {
                if (item.Contains(". "))
                {
                    _titles.Add(item);
                    count++;
                }
                else
                {
                    _parameters.Add(new string[]
                    {
                        Convert.ToString(count),
                        item.Trim()
                    });
                }
            }
            count = 0;

            using (StreamWriter sw = new StreamWriter(_pathTxt2))
            {
                foreach (var item in _parameters)
                {
                    sw.WriteLine(item[0] + " - " + item[1]);
                }
            }
        }

        /// <summary>
        ///   Metodo per riempire la LISTBOX
        /// </summary>
        /// <remarks>
        /// </remarks>
        /// 
        private ArrayList ListForListBox(string value)
        {
            ArrayList listParameters = new ArrayList();
            
            bool go = false;

            string numberStr = value.Substring(0, value.IndexOf("."));

            foreach (string[] par in _parameters)
            {
                if (par[0] == numberStr)
                {
                    go = true;
                }

                if (go == true && par[0] != numberStr)
                {
                    break;
                }

                if (go)
                {
                    listParameters.Add(par[1]);
                }
            }

            return listParameters;
        }

    }  // class

}  // namespace

