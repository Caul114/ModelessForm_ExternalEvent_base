#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;

#endregion

namespace ModelessForm_ExternalEvent
{
    /// <summary>
    /// Implementa l'interfaccia del componente aggiuntivo Revit IExternalCommand
    /// </summary>
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    public class Command : IExternalCommand
    {
        /// <summary>
        /// Implementa questo metodo come comando esterno per Revit.
        /// </summary>
        /// <param name="commandData">Un oggetto che viene passato all'applicazione esterna 
        /// che contiene i dati relativi al comando, 
        /// come l'oggetto dell'applicazione e la vista attiva.</param>
        /// <param name="message">Un messaggio che può essere impostato 
        /// dall'applicazione esterna che verrà visualizzato 
        /// se viene restituito un errore o un annullamento dal comando esterno.</param>
        /// <param name="elements">Un insieme di elementi a cui l'applicazione esterna può aggiungere elementi 
        /// che devono essere evidenziati in caso di errore o cancellazione.</param>
        /// <returns>Restituisce lo stato del comando esterno. 
        /// Un risultato Riuscito significa che il metodo esterno dell'API ha funzionato come previsto. 
        /// Annullato può essere utilizzato per indicare che l'utente ha annullato l'operazione esterna ad un certo punto. 
        /// L'errore deve essere restituito se l'applicazione non è in grado di procedere con l'operazione.</returns>
        /// 
        public virtual Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                App.thisApp.ShowForm(commandData.Application);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
