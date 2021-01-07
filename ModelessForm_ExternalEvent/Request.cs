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
using System.Threading;

namespace ModelessForm_ExternalEvent
{
    /// <summary>
    ///   Un elenco di richieste disponibili nella finestra di dialogo
    /// </summary>
    /// 
    public enum RequestId : int
    {
        /// <summary>
        /// None
        /// </summary>
        None = 0,
        /// <summary>
        /// "Picker" request
        /// </summary>
        Id = 1,
    }

    /// <summary>
    ///   Una classe attorno a una variabile che contiene la richiesta corrente.
    /// </summary>
    /// <remarks>
    ///   L'accesso è reso thread-safe, anche se non ne abbiamo necessariamente bisogno 
    ///   se disabilitiamo sempre il dialogo tra le singole richieste.
    /// </remarks>
    /// 
    public class Request
    {
        // Memorizzare il valore come un Int semplice rende più semplice l'utilizzo del meccanismo di interblocco
        private int m_request = (int)RequestId.None;

        /// <summary>
        ///   Take - Il gestore Idling lo chiama per ottenere l'ultima richiesta.
        /// </summary>
        /// <remarks>
        ///   Questo non è un getter! Prende la richiesta e la sostituisce 
        ///   con "Nessuno" per indicare che la richiesta è stata "trasmessa".
        /// </remarks>
        /// 
        public RequestId Take()
        {
            return (RequestId)Interlocked.Exchange(ref m_request, (int)RequestId.None);
        }

        /// <summary>
        ///   Make - La finestra di dialogo lo chiama quando l'utente preme un pulsante di comando lì.
        /// </summary>
        /// <remarks>
        ///   Sostituisce qualsiasi richiesta effettuata in precedenza.
        /// </remarks>
        /// 
        public void Make(RequestId request)
        {
            Interlocked.Exchange(ref m_request, (int)request);
        }
    }
}
