﻿<%@ Application Language="C#" %>

<script runat="server">
        

    void Application_Start(object sender, EventArgs e) 
    {
        // Codice eseguito all'avvio dell'applicazione
        ScriptManager.ScriptResourceMapping.AddDefinition("jquery",
        new ScriptResourceDefinition
        {
            Path = "~/script/jquery-1.6.2.min.js",
            DebugPath = "~/script/jquery-1.6.2.min.js",
            CdnPath = "http://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.4.1.min.js",
            CdnDebugPath = "http://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.4.1.js"
        });
    }
    
    void Application_End(object sender, EventArgs e) 
    {
        //  Codice eseguito all'arresto dell'applicazione

    }
        
    void Application_Error(object sender, EventArgs e) 
    { 
        // Codice eseguito in caso di errore non gestito

    }

    void Session_Start(object sender, EventArgs e) 
    {
        // Codice eseguito all'avvio di una nuova sessione
        
    }

    void Session_End(object sender, EventArgs e) 
    {
        // Codice eseguito al termine di una sessione. 
        // Nota: l'evento Session_End viene generato solo quando la modalità sessionstate
        // è impostata su InProc nel file Web.config. Se la modalità è impostata su StateServer 
        // o SQLServer, l'evento non viene generato.

    }
       
</script>
