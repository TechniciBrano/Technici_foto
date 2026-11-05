/* global msalRedirectBridge */
"use strict";
msalRedirectBridge.broadcastResponseToMainFrame().catch(() => {
  document.querySelector("p").textContent = "Přihlášení se nepodařilo předat Outlooku. Zavřete toto okno a kontaktujte IT.";
});
