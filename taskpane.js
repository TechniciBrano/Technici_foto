/* global Office, msal */
"use strict";

(() => {
  const config = window.BRANO_REPORT_CONFIG;
  const button = document.getElementById("report-button");
  const status = document.getElementById("status");
  const note = document.getElementById("note");
  let client;
  let selected;
  let busy = false;
  let finished = false;

  function setStatus(message, kind = "info") {
    // Zprávy i chyby vkládáme jako text. Obsah podezřelého e-mailu nevykreslujeme.
    status.textContent = message;
    status.dataset.kind = kind;
  }

  function validateConfig() {
    const guid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    if (!config || !guid.test(config.clientId) || !guid.test(config.tenantId)) {
      throw new Error("Doplněk ještě není nakonfigurován. IT musí do config.js doplnit Application (client) ID a Directory (tenant) ID.");
    }
    if (!Array.isArray(config.recipients) || config.recipients.length < 1 ||
        config.recipients.some(address => !/^[^\s@]+@brano\.eu$/i.test(address))) {
      throw new Error("Není správně nastavena adresa příjemců hlášení. Kontaktujte IT.");
    }
  }

  function readSelection() {
    const mailbox = Office.context.mailbox;
    const item = mailbox.item;
    if (!item || item.itemType !== Office.MailboxEnums.ItemType.Message || !item.itemId) {
      throw new Error("Nejdříve vyberte přijatý e-mail ve své firemní schránce a potom otevřete doplněk.");
    }
    const sender = item.from || item.sender || {};
    return {
      itemId: item.itemId,
      restId: mailbox.convertToRestId(item.itemId, Office.MailboxEnums.RestVersion.v2_0),
      subject: item.subject || "(bez předmětu)",
      sender: sender.emailAddress || "(neuveden)",
      internetMessageId: item.internetMessageId || "(neuvedeno)",
      reporter: mailbox.userProfile.emailAddress
    };
  }

  function ensureSameSelection() {
    const item = Office.context.mailbox.item;
    if (!item || item.itemId !== selected.itemId) {
      throw new Error("Vybraný e-mail se změnil. Zavřete panel a otevřete jej znovu u zprávy, kterou chcete nahlásit.");
    }
  }

  async function acquireToken() {
    const request = {
      scopes: ["https://graph.microsoft.com/Mail.Send"],
      loginHint: selected.reporter
    };
    let result;
    try {
      result = await client.acquireTokenSilent(request);
    } catch (error) {
      const interactiveCodes = ["no_account_error", "interaction_required", "consent_required", "login_required"];
      if (!(error instanceof msal.InteractionRequiredAuthError) && !interactiveCodes.includes(error.errorCode)) {
        throw error;
      }
      setStatus("Pro dokončení hlášení potvrďte firemní přihlášení v zobrazeném okně.");
      result = await client.acquireTokenPopup(request);
    }
    if (!result.accessToken) {
      throw new Error("Nepodařilo se získat oprávnění k odeslání. Kontaktujte IT.");
    }
    const tokenTenant = result.tenantId || (result.account && result.account.tenantId);
    if (tokenTenant && tokenTenant.toLowerCase() !== config.tenantId.toLowerCase()) {
      throw new Error("Přihlášení neodpovídá firemnímu prostředí BRANO. Kontaktujte IT.");
    }
    return result.accessToken;
  }

  function makeComment() {
    // Graph comment vloží před původní zprávu. Prostý text neobsahuje HTML.
    return [
      "HLÁŠENÍ PODEZŘELÉHO E-MAILU – BRANO IT",
      "Nahlásil/a: " + selected.reporter,
      "Čas hlášení (UTC): " + new Date().toISOString(),
      "Původní odesílatel: " + selected.sender,
      "Původní předmět: " + selected.subject,
      "Původní Message-ID: " + selected.internetMessageId,
      "",
      "Poznámka uživatele:",
      note.value.trim().slice(0, 2000) || "Bez poznámky."
    ].join("\r\n");
  }

  async function graphError(response) {
    let code = "";
    try {
      const data = await response.json();
      code = String(data.error && data.error.code || "").slice(0, 100);
    } catch (_) { /* Chyba může mít i jiné než JSON tělo. */ }
    const prefix = "Microsoft hlášení nepřijal (HTTP " + response.status + (code ? ", " + code : "") + "). ";
    if (response.status === 401 || response.status === 403) {
      return prefix + "IT musí ověřit přihlášení, oprávnění Mail.Send a souhlas správce aplikace.";
    }
    if (response.status === 404) {
      return prefix + "Zpráva nebyla nalezena v přihlášené schránce. Vyberte zprávu ve své hlavní schránce a otevřete doplněk znovu.";
    }
    if (response.status === 429) {
      return prefix + "Služba je dočasně vytížena. Zkuste hlášení později.";
    }
    return prefix + "Kontaktujte IT a předejte mu tento kód chyby.";
  }

  async function reportMessage() {
    if (busy || finished || !client) return;
    busy = true;
    button.disabled = true;
    note.disabled = true;
    let requestStarted = false;
    let definitiveRejection = false;
    let timeoutId;
    try {
      ensureSameSelection();
      setStatus("Ověřování firemního přihlášení…");
      const token = await acquireToken();
      // Během přihlášení mohl uživatel přepnout na jinou zprávu.
      ensureSameSelection();
      setStatus("Předávání hlášení IT…");
      const controller = new AbortController();
      timeoutId = window.setTimeout(() => controller.abort(), 30000);
      const payload = {
        comment: makeComment(),
        toRecipients: config.recipients.map(address => ({ emailAddress: { address } }))
      };
      requestStarted = true;
      // Záměrně jediný POST. Při síťové chybě požadavek automaticky neopakujeme.
      const response = await fetch("https://graph.microsoft.com/v1.0/me/messages/" + encodeURIComponent(selected.restId) + "/forward", {
        method: "POST",
        headers: { Authorization: "Bearer " + token, "Content-Type": "application/json" },
        body: JSON.stringify(payload),
        signal: controller.signal,
        credentials: "omit",
        redirect: "error"
      });
      if (response.status !== 202) {
        definitiveRejection = response.status >= 400 && response.status < 500 && response.status !== 408;
        if (definitiveRejection) throw new Error(await graphError(response));
        throw new Error("Neočekávaná odpověď služby HTTP " + response.status + ".");
      }
      finished = true;
      button.textContent = "Hlášení předáno IT";
      setStatus("Microsoft přijal hlášení k odeslání na obě IT adresy. Kopii přeposlané zprávy najdete v Odeslané poště.", "success");
    } catch (error) {
      if (requestStarted && !definitiveRejection) {
        finished = true;
        button.textContent = "Stav odeslání není potvrzen";
        setStatus("Nepodařilo se potvrdit výsledek odeslání. Hlášení už mohlo být přijato. Než ho zopakujete, ověřte Odeslanou poštu nebo kontaktujte IT.", "error");
      } else {
        const code = error.errorCode ? " (" + String(error.errorCode).slice(0, 100) + ")" : "";
        setStatus((error.errorCode ? "Přihlášení nebo oprávnění aplikace se nepodařilo ověřit. Kontaktujte IT." : error.message || "Hlášení se nepodařilo odeslat.") + code, "error");
      }
    } finally {
      window.clearTimeout(timeoutId);
      busy = false;
      button.disabled = finished;
      note.disabled = finished;
    }
  }

  async function initialize(info) {
    try {
      if (info.host !== Office.HostType.Outlook) throw new Error("Tento panel otevřete jako doplněk v Outlooku u vybraného e-mailu.");
      selected = readSelection();
      document.getElementById("message-subject").textContent = selected.subject;
      document.getElementById("message-sender").textContent = selected.sender;
      if (config && Array.isArray(config.recipients)) {
        document.getElementById("recipients").textContent = "Příjemci: " + config.recipients.join(", ");
      }
      validateConfig();
      if (!Office.context.requirements.isSetSupported("NestedAppAuth", "1.1")) {
        throw new Error("Tato verze Outlooku nepodporuje firemní přihlášení doplňku. Použijte aktualizovaný Outlook nebo Outlook na webu.");
      }
      if (typeof msal === "undefined" || typeof msal.createNestablePublicClientApplication !== "function") {
        throw new Error("Nepodařilo se načíst přihlašovací knihovnu doplňku. Kontaktujte IT.");
      }
      client = await msal.createNestablePublicClientApplication({
        auth: {
          clientId: config.clientId,
          authority: "https://login.microsoftonline.com/" + config.tenantId,
          redirectUri: new URL("auth.html", window.location.href).href
        },
        cache: { cacheLocation: "memoryStorage" }
      });
      button.disabled = false;
      setStatus("Připraveno k nahlášení vybraného e-mailu.");
    } catch (error) {
      setStatus(error.message || "Doplněk se nepodařilo spustit. Kontaktujte IT.", "error");
    }
  }

  button.addEventListener("click", reportMessage);
  if (typeof Office === "undefined") {
    setStatus("Nepodařilo se načíst knihovnu Outlooku. Kontaktujte IT.", "error");
  } else {
    Office.onReady(initialize);
  }
})();
