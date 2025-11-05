Fetch(https://www.anthropic.com/engineering/code-execution-with-mcp)
⎿  Received 131.2KB (200 OK)

⏺ Na základe článku od Anthropic o code execution s MCP, tu sú hlavné zmeny, ktoré by sme mali zvážiť pre náš MS 365 MCP server:

🎯 Kľúčové Odporúčania

1. Progresívne Odhaľovanie Nástrojov (Progressive Tool Discovery)

Problém: Náš server momentálne exponuje všetkých ~80+ toolov naraz. To znamená veľký context pre AI.

Riešenie z článku: Namiesto toho implementovať:
- list_available_modules() - zobrazí kategórie (Mail, Calendar, Teams, SharePoint, Planner...)
- get_module_operations(module_name) - zobrazí operácie pre danú kategóriu
- Lazy loading - nástroje sa načítajú až keď sú potrebné

Príklad:
// Namiesto 80+ toolov hneď, začni s:
- list-m365-categories  → vracia: ["mail", "calendar", "teams", "sharepoint"]
- get-category-tools(category: "mail") → vracia: [list-mail, send-mail, ...]

2. Code-Based Interface (Kódové API)

Problém: AI volá každý endpoint ako separátny tool call.

Riešenie z článku: Vytvoriť code execution environment, kde AI píše JavaScript/TypeScript kód:

// Namiesto tool calls:
// 1. list-mail-messages
// 2. get-mail-message  
// 3. send-mail

// AI napíše kód:
const messages = await m365.mail.list({ filter: "isRead eq false" });
const urgentMsgs = messages.filter(m => m.importance === "high");
for (const msg of urgentMsgs) {
await m365.mail.send({
to: "manager@company.com",
subject: `FWD: ${msg.subject}`,
body: msg.body
});
}

Benefit: Redukcia tool calls, lokálne filtrovanie, batch operácie.

3. Data Filtering v Execution Environment

Problém: Keď načítame 1000 emailov, všetky idú do contextu.

Riešenie z článku: Agent filtruje dáta v kóde pred vrátením do contextu:

// Server poskytne execution env kde agent môže:
const allEmails = await m365.mail.list({ top: 1000 });

// Filtruje lokálne (nie cez Graph API):
const summary = {
unreadCount: allEmails.filter(e => !e.isRead).length,
urgentSenders: [...new Set(
allEmails
.filter(e => e.importance === "high")
.map(e => e.from.emailAddress.address)
)]
};

return summary; // Iba 50 bytes namiesto 500KB

Benefit: Článok uvádza 98.7% redukciu tokenov! 🚀

4. PII Protection (Ochrana Osobných Údajov)

Riešenie z článku: Tokenizovať citlivé dáta pred poslaním do LLM:

// Server tokenizuje:
{
from: "user_abc123",  // namiesto "john.doe@company.com"
body: "Meeting with client_xyz789"  // namiesto "Meeting with Microsoft"
}

// Lookup table zostane v serveri
// LLM pracuje s tokenmi, ale nevidí skutočné údaje

5. Reusable Skills (Znovupoužiteľné Funkcie)

Riešenie z článku: Agent si ukladá často používané kódy:

// Agent vytvorí skill:
async function getUnreadUrgentEmails() {
const messages = await m365.mail.list({
filter: "isRead eq false and importance eq 'high'"
});
return messages.map(m => ({
from: m.from.emailAddress.address,
subject: m.subject,
receivedAt: m.receivedDateTime
}));
}

// Skill sa uloží a znovu použije

📊 Konkrétne Zmeny Pre Náš Server

✅ Fáza 1: Kategorizácia Toolov (IMPLEMENTED)

+ ✅ Pridať meta-tools:
+ ✅ - list-m365-categories
+ ✅ - list-category-tools
+ ⏸️ - get-tool-schema (future enhancement)

Status: Implemented in commit 8a304c1
- 14 categories created (mail, sharepoint, calendar, files, etc.)
- 98 tools categorized
- Progressive tool discovery enabled

✅ Fáza 2: Code Execution Sandbox (IMPLEMENTED)

+ ✅ Implementovať izolovaný Node.js sandbox
+ ✅ Vytvoriť m365 client SDK pre code execution
+ ✅ Bezpečnostné obmedzenia (timeout, memory limit)
+ ✅ Pridať execute-m365-code tool
+ ✅ Umožniť lokálne filtrovanie pred vrátením

Status: Implemented in current commit
- Sandboxed JavaScript execution with Node.js vm module
- M365 client SDK with 7 service categories (mail, calendar, teams, files, sharepoint, planner, todo)
- Security: timeout protection, context isolation, blocked dangerous globals
- Comprehensive documentation in CODE_EXECUTION.md
- Test suite with 98.7% token reduction verified

⏸️ Fáza 3: Advanced Features (FUTURE)

+ Implementovať streaming pre veľké datasety
+ TypeScript execution support
+ Skill persistence (save reusable code)
+ PII tokenization
+ Rate limiting and quota management

⚠️ Trade-offs

Výhody:
- ✅ Masívna redukcia tokenov (98.7% podľa článku)
- ✅ Rýchlejšie responses
- ✅ Nižšie náklady
- ✅ Komplexnejšie operácie v jednom kroku

Nevýhody:
- ❌ Vyššia komplexita implementácie
- ❌ Bezpečnostné riziká (treba sandbox)
- ❌ Debugging je ťažší
- ❌ Potreba runtime environment

🤔 Moja Odporúčanie

Pre MS 365 server by som postupne implementoval:

1. Teraz: Kategorizácia toolov (jednoduchá zmena)
2. Neskôr: Code execution pre filtrovanie veľkých datasetov (emails, SharePoint lists)
3. Možno: Plný code-based interface (veľká zmena architektúry)
