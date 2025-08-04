# NPP Auditní nástroj - Excel Add-in

**Aktuální verze:** v0.6 (Dokončená) - s novým NEXIA logem a finální úpravami  
**Předchozí verze:** v0.5 (NPP-v0.5) - s moderním grafickým rozhraním  
**Starší verze:** v0.4 (NPP-v0.4), v0.3, v0.2, v0.1 Alpha

## Popis projektu

Profesionální auditní nástroj pro Excel poskytující různé metody statistického vzorkování:

### 🎯 Dostupné metody vzorkování:
1. **Náhodná peněžní procházka (NPP)** - Monetary Unit Sampling
2. **Náhodný generátor čísel** - Random Number Generator
3. *(Plánováno: další auditní techniky)*

### ✨ Hlavní funkce:
- ✅ Volba mezi různými metodami vzorkování  
- ✅ Excel integrace s automatickým vytvářením tabulek
- ✅ Barevné zvýraznění vybraných vzorků (žlutá pro vzorky, růžová pro významnost)
- ✅ Kompletní auditní dokumentace se vzorci
- ✅ NEXIA AP branding a profesionální UI
- ✅ Detekce významnosti částek
- ✅ Účetní formátování čísel
- ✅ Automatické získání jména přihlášeného uživatele
- ✅ Záznam času generace vzorku
- ✅ Intuitivní terminologie pro uživatele
- ✅ **Volba umístění parametrů** - na stejný list nebo nový list
- ✅ **Automatická detekce Excel limitů** - ochrana před překročením limitu řádků
- ✅ **Inteligentní správa velkých datových sad** - automatické vytváření nových listů
- ✅ **Moderní grafické rozhraní** - card-based design s gradientovým pozadím
- ✅ **Fixed status bar** - zprávy o výsledcích na spodku okna bez posouvání obsahu
- ✅ **Animace a efekty** - plynulé přechody a interaktivní komponenty

## Changelog

### 🚀 Plánované verze:

### v0.7 (Budoucí) 📋
- 🔄 Další auditní techniky a metody vzorkování
- � Rozšířená lokalizace
- 📈 Pokročilé reporty a analýzy

---

### ✅ Dokončené verze:

### v0.6 (Dokončená) ✅
- ✅ **Nové NEXIA logo** - Implementace "Nové logo Nexia.png" v hlavičce add-inu
- ✅ **Nové ikony add-inu** - Vytvoření všech velikostí ikon (16, 32, 64, 80, 128 px) z "Nexiaa.png"
- ✅ **Přejmenování aplikace** - Z "NPP" na "NEXIA - Vzorkování" ve všech kontextech
- ✅ **Vylepšený branding** - Konzistentní NEXIA identita napříč celým add-inem
- ✅ **Tučný nadpis** - "Nástroj pro vzorkování" s optimálním spacingem
- ✅ **Cache management** - Řešení problémů s načítáním nových ikon

### v0.5 (Dokončená) ✅
- ✅ **Moderní grafické rozhraní** - Kompletní redesign s card-based layoutem
- ✅ **Gradient pozadí** - Profesionální fialovo-modrá barevná paleta
- ✅ **Animované komponenty** - Hover efekty, přechody a interaktivní prvky
- ✅ **Fixed status bar** - Zprávy na spodku okna bez posouvání obsahu
- ✅ **Vylepšené input fieldy** - Moderní styling s focus efekty
- ✅ **Animované buttony** - Gradient efekty s hover animacemi
- ✅ **Processing indikátory** - Průběh dlouhých operací
- ✅ **Responsive design** - Optimalizace pro různé velikosti okna
- ✅ **Diskrétní autor info** - Minimalistické umístění na spodku
- ✅ **Emoji v help textech** - Přívětivé a intuitivní UX

### v0.4 (Dokončená) ✅
- ✅ **Volba umístění parametrů** - Radio buttony pro výběr mezi stejným listem nebo novým listem
- ✅ **Detekce Excel limitů** - Automatická kontrola překročení 1 000 000 řádků
- ✅ **Automatické vytváření nových listů** - Pro parametry při dosažení limitů nebo na přání uživatele
- ✅ **Inteligentní pojmenování listů** - "NPP parametry YYYYMMDDHHMM" nebo "NGČ parametry YYYYMMDDHHMM"
- ✅ **Ochrana dat** - Data zůstávají na původním listu, parametry na novém
- ✅ **Informativní zprávy** - Uživatel je informován o umístění parametrů a důvodu

### v0.3 (Dokončená) ✅
- ✅ Přidána volba metody vzorkování
- ✅ Implementován náhodný generátor čísel s cyklickým výběrem
- ✅ Rozšířené UI s výběrem metod
- ✅ Seeded random generator pro reprodukovatelnost
- ✅ Unifikovaná žlutá barva pro všechny vzorky
- ✅ Kompletní auditní dokumentace pro obě metody
- ✅ Automatické získání jména přihlášeného uživatele
- ✅ Záznam času generace s přesným timestampem
- ✅ Intuitivní terminologie ("Unikátní číselný název generované sekvence")
- ✅ Automatické formátování číselných hodnot během psaní

### v0.2 (Dokončená) ✅
- ✅ Implementace duálních metod vzorkování (NPP + Random)
- ✅ Cyklický systematický výběr pro Random metodu
- ✅ Barevné rozlišení vzorků (žlutá/růžová)
- ✅ Kompletní auditní trail s parametry generace
- ✅ Základní informace o uživateli a času

### v0.1 Alpha (Dokončená)
- ✅ Základní NPP algoritmus
- ✅ Excel integrace
- ✅ Auditní dokumentace

---

# Build Excel add-ins using Office Add-ins Development Kit

Excel add-ins are integrations built by third parties into Excel by using [Excel JavaScript API](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) and [Office Platform capabilities](https://learn.microsoft.com/en-us/office/dev/add-ins/overview/office-add-ins).

## How to run this project

### Prerequisites

- Node.js (the latest LTS version). Visit the [Node.js site](https://nodejs.org/) to download and install the right version for your operating system. To verify that you've already installed these tools, run the commands `node -v` and `npm -v` in your terminal.
- Office connected to a Microsoft 365 subscription. You might qualify for a Microsoft 365 E5 developer subscription through the [Microsoft 365 Developer Program](https://developer.microsoft.com/microsoft-365/dev-program), see [FAQ](https://learn.microsoft.com/office/developer-program/microsoft-365-developer-program-faq#who-qualifies-for-a-microsoft-365-e5-developer-subscription-) for details. Alternatively, you can [sign up for a 1-month free trial](https://www.microsoft.com/microsoft-365/try?rtc=1) or [purchase a Microsoft 365 plan](https://www.microsoft.com/microsoft-365/buy/compare-all-microsoft-365-products).

### Run the add-in using Office Add-ins Development Kit extension

1. **Open the Office Add-ins Development Kit**
    
    In the **Activity Bar**, select the **Office Add-ins Development Kit** icon to open the extension.

1. **Preview Your Office Add-in (F5)**

    Select **Preview Your Office Add-in(F5)** to launch the add-in and debug the code. In the Quick Pick menu, select the option **Excel Desktop (Edge Chromium)**.

    The extension then checks that the prerequisites are met before debugging starts. Check the terminal for detailed information if there are issues with your environment. After this process, the Excel desktop application launches and sideloads the add-in.

1. **Stop Previewing Your Office Add-in**

    Once you are finished testing and debugging the add-in, select **Stop Previewing Your Office Add-in**. This closes the web server and removes the add-in from the registry and cache.

## Use the add-in project

The add-in project that you've created contains sample code for a basic task pane add-in.

## Explore the add-in code

To explore an Office add-in project, you can start with the key files listed below.

- The `./manifest.xml` file in the root directory of the project defines the settings and capabilities of the add-in.  <br>You can check whether your manifest file is valid by selecting **Validate Manifest File** option from the Office Add-ins Development Kit.
- The `./src/taskpane/taskpane.html` file contains the HTML markup for the task pane.
- The `./src/taskpane/**/*.tsx` file contains the react code and Office JavaScript API code that facilitates interaction between the task pane and the Excel application.

## Troubleshooting

If you have problems running the add-in, take these steps.

- Close any open instances of Excel.
- Close the previous web server started for the add-in with the **Stop Previewing Your Office Add-in** Office Add-ins Development Kit extension option.

If you still have problems, see [troubleshoot development errors](https://learn.microsoft.com//office/dev/add-ins/testing/troubleshoot-development-errors) or [create a GitHub issue](https://aka.ms/officedevkitnewissue) and we'll help you.  

For information on running the add-in on Excel on the web, see [Sideload Office Add-ins to Office on the web](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-office-add-ins-for-testing).

For information on debugging on older versions of Office, see [Debug add-ins using developer tools in Microsoft Edge Legacy](https://learn.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-devtools-edge-legacy).

## Make code changes

All the information about Office Add-ins is found in our [official documentation](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins). You can also explore more samples in the Office Add-ins Development Kit. Select **View Samples** to see more samples of real-world scenarios.

If you edit the manifest as part of your changes, use the **Validate Manifest File** option in the Office Add-ins Development Kit. This shows you errors in the manifest syntax.

## Engage with the team

Did you experience any problems? [Create an issue](https://aka.ms/officedevkitnewissue) and we'll help you out.

Want to learn more about new features and best practices for the Office platform? [Join the Microsoft Office Add-ins community call](https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins-community-call).

## Copyright

Copyright (c) 2024 Microsoft Corporation. All rights reserved.

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**