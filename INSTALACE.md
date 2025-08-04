# NEXIA - Vzorkování | Instalační příručka

## 📦 Instalační balíček v0.65

Tento balíček obsahuje kompletní auditní nástroj pro statistické vzorkování v Excelu.

### 🎯 Co je v balíčku:
- `manifest.xml` - hlavní konfigurační soubor add-inu
- `dist/` - produkční verze aplikace
- `assets/` - ikony a loga
- `INSTALACE.md` - tento návod

### 📋 Systémové požadavky:
- Microsoft Excel 2016 nebo novější ✅
- Windows 10/11 nebo macOS 10.14+ ✅
- Připojení k internetu (pouze při prvním načtení) ✅

**Tvoje verze:** Excel pro Microsoft 365 (Version 2507) - **PLNĚ PODPOROVÁNA** ✅

### 🚀 Způsoby instalace:

#### **Metoda 1: Instalační odkaz (NEJJEDNODUŠŠÍ) ⚡**

**Klikni na tento odkaz a add-in se automaticky nainstaluje:**

🔗 **[NAINSTALOVAT NEXIA - VZORKOVÁNÍ](ms-excel:ofv|u|https://sirkolombus.github.io/nexia-sampling/manifest.xml)**

🌐 **[INSTALAČNÍ WEBOVÁ STRÁNKA](https://sirkolombus.github.io/nexia-sampling/install.html)**

*Odkaz automaticky otevře Excel s nabídkou instalace add-inu. Stačí potvrdit instalaci.*

#### **Metoda 2: Ruční nahrání souboru**

**Český Excel pro Microsoft 365 (tvoje verze):**
1. **Otevři Excel a vytvoř nový sešit**
2. **Klikni na kartu "Vložit"** (v horním menu)
3. **V sekci "Doplňky" klikni na "Získat doplňky"**
4. **V dolní části okna klikni na "Moje doplňky"**
5. **Klikni na "Nahrát můj doplněk"** (vpravo nahoře)
6. **Vyber soubor** `manifest.xml` z rozbalené složky
7. **Klikni "Nahrát"**

**Anglický Excel:**
1. **Otevři Excel**
2. **Jdi na** Insert → Get Add-ins → My Add-ins
3. **Klikni na** "Upload My Add-in"
4. **Vyber soubor** `manifest.xml` z tohoto balíčku
5. **Klikni** "Upload"

Add-in se objeví na kartě **Domů** (Home) jako tlačítko **"NEXIA - Vzorkování"**.

#### **Metoda 3: Office 365 Admin Center (Pro organizace)**

1. **Přihlas se** do Microsoft 365 Admin Center
2. **Jdi na** Settings → Integrated apps
3. **Klikni** "Upload custom apps"
4. **Nahraj** `manifest.xml` soubor
5. **Nastav** dostupnost pro požadované uživatele/skupiny

#### **Metoda 4: Centralized Deployment**

Pro centrální nasazení v organizaci kontaktuj IT administrátora s tímto balíčkem.

### 🌐 Hosting souborů

Add-in očekává soubory na: `https://sirkolombus.github.io/nexia-sampling/`

**Pro vlastní hosting:**
1. Nahraj obsah `dist/` složky na svůj webserver
2. Upraví všechny URL v `manifest.xml` na svoji doménu
3. Ujisti se, že server podporuje HTTPS

### ✅ Ověření instalace:

1. **Otevři Excel**
2. **Na kartě Home** najdi tlačítko **"NEXIA - Vzorkování"**
3. **Klikni na tlačítko** - mělo by se otevřít okno s NEXIA logem
4. **Vyzkoušej** základní funkci vzorkování

### 🔧 Řešení problémů:

**Add-in se nezobrazuje:**
- Ověř, že Excel podporuje add-iny
- Zkontroluj připojení k internetu
- Restartuj Excel

**Chyba při načítání:**
- Ověř dostupnost hosting URL
- Zkontroluj, že server podporuje HTTPS
- Vymaž cache Excelu

**Ikony se nezobrazují:**
- Ověř dostupnost ikon na hosting URL
- Vymaž Office cache
- Restartuj Excel

### 📞 Podpora:

Pro technickou podporu kontaktuj NEXIA AP tým.

---

**NEXIA - Vzorkování v0.65**  
© 2025 NEXIA AP | Auditní nástroj pro Excel
