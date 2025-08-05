/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 * 
 * Auditní nástroj pro náhodnou peněžní procházku
 * Autorem je Tomáš Pavlovič ve spolupráci s GitHub Copilot
 */

/* global console, document, Excel, Office */

// Globální proměnné pro uchování dat
let selectedRange: string = "";
let totalSum: number = 0;
let amountColumnIndex: number = -1;
let calculatedSampleSize: number = 0;

declare const Office: any;
declare const Excel: any;

// Funkce pro formátování čísel účetním způsobem
function formatNumber(num: number): string {
  return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, " ");
}

// Funkce pro parsování účetně formátovaného čísla
function parseAccountingNumber(value: string): number {
  return parseInt(value.replace(/\s/g, "")) || 0;
}

// Funkce pro zobrazování zpráv ve fixed status baru
function showMessage(message: string, type: "info" | "error" | "success" | "processing" = "info") {
  const statusBar = document.getElementById("status-bar");
  if (!statusBar) return;
  
  // Vytvoření status message elementu
  const messageDiv = document.createElement("div");
  messageDiv.className = `status-message ${type}`;
  messageDiv.innerHTML = message;
  
  // Přidání do status baru
  statusBar.appendChild(messageDiv);
  
  // Automatické odstranění po 4 sekundách (kromě processing zpráv)
  if (type !== "processing") {
    setTimeout(() => {
      if (messageDiv.parentNode) {
        messageDiv.parentNode.removeChild(messageDiv);
      }
    }, 4000);
  }
  
  return messageDiv; // Vrácení elementu pro manuální odstranění
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    // Event listenery
    document.getElementById("calculate-sample").onclick = calculateSample;
    document.getElementById("generate-sample").onclick = generateSample;
    
    // Event listener pro změnu metody vzorkování
    document.getElementById("sampling-method").onchange = onSamplingMethodChange;
    
    // Event listener pro automatické formátování prováděcí významnosti
    const materialityInput = document.getElementById("materiality") as HTMLInputElement;
    materialityInput.addEventListener("blur", function() {
      const value = parseAccountingNumber(this.value);
      if (value > 0) {
        this.value = formatNumber(value);
      }
    });
    
    // Event listener pro formátování během psaní (input event)
    materialityInput.addEventListener("input", function() {
      const cursorPosition = this.selectionStart;
      const value = parseAccountingNumber(this.value);
      if (value > 0) {
        const formattedValue = formatNumber(value);
        this.value = formattedValue;
        // Obnovení pozice kurzoru
        const newPosition = Math.min(cursorPosition || 0, formattedValue.length);
        this.setSelectionRange(newPosition, newPosition);
      }
    });
    
    // Event listener pro automatické formátování prováděcí významnosti pro náhodný výběr
    const randomMaterialityInput = document.getElementById("random-materiality") as HTMLInputElement;
    randomMaterialityInput.addEventListener("blur", function() {
      const value = parseAccountingNumber(this.value);
      if (value > 0) {
        this.value = formatNumber(value);
      }
    });
    
    // Event listener pro formátování během psaní pro náhodný výběr
    randomMaterialityInput.addEventListener("input", function() {
      const cursorPosition = this.selectionStart;
      const value = parseAccountingNumber(this.value);
      if (value > 0) {
        const formattedValue = formatNumber(value);
        this.value = formattedValue;
        // Obnovení pozice kurzoru
        const newPosition = Math.min(cursorPosition || 0, formattedValue.length);
        this.setSelectionRange(newPosition, newPosition);
      }
    });
    
    // Inicializace UI podle výchozí metody
    onSamplingMethodChange();
    
    // Sledování změn výběru v Excelu
    Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      worksheet.onSelectionChanged.add(updateSelectedRange);
      await context.sync();
    }).catch((error) => {
      console.error("Chyba při registraci event handleru:", error);
    });
  }
});

// Funkce pro změnu metody vzorkování
function onSamplingMethodChange() {
  const methodSelect = document.getElementById("sampling-method") as HTMLSelectElement;
  const nppParameters = document.getElementById("npp-parameters");
  const randomParameters = document.getElementById("random-parameters");
  const calculateButtonText = document.getElementById("calculate-button-text");
  
  if (methodSelect.value === "npp") {
    nppParameters.style.display = "block";
    randomParameters.style.display = "none";
    calculateButtonText.textContent = "Vypočítat statistický vzorek";
  } else {
    nppParameters.style.display = "none";
    randomParameters.style.display = "block";
    calculateButtonText.textContent = "Vypočítat statistický vzorek pro náhodný výběr";
  }
}

// Aktualizace informací o vybrané oblasti
async function updateSelectedRange() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("address");
      await context.sync();
      
      selectedRange = range.address;
      document.getElementById("selected-range").textContent = selectedRange;
    });
  } catch (error) {
    document.getElementById("selected-range").textContent = "Vyberte oblast dat";
  }
}

// Pomocná funkce pro získání jména uživatele
async function getCurrentUser(context: Excel.RequestContext): Promise<string> {
  try {
    // Pokus o získání jména uživatele z Office kontextu
    // Office.context.mailbox je dostupné v Outlook, ale ne v Excelu
    // Pokusíme se o různé způsoby získání uživatele
    
    // Metoda 1: Pokus přes Office.context (pokud je dostupné)
    if (typeof Office !== 'undefined' && Office.context) {
      // Pro Excel můžeme zkusit získat informace o dokumentu
      if (Office.context.document && Office.context.document.settings) {
        // Některé verze Office mohou mít přístup k uživatelským informacím
      }
    }
    
    // Metoda 2: Pokus přes Excel aplikaci a vlastnosti dokumentu
    const workbook = context.workbook;
    workbook.load(["name"]);
    
    // Pokus o získání vlastností dokumentu, které mohou obsahovat informace o uživateli
    const properties = workbook.properties;
    properties.load(["author", "lastAuthor"]);
    
    await context.sync();
    
    // Použijeme autora dokumentu jako uživatele
    if (properties.lastAuthor && properties.lastAuthor.trim() !== "") {
      return properties.lastAuthor;
    } else if (properties.author && properties.author.trim() !== "") {
      return properties.author;
    }
    
    // Metoda 3: Pokus přes Office.context.user (pokud je dostupné)
    if (typeof Office !== 'undefined' && Office.context && Office.context.user) {
      if (Office.context.user.displayName) {
        return Office.context.user.displayName;
      }
    }
    
    // Výchozí hodnota, pokud se nepodařilo získat jméno
    return "Microsoft Office uživatel";
  } catch (error) {
    console.warn("Nepodařilo se získat jméno uživatele:", error);
    return "Neznámý uživatel";
  }
}

// Konstanty pro Excel limity
const EXCEL_MAX_ROWS = 1048576; // Maximální počet řádků v Excel listu
const SAFE_ROW_LIMIT = 1000000; // Bezpečný limit, kdy začneme vytvářet nový list

// Pomocná funkce pro kontrolu, zda vytváření parametrů překročí limit řádků
function willExceedRowLimit(dataStartRow: number, parameterRows: number, dataRows: number): boolean {
  const totalRequiredRows = dataStartRow + parameterRows + dataRows;
  return totalRequiredRows > SAFE_ROW_LIMIT;
}

// Pomocná funkce pro získání uživatelské volby umístění parametrů
function getUserParamsLocationChoice(): 'same' | 'new' {
  const sameSheetRadio = document.getElementById("params-same-sheet") as HTMLInputElement;
  return sameSheetRadio.checked ? 'same' : 'new';
}

// Pomocná funkce pro vytvoření nového listu pro parametry
async function createParametersSheet(context: Excel.RequestContext, methodName: string): Promise<Excel.Worksheet> {
  try {
    const workbook = context.workbook;
    const timestamp = new Date().toLocaleString('cs-CZ', {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit'
    }).replace(/[^\d]/g, '');
    
    // Název listu podle metody
    const sheetName = `${methodName} parametry ${timestamp}`;
    
    // Vytvoření nového listu
    const newSheet = workbook.worksheets.add(sheetName);
    await context.sync();
    
    return newSheet;
  } catch (error) {
    console.error("Chyba při vytváření nového listu:", error);
    throw error;
  }
}

// Převod písmena sloupce na číslo (A=1, B=2, atd.)
function columnLetterToNumber(letter: string): number {
  letter = letter.toUpperCase();
  let result = 0;
  for (let i = 0; i < letter.length; i++) {
    result = result * 26 + (letter.charCodeAt(i) - 64);
  }
  return result;
}

// Parsování sloupce (písmeno nebo číslo)
function parseColumnInput(input: string): number {
  const trimmed = input.trim();
  if (/^\d+$/.test(trimmed)) {
    return parseInt(trimmed);
  } else if (/^[A-Za-z]+$/.test(trimmed)) {
    return columnLetterToNumber(trimmed);
  }
  return -1;
}

// Výpočet statistického vzorku nebo příprava náhodného výběru
async function calculateSample() {
  const methodSelect = document.getElementById("sampling-method") as HTMLSelectElement;
  
  // Zobrazení processing zprávy
  const processingMsg = showMessage("⚙️ Počítám statistický vzorek...", "processing");
  
  try {
    if (methodSelect.value === "npp") {
      await calculateNPPSample();
    } else {
      await prepareRandomSample();
    }
  } finally {
    // Odstranění processing zprávy
    if (processingMsg && processingMsg.parentNode) {
      processingMsg.parentNode.removeChild(processingMsg);
    }
  }
}

// Výpočet NPP vzorku (původní logika)
async function calculateNPPSample() {
  try {
    const amountColumnInput = (document.getElementById("amount-column") as HTMLInputElement).value;
    const confidenceFactor = parseFloat((document.getElementById("confidence-factor") as HTMLInputElement).value);
    const materiality = parseAccountingNumber((document.getElementById("materiality") as HTMLInputElement).value);
    
    if (!amountColumnInput || !confidenceFactor || !materiality) {
      showMessage("Prosím vyplňte všechna pole.", "error");
      return;
    }
    
    amountColumnIndex = parseColumnInput(amountColumnInput);
    if (amountColumnIndex < 1) {
      showMessage("Neplatný formát sloupce. Použijte písmeno (A, B, C...) nebo číslo (1, 2, 3...).", "error");
      return;
    }
    
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["values", "rowCount", "columnCount", "address"]);
      await context.sync();
      
      if (amountColumnIndex > range.columnCount) {
        showMessage(`Sloupec ${amountColumnInput} je mimo vybranou oblast dat.`, "error");
        return;
      }
      
      // Výpočet celkové sumy
      totalSum = 0;
      const values = range.values;
      
      for (let i = 0; i < values.length; i++) {
        const cellValue = values[i][amountColumnIndex - 1];
        if (typeof cellValue === "number") {
          totalSum += Math.abs(cellValue);
        }
      }
      
      // Statistický výpočet vzorku
      const statisticalSample = Math.ceil((Math.abs(totalSum) * confidenceFactor) / materiality);
      calculatedSampleSize = statisticalSample;
      
      // Zobrazení výsledku
      document.getElementById("calculated-sample").textContent = formatNumber(statisticalSample);
      (document.getElementById("final-sample-size") as HTMLInputElement).value = statisticalSample.toString();
      document.getElementById("sample-result").style.display = "block";
      
      selectedRange = range.address;
    });
    
  } catch (error) {
    console.error(error);
    showMessage("Chyba při výpočtu vzorku: " + error.message, "error");
  }
}

// Příprava náhodného výběru (nyní statisticky založeného)
async function prepareRandomSample() {
  try {
    const amountColumnInput = (document.getElementById("amount-column") as HTMLInputElement).value;
    const confidenceFactor = parseFloat((document.getElementById("random-confidence") as HTMLInputElement).value);
    const materiality = parseAccountingNumber((document.getElementById("random-materiality") as HTMLInputElement).value);
    
    if (!amountColumnInput || !confidenceFactor || !materiality) {
      showMessage("Prosím vyplňte všechna pole.", "error");
      return;
    }
    
    amountColumnIndex = parseColumnInput(amountColumnInput);
    if (amountColumnIndex < 1) {
      showMessage("Neplatný formát sloupce. Použijte písmeno (A, B, C...) nebo číslo (1, 2, 3...).", "error");
      return;
    }
    
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(["values", "rowCount", "columnCount", "address"]);
      await context.sync();
      
      if (amountColumnIndex > range.columnCount) {
        showMessage(`Sloupec ${amountColumnInput} je mimo vybranou oblast dat.`, "error");
        return;
      }
      
      // Výpočet celkové sumy
      totalSum = 0;
      const values = range.values;
      
      for (let i = 0; i < values.length; i++) {
        const cellValue = values[i][amountColumnIndex - 1];
        if (typeof cellValue === "number") {
          totalSum += Math.abs(cellValue);
        }
      }
      
      // Statistický výpočet vzorku (stejně jako u NPP)
      const statisticalSample = Math.ceil((Math.abs(totalSum) * confidenceFactor) / materiality);
      calculatedSampleSize = statisticalSample;
      
      // Zobrazení výsledku
      document.getElementById("calculated-sample").textContent = formatNumber(statisticalSample);
      (document.getElementById("final-sample-size") as HTMLInputElement).value = statisticalSample.toString();
      document.getElementById("sample-result").style.display = "block";
      
      selectedRange = range.address;
      
      showMessage(`Statisticky vypočítáno ${formatNumber(statisticalSample)} vzorků pro náhodný výběr z celkové sumy ${formatNumber(Math.abs(totalSum))} Kč.`, "success");
    });
    
  } catch (error) {
    console.error(error);
    showMessage("Chyba při výpočtu statistického vzorku pro náhodný výběr: " + error.message, "error");
  }
}

// Generování výběru vzorku
async function generateSample() {
  const methodSelect = document.getElementById("sampling-method") as HTMLSelectElement;
  
  // Zobrazení processing zprávy
  const processingMsg = showMessage("🔄 Generuji vzorek, prosím čekejte...", "processing");
  
  try {
    if (methodSelect.value === "npp") {
      await generateNPPSample();
    } else {
      await generateRandomSample();
    }
  } finally {
    // Odstranění processing zprávy
    if (processingMsg && processingMsg.parentNode) {
      processingMsg.parentNode.removeChild(processingMsg);
    }
  }
}

// Generování NPP vzorku (původní logika)
async function generateNPPSample() {
  try {
    const finalSampleSize = parseInt((document.getElementById("final-sample-size") as HTMLInputElement).value);
    const confidenceFactor = parseFloat((document.getElementById("confidence-factor") as HTMLInputElement).value);
    const materiality = parseAccountingNumber((document.getElementById("materiality") as HTMLInputElement).value);
    const amountColumnInput = (document.getElementById("amount-column") as HTMLInputElement).value;
    
    if (!finalSampleSize || finalSampleSize < 1) {
      showMessage("Prosím zadejte platný počet vzorků.", "error");
      return;
    }
    
    if (!selectedRange) {
      showMessage("Prosím nejdříve vyberte oblast dat a spočítejte vzorek.", "error");
      return;
    }
    
    // Zjištění, zda byl počet vzorků změněn
    const sampleTypeMessage = finalSampleSize === calculatedSampleSize ? 
      "Použit statisticky stanovený počet vzorků" : 
      "Použit chtěný počet vzorků";
    
    await Excel.run(async (context) => {
      // Získání aktivního worksheetu správným způsobem
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(selectedRange);
      range.load(["values", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
      await context.sync();
      
      const values = range.values;
      
      // Překalkulace totalSum z aktuálních dat pro zajištění správnosti
      totalSum = 0;
      for (let i = 0; i < values.length; i++) {
        const cellValue = values[i][amountColumnIndex - 1];
        if (typeof cellValue === "number") {
          totalSum += Math.abs(cellValue);
        }
      }
      
      const step = Math.abs(totalSum) / finalSampleSize;
      const randomStart = Math.random() * step;
      
      // Získání informací o uživateli a času
      const currentUser = await getCurrentUser(context);
      const generationTime = new Date().toLocaleString('cs-CZ', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit'
      });
      
      // Zápis parametrů nad data
      const parameterData = [
        ["PARAMETRY VÝBĚRU VZORKU - NPP", ""],
        ["", ""],
        ["INFORMACE O GENERACI:", ""],
        ["Uživatel:", currentUser],
        ["Čas generace:", generationTime],
        ["", ""],
        ["VSTUPNÍ PARAMETRY:", ""],
        ["Oblast dat:", selectedRange],
        ["Sloupec s obraty:", amountColumnInput],
        ["Faktor spolehlivosti:", confidenceFactor],
        ["Prováděcí významnost:", formatNumber(materiality)],
        ["Celková suma obratů:", formatNumber(Math.abs(totalSum))],
        ["", ""],
        ["VÝPOČTY A VZORCE:", ""],
        ["Statistický odhad vzorku:", `${formatNumber(calculatedSampleSize)} = (Faktor spolehlivosti × Celková suma) / Prováděcí významnost = (${confidenceFactor} × ${formatNumber(Math.abs(totalSum))}) / ${formatNumber(materiality)}`],
        ["Použitý počet vzorků:", formatNumber(finalSampleSize)],
        ["Typ vzorku:", sampleTypeMessage],
        ["Krok vzorkování:", `${formatNumber(Math.round(step))} = Celková suma / Počet vzorků = ${formatNumber(Math.abs(totalSum))} / ${finalSampleSize}`],
        ["Náhodný start:", `${formatNumber(Math.round(randomStart))} = Náhodné číslo × Krok = ${(randomStart/step).toFixed(4)} × ${formatNumber(Math.round(step))}`],
        ["Poznámka k náhodnosti:", "Náhodné číslo je generováno funkcí Math.random() JavaScriptu, která vytváří pseudonáhodná čísla v rozsahu 0-1 s rovnoměrným rozdělením pravděpodobnosti"]
      ];
      
      const parameterRows = parameterData.length;
      const dataStartRow = range.rowIndex;
      
      // Kontrola uživatelské volby a Excel limitů
      const userChoice = getUserParamsLocationChoice();
      const wouldExceedLimit = willExceedRowLimit(dataStartRow, parameterRows, range.rowCount);
      const useNewSheet = userChoice === 'new' || wouldExceedLimit;
      
      let targetWorksheet = worksheet;
      let targetStartRow = dataStartRow;
      let dataWorksheet = worksheet;
      let dataStartRowFinal = dataStartRow;
      
      if (useNewSheet) {
        // Vytvoření nového listu pro parametry
        targetWorksheet = await createParametersSheet(context, "NPP");
        targetStartRow = 0;
        
        // Data zůstávají na původním listu
        dataWorksheet = worksheet;
        dataStartRowFinal = dataStartRow;
        
        if (wouldExceedLimit) {
          showMessage("⚠️ Detekován Excel limit řádků! Parametry byly automaticky vygenerovány na nový list.", "info");
        } else {
          showMessage("📋 Parametry byly vygenerovány na nový list podle vaší volby.", "info");
        }
      } else {
        // Standardní režim - parametry nad data
        // Vložení řádků pro parametry
        const insertRange = worksheet.getRangeByIndexes(dataStartRow, 0, parameterRows, range.columnCount + 1);
        insertRange.insert(Excel.InsertShiftDirection.down);
        dataStartRowFinal = dataStartRow + parameterRows;
      }
      
      // Zápis parametrů na cílový list
      const parameterRange = targetWorksheet.getRangeByIndexes(targetStartRow, 0, parameterRows, 2);
      parameterRange.values = parameterData;
      parameterRange.format.font.bold = true;
      
      // Aktualizace oblasti dat (podle situace)
      const newDataRange = dataWorksheet.getRangeByIndexes(dataStartRowFinal, range.columnIndex, range.rowCount, range.columnCount);
      newDataRange.load(["values"]);
      await context.sync();
      
      // Přidání sloupce "Výběr" vedle dat na datovém listu
      const selectionColumnIndex = range.columnIndex + range.columnCount;
      const selectionHeaderRange = dataWorksheet.getRangeByIndexes(dataStartRowFinal, selectionColumnIndex, 1, 1);
      selectionHeaderRange.values = [["NPP - Výběr"]];
      selectionHeaderRange.format.font.bold = true;
      
      // Výpočet výběru pomocí náhodné peněžní procházky
      let cumulativeSum = randomStart;
      let currentTarget = step;
      const selectionResults: string[][] = [];
      
      for (let i = 0; i < values.length; i++) {
        const cellValue = values[i][amountColumnIndex - 1];
        let isSelected = "ne";
        
        if (typeof cellValue === "number") {
          const absoluteValue = Math.abs(cellValue);
          
          // Kontrola významnosti - pokud přesahuje prováděcí významnost
          if (absoluteValue > materiality) {
            isSelected = "Ano - významnost";
            // Neřadíme do kumulativního součtu
          } else {
            cumulativeSum += absoluteValue;
            
            if (cumulativeSum >= currentTarget) {
              isSelected = "Ano - NPP";
              currentTarget += step;
            }
          }
        }
        
        selectionResults.push([isSelected]);
      }
      
      // Zápis výsledků výběru na datový list
      const selectionRange = dataWorksheet.getRangeByIndexes(dataStartRowFinal, selectionColumnIndex, values.length, 1);
      selectionRange.values = selectionResults;
      
      // Zvýraznění vybraných řádků na datovém listu
      for (let i = 0; i < selectionResults.length; i++) {
        if (selectionResults[i][0].startsWith("Ano")) {
          const rowRange = dataWorksheet.getRangeByIndexes(dataStartRowFinal + i, range.columnIndex, 1, range.columnCount + 1);
          if (selectionResults[i][0] === "Ano - významnost") {
            rowRange.format.fill.color = "#FFB6C1"; // světle růžová pro významnost
          } else {
            rowRange.format.fill.color = "#FFFF99"; // světle žlutá pro NPP
          }
        }
      }
      
      // Vytvoření nebo rozšíření tabulky na datovém listu
      try {
        // Oblast dat včetně nového sloupce
        const tableRange = dataWorksheet.getRangeByIndexes(
          dataStartRowFinal, 
          range.columnIndex, 
          values.length, 
          range.columnCount + 1
        );
        
        // Kontrola, zda již existuje tabulka v této oblasti
        const tables = dataWorksheet.tables;
        tables.load("items");
        await context.sync();
        
        let existingTable = null;
        for (let j = 0; j < tables.items.length; j++) {
          const table = tables.items[j];
          table.load(["range"]);
          await context.sync();
          
          // Kontrola překryvu s naší oblastí
          const tableRange2 = table.range;
          if (tableRange2.rowIndex <= dataStartRowFinal && 
              tableRange2.rowIndex + tableRange2.rowCount >= dataStartRowFinal &&
              tableRange2.columnIndex <= range.columnIndex &&
              tableRange2.columnIndex + tableRange2.columnCount >= range.columnIndex) {
            existingTable = table;
            break;
          }
        }
        
        if (existingTable) {
          // Rozšíření existující tabulky o nový sloupec
          existingTable.resize(tableRange);
          
          // Přejmenování posledního sloupce na "NPP - Výběr"
          existingTable.load("columns");
          await context.sync();
          
          const lastColumn = existingTable.columns.getItemAt(existingTable.columns.items.length - 1);
          lastColumn.name = "NPP - Výběr";
          
          // Aplikace autofilter a filtrování hodnot "ne"
          await applyFilterToTable(existingTable, context);
          
          const paramLocation = useNewSheet ? `Parametry na listu: ${targetWorksheet.name}` : "Parametry nad daty";
          showMessage(`Existující tabulka byla rozšířena o sloupec 'NPP - Výběr'. NPP vzorky jsou žluté, významné částky růžové. ${paramLocation}`, "success");
        } else {
          // Vytvoření nové tabulky na datovém listu
          const newTable = dataWorksheet.tables.add(tableRange, true);
          newTable.name = "AuditniVzorek_" + Date.now();
          
          // Přejmenování posledního sloupce na "NPP - Výběr"
          newTable.load("columns");
          await context.sync();
          
          const lastColumn = newTable.columns.getItemAt(newTable.columns.items.length - 1);
          lastColumn.name = "NPP - Výběr";
          
          // Aplikace autofilter a filtrování hodnot "ne"
          await applyFilterToTable(newTable, context);
          
          const paramLocation = useNewSheet ? `Parametry na listu: ${targetWorksheet.name}` : "Parametry nad daty";
          showMessage(`Byla vytvořena nová tabulka s výsledky vzorkování. NPP vzorky jsou žluté, významné částky růžové. ${paramLocation}`, "success");
        }
        
      } catch (tableError) {
        console.warn("Chyba při práci s tabulkou:", tableError);
        showMessage("Výběr vzorku byl úspěšně vygenerován! NPP vzorky jsou žluté, významné částky růžové.", "success");
      }
      
      await context.sync();
    });
    
  } catch (error) {
    console.error(error);
    showMessage("Chyba při generování vzorku: " + error.message, "error");
  }
}

// Generování náhodného výběru podle pořadových čísel řádků (cyklicky)
async function generateRandomSample() {
  try {
    const finalSampleSize = parseInt((document.getElementById("final-sample-size") as HTMLInputElement).value);
    const seedInput = (document.getElementById("random-seed") as HTMLInputElement).value;
    const confidenceFactor = parseFloat((document.getElementById("random-confidence") as HTMLInputElement).value);
    const materiality = parseAccountingNumber((document.getElementById("random-materiality") as HTMLInputElement).value);
    const amountColumnInput = (document.getElementById("amount-column") as HTMLInputElement).value;
    
    if (!finalSampleSize || finalSampleSize < 1) {
      showMessage("Prosím zadejte platný počet vzorků.", "error");
      return;
    }
    
    if (!selectedRange) {
      showMessage("Prosím nejdříve vyberte oblast dat a vypočítejte statistický vzorek.", "error");
      return;
    }
    
    // Zjištění, zda byl počet vzorků změněn
    const sampleTypeMessage = finalSampleSize === calculatedSampleSize ? 
      "Použit statisticky stanovený počet vzorků" : 
      "Použit chtěný počet vzorků";
    
    await Excel.run(async (context) => {
      // Získání aktivního worksheetu správným způsobem
      const worksheet = context.workbook.worksheets.getActiveWorksheet();
      const range = worksheet.getRange(selectedRange);
      range.load(["values", "rowCount", "columnCount", "rowIndex", "columnIndex"]);
      await context.sync();
      
      const values = range.values;
      const totalRows = values.length;
      
      // Překalkulace totalSum z aktuálních dat pro zajištění správnosti
      totalSum = 0;
      for (let i = 0; i < values.length; i++) {
        const cellValue = values[i][amountColumnIndex - 1];
        if (typeof cellValue === "number") {
          totalSum += Math.abs(cellValue);
        }
      }
      
      // Počet řádků dat (bez záhlaví)
      const dataRows = totalRows - 1; // Předpokládáme, že první řádek je záhlaví
      
      if (finalSampleSize >= dataRows) {
        showMessage("Počet vzorků nemůže být větší nebo roven počtu datových řádků.", "error");
        return;
      }
      
      // Vytvoření generátoru s unikátním názvem nebo použití Math.random
      let randomGenerator: () => number;
      let sequenceNameUsed: string;
      
      if (seedInput && !isNaN(parseInt(seedInput))) {
        const sequenceName = parseInt(seedInput);
        sequenceNameUsed = sequenceName.toString();
        // Simple seeded random generator (Linear Congruential Generator)
        let seedValue = sequenceName;
        randomGenerator = () => {
          seedValue = (seedValue * 1664525 + 1013904223) % 4294967296;
          return seedValue / 4294967296;
        };
      } else {
        const generatedName = Math.floor(Math.random() * 1000000);
        sequenceNameUsed = "Náhodný název: " + generatedName;
        randomGenerator = Math.random;
      }
      
      // Výpočet kroku a náhodného startu pro pořadová čísla řádků
      const step = Math.floor(dataRows / finalSampleSize);
      const randomStartIndex = Math.floor(randomGenerator() * dataRows) + 1; // +1 protože řádek 0 je záhlaví
      
      // Získání informací o uživateli a času
      const currentUser = await getCurrentUser(context);
      const generationTime = new Date().toLocaleString('cs-CZ', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit'
      });
      
      // Zápis parametrů nad data
      const parameterData = [
        ["PARAMETRY VÝBĚRU VZORKU - NGČ", ""],
        ["", ""],
        ["INFORMACE O GENERACI:", ""],
        ["Uživatel:", currentUser],
        ["Čas generace:", generationTime],
        ["", ""],
        ["VSTUPNÍ PARAMETRY:", ""],
        ["Oblast dat:", selectedRange],
        ["Sloupec s obraty:", amountColumnInput],
        ["Metoda vzorkování:", "Náhodný generátor čísel (cyklický systematický výběr)"],
        ["Faktor spolehlivosti:", confidenceFactor],
        ["Prováděcí významnost:", formatNumber(materiality)],
        ["Celková suma obratů:", formatNumber(Math.abs(totalSum))],
        ["", ""],
        ["VÝPOČTY A VZORCE:", ""],
        ["Statistický odhad vzorku:", `${formatNumber(calculatedSampleSize)} = (Faktor spolehlivosti × Celková suma) / Prováděcí významnost = (${confidenceFactor} × ${formatNumber(Math.abs(totalSum))}) / ${formatNumber(materiality)}`],
        ["Použitý počet vzorků:", formatNumber(finalSampleSize)],
        ["Typ vzorku:", sampleTypeMessage],
        ["Celkem řádků:", `${totalRows} (včetně záhlaví)`],
        ["Datových řádků:", `${dataRows} (bez záhlaví)`],
        ["Krok výběru řádků:", `${step} = floor(${dataRows} / ${finalSampleSize})`],
        ["Náhodný start (řádek):", `${randomStartIndex} (index ${randomStartIndex - 1} v poli)`],
        ["Použitý název sekvence:", sequenceNameUsed],
        ["Poznámka:", "Cyklický systematický výběr - při překročení konce dat se pokračuje od začátku"]
      ];
      
      const parameterRows = parameterData.length;
      const dataStartRow = range.rowIndex;
      
      // Kontrola uživatelské volby a Excel limitů
      const userChoice = getUserParamsLocationChoice();
      const wouldExceedLimit = willExceedRowLimit(dataStartRow, parameterRows, range.rowCount);
      const useNewSheet = userChoice === 'new' || wouldExceedLimit;
      
      let targetWorksheet = worksheet;
      let targetStartRow = dataStartRow;
      let dataWorksheet = worksheet;
      let dataStartRowFinal = dataStartRow;
      
      if (useNewSheet) {
        // Vytvoření nového listu pro parametry
        targetWorksheet = await createParametersSheet(context, "NGČ");
        targetStartRow = 0;
        
        // Data zůstávají na původním listu
        dataWorksheet = worksheet;
        dataStartRowFinal = dataStartRow;
        
        if (wouldExceedLimit) {
          showMessage("⚠️ Detekován Excel limit řádků! Parametry byly automaticky vygenerovány na nový list.", "info");
        } else {
          showMessage("📋 Parametry byly vygenerovány na nový list podle vaší volby.", "info");
        }
      } else {
        // Standardní režim - parametry nad data
        // Vložení řádků pro parametry
        const insertRange = worksheet.getRangeByIndexes(dataStartRow, 0, parameterRows, range.columnCount + 1);
        insertRange.insert(Excel.InsertShiftDirection.down);
        dataStartRowFinal = dataStartRow + parameterRows;
      }
      
      // Zápis parametrů na cílový list
      const parameterRange = targetWorksheet.getRangeByIndexes(targetStartRow, 0, parameterRows, 2);
      parameterRange.values = parameterData;
      parameterRange.format.font.bold = true;
      
      // Přidání sloupce "Random - Výběr" vedle dat na datovém listu
      const selectionColumnIndex = range.columnIndex + range.columnCount;
      const selectionHeaderRange = dataWorksheet.getRangeByIndexes(dataStartRowFinal, selectionColumnIndex, 1, 1);
      selectionHeaderRange.values = [["Random - Výběr"]];
      selectionHeaderRange.format.font.bold = true;
      
      // Cyklický výběr řádků
      const selectedIndices = new Set<number>();
      let currentIndex = randomStartIndex - 1; // Převod na 0-based index (záhlaví = 0)
      let selectedCount = 0;
      const maxIterations = dataRows * 2; // Ochrana před nekonečnou smyčkou
      let iterations = 0;
      
      while (selectedCount < finalSampleSize && iterations < maxIterations) {
        // Kontrola, zda je index v rozsahu datových řádků (1 až dataRows)
        if (currentIndex > 0 && currentIndex < totalRows && !selectedIndices.has(currentIndex)) {
          selectedIndices.add(currentIndex);
          selectedCount++;
        }
        
        // Posun o krok
        currentIndex += step;
        
        // Cyklické přetočení - když překročíme konec dat, vraťme se na začátek
        if (currentIndex >= totalRows) {
          currentIndex = (currentIndex - totalRows) + 1; // +1 aby se přeskočilo záhlaví
        }
        
        iterations++;
      }
      
      // Vytvoření výsledků výběru
      const selectionResults: string[][] = [];
      
      for (let i = 0; i < values.length; i++) {
        let isSelected = "ne";
        
        if (i === 0) {
          // První řádek je záhlaví - nevybíráme
          isSelected = "záhlaví";
        } else {
          // Kontrola významnosti
          const cellValue = values[i][amountColumnIndex - 1];
          if (typeof cellValue === "number" && Math.abs(cellValue) > materiality) {
            isSelected = "Ano - významnost";
          } else if (selectedIndices.has(i)) {
            isSelected = "Ano - Random";
          }
        }
        
        selectionResults.push([isSelected]);
      }
      
      // Zápis výsledků výběru na datový list
      const selectionRange = dataWorksheet.getRangeByIndexes(dataStartRowFinal, selectionColumnIndex, values.length, 1);
      selectionRange.values = selectionResults;
      
      // Zvýraznění vybraných řádků na datovém listu
      for (let i = 0; i < selectionResults.length; i++) {
        if (selectionResults[i][0].startsWith("Ano")) {
          const rowRange = dataWorksheet.getRangeByIndexes(dataStartRowFinal + i, range.columnIndex, 1, range.columnCount + 1);
          if (selectionResults[i][0] === "Ano - významnost") {
            rowRange.format.fill.color = "#FFB6C1"; // světle růžová pro významnost
          } else {
            rowRange.format.fill.color = "#FFFF99"; // světle žlutá pro náhodný výběr (stejně jako NPP)
          }
        }
      }
      
      // Vytvoření tabulky na datovém listu
      try {
        const tableRange = dataWorksheet.getRangeByIndexes(
          dataStartRowFinal, 
          range.columnIndex, 
          values.length, 
          range.columnCount + 1
        );
        
        const newTable = dataWorksheet.tables.add(tableRange, true);
        newTable.name = "RandomVzorek_" + Date.now();
        
        // Přejmenování posledního sloupce
        newTable.load("columns");
        await context.sync();
        
        const lastColumn = newTable.columns.getItemAt(newTable.columns.items.length - 1);
        lastColumn.name = "Random - Výběr";
        
        // Aplikace autofilter
        await applyFilterToTable(newTable, context);
        
        const paramLocation = useNewSheet ? `Parametry na listu: ${targetWorksheet.name}` : "Parametry nad daty";
        showMessage(`Byla vytvořena tabulka s cyklickým náhodným výběrem ${selectedCount} vzorků. Vybrané vzorky jsou žluté, významné částky růžové. ${paramLocation}`, "success");
        
      } catch (tableError) {
        console.warn("Chyba při práci s tabulkou:", tableError);
        const paramLocation = useNewSheet ? `Parametry na listu: ${targetWorksheet.name}` : "Parametry nad daty";
        showMessage(`Cyklický náhodný výběr ${selectedCount} vzorků byl úspěšně vygenerován! Vybrané vzorky jsou žluté, významné částky růžové. ${paramLocation}`, "success");
      }
      
      await context.sync();
    });
    
  } catch (error) {
    console.error(error);
    showMessage("Chyba při generování náhodného vzorku: " + error.message, "error");
  }
}

export { calculateSample, generateSample };

// Pomocná funkce pro aplikaci základního autofilter na tabulku
async function applyFilterToTable(table: Excel.Table, context: Excel.RequestContext) {
  try {
    // Načtení informací o tabulce a aplikace autofilter
    const tableRange = table.getRange();
    await context.sync();
    
    // Aktivace autofilter na celé tabulce (bez konkrétního filtru)
    const worksheet = tableRange.worksheet;
    worksheet.autoFilter.apply(tableRange);
    await context.sync();
    
  } catch (filterError) {
    console.warn("Chyba při aplikaci autofilter:", filterError);
    // Autofilter se nepodařilo aktivovat, ale to není kritická chyba
  }
}
