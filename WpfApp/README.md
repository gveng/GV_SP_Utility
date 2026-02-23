# GV_SP_Utility (WpfApp)

Applicazione WPF in C# per l'analisi e la visualizzazione di file Touchstone (parametri S), con funzionalità di plotting avanzate e tabelle dati.

## Funzionalità Principali

### Gestione File
- **Caricamento Multiplo**: Supporto per l'apertura simultanea di più file Touchstone (`.s*p`, `.spd`).
- **Parsing**: Supporto per i formati dati:
  - **RI**: Reale-Immaginario
  - **MA**: Magnitudine-Angolo (Gradi)
  - _Nota: Altri formati (es. DB) non sono attualmente supportati._
- **Rilevamento Porte**: Deduzione automatica del numero di porte basata sulla struttura dei dati nel file.

### Visualizzazione Dati
- **Tabelle Dati**: Visualizzazione tabellare dettagliata per ogni file caricato, accessibile dal menu `Tabelle`.
- **Dati Calcolati**: Conversione automatica e visualizzazione di Magnitudine (lineare e dB), Fase (gradi), Parte Reale e Immaginaria.

### Grafici (Plotting)
- **Finestra Grafici Dedicata**: Interfaccia flessibile per il confronto di parametri S.
- **Selezione Multipla**: Possibilità di selezionare e sovrapporre curve (es. S11, S21) provenienti da file diversi.
- **Assi Personalizzabili**:
  - Scala Lineare o Logaritmica per entrambi gli assi (X: Frequenza, Y: Magnitudine dB).
  - Autoscaling (default) o limiti manuali (Min/Max).
- **Interattività**: Zoom e Pan (grazie alla libreria OxyPlot).
- **Pulizia Automatica**: Il grafico si svuota automaticamente se nessun parametro è selezionato o se il file di origine viene rimosso.

## Requisiti di Sistema

- **Framework**: .NET 8.0 (Windows)
- **Sistema Operativo**: Windows 10/11

## Istruzioni per l'Uso

1. **Avvio**:
   - Aprire la soluzione in Visual Studio o VS Code.
   - Eseguire il progetto `WpfApp`.
   - Alternativamente, da riga di comando: `dotnet run --project ./WpfApp`

2. **Caricamento Dati**:
   - Menu `File` -> `Apri File...` per selezionare i file `.s2p`, `.s1p`, ecc.

3. **Grafici**:
   - Menu `Grafici` -> `Apri finestra grafici`.
   - Nella finestra grafici, espandere i file nella colonna di sinistra e spuntare i parametri desiderati (es. `S11`, `S21`).
   - Usare il pulsante `Impostazioni` per cambiare scala (Log/Lin) o limiti degli assi.

## Struttura del Progetto

- **MainWindow.xaml**: Finestra principale, gestione menu e lista file caricati.
- **GraphWindow.xaml**: Logica di visualizzazione dei grafici (basata su OxyPlot).
- **TableWindow.xaml**: Visualizzazione tabellare dei dati grezzi e calcolati.
- **SettingsWindow.xaml**: Dialogo per la configurazione degli assi del grafico.
- **TouchstoneParser.cs**: Motore di parsing e modelli dati per i file Touchstone.

## Dipendenze

- **OxyPlot.Wpf**: Libreria per la generazione dei grafici.

## Note Tecniche

Il parser normalizza tutte le frequenze in Hz. I dati vengono memorizzati internamente come numeri complessi, permettendo conversioni on-the-fly tra le varie rappresentazioni (Mag/Phase/dB).
