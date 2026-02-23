# GV_SP_Utility (WpfApp)

Applicazione WPF in C# per l'analisi e la visualizzazione di file Touchstone (parametri S), con funzionalità di plotting avanzate (dominio della frequenza e del tempo - TDR) e tabelle dati.

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
    - **Esportazione Excel**: Possibilità di esportare i dati in formato CSV o Excel nativo (`.xlsx`).
- **Dati Calcolati**: Conversione automatica e visualizzazione di Magnitudine (lineare e dB), Fase (gradi), Parte Reale e Immaginaria.

### Grafici (Plotting)
- **Menu Grafici**: Accesso a grafici nel dominio della frequenza e nel dominio del tempo (TDR).
  - **Dominio della Frequenza**: Confronto di parametri S.
    - **Visualizzazione**: Magnitudine (dB o Lineare) vs Frequenza.
    - **Impedenza**: Calcolo e plot del modulo dell'impedenza ($|Z|$). Supporto automatico per metodo Riflessione ($S_{11}$) e metodo Series-Thru ($S_{21}$).
    - **Controllo Assi**: Scala Lineare o Logaritmica per entrambi gli assi X e Y.
    - Selezione Multipla e sovrapposizione curve.
    - Zoom e Pan interattivi.
    - Legenda mobile.
  - **Dominio del Tempo (TDR)**: Analisi di riflettometria nel tempo (Impedenza vs Tempo).
    - **Calcolo TDR**: Conversione da Parametri S a profilo di impedenza tramite IFFT.
    - **Controlli**:
      - **Rise Time**: Filtro gaussiano per simulare il tempo di salita del segnale.
      - **Windowing**: Finestre (Rectangular, Hamming, Hanning, Blackman) per ridurre il ringing.
      - **Delay**: Offset temporale.
      - **Impedenza di sistema ($Z_0$)**: Impostabile (es. 50 $\Omega$).
    - Visualizzazione interattiva del profilo di impedenza nel tempo.

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
   - **Frequenza**: Menu `Grafici` -> `Apri Freuency Domain`.
   - **TDR**: Menu `Grafici` -> `Apri Time Domain (TDR)`.
     - Selezionare il parametro desiderato (es. $S_{11}$, $S_{22}$) dal pannello laterale.
     - Impostare Rise Time e altri parametri (es. Windowing).
     - Premere `Aggiorna Grafico`.

## Struttura del Progetto

- **MainWindow.xaml**: Finestra principale, gestione menu.
- **TdrWindow.xaml**: Finestra per l'analisi TDR.
- **TdrCalculator.cs**: Logica di calcolo TDR (IFFT, Windowing, Step Response).
- **GraphWindow.xaml**: Logica di visualizzazione dei grafici (basata su OxyPlot).
- **TableWindow.xaml**: Visualizzazione tabellare dei dati grezzi e calcolati.
- **SettingsWindow.xaml**: Dialogo per la configurazione degli assi del grafico.
- **TouchstoneParser.cs**: Motore di parsing e modelli dati per i file Touchstone.

## Dipendenze

- **OxyPlot.Wpf**: Libreria per la generazione dei grafici.
- **ClosedXML**: Libreria per la gestione dei file Excel (.xlsx).

## Note Tecniche

Il parser normalizza tutte le frequenze in Hz. I dati vengono memorizzati internamente come numeri complessi, permettendo conversioni on-the-fly tra le varie rappresentazioni (Mag/Phase/dB).
