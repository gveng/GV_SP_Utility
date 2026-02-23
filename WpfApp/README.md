# WpfApp

Applicazione WPF in C# con una finestra principale e menu File (Apri File, Chiudi applicazione), gestione di file Touchstone, tabelle e grafico delle magnitudini.

## Avvio

Compila ed esegui il progetto con Visual Studio o con `dotnet run --project ./WpfApp`.

## Funzionalità
- Apertura multipla di file Touchstone (.s*p, .spd) con parsing dei formati RI e MA.
- Visualizzazione tabelle dati e magnitudini (in dB) per ciascun file dal menu Tabelle (generate solo su richiesta).
- Calcolo automatico delle magnitudini (da RI o MA) per ogni parametro S.
- Finestra Grafici per selezionare parametri anche da file diversi e tracciare magnitudini vs frequenza su sfondo bianco con assi etichettati.
- Menu File con voci Apri File e Chiudi applicazione.

## Requisiti
- .NET 8.0 o superiore

## Struttura
- App.xaml / App.xaml.cs
- MainWindow.xaml / MainWindow.xaml.cs
- TouchstoneParser.cs (parser e modelli dati)
- GraphWindow.xaml / GraphWindow.xaml.cs (visualizzazione grafico)

## Note
Il parser deduce il numero di porte dal conteggio dei parametri per ciascun blocco di frequenza. I dati sono normalizzati in Hz e memorizzati come numeri complessi, con magnitudini calcolate automaticamente.
