# Ottimizzatore Taglio Barre

Un'applicazione desktop con interfaccia grafica per ottimizzare il taglio di barre e profilati, minimizzando gli scarti e riducendo i costi di produzione.

## Caratteristiche

- **Interfaccia grafica intuitiva** - Facile da usare, non richiede competenze tecniche
- **Ottimizzazione intelligente** - Algoritmo First Fit Decreasing per minimizzare gli scarti
- **Supporto multi-barra** - Gestione di barre di diverse lunghezze
- **Scenari multipli** - Genera e confronta diverse soluzioni di taglio
- **Export professionale** - Esportazione in PDF ed Excel
- **Calcolo automatico** - Determina il fabbisogno minimo di barre

## Screenshot

![Screenshot dell'applicazione](screenshot.png)

## Requisiti

- Python 3.7 o superiore
- Sistema operativo: Windows, macOS, Linux

## Installazione

### Metodo 1: Utilizzo diretto (richiede Python)

1. Clona il repository:
```bash
git clone https://github.com/TUO_USERNAME/ottimizzatore-taglio-barre.git
cd ottimizzatore-taglio-barre
```

2. Installa le dipendenze:
```bash
pip install -r requirements.txt
```

3. Esegui l'applicazione:
```bash
python ottimizzatore_taglio.py
```

### Metodo 2: Creazione eseguibile

Puoi creare un eseguibile standalone con PyInstaller:

```bash
pip install pyinstaller
python build_exe.py
```

L'eseguibile sara disponibile nella cartella `dist/`.

## Utilizzo

### 1. Configurazione iniziale

- **Barre disponibili**: Inserisci le lunghezze e quantita delle barre che hai a disposizione
- **Spessore lama**: Imposta lo spessore della lama di taglio (default: 3mm)

### 2. Inserimento pezzi richiesti

- Aggiungi i pezzi da tagliare specificando lunghezza e quantita
- Utilizza i pulsanti "Aggiungi" per inserire i dati

### 3. Ottimizzazione

- Clicca su "Ottimizza" per calcolare la soluzione migliore
- Genera piu scenari per confrontare diverse opzioni
- Visualizza scarti e numero di barre utilizzate

### 4. Esportazione

- **PDF**: Report completo con schema di taglio per ogni barra
- **Excel**: Tabella dettagliata per analisi e archivio

## Funzionalita avanzate

### Calcola fabbisogno

Calcola automaticamente il numero minimo di barre necessarie per i tuoi pezzi, considerando:
- Tutte le barre disponibili nel magazzino
- Ottimizzazione degli scarti
- Costi di acquisto

### Gestione scenari

- Salva fino a 10 scenari diversi
- Confronta le soluzioni
- Scegli la migliore per la tua produzione

## Esempio pratico

Hai in magazzino:
- 10 barre da 6000mm
- 5 barre da 4000mm

Devi tagliare:
- 15 pezzi da 1200mm
- 20 pezzi da 800mm
- 10 pezzi da 500mm

L'applicazione ti dira:
1. Quante barre ti servono
2. Come tagliare ogni barra
3. Quanto scarto avrai
4. Il report in PDF da portare in officina

## Tecnologie utilizzate

- **Python 3** - Linguaggio di programmazione
- **Tkinter** - Interfaccia grafica
- **ReportLab** - Generazione PDF
- **OpenPyXL** - Gestione file Excel
- **PyInstaller** - Creazione eseguibili

## Algoritmo di ottimizzazione

L'applicazione utilizza l'algoritmo **First Fit Decreasing (FFD)**:

1. Ordina i pezzi per lunghezza decrescente
2. Per ogni pezzo, cerca la prima barra con spazio sufficiente
3. Se non trova spazio, usa una nuova barra
4. Minimizza lo scarto totale

L'algoritmo include variazioni casuali controllate per generare scenari diversi ad ogni esecuzione.

## Contribuire

Le contribuzioni sono benvenute! Per contribuire:

1. Fai un fork del progetto
2. Crea un branch per la tua feature (`git checkout -b feature/NuovaFunzionalita`)
3. Committa le modifiche (`git commit -m 'Aggiunta NuovaFunzionalita'`)
4. Push al branch (`git push origin feature/NuovaFunzionalita`)
5. Apri una Pull Request

## Segnalazione bug

Se trovi un bug, apri una [Issue](https://github.com/TUO_USERNAME/ottimizzatore-taglio-barre/issues) descrivendo:
- Il problema riscontrato
- I passi per riprodurlo
- Il comportamento atteso
- Screenshot (se applicabili)

## Roadmap

- [ ] Supporto per formati di esportazione aggiuntivi (CSV, JSON)
- [ ] Visualizzazione grafica 3D dei tagli
- [ ] Database storico dei lavori
- [ ] Calcolo automatico dei costi
- [ ] Supporto multi-lingua
- [ ] API REST per integrazione con altri software

## Licenza

Questo progetto e distribuito sotto licenza MIT. Vedi il file [LICENSE](LICENSE) per maggiori dettagli.

## Autore

Creato con passione per ottimizzare il lavoro in officina.

## Supporto

Se trovi utile questo progetto, considera di:
- Mettere una stella su GitHub
- Condividerlo con colleghi
- Segnalare miglioramenti

---

**Nota**: Questo software e fornito "cosi com'e", senza garanzie di alcun tipo. Verifica sempre i calcoli prima dell'uso in produzione.
