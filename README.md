# Arca Outlook Tools

## Funzionalità
Il componente aggiuntivo per Microsoft Outlook permette di:
- Allegare file provenienti dai DMS di Arca Evolution direttamente dalla mail di Outlook

## Installazione
Aprire il file *Installer.msi* e seguire la procedura guidata per installare il componente aggiuntivo.\
Nel caso in cui il componente sia già installato, si consiglia di disinstallarlo tramite Pannello di Controllo.

## Configurazione
Una volta installato il componente aggiuntivo, aprire Outlook e fare click sul bottone *Impostazioni* contenuto nel riquadro *ArcaTools* del Tab *Home*.
In questa finestra va configurata la connessione al Database di Arca Evolution.
Una volta aver immesso tutte le informazioni necessarie fare click su *Salva* per memorizzare la configurazione.

*NB: Per inserire manualmente il nome del server al quale collegarsi fare click su 'Manuale' affianco all'etichetta Server nella maschera delle impostazioni.*

## Utilizzo
Per allegare un file presente nei DMS di Arca Evolution in una mail di Outlook, cliccare sul bottone *Allega DMS* presente nel riquadro *ArcaTools* della finestra di scrittura di una mail.\
Verrà aperta una finestra contenente i filtri applicabili per la ricerca dei file dei DMS, una volta scelti i filtri da applicare premere su *Avanti*.\
Nella pagina seguente verranno mostrati tutti i file presenti nei DMS che rispettano i filtri applicati. Una volta selezionati i file interessati fare click su *Inserisci* per allegarli alla mail corrente.

## Licenza
L'utilizzo del componente aggiuntivo prevede l'acquisto di una licenza d'uso con sottoscrizione annuale. La licenza è legata alla cassetta postale di Microsoft Outlook con la quale il componente aggiuntivo viene registrato. Aprendo Outlook su un altro dispositivo avente la stessa casella postale sarà possibile utilizzare la stessa licenza.\
Il componente aggiuntivo poù essere provato gratuitamente per un periodo limitato di 14 giorni, durante i quali saranno disponibili tutte le funzionalità. Una volta terminato il periodo di prova, non sarà più possibile allegare i file dei DMS alle email.

Per acquistare o avere maggiori informazioni riguardo la licenza contattare Microdesign Srl tramite l'indirizzo [ufficioamministrativo@microdesign.it](mailto:ufficioamministrativo@microdesign.it).

## Requisiti
SO: Microsoft Windows 10/11

Outlook: 2016/2019/2021/365

## Troubleshooting
### Compoente disattivato
Nel caso in cui all'avvio di Outlook venga mostrato un messaggio che indica che il componente *ArcaOutlook* sia stato disattivato, effettuare i seguenti passaggi per abilitarlo:
- Da Microsoft Outlook cliccare su *File* e poi su *Componenti aggiuntivi COM lenti e disabilitati*.
- Individuare il componente aggiuntivo *ArcaOutlook* e fare click su *Opzioni*
- Selezionare la voce *Non monitorare questo componente aggiuntivo per i prossimi 30 giorni*.
- Fare click su *Applica*

### Configurazione manuale
Nel caso in cui sia necessario configurare manualmente la connessione al Database di Arca Evolution, effettuare i seguenti passaggi:
- Recarsi nella cartella *ArcaOutlook* contenuta in *AppData/Local* *(es. C:\Users\Utente\AppData\Local\ArcaOutlook)*
- Aprire il file denominato *vsto.config.json*
- Compilare i campi necessari alla connessione
- Salvare il file facendo attenzione che l'estensione sia *.json*

*NB: Questa procedura è da svolgere solo nel caso in cui non sia possibile configurare il componente aggiuntivo attraverso la finestra di configurazione presente in Outlook.*
