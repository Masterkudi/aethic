// VARIABILI GLOBALI
let currentWorkbook = null; // tiene traccia del workbook Excel corrente.
let cryptoDataCount = 0; // conta il numero di dati delle criptovalute recuperati.
let previousBlobURL = null; // tiene traccia dell'URL del blob del file Excel precedente.

// Quando la pagina web viene caricata, viene registrato un evento listener che esegue la funzione principale del codice.
window.addEventListener('load', function () {
    console.log('Pagina caricata');

    // Oggetto di stato contenente i dati dei prezzi e delle variazioni dei prezzi delle criptovalute, contiene due proprietà:
    const state = {
        prices: {}, // memorizza i prezzi correnti delle criptovalute.
        priceChanges: {} // memorizza le variazioni di prezzo delle criptovalute.
    };

    // Funzione asincrona per recuperare il prezzo di una criptovaluta da un API
    async function fetchCryptoPrice(symbol, priceElementId, priceChangeElementId) {
        try {
            console.log(`Recupero il prezzo di ${symbol}`);
            // Effettua una richiesta HTTP alla API per ottenere il prezzo della criptovaluta
            const response = await fetch(`https://api.binance.com/api/v3/ticker/price?symbol=${symbol}USDT`);
            const data = await response.json();
            console.log(`Dati ricevuti per ${symbol}:`, data);

            // Processa i dati ricevuti
            const price = parseFloat(data.price); // Prezzo della criptovaluta
            const formattedPrice = formatPrice(price); // Formatta il prezzo in modo corretto
            const currentPriceElement = document.getElementById(priceElementId);
            currentPriceElement.textContent = `${formattedPrice} USDT`; // Aggiorna l'elemento HTML con il prezzo corrente
            console.log(`Prezzo corrente di ${symbol}: ${formattedPrice} USDT`);

            const priceChange = state.prices[symbol] !== undefined ? price - state.prices[symbol] : 0; // Calcola la variazione di prezzo e la memorizza nell'oggetto state.
            state.priceChanges[symbol] = priceChange.toFixed(2); // Memorizza la variazione di prezzo
            const priceChangeElement = document.getElementById(priceChangeElementId);
            priceChangeElement.textContent = `${state.priceChanges[symbol]} USDT`; // Aggiorna l'elemento HTML con la variazione di prezzo
            console.log(`Variazione di prezzo di ${symbol}: ${state.priceChanges[symbol]} USDT`);

            state.prices[symbol] = price; // Memorizza il prezzo corrente
            cryptoDataCount++; // Incrementa il contatore dei dati delle criptovalute
            console.log(`Dati recuperati per ${cryptoDataCount}/3 criptovalute`);

            // Incrementa il contatore cryptoDataCount e, se tutte e tre le criptovalute sono state caricate, chiama la funzione updateExcel().
            if (cryptoDataCount === 3) {
                console.log('Dati recuperati per tutte le criptovalute');
                updateExcel(); // Aggiorna il file Excel
            }
        }
        // In caso di errore, chiama la funzione handleError(). 
        catch (error) {
            console.error(`Errore durante il recupero del prezzo di ${symbol}:`, error);
            return handleError(error, symbol, priceElementId, priceChangeElementId); // Gestisce gli errori durante il recupero dei dati
        }
    }

    // Formatta il prezzo in modo corretto, con un minimo di 2 cifre decimali e un massimo di 2.
    function formatPrice(price) {
        return price.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    }

    // Gestisce gli errori che si verificano durante il recupero dei prezzi delle criptovalute.
    function handleError(error, symbol, priceElementId, priceChangeElementId) {
        const errorMessage = `An error occurred while fetching ${symbol} price: ${error.message}`;
        // Stampa un messaggio di errore nella console.
        console.error(errorMessage);

        // Aggiungo un messaggio di errore nell'interfaccia utente per informare l'utente dell'eventuale errore
        const errorElement = document.getElementById('error-message');
        errorElement.textContent = errorMessage;

        // Esegue un timeout di 5 secondi prima di riprovare a recuperare i dati.
        setTimeout(() => {
            console.log(`Riprovo a recuperare il prezzo di ${symbol}`);
            fetchCryptoPrice(symbol, priceElementId, priceChangeElementId);
        }, 5000);
    }

    function updateExcel() {
        console.log('Aggiorno il file Excel');

        // Se non esiste un workbook corrente, lo crea
        if (!currentWorkbook) {
            currentWorkbook = createWorkbook();
        }

        // Converte i nuovi dati in formato Excel
        const newData = convertToExcel();

        // Aggiunge i nuovi dati al foglio di lavoro esistente anziché sovrascriverli
        const currentWS = currentWorkbook.Sheets['Crypto Prices'];
        XLSX.utils.sheet_add_json(currentWS, newData, { skipHeader: true, origin: -1 });

        // Dopo aver aggiornato il workbook, genera il Blob e scarica il file Excel
        if (currentWorkbook) {
            const wbout = XLSX.write(currentWorkbook, { bookType: 'xlsx', type: 'array' });
            const blob = new Blob([wbout], { type: 'application/octet-stream' });
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'crypto_prices.xlsx';
            a.click();

            window.URL.revokeObjectURL(url);

            // Chiamata alla funzione deletePreviousExcel() per eliminare il file Excel precedente
            deletePreviousExcel();
        }
    }

    // Crea un nuovo workbook Excel e un nuovo foglio di lavoro.
    function createWorkbook() {
        console.log('Creo un nuovo workbook');
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet([

        ]);
        ws['!cols'] = [{ wpx: 120 }, { wpx: 120 }, { wpx: 120 }]; // Imposta la larghezza delle colonne.
        XLSX.utils.book_append_sheet(wb, ws, 'Crypto Prices');
        return wb; // Restituisce il nuovo workbook.
    }


    function convertToExcel() {
        console.log('Converto i dati in formato Excel');
        const data = [['Cryptocurrency', 'Current Price (USDT)', 'Price Change (USDT)']];

        Object.keys(state.prices).forEach(symbol => {
            const price = state.prices[symbol] || 'N/A';
            const priceChange = state.priceChanges[symbol] || 'N/A';
            data.push([symbol, price, priceChange]);
        });

        return data;
    }

    // Funzione per recuperare i prezzi di tutte le criptovalute
    function fetchAllPrices() {
        console.log('Recupero i prezzi di tutte le criptovalute');
        cryptoDataCount = 0; // Resetta il contatore dei dati delle criptovalute
        Promise.all([
            fetchCryptoPrice('BTC', 'bitcoin-price', 'bitcoin-price-change'),
            fetchCryptoPrice('ETH', 'ethereum-price', 'ethereum-price-change'),
            fetchCryptoPrice('SOL', 'solana-price', 'solana-price-change')
        ])
            .catch(error => {
                console.error('Errore durante il recupero dei prezzi delle criptovalute:', error);
                // Gestisci l'errore in modo appropriato, ad esempio mostrando un messaggio di errore all'utente
                const errorElement = document.getElementById('error-message');
                errorElement.textContent = 'Si è verificato un errore durante il recupero dei prezzi delle criptovalute. Riprova più tardi.';
            });
    }

    // Esegui fetchAllPrices ogni 5 secondi per aggiornare i prezzi
    setInterval(fetchAllPrices, 5000);


    // Funzione per eliminare il file Excel precedente
    function deletePreviousExcel() {
        if (previousBlobURL) {
            window.URL.revokeObjectURL(previousBlobURL); // Elimina l'URL del blob precedente
            previousBlobURL = null; // Resetta la variabile dell'URL del blob precedente

            // Ottieni l'URL del file Excel precedente dallo storage locale
            const previousFileUrl = localStorage.getItem('previousExcelUrl');
            if (previousFileUrl) {
                // Effettua una chiamata al server per eliminare il file
                fetch(previousFileUrl, {
                    method: 'DELETE'
                })
                    .then(response => {
                        if (response.ok) {
                            // Rimuovi l'URL del file precedente dallo storage locale
                            localStorage.removeItem('previousExcelUrl');
                            console.log('Il file Excel precedente è stato eliminato con successo.');
                        } else {
                            console.error('Errore durante l\'eliminazione del file Excel precedente:', response.statusText);
                        }
                    })
                    .catch(error => {
                        console.error('Si è verificato un errore durante l\'eliminazione del file Excel precedente:', error);
                    });
            }
        }
    }

});
