const pino = require('pino');

async function getGroups() {
    // Gunakan dynamic import seperti sebelumnya
    const baileys = await import('@whiskeysockets/baileys');
    const makeWASocket = baileys.default ? baileys.default : baileys.makeWASocket;
    const { useMultiFileAuthState } = baileys;

    // Gunakan sesi login yang sudah tersimpan
    const { state } = await useMultiFileAuthState('auth_info_baileys');
    
    const sock = makeWASocket({
        auth: state,
        printQRInTerminal: false,
        logger: pino({ level: 'silent' })
    });

    sock.ev.on('connection.update', async (update) => {
        const { connection } = update;
        
        if (connection === 'open') {
            console.log('\n[V] Terhubung! Sedang mengambil daftar grup...\n');
            
            try {
                // Mengambil data semua grup di mana kamu menjadi anggota
                const groups = await sock.groupFetchAllParticipating();
                
                console.log('=== DAFTAR ID GRUP WHATSAPP ===');
                for (const jid in groups) {
                    const groupName = groups[jid].subject;
                    console.log(`Nama Grup : ${groupName}`);
                    console.log(`ID Grup   : "${jid}"`);
                    console.log('-------------------------------');
                }
                
                console.log('\nSilakan copy ID Grup di atas dan paste ke dalam config.json');
                process.exit(0);
                
            } catch (error) {
                console.error('Gagal mengambil grup:', error);
                process.exit(1);
            }
        }
    });
}

getGroups();