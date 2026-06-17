const pino = require('pino');
const fs = require('fs');
const qrcode = require('qrcode-terminal');
require('dotenv').config();

const CONFIG_FILE = 'config.json';
const IMAGE_FILE = 'temp_report.png';
const NOTIF_NUMBER = process.env.NOTIF_NUMBER; // Ambil dari .env

// Suppress Baileys debug output - intercept pada level stdout/stderr
const originalStdoutWrite = process.stdout.write.bind(process.stdout);
const originalStderrWrite = process.stderr.write.bind(process.stderr);

const suppressed = ['Closing session', 'SessionEntry'];

process.stdout.write = function(chunk, encoding, callback) {
    const str = chunk.toString();
    if (!suppressed.some(term => str.includes(term))) {
        return originalStdoutWrite(chunk, encoding, callback);
    }
    if (typeof callback === 'function') callback();
    return true;
};

process.stderr.write = function(chunk, encoding, callback) {
    const str = chunk.toString();
    if (!suppressed.some(term => str.includes(term))) {
        return originalStderrWrite(chunk, encoding, callback);
    }
    if (typeof callback === 'function') callback();
    return true;
};

async function connectToWA() {
    // DYNAMIC IMPORT UNTUK BAILEYS
    const baileys = await import('@whiskeysockets/baileys');
    const makeWASocket = baileys.default ? baileys.default : baileys.makeWASocket;
    const { useMultiFileAuthState, DisconnectReason } = baileys;

    const { state, saveCreds } = await useMultiFileAuthState('auth_info_baileys');
    
    // Custom logger yang tidak menampilkan output
    const silentLogger = {
        trace: () => {},
        debug: () => {},
        info: () => {},
        warn: () => {},
        error: () => {},
        fatal: () => {},
        child: () => silentLogger
    };
    
    const sock = makeWASocket({
        auth: state,
        printQRInTerminal: false, // Kita matikan bawaannya karena akan kita cetak manual
        logger: silentLogger
    });

    sock.ev.on('creds.update', saveCreds);

    sock.ev.on('connection.update', async (update) => {
        const { connection, lastDisconnect, qr } = update;
        
        if (qr) {
            console.log("\n>>> Silakan scan QR Code ini menggunakan WhatsApp di HP kamu <<<\n");
            // Cetak QR secara manual di terminal
            qrcode.generate(qr, { small: true }); 
        }

        if (connection === 'close') {
            const shouldReconnect = lastDisconnect.error?.output?.statusCode !== DisconnectReason.loggedOut;
            console.log('Koneksi terputus. Reconnecting:', shouldReconnect);
            
            // Kirim notifikasi error jika logout
            if (!shouldReconnect) {
                try {
                    const errorMsg = 'ALERT: WhatsApp Bot gagal login. Silakan scan QR Code lagi.';
                    await sock.sendMessage(NOTIF_NUMBER, { text: errorMsg });
                    console.log('Notifikasi error terkirim ke admin.');
                } catch (err) {
                    console.log('Gagal mengirim notifikasi error:', err.message);
                }
            }
            
            if (shouldReconnect) connectToWA();
        } else if (connection === 'open') {
            console.log('\n[V] Terhubung ke WhatsApp!');
            console.log('⏳ Menunggu socket siap...');
            await new Promise(resolve => setTimeout(resolve, 2000)); // Tunggu 2 detik
            await kirimReport(sock);
        }
    });
}

async function kirimReport(sock) {
    try {
        const config = JSON.parse(fs.readFileSync(CONFIG_FILE, 'utf8'));
        
        // Cari semua file temp_report_*.png
        const files = fs.readdirSync('.')
            .filter(file => file.match(/^temp_report_\d+\.png$/))
            .sort((a, b) => {
                const numA = parseInt(a.match(/\d+/)[0]);
                const numB = parseInt(b.match(/\d+/)[0]);
                return numA - numB;
            });
        
        if (files.length === 0) {
            console.log(`[!] Tidak ada file screenshot yang ditemukan.`);
            
            // Kirim notifikasi error
            try {
                const errorMsg = 'ERROR: Tidak ada file screenshot yang ditemukan. Periksa dashboard Excel.';
                await sock.sendMessage(NOTIF_NUMBER, { text: errorMsg });
            } catch (err) {
                console.log('Gagal mengirim notifikasi error:', err.message);
            }
            
            process.exit(1);
        }
        
        // Iterasi setiap grup tujuan
        for (const jid of config.grup_tujuan) {
            console.log(`\n=== Mengirim ke grup: ${jid} ===`);
            
            try {
                // Kirim semua screenshot ke grup ini
                for (const file of files) {
                    try {
                        const imageBuffer = fs.readFileSync(file);
                        console.log(`  Mengirim screenshot: ${file}...`);
                        await sock.sendMessage(jid, { 
                            image: imageBuffer
                        });
                        
                        await new Promise(resolve => setTimeout(resolve, 5000));
                    } catch (err) {
                        console.error(`  ❌ Error mengirim ${file}: ${err.message}`);
                    }
                }
                
                // Kirim caption setelah semua dashboard
                if (fs.existsSync('caption.txt')) {
                    try {
                        const captionText = fs.readFileSync('caption.txt', 'utf8');
                        console.log(`  Mengirim caption...`);
                        await sock.sendMessage(jid, { 
                            text: captionText
                        });
                        
                        await new Promise(resolve => setTimeout(resolve, 5000));
                    } catch (err) {
                        console.error(`  ❌ Error mengirim caption: ${err.message}`);
                    }
                }
                
                console.log(`✅ Selesai mengirim ke grup: ${jid}`);
            } catch (err) {
                console.error(`❌ Error proses grup ${jid}: ${err.message}`);
            }
        }
        
        console.log('\nSemua report selesai diproses!');
        
        // Kirim notifikasi sukses
        try {
            const successMsg = `✅ SUKSES: Report telah dikirim ke ${config.grup_tujuan.length} grup.\n⏰ Waktu: ${new Date().toLocaleString('id-ID')}`;
            await sock.sendMessage(NOTIF_NUMBER, { text: successMsg });
            console.log('Notifikasi sukses terkirim ke admin.');
        } catch (err) {
            console.log('Gagal mengirim notifikasi sukses:', err.message);
        }
        
        process.exit(0); 
        
    } catch (error) {
        console.error('❌ Terjadi kesalahan saat mengirim pesan:');
        console.error('   Error:', error.message);
        if (error.stack) {
            console.error('   Stack:', error.stack);
        }
        
        // Kirim notifikasi error
        try {
            const errorMsg = `❌ ERROR: ${error.message}\n⏰ Waktu: ${new Date().toLocaleString('id-ID')}`;
            await sock.sendMessage(NOTIF_NUMBER, { text: errorMsg });
        } catch (err) {
            console.log('Gagal mengirim notifikasi error:', err.message);
        }
        
        process.exit(1);
    }
}

connectToWA();