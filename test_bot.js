const fs = require('fs');
const path = require('path');

console.log("====================================================");
console.log("    GLOBALUNIDO BOT - SYSTEM DIAGNOSTIC PANEL       ");
console.log("====================================================");
console.log("");

let errors = 0;

function checkDependency(name) {
    try {
        require(name);
        console.log(`[✔] ${name}: INSTALLED`);
    } catch (e) {
        console.log(`[❌] ${name}: NOT INSTALLED (${e.message})`);
        errors++;
    }
}

console.log("1. Checking System Dependencies...");
checkDependency('whatsapp-web.js');
checkDependency('qrcode-terminal');
checkDependency('xlsx');
checkDependency('puppeteer');
checkDependency('ffmpeg-static');
checkDependency('@google/generative-ai');
console.log("");

console.log("2. Checking File System & Directories...");
const postsDir = 'C:\\Users\\Admin\\Desktop\\GLOBALUNIDO_Posts';
if (fs.existsSync(postsDir)) {
    console.log(`[✔] Post Output Directory: FOUND (${postsDir})`);
} else {
    console.log(`[i] Post Output Directory: MISSING (Will be auto-created on post)`);
}

const defaultMusicDir = path.join(__dirname, 'default_music');
if (fs.existsSync(defaultMusicDir)) {
    const files = fs.readdirSync(defaultMusicDir).filter(f => f.toLowerCase().endsWith('.mp3'));
    console.log(`[✔] Creative Music Directory: FOUND (${files.length} tracks ready)`);
} else {
    console.log(`[i] Creative Music Directory: MISSING (Will download on startup)`);
}

const sessionDir = 'C:\\Users\\Admin\\.gemini\\antigravity\\scratch\\instagram_session';
if (fs.existsSync(sessionDir)) {
    console.log(`[✔] Instagram Saved Session: FOUND (${sessionDir})`);
} else {
    console.log(`[!] Instagram Saved Session: NOT FOUND (Run Login_Instagram.bat to set up)`);
}
console.log("");

console.log("3. Testing Code Structure & Compiling...");
try {
    const { postToInstagram, postToInstagramStory, sendToInstagramGroup } = require('./instagram');
    if (typeof postToInstagram === 'function' && typeof postToInstagramStory === 'function' && typeof sendToInstagramGroup === 'function') {
        console.log(`[✔] instagram.js compile and exports: SUCCESSFUL`);
    } else {
        throw new Error('Exported functions are not valid');
    }
} catch (e) {
    console.log(`[❌] instagram.js: COMPILE ERROR (${e.message})`);
    errors++;
}

console.log("");
console.log("====================================================");
if (errors === 0) {
    console.log("🏆 DIAGNOSTIC STATUS: 100% HEALTHY & READY TO RUN!");
    console.log("Your automation engine is completely prepared and clean.");
} else {
    console.log(`⚠ DIAGNOSTIC STATUS: FOUND ${errors} PROBLEMS.`);
    console.log("Please make sure all setups are complete.");
}
console.log("====================================================");
