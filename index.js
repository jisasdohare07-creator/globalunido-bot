const { Client, LocalAuth, MessageMedia } = require('whatsapp-web.js');
const qrcode = require('qrcode-terminal');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');
const puppeteer = require('puppeteer');
const https = require('https');
const { exec } = require('child_process');
const ffmpegPath = require('ffmpeg-static');
const { GoogleGenerativeAI } = require('@google/generative-ai');
const { postToInstagram, postToInstagramStory, sendToInstagramGroup } = require('./instagram');

const isWin = process.platform === "win32";
const EXCEL_PATH = isWin ? 'C:\\Users\\Admin\\Desktop\\WhatsApp_Loads.xlsx' : path.join(__dirname, 'WhatsApp_Loads.xlsx');

// Function to save loads into a local Excel file on Desktop
function saveToExcel(channelName, originalNumber, originalEmail, fullMessage) {
    try {
        let workbook;
        let worksheet;
        let data = [];

        const newRow = {
            "Date & Time": new Date().toLocaleString('en-IN', { timeZone: 'Asia/Kolkata' }),
            "Channel Name": channelName,
            "Original Contact": originalNumber || "N/A",
            "Original Email": originalEmail || "N/A",
            "Full Message": fullMessage
        };

        if (fs.existsSync(EXCEL_PATH)) {
            workbook = XLSX.readFile(EXCEL_PATH);
            worksheet = workbook.Sheets["Loads"];
            if (worksheet) {
                data = XLSX.utils.sheet_to_json(worksheet);
            }
        } else {
            workbook = XLSX.utils.book_new();
        }

        data.push(newRow);
        worksheet = XLSX.utils.json_to_sheet(data);

        // Adjust column widths
        worksheet['!cols'] = [
            { wch: 22 }, // Date & Time
            { wch: 30 }, // Channel Name
            { wch: 20 }, // Original Contact
            { wch: 30 }, // Original Email
            { wch: 60 }  // Full Message
        ];

        // Append sheet to workbook (recreating/updating it)
        workbook.Sheets["Loads"] = worksheet;
        if (!workbook.SheetNames.includes("Loads")) {
            workbook.SheetNames.push("Loads");
        }

        XLSX.writeFile(workbook, EXCEL_PATH);
        console.log(`[✔] Load saved to Excel: ${EXCEL_PATH}`);
    } catch (error) {
        console.error('[!] Error saving to Excel:', error.message);
    }
}

// Global Map to track loader inquiries (Conversational Memory)
global.pendingInquiries = global.pendingInquiries || new Map();

const DRIVERS_FILE = path.join(__dirname, 'drivers.json');

function loadDrivers() {
    try {
        if (fs.existsSync(DRIVERS_FILE)) {
            return JSON.parse(fs.readFileSync(DRIVERS_FILE, 'utf8'));
        }
    } catch (e) {
        console.error('[!] Error loading drivers:', e.message);
    }
    return [];
}

function saveDrivers(drivers) {
    try {
        fs.writeFileSync(DRIVERS_FILE, JSON.stringify(drivers, null, 2), 'utf8');
    } catch (e) {
        console.error('[!] Error saving drivers:', e.message);
    }
}

// Function to parse cities, weight, and vehicle from load text
function parseLoadText(text) {
    let fromPlace = "Anywhere";
    let toPlace = "Anywhere";
    let materialInfo = "Industrial Cargo";
    let vehicleInfo = "Any Truck Required";

    const lowerText = text.toLowerCase();

    // Parse Route (e.g. Latur to Mumbai or Latur se Pune or Latur - Mumbai)
    const routeRegex = /([a-zA-Z\u0900-\u097F\s]{3,20})\s+(?:to|se|-|👉|से)\s+([a-zA-Z\u0900-\u097F\s]{3,20})/i;
    const match = text.match(routeRegex);
    if (match) {
        fromPlace = match[1].trim();
        toPlace = match[2].trim();
    }

    // Parse Vehicle Type
    const vehicleKeywords = ['14 wheeler', '10 wheeler', '12 wheeler', 'open', 'container', 'trailer', 'lpt', 'tata ace', 'bolero', 'chota hathi', 'chhota hathi', 'tempo', 'eicher', 'hcv', 'lcv'];
    for (const kw of vehicleKeywords) {
        if (lowerText.includes(kw)) {
            vehicleInfo = kw.toUpperCase();
            break;
        }
    }

    // Parse Weight/Material Info
    const weightRegex = /(\d+(?:\.\d+)?\s*(?:ton|tons|mt|kg))/i;
    const weightMatch = text.match(weightRegex);
    if (weightMatch) {
        materialInfo = weightMatch[1].toUpperCase();
    }

    return { fromPlace, toPlace, materialInfo, vehicleInfo };
}

// Function to generate premium 1080x1080 social media image for Instagram/WhatsApp
async function generateCargoCard(originalText, filename) {
    let cardBrowser;
    try {
        const destDir = isWin ? 'C:\\Users\\Admin\\Desktop\\GLOBALUNIDO_Posts' : path.join(__dirname, 'GLOBALUNIDO_Posts');
        if (!fs.existsSync(destDir)) {
            fs.mkdirSync(destDir, { recursive: true });
        }

        const parsed = parseLoadText(originalText);
        const formattedMsg = originalText.replace(/\n/g, '<br>');

        const htmlContent = `
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@400;600;800&family=Noto+Sans+Devanagari:wght@400;700&display=swap');
  body {
    margin: 0;
    padding: 0;
    width: 1080px;
    height: 1080px;
    background: linear-gradient(135deg, #090d16 0%, #111827 50%, #1e1b4b 100%);
    font-family: 'Outfit', 'Noto Sans Devanagari', sans-serif;
    color: #ffffff;
    display: flex;
    justify-content: center;
    align-items: center;
    overflow: hidden;
  }
  .card {
    width: 960px;
    height: 960px;
    background: rgba(255, 255, 255, 0.02);
    border: 1px solid rgba(255, 255, 255, 0.08);
    border-radius: 40px;
    box-shadow: 0 25px 60px -10px rgba(0, 0, 0, 0.7);
    backdrop-filter: blur(25px);
    padding: 60px;
    box-sizing: border-box;
    display: flex;
    flex-direction: column;
    justify-content: space-between;
    position: relative;
  }
  .header {
    display: flex;
    justify-content: space-between;
    align-items: center;
  }
  .logo {
    font-size: 45px;
    font-weight: 800;
    letter-spacing: 2px;
    background: linear-gradient(to right, #60a5fa, #a78bfa);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
  }
  .badge {
    background: rgba(96, 165, 250, 0.15);
    border: 1px solid rgba(96, 165, 250, 0.3);
    padding: 10px 25px;
    border-radius: 20px;
    font-size: 20px;
    font-weight: 600;
    color: #93c5fd;
    text-transform: uppercase;
    letter-spacing: 1px;
  }
  .title {
    font-size: 42px;
    font-weight: 800;
    color: #38bdf8;
    letter-spacing: 1px;
    margin-top: 10px;
  }
  .message-box {
    background: rgba(255, 255, 255, 0.03);
    border: 1px solid rgba(255, 255, 255, 0.05);
    border-radius: 24px;
    padding: 40px;
    font-size: 34px;
    line-height: 1.5;
    font-weight: 600;
    color: #f1f5f9;
    max-height: 520px;
    overflow: hidden;
    text-overflow: ellipsis;
    box-shadow: inset 0 2px 4px rgba(0,0,0,0.3);
  }
  .footer {
    display: flex;
    justify-content: space-between;
    align-items: center;
    border-top: 1px solid rgba(255, 255, 255, 0.1);
    padding-top: 30px;
  }
  .contact-info {
    display: flex;
    flex-direction: column;
  }
  .contact-label {
    font-size: 16px;
    color: #94a3b8;
    text-transform: uppercase;
    letter-spacing: 1px;
  }
  .contact-value {
    font-size: 36px;
    font-weight: 800;
    color: #60a5fa;
  }
  .call-now {
    background: linear-gradient(135deg, #3b82f6 0%, #8b5cf6 100%);
    padding: 18px 40px;
    border-radius: 25px;
    font-size: 22px;
    font-weight: 800;
    color: #ffffff;
    box-shadow: 0 10px 20px rgba(59, 130, 246, 0.3);
    letter-spacing: 1px;
  }
</style>
</head>
<body>
<div class="card">
  <div class="header">
    <div class="logo">GLOBALUNIDO</div>
    <div class="badge">🚚 CONFIRMED LOAD</div>
  </div>
  
  <div class="title">LOAD REQUIREMENT ALERT</div>
  
  <div class="message-box">
    ${formattedMsg}
  </div>

  <div class="footer">
    <div class="contact-info">
      <div class="contact-label">FOR BOOKING CALL</div>
      <div class="contact-value">📞 ${MY_NUMBER}</div>
    </div>
    <div class="call-now">CONTACT US NOW</div>
  </div>
</div>
</body>
</html>
        `;

        cardBrowser = await puppeteer.launch({
            headless: true,
            args: ['--no-sandbox', '--disable-setuid-sandbox']
        });
        const page = await cardBrowser.newPage();
        await page.setViewport({ width: 1080, height: 1080 });
        await page.setContent(htmlContent, { waitUntil: 'networkidle0' });

        const imagePath = path.join(destDir, filename);
        await page.screenshot({ path: imagePath, type: 'png' });
        console.log(`[✔] Cargo card image created at: ${imagePath}`);
        return imagePath;
    } catch (e) {
        console.error('[!] Error generating cargo card image:', e.message);
    } finally {
        if (cardBrowser) await cardBrowser.close();
    }
    return null;
}

// Array of 5 stable, high-quality, creative copyright-free background music tracks
const CREATIVE_SONGS = [
    { name: 'creative_track_1.mp3', url: 'https://www.soundhelix.com/examples/mp3/SoundHelix-Song-1.mp3' },
    { name: 'creative_track_2.mp3', url: 'https://www.soundhelix.com/examples/mp3/SoundHelix-Song-2.mp3' },
    { name: 'creative_track_3.mp3', url: 'https://www.soundhelix.com/examples/mp3/SoundHelix-Song-3.mp3' },
    { name: 'creative_track_4.mp3', url: 'https://www.soundhelix.com/examples/mp3/SoundHelix-Song-4.mp3' },
    { name: 'creative_track_5.mp3', url: 'https://www.soundhelix.com/examples/mp3/SoundHelix-Song-5.mp3' }
];

// Function to automatically download all 5 creative background tracks on startup
function downloadAllDefaultMusic() {
    try {
        const defaultMusicDir = path.join(__dirname, 'default_music');
        if (!fs.existsSync(defaultMusicDir)) {
            fs.mkdirSync(defaultMusicDir, { recursive: true });
        }

        CREATIVE_SONGS.forEach(song => {
            const destPath = path.join(defaultMusicDir, song.name);
            if (!fs.existsSync(destPath)) {
                console.log(`[i] Auto-downloading creative background song: ${song.name}...`);
                const file = fs.createWriteStream(destPath);
                https.get(song.url, response => {
                    response.pipe(file);
                    file.on('finish', () => {
                        file.close();
                        console.log(`[✔] Creative song downloaded: ${song.name}`);
                    });
                }).on('error', err => {
                    fs.unlink(destPath, () => {});
                    console.error(`[!] Failed to download ${song.name}:`, err.message);
                });
            }
        });
    } catch (e) {
        console.error('[!] Error in downloadAllDefaultMusic:', e.message);
    }
}

// Function to generate premium 15-second cargo MP4 video with random background music using FFmpeg
async function generateCargoVideo(imagePath, videoFilename) {
    return new Promise((resolve) => {
        try {
            const defaultMusicDir = path.join(__dirname, 'default_music');
            if (!fs.existsSync(defaultMusicDir)) fs.mkdirSync(defaultMusicDir, { recursive: true });

            // Trigger download just in case
            downloadAllDefaultMusic();

            // Scan default_music folder for MP3 songs
            let songs = [];
            if (fs.existsSync(defaultMusicDir)) {
                const files = fs.readdirSync(defaultMusicDir);
                songs = files.filter(f => f.toLowerCase().endsWith('.mp3')).map(f => path.join(defaultMusicDir, f));
            }

            if (songs.length === 0) {
                console.log('[!] Creative background music files not fully downloaded yet! Skipping video generation.');
                return resolve(null);
            }

            // Pick a random song from the creative tracks list
            const randomSong = songs[Math.floor(Math.random() * songs.length)];
            console.log(`[+] Selected background music track: ${path.basename(randomSong)}`);

            const destDir = isWin ? 'C:\\Users\\Admin\\Desktop\\GLOBALUNIDO_Posts' : path.join(__dirname, 'GLOBALUNIDO_Posts');
            const videoPath = path.join(destDir, videoFilename);

            // FFmpeg command to loop the cargo image, merge with selected audio, cut at exactly 15 seconds, and encode as Instagram-compatible H.264 MP4
            const cmd = `"${ffmpegPath}" -y -loop 1 -i "${imagePath}" -i "${randomSong}" -c:v libx264 -t 15 -pix_fmt yuv420p -c:a aac -b:a 192k -shortest "${videoPath}"`;

            console.log('[i] Rendering premium video with background music using FFmpeg...');
            exec(cmd, (error, stdout, stderr) => {
                if (error) {
                    console.error('[!] FFmpeg video render failed:', error.message);
                    return resolve(null);
                }
                console.log(`[✔] Premium cargo video generated successfully: ${videoPath}`);
                resolve(videoPath);
            });
        } catch (e) {
            console.error('[!] Error during video generation:', e.message);
            resolve(null);
        }
    });
}

// Function to generate extremely professional social media caption using Gemini AI (fallback to template if no key is set)
async function generateGeminiPostCaption(originalText) {
    try {
        const apiKey = process.env.GEMINI_API_KEY || global.GEMINI_API_KEY || "";
        if (!apiKey) {
            console.log('[i] No Gemini API key provided. Using built-in premium copywriter template...');
            return `🚨 URGENT CARGO REQUIREMENT ALERT! 🚨\n\n${originalText}\n\nFor bookings, please contact GlobalUnido immediately at 📞 ${MY_NUMBER}.\n\n#logistics #transport #transportindia #truckload #globalunido #freightforwarder #indianlogistics #trucks`;
        }

        console.log('[i] Calling Gemini AI (gemini-1.5-flash) to write professional logistics copy...');
        const genAI = new GoogleGenerativeAI(apiKey);
        const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });
        const prompt = `You are a professional logistics social media copywriter for GLOBALUNIDO. Take this raw logistics cargo load requirement and rewrite it into a highly professional, engaging, clear, and extremely premium social media post caption (suitable for Instagram Reels, Posts, and Stories).
Highlight the contact booking number: 📞 ${MY_NUMBER}.
Use premium emojis, clean layouts, bullet points, and trending hashtags (e.g. #logistics, #transport, #transportindia).
Include both English and clear Devanagari Hindi phrases so it is perfectly tailored for Indian drivers and loaders.
Raw Cargo Requirement Details:
"${originalText}"`;

        const result = await model.generateContent(prompt);
        const responseText = result.response.text();
        console.log('[✔] Successfully generated premium caption using Gemini AI!');
        return responseText;
    } catch (e) {
        console.error('[!] Gemini AI copywriting failed, using fallback:', e.message);
        return `🚨 URGENT CARGO REQUIREMENT ALERT! 🚨\n\n${originalText}\n\nFor bookings, please contact GlobalUnido immediately at 📞 ${MY_NUMBER}.\n\n#logistics #transport #transportindia #truckload #globalunido #freightforwarder #indianlogistics #trucks`;
    }
}

// Generic helper to handle a confirmed load (forwarding, logging, image generation, phone delivery, and posting)
async function processConfirmedLoad(channelName, originalNumber, originalEmail, originalText, modifiedText, isUrgent = false, loaderWID = null) {
    try {
        const chats = await client.getChats();
        const targetGroup = chats.find(c => c.name === TARGET_GROUP_NAME && c.isGroup);

        const urgentPrefix = isUrgent ? `🚨🚨 *URGENT LOAD / बेहद जरूरी लोड* 🚨🚨\n\n` : ``;
        const footer = `\n\n*अगर आपको यह लोड चाहिए, तो तुरंत संपर्क करें (कॉल करें):* 📞 ${MY_NUMBER}`;
        const textToSend = urgentPrefix + modifiedText + footer;

        let groupMsg = null;

        // 1. Forward to WhatsApp Group
        if (targetGroup) {
            groupMsg = await client.sendMessage(targetGroup.id._serialized, textToSend);
            console.log(`[✔] Confirmed load forwarded to group: ${TARGET_GROUP_NAME}`);
        } else {
            console.log(`[❌] Error: Could not find target group named "${TARGET_GROUP_NAME}".`);
        }

        // 1.5. Forward to Instagram DM Group
        sendToInstagramGroup(TARGET_GROUP_NAME, textToSend).then(() => {
            console.log(`[✔] Instagram background DM group forward finished.`);
        }).catch(e => {
            console.error(`[!] Instagram DM group forward error:`, e.message);
        });

        // 2. Save to Excel
        saveToExcel(channelName, originalNumber, originalEmail, textToSend);

        // 3. Generate Cargo Card Image
        const timestamp = Date.now();
        const imageFilename = `load_${timestamp}.png`;
        const imagePath = await generateCargoCard(originalText, imageFilename);

        if (imagePath) {
            // 4. Generate Premium Cargo Video with Music
            const videoFilename = `load_${timestamp}.mp4`;
            const videoPath = await generateCargoVideo(imagePath, videoFilename);
            const postFilePath = videoPath || imagePath; // Fallback to image if video failed

            const userWID = `91${MY_NUMBER}@c.us`;

            // 5. Send Image to User's Phone Chat on WhatsApp (For easy downloading/reading)
            try {
                const imgMedia = MessageMedia.fromFilePath(imagePath);
                await client.sendMessage(userWID, imgMedia, { caption: `यहाँ आपकी नई लोड इमेज है! 📊🚚\n\n${textToSend}` });
                console.log(`[✔] Cargo card image delivered to user's phone on WhatsApp.`);
            } catch (whatsappImgErr) {
                console.error('[!] Failed to send generated image to user phone:', whatsappImgErr.message);
            }

            // 6. Send Video with Music to User's Phone Chat on WhatsApp (For listening/viewing)
            if (videoPath) {
                try {
                    const vidMedia = MessageMedia.fromFilePath(videoPath);
                    await client.sendMessage(userWID, vidMedia, { caption: `यहाँ आपका नया लोड म्यूजिक वीडियो है! 🎵🎥` });
                    console.log(`[✔] Cargo music video delivered to user's phone on WhatsApp.`);
                } catch (whatsappVidErr) {
                    console.error('[!] Failed to send generated video to user phone:', whatsappVidErr.message);
                }
            }

            // 7. Post to WhatsApp Status! (Automatically uploads video/image to user's WhatsApp Status)
            try {
                const statusMedia = MessageMedia.fromFilePath(postFilePath);
                const statusCaption = `🚨 नया कन्फर्म लोड आया है! 🚚\n\n${originalText.substring(0, 100)}...\n\nगाड़ी लगाने के लिए तुरंत कॉल करें: 📞 ${MY_NUMBER}`;
                await client.sendMessage('status@broadcast', statusMedia, { caption: statusCaption });
                console.log(`[✔] Cargo load successfully published to your WhatsApp Status!`);
            } catch (whatsappStatusErr) {
                console.error('[!] Failed to publish to WhatsApp Status:', whatsappStatusErr.message);
            }

            // 8. Generate professional caption using Gemini AI!
            const captionText = await generateGeminiPostCaption(originalText);

            // 9. Post to Instagram Feed & Stories (Background Auto-posts)
            // Post to Feed
            postToInstagram(postFilePath, captionText).then(() => {
                console.log(`[✔] Instagram background Feed auto-post finished.`);
            }).catch(e => {
                console.error(`[!] Instagram Feed auto-post error:`, e.message);
            });

            // Post to Story
            postToInstagramStory(postFilePath).then(() => {
                console.log(`[✔] Instagram background Story auto-post finished.`);
            }).catch(e => {
                console.error(`[!] Instagram Story auto-post error:`, e.message);
            });
        }

        // 6. Smart Driver Matchmaking
        try {
            const drivers = loadDrivers();
            const matchedDrivers = [];
            const loadWords = originalText.toLowerCase().replace(/[^a-zA-Z0-9\s]/g, '').split(/\s+/).filter(w => w.length >= 4);
            
            for (const driver of drivers) {
                const driverMsgClean = driver.message.toLowerCase();
                const hasMatch = loadWords.some(word => {
                    const ignores = ['load', 'gadi', 'khali', 'truck', 'chahiye', 'available', 'required', 'transport', 'service', 'lines', 'roadways'];
                    if (ignores.includes(word)) return false;
                    return driverMsgClean.includes(word);
                });
                
                if (hasMatch) {
                    matchedDrivers.push(driver);
                }
            }

            for (const driver of matchedDrivers) {
                const driverNotification = `नमस्ते भाई साहब! आपकी गाड़ी के लिए एक *कन्फर्म लोड* मिला है! 🚚\n\n*लोड डिटेल्स:*\n${textToSend}\n\n*तुरंत संपर्क करें (कॉल करें):* 📞 ${MY_NUMBER}`;
                await client.sendMessage(driver.sender, driverNotification);
                console.log(`[✔] Auto-match: Sent load alert to driver ${driver.sender}`);
            }
        } catch (matchErr) {
            console.error('[!] Error in driver matchmaking:', matchErr.message);
        }

        return groupMsg;
    } catch (e) {
        console.error('[!] Error in processConfirmedLoad:', e.message);
    }
    return null;
}

// 🔴 USER SETTINGS / सेटिंग्स (यहाँ अपने डिटेल्स भरें)
// ==========================================

// 1. Channel Names (उन सभी चैनल्स के नाम यहाँ लिखें जहाँ से मैसेज उठाना है)
const DEFAULT_CHANNELS = [
    "Transport Parivar Corporation",
    "BANNA TRANSPORT & CONSTRUCTION",
    "shree shyam transport 🙏🏻🙏🏻",
    "National Freight Transport (NFS)",
    "𝑵𝑨𝑴𝑶 𝑻𝑹𝑨𝑵𝑺𝑷𝑶𝑹𝑻",
    "𝑴𝑨𝑯𝑨𝑲𝑨𝑳 𝑻𝑹𝑨𝑵𝑺𝑷𝑶𝑹𝑻",
    "ats transport service",
    "Sri Velavan Transport 🚚",
    "Mama Sarkar Group",
    "𝗦𝗼𝗻𝗮𝗹𝗶 𝗧𝗿𝗮𝗻𝘀𝗽𝗼𝗿𝘁,𝗦𝗲𝗿𝘃𝗶𝗰𝗲 𝗥𝗮𝗻𝗷𝗮𝗻𝗴𝗮𝗼𝗻,𝗠𝗶𝗱𝗰",
    "KK SAHA TR",
    "Top Logistics All India Truck & Container Load",
    "GADI WALA TRANSPORT SERVICE",
    "Ashirwad Transport Service",
    "Mayur transport and tempo services latur",
    "Traffic Thane Mumbai Zk",
    "Chavan Transport Agency ( CTA INDIA LOAD ) CTA  Transport & Logistics",
    "RB ROAD LINES LOADING ADDA (RBR)",
    "MAHADEV TRANSPORT INDORE",
    "Harsddhi ROADLINES",
    "Rolex⚜️ transport",
    "Siddheshwar Transport",
    "Shayam Transport",
    "SSR TRANSPORTATION 🔱",
    "Abhi Transporter 🚛",
    "Shukla Logistic Chhattisgarh",
    "GADI WALA TRANSPORT SERVICE 🚚🛳️🚆✈️🚛",
    "Shukla logistic Vapi 📌",
    "आर्वी हेवी रोडलाईन्स महाराष्ट्र",
    "सोलापूर 2 🚛🚚",
    "Shukla logistic Madhya Pradesh 📌",
    "~MAHAKAL TRANSPORT ~ 01"
];

const SOURCE_CHANNELS = process.env.SOURCE_CHANNELS ? process.env.SOURCE_CHANNELS.split(",") : DEFAULT_CHANNELS;

// 2. Target Group Name (जिस ग्रुप में मैसेज भेजना है उसका बिल्कुल सही नाम)
const TARGET_GROUP_NAME = process.env.TARGET_GROUP_NAME || "GlobalUnido loading requirements";

// 3. YOUR DETAILS (आपकी डिटेल्स जो हर मैसेज में ऑटोमैटिकली लगानी हैं)
const MY_NUMBER = process.env.MY_NUMBER || "8200210397";
const MY_EMAIL = process.env.MY_EMAIL || "supportglobalunido@gmail.com";
const MY_COMPANY = process.env.MY_COMPANY || "GLOBALUNIDO";

// 4. GEMINI AI API KEY (यहाँ अपनी जेमिनी API की डालें ताकि जेमिनी प्रोफेशनल कॉपीराइटिंग कर सके, खाली रखने पर डिफॉल्ट कॉपीराइटिंग होगी)
const GEMINI_API_KEY = process.env.GEMINI_API_KEY || "";
global.GEMINI_API_KEY = GEMINI_API_KEY;

// ==========================================
// 🔵 CLOUD WEB SERVER & DYNAMIC QR DISPLAY (Render.com 24/7 Support)
// ==========================================
const express = require('express');
const app = express();
const PORT = process.env.PORT || 3000;

let latestQRCode = null;
let isBotLoggedIn = false;

app.get('/', (req, res) => {
    if (isBotLoggedIn) {
        res.send(`
            <html>
            <head>
                <title>GLOBALUNIDO Bot: Active</title>
                <style>
                    body { font-family: Arial, sans-serif; text-align: center; background: #0f172a; color: white; padding-top: 100px; }
                    .card { background: #1e293b; padding: 40px; display: inline-block; border-radius: 15px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); }
                    h1 { color: #10b981; }
                </style>
            </head>
            <body>
                <div class="card">
                    <h1>🏆 GLOBALUNIDO Bot Is Active & Running! 🚀</h1>
                    <p>Your 24/7 cloud automation is logged in and working perfectly in the background.</p>
                </div>
            </body>
            </html>
        `);
    } else if (latestQRCode) {
        res.send(`
            <html>
            <head>
                <title>Scan WhatsApp QR</title>
                <style>
                    body { font-family: Arial, sans-serif; text-align: center; background: #0f172a; color: white; padding-top: 50px; }
                    .card { background: #1e293b; padding: 30px; display: inline-block; border-radius: 15px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); }
                    .qr-container { background: white; padding: 20px; display: inline-block; border-radius: 10px; margin-top: 20px; }
                    h1 { color: #f59e0b; }
                </style>
            </head>
            <body>
                <div class="card">
                    <h1>Scan to Login GLOBALUNIDO Bot 🚚</h1>
                    <p>Open WhatsApp on your phone, go to Linked Devices, and scan this QR code:</p>
                    <div class="qr-container">
                        <img src="https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=${encodeURIComponent(latestQRCode)}" />
                    </div>
                    <p style="color: #94a3b8; margin-top: 20px; font-size: 14px;">This page refreshes automatically. Scan within 20 seconds of load!</p>
                </div>
                <script>
                    setInterval(() => { location.reload(); }, 5000);
                </script>
            </body>
            </html>
        `);
    } else {
        res.send(`
            <html>
            <head>
                <title>Initializing Client...</title>
                <style>
                    body { font-family: Arial, sans-serif; text-align: center; background: #0f172a; color: white; padding-top: 100px; }
                </style>
            </head>
            <body>
                <h1>⏳ Initializing WhatsApp Web Client...</h1>
                <p>Generating a secure QR code connection. Please wait, this page refreshes automatically...</p>
                <script>
                    setInterval(() => { location.reload(); }, 3000);
                </script>
            </body>
            </html>
        `);
    }
});

app.listen(PORT, () => {
    console.log(`[✔] Cloud Web Server listening on port ${PORT}`);
});

// ==========================================
// 🔵 AUTOMATION LOGIC / कोडिंग (नीचे कुछ मत बदलें)
// ==========================================

const client = new Client({
    authStrategy: new LocalAuth(), // Saves session so you don't have to scan QR every time
    puppeteer: {
        headless: true, // Runs in background
        args: ['--no-sandbox', '--disable-setuid-sandbox']
    }
});

client.on('qr', (qr) => {
    console.log('\n[!] Please SCAN the QR Code below with your WhatsApp:\n');
    qrcode.generate(qr, { small: true });
    latestQRCode = qr;
    isBotLoggedIn = false;
});

client.on('ready', () => {
    console.log('\n[✔] WhatsApp Web is Ready & Automation Started!');
    console.log(`[i] Listening for messages from these channels: ${SOURCE_CHANNELS.join(", ")}`);
    console.log(`[i] Will forward to group: "${TARGET_GROUP_NAME}"\n`);
    
    isBotLoggedIn = true;
    latestQRCode = null;

    // Auto-download creative background tracks on startup
    downloadAllDefaultMusic();
});

client.on('disconnected', () => {
    console.log('[!] WhatsApp Client was disconnected.');
    isBotLoggedIn = false;
    latestQRCode = null;
});

client.on('message_create', async (msg) => {
    try {
        // Wait for chat to load
        const chat = await msg.getChat();

        // Direct matching of all chats to maximize coverage and attendee accuracy!
        const isTargetChannel = SOURCE_CHANNELS.some(channelName => 
            chat.name && chat.name.toLowerCase().includes(channelName.toLowerCase())
        );

        if (isTargetChannel) {
            console.log(`\n[+] New message received in channel: ${chat.name}`);

            let originalText = msg.body;

            // Extract original details BEFORE replacing them
            const phoneRegexForExtraction = /(?:\+91[\-\s]?)?(?:\d{5}[\-\s]?\d{5}|\d{10})/g;
            const foundPhones = originalText.match(phoneRegexForExtraction);
            const originalNumber = foundPhones ? foundPhones.join(', ') : 'N/A';

            const emailRegexForExtraction = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
            const foundEmails = originalText.match(emailRegexForExtraction);
            const originalEmail = foundEmails ? foundEmails.join(', ') : 'N/A';

            let modifiedText = originalText;

            // 1. SMART PHONE NUMBER REPLACEMENT
            const phoneRegex = /(?:\+91[\-\s]?)?(?:\d{5}[\-\s]?\d{5}|\d{10})/g;
            modifiedText = modifiedText.replace(phoneRegex, MY_NUMBER);

            // 2. SMART EMAIL REPLACEMENT
            const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
            modifiedText = modifiedText.replace(emailRegex, MY_EMAIL);

            // 3. SMART COMPANY NAME REPLACEMENT
            const escapeChannelName = chat.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            const channelNameRegex = new RegExp(escapeChannelName, 'gi');
            modifiedText = modifiedText.replace(channelNameRegex, MY_COMPANY);

            // Also replace common keywords just in case they use variations in the text
            const commonWordsRegex = /TRANSPORT PARIVAR|BANNA TRANSPORT|shree shyam transport|National Freight Transport|NFS|𝑵𝑨𝑴𝑶 𝑻𝑹𝑨𝑵𝑺𝑷𝑶𝑹𝑻|𝑴𝑨𝑯𝑨𝑲𝑨𝑳 𝑻𝑹𝑨𝑵𝑺𝑷𝑶𝑹𝑻|Shukla Logistic|Shukla logistic|आर्वी हेवी ROADLINES|सोलापूर|MAHAKAL TRANSPORT/gi;
            modifiedText = modifiedText.replace(commonWordsRegex, MY_COMPANY);

            console.log(`[+] Modified text successfully.`);

            // --- LOADER AUTO-INQUIRY SYSTEM (ONLY SEND INQUIRY, DO NOT SAVE/FORWARD YET) ---
            let inquirySent = false;
            if (foundPhones && foundPhones.length > 0) {
                const rawPhone = foundPhones[0];
                const cleanPhone = rawPhone.replace(/\D/g, ''); // keep only digits
                
                // Standardize to 10 digits
                let tenDigitPhone = cleanPhone;
                if (cleanPhone.length > 10) {
                    if (cleanPhone.startsWith('91')) {
                        tenDigitPhone = cleanPhone.slice(2);
                    } else if (cleanPhone.startsWith('0')) {
                        tenDigitPhone = cleanPhone.slice(1);
                    } else {
                        tenDigitPhone = cleanPhone.slice(-10);
                    }
                }

                if (tenDigitPhone.length === 10 && tenDigitPhone !== MY_NUMBER) {
                    const loaderWID = `91${tenDigitPhone}@c.us`;
                    const inquiryText = `नमस्ते सर! क्या यह लोड अभी खाली (Available) है? 🚚\n\n*लोड डिटेल्स:*\n${originalText}`;
                    
                    inquirySent = true;
                    // Short delay of 3 seconds to feel human/natural
                    setTimeout(async () => {
                        try {
                            await client.sendMessage(loaderWID, inquiryText);
                            global.pendingInquiries.set(loaderWID, {
                                timestamp: Date.now(),
                                channelName: chat.name,
                                originalNumber: originalNumber,
                                originalEmail: originalEmail,
                                originalText: originalText,
                                modifiedText: modifiedText,
                                autoForwarded: false,
                                groupMsg: null
                            });
                            console.log(`[✔] Auto-inquiry sent to loader: ${loaderWID} (Load held in memory pending confirmation)`);

                            // ⏳ Set 5-Minute Auto-Forward Timeout (To prevent losing high-value loads!)
                            setTimeout(async () => {
                                try {
                                    if (global.pendingInquiries.has(loaderWID)) {
                                        const inquiryObj = global.pendingInquiries.get(loaderWID);
                                        if (!inquiryObj.autoForwarded) {
                                            inquiryObj.autoForwarded = true;
                                            console.log(`[⏳] 5 Minutes Timeout: Loader ${loaderWID} did not respond. Auto-forwarding load details now to prevent loss!`);
                                            
                                            // Process and capture groupMsg so we can reply/cancel later if needed
                                            const groupMsg = await processConfirmedLoad(
                                                inquiryObj.channelName,
                                                inquiryObj.originalNumber,
                                                inquiryObj.originalEmail,
                                                inquiryObj.originalText,
                                                inquiryObj.modifiedText,
                                                false, // not explicitly urgent on auto-forward
                                                loaderWID
                                            );
                                            inquiryObj.groupMsg = groupMsg;
                                        }
                                    }
                                } catch (timeoutErr) {
                                    console.error('[!] Error in 5-minute auto-forward timeout:', timeoutErr.message);
                                }
                            }, 5 * 60 * 1000); // 5 minutes

                        } catch (e) {
                            console.error(`[!] Failed to send auto-inquiry to loader ${loaderWID}:`, e.message);
                        }
                    }, 3000);
                }
            }

            // If no valid phone number was found to verify, forward it straight away to avoid losing the load!
            if (!inquirySent) {
                await processConfirmedLoad(chat.name, originalNumber, originalEmail, originalText, modifiedText, true); // direct confirmed loads are always marked as urgent!
            }
        } else if (!msg.fromMe) {
            // Conversational Chatbot Logic for Loader Auto-Inquiries
            if (global.pendingInquiries && global.pendingInquiries.has(msg.from)) {
                const replyText = msg.body.toLowerCase().trim();

                // Simple yes/no detection in Hindi and English
                const yesKeywords = ['yes', 'available', 'haa', 'ha', 'y', 'hai', 'है', 'हाँ', 'h', 'avail', 'khali', 'खाली', 'he', 'hoga', 'hai sir', 'ha sir'];
                const noKeywords = ['no', 'nhi', 'ni', 'nahi', 'booked', 'n', 'नहीं', 'नही', 'बुक', 'ho gaya', 'nahi hai', 'no sir', 'nhi sir'];

                const isYes = yesKeywords.some(kw => replyText.includes(kw));
                const isNo = noKeywords.some(kw => replyText.includes(kw));

                if (isYes) {
                    const inquiry = global.pendingInquiries.get(msg.from);
                    
                    if (inquiry.autoForwarded) {
                        console.log(`[✔] Loader (${msg.from}) confirmed load LATE (after 5-minute timeout). Re-posting details with URGENT tag for double boost!`);
                        
                        // Repost with URGENT tag
                        await processConfirmedLoad(inquiry.channelName, inquiry.originalNumber, inquiry.originalEmail, inquiry.originalText, inquiry.modifiedText, true, msg.from);

                        const positiveReply = `बहुत-बहुत धन्यवाद सर! लोड को URGENT (बेहद जरूरी) मार्क करके दोबारा सभी ग्रुप्स और इंस्टाग्राम पर पब्लिश कर दिया गया है! 👍`;
                        await msg.reply(positiveReply);
                    } else {
                        // Regular flow before timeout - Mark as URGENT!
                        await processConfirmedLoad(inquiry.channelName, inquiry.originalNumber, inquiry.originalEmail, inquiry.originalText, inquiry.modifiedText, true, msg.from);
                        
                        const positiveReply = `ठीक है सर, लोड को कन्फर्म करके बेहद जरूरी (URGENT) मार्क करके सभी जगह पब्लिश कर दिया गया है। जैसे ही गाड़ी मिलेगी हम तुरंत भेज देंगे। धन्यवाद! 👍`;
                        await msg.reply(positiveReply);
                        console.log(`[✔] Loader (${msg.from}) confirmed load before timeout. Auto-response sent.`);
                    }

                    global.pendingInquiries.delete(msg.from); // Clear state
                    return; // stop execution
                } else if (isNo) {
                    const inquiry = global.pendingInquiries.get(msg.from);
                    
                    if (inquiry.autoForwarded) {
                        console.log(`[⚠️ ALERT] Loader (${msg.from}) replied NO/BOOKED late. Sending immediate BOOKED correction to groups!`);
                        
                        // 1. Reply directly to the original WhatsApp group post to show it's booked!
                        if (inquiry.groupMsg) {
                            try {
                                await inquiry.groupMsg.reply(`❌❌ *LOAD BOOKED / लोड बुक हो चुका है!* ❌❌\n\nड्राइवर भाई साहब ध्यान दें, यह लोड अब बुक हो चुका है। कृपया इसके लिए कॉल न करें!`);
                                console.log('[✔] Posted direct WhatsApp reply correction marking load as Booked!');
                            } catch (replyErr) {
                                console.error('[!] Failed to reply directly to message:', replyErr.message);
                            }
                        } else {
                            // Fallback if message object was not captured
                            const chats = await client.getChats();
                            const targetGroup = chats.find(c => c.name === TARGET_GROUP_NAME && c.isGroup);
                            if (targetGroup) {
                                await client.sendMessage(targetGroup.id._serialized, `❌❌ *LOAD BOOKED / लोड बुक हो चुका है!* ❌❌\n\n*विवरण (Details):*\n${inquiry.modifiedText}`);
                            }
                        }

                        // 2. Post BOOKED correction to Instagram DM group
                        await sendToInstagramGroup(TARGET_GROUP_NAME, `❌❌ *LOAD BOOKED / लोड बुक हो चुका है!* ❌❌\n\nयह लोड अब बुक हो चुका है।`);

                        // 3. Send alert message to the user
                        const userWID = `91${MY_NUMBER}@c.us`;
                        await client.sendMessage(userWID, `⚠️ *ALERT / अलर्ट* ⚠️\nलोडर ने अभी बताया कि नीचे दिया गया लोड बुक हो चुका है! मैंने ग्रुप और इंस्टाग्राम पर तुरंत *BOOKED* का मैसेज पोस्ट कर दिया है ताकि कोई ड्राइवर कॉल न करे।\n\n*विवरण:*\n${inquiry.modifiedText}`);

                    } else {
                        // Cancelled before timeout - No posts made
                        console.log(`[✔] Loader (${msg.from}) replied NO before timeout. Cancelled load forwarding.`);
                    }

                    const negativeReply = `कोई बात नहीं सर! आपको कभी भी भविष्य में गाड़ी (Truck) चाहिए हो तो हमें इस नंबर पर संपर्क करें: 📞 ${MY_NUMBER}`;
                    await msg.reply(negativeReply);
                    global.pendingInquiries.delete(msg.from); // Clear state
                    return; // stop execution
                } else {
                    // Clear state to avoid locking normal conversation
                    global.pendingInquiries.delete(msg.from);
                }
            }

            const command = msg.body.toLowerCase().trim();

            // WhatsApp Command to Send Excel Sheet directly to the user's phone!
            if (command === 'excel' || command === 'send excel' || command === 'report') {
                if (fs.existsSync(EXCEL_PATH)) {
                    try {
                        const media = MessageMedia.fromFilePath(EXCEL_PATH);
                        await client.sendMessage(msg.from, media, { caption: "Here is your WhatsApp Loads Excel Sheet! 📊" });
                        console.log(`[✔] Sent Excel file to ${msg.from}`);
                    } catch (error) {
                        console.error('[!] Failed to send Excel file:', error.message);
                        await msg.reply("Sorry, I could not send the Excel file right now.");
                    }
                } else {
                    await msg.reply("No Excel file found yet on Desktop. Please wait for some loads to be forwarded first!");
                }
                return; // stop execution
            }

            // --- DRIVER INQUIRY ALARM NOTIFICATION ---
            const driverInquiryKeywords = [
                'mujhe chahiye', 'khali hai', 'gadi hai', 'dilado', 'load chahiye', 'available', 'interested', 
                'booking', 'book krna', 'contact me', 'call me', 'gadi available', 'gadi he', 'gadi khali',
                'chahiye', 'loading', 'mujhe do', 'gaddi hai', 'gaadi khali'
            ];
            
            const isDriverQuery = driverInquiryKeywords.some(kw => command.includes(kw));
            if (isDriverQuery) {
                // Play highly audible loops of system sound alert to catch user's attention!
                const alarmCmd = `powershell -Command "[console]::beep(1200, 300); Start-Sleep -Milliseconds 100; [console]::beep(1200, 300); Start-Sleep -Milliseconds 100; [console]::beep(1200, 600)"`;
                exec(alarmCmd);
                console.log(`[🔔 ALARM] Played driver inquiry alarm for message: "${msg.body}" from ${msg.from}`);
            }

            // --- SMART DRIVER REGISTRATION ---
            const driverKeywords = ['load chahiye', 'gadi khali', 'truck available', 'truck khali', 'load required', 'gadi hai', 'गाड़ी खाली', 'गाडी खाली', 'लोड चाहिए', 'gaadi khali'];
            const lowerMsg = msg.body.toLowerCase();
            const isDriver = driverKeywords.some(kw => lowerMsg.includes(kw)) || (lowerMsg.includes('to') && (lowerMsg.includes('load') || lowerMsg.includes('gadi') || lowerMsg.includes('gaadi')));

            if (isDriver) {
                let drivers = loadDrivers();
                // Remove old request from this driver to prevent duplicate messages
                drivers = drivers.filter(d => d.sender !== msg.from);
                drivers.push({
                    sender: msg.from,
                    message: msg.body,
                    timestamp: Date.now()
                });
                saveDrivers(drivers);
                
                await msg.reply("ठीक है भाई साहब! आपकी गाड़ी की लोकेशन और रूट मैंने नोट कर ली है। जैसे ही हमारे पास आपकी गाड़ी के लिए कोई कन्फर्म लोड आएगा, सिस्टम आपको तुरंत पर्सनल व्हाट्सएप मैसेज भेज देगा! 🚚👍");
                console.log(`[✔] Registered/updated driver: ${msg.from}`);
                return; // stop further execution
            }

            // Auto-Responder for Private Chats
            const keywords = ['transport', 'load', 'truck', 'tempo', 'delivery', 'maal', 'माल', 'ट्रांसपोर्ट', 'गाड़ी', 'gadi', 'freight'];
            
            // Check if message contains any keyword
            const containsKeyword = keywords.some(keyword => lowerMsg.includes(keyword));
            
            if (containsKeyword) {
                // To avoid spamming, check if we already sent them an invite in this session
                if (!global.invitedUsers) global.invitedUsers = new Set();
                
                if (!global.invitedUsers.has(msg.from)) {
                    const chats = await client.getChats();
                    const targetGroup = chats.find(c => c.name === TARGET_GROUP_NAME && c.isGroup);
                    
                    if (targetGroup) {
                        try {
                            const inviteCode = await targetGroup.getInviteCode();
                            const inviteLink = `https://chat.whatsapp.com/${inviteCode}`;
                            
                            const replyMessage = `नमस्ते! *${MY_COMPANY}* में आपका स्वागत है। 🚚\n\nट्रांसपोर्ट और लोडिंग की रोज़ाना अपडेट्स और पूछताछ के लिए कृपया हमारे ऑफिशियल ग्रुप से जुड़ें:\n👉 ${inviteLink}`;
                            
                            await msg.reply(replyMessage);
                            global.invitedUsers.add(msg.from);
                            console.log(`[+] Sent auto-reply & invite link to ${msg.from}`);
                        } catch (err) {
                            console.error('\n[!] ERROR: Could not get Group Invite Link! Make sure your WhatsApp number is an ADMIN of the group "GlobalUnido loading requirements"!\n');
                        }
                    }
                }
            }
        }
    } catch (err) {
        console.error('[!] Error processing message:', err.message);
    }
});

// Start the client
client.initialize();
