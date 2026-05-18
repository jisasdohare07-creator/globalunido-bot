const { Client, RemoteAuth, MessageMedia } = require('whatsapp-web.js');
const { MongoStore } = require('wwebjs-mongo');
const mongoose = require('mongoose');
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
        console.log(`[âœ”] Load saved to Excel: ${EXCEL_PATH}`);
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

const CONTACTS_FILE = path.join(__dirname, 'contacts.json');

function loadContacts() {
    try {
        if (fs.existsSync(CONTACTS_FILE)) {
            return JSON.parse(fs.readFileSync(CONTACTS_FILE, 'utf8'));
        }
    } catch (e) {
        console.error('[!] Error loading contacts:', e.message);
    }
    return {};
}

function saveContacts(contacts) {
    try {
        fs.writeFileSync(CONTACTS_FILE, JSON.stringify(contacts, null, 2), 'utf8');
    } catch (e) {
        console.error('[!] Error saving contacts:', e.message);
    }
}

global.chatSessions = new Map();


// Function to parse cities, weight, and vehicle from load text
function parseLoadText(text) {
    let fromPlace = "Anywhere";
    let toPlace = "Anywhere";
    let materialInfo = "Industrial Cargo";
    let vehicleInfo = "Any Truck Required";

    const lowerText = text.toLowerCase();

    // Parse Route (e.g. Latur to Mumbai or Latur se Pune or Latur - Mumbai)
    const routeRegex = /([a-zA-Z\u0900-\u097F\s]{3,20})\s+(?:to|se|-|ðŸ‘‰|à¤¸à¥‡)\s+([a-zA-Z\u0900-\u097F\s]{3,20})/i;
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
    <div class="badge">ðŸšš CONFIRMED LOAD</div>
  </div>
  
  <div class="title">LOAD REQUIREMENT ALERT</div>
  
  <div class="message-box">
    ${formattedMsg}
  </div>

  <div class="footer">
    <div class="contact-info">
      <div class="contact-label">FOR BOOKING CALL</div>
      <div class="contact-value">ðŸ“ž ${MY_NUMBER}</div>
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
        console.log(`[âœ”] Cargo card image created at: ${imagePath}`);
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
                        console.log(`[âœ”] Creative song downloaded: ${song.name}`);
                    });
                }).on('error', err => {
                    fs.unlink(destPath, () => { });
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
                console.log(`[âœ”] Premium cargo video generated successfully: ${videoPath}`);
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
            return `ðŸš¨ URGENT CARGO REQUIREMENT ALERT! ðŸš¨\n\n${originalText}\n\nFor bookings, please contact GlobalUnido immediately at ðŸ“ž ${MY_NUMBER}.\n\n#logistics #transport #transportindia #truckload #globalunido #freightforwarder #indianlogistics #trucks`;
        }

        console.log('[i] Calling Gemini AI (gemini-1.5-flash) to write professional logistics copy...');
        const genAI = new GoogleGenerativeAI(apiKey);
        const model = genAI.getGenerativeModel({ model: "gemini-1.5-flash" });
        const prompt = `You are a professional logistics social media copywriter for GLOBALUNIDO. Take this raw logistics cargo load requirement and rewrite it into a highly professional, engaging, clear, and extremely premium social media post caption (suitable for Instagram Reels, Posts, and Stories).
Highlight the contact booking number: ðŸ“ž ${MY_NUMBER}.
Use premium emojis, clean layouts, bullet points, and trending hashtags (e.g. #logistics, #transport, #transportindia).
Include both English and clear Devanagari Hindi phrases so it is perfectly tailored for Indian drivers and loaders.
Raw Cargo Requirement Details:
"${originalText}"`;

        const result = await model.generateContent(prompt);
        const responseText = result.response.text();
        console.log('[âœ”] Successfully generated premium caption using Gemini AI!');
        return responseText;
    } catch (e) {
        console.error('[!] Gemini AI copywriting failed, using fallback:', e.message);
        return `ðŸš¨ URGENT CARGO REQUIREMENT ALERT! ðŸš¨\n\n${originalText}\n\nFor bookings, please contact GlobalUnido immediately at ðŸ“ž ${MY_NUMBER}.\n\n#logistics #transport #transportindia #truckload #globalunido #freightforwarder #indianlogistics #trucks`;
    }
}

// Generic helper to handle a confirmed load (forwarding, logging, image generation, phone delivery, and posting)
async function processConfirmedLoad(channelName, originalNumber, originalEmail, originalText, modifiedText, isUrgent = false, loaderWID = null) {
    try {
        const chats = await client.getChats();
        const targetGroup = chats.find(c => c.name === TARGET_GROUP_NAME && c.isGroup);

        const urgentPrefix = isUrgent ? `ðŸš¨ðŸš¨ *URGENT LOAD / à¤¬à¥‡à¤¹à¤¦ à¤œà¤°à¥‚à¤°à¥€ à¤²à¥‹à¤¡* ðŸš¨ðŸš¨\n\n` : ``;
        const footer = `\n\n*à¤…à¤—à¤° à¤†à¤ªà¤•à¥‹ à¤¯à¤¹ à¤²à¥‹à¤¡ à¤šà¤¾à¤¹à¤¿à¤, à¤¤à¥‹ à¤¤à¥à¤°à¤‚à¤¤ à¤¸à¤‚à¤ªà¤°à¥à¤• à¤•à¤°à¥‡à¤‚ (à¤•à¥‰à¤² à¤•à¤°à¥‡à¤‚):* ðŸ“ž ${MY_NUMBER}`;
        const textToSend = urgentPrefix + modifiedText + footer;

        let groupMsg = null;

        // 1. Forward to WhatsApp Group
        if (targetGroup) {
            groupMsg = await client.sendMessage(targetGroup.id._serialized, textToSend);
            console.log(`[âœ”] Confirmed load forwarded to group: ${TARGET_GROUP_NAME}`);
        } else {
            console.log(`[âŒ] Error: Could not find target group named "${TARGET_GROUP_NAME}".`);
        }

        // 1.5. Forward to Instagram DM Group
        sendToInstagramGroup(TARGET_GROUP_NAME, textToSend).then(() => {
            console.log(`[âœ”] Instagram background DM group forward finished.`);
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
                await client.sendMessage(userWID, imgMedia, { caption: `à¤¯à¤¹à¤¾à¤ à¤†à¤ªà¤•à¥€ à¤¨à¤ˆ à¤²à¥‹à¤¡ à¤‡à¤®à¥‡à¤œ à¤¹à¥ˆ! ðŸ“ŠðŸšš\n\n${textToSend}` });
                console.log(`[âœ”] Cargo card image delivered to user's phone on WhatsApp.`);
            } catch (whatsappImgErr) {
                console.error('[!] Failed to send generated image to user phone:', whatsappImgErr.message);
            }

            // 6. Send Video with Music to User's Phone Chat on WhatsApp (For listening/viewing)
            if (videoPath) {
                try {
                    const vidMedia = MessageMedia.fromFilePath(videoPath);
                    await client.sendMessage(userWID, vidMedia, { caption: `à¤¯à¤¹à¤¾à¤ à¤†à¤ªà¤•à¤¾ à¤¨à¤¯à¤¾ à¤²à¥‹à¤¡ à¤®à¥à¤¯à¥‚à¤œà¤¿à¤• à¤µà¥€à¤¡à¤¿à¤¯à¥‹ à¤¹à¥ˆ! ðŸŽµðŸŽ¥` });
                    console.log(`[âœ”] Cargo music video delivered to user's phone on WhatsApp.`);
                } catch (whatsappVidErr) {
                    console.error('[!] Failed to send generated video to user phone:', whatsappVidErr.message);
                }
            }

            // 7. Post to WhatsApp Status! (Automatically uploads video/image to user's WhatsApp Status)
            try {
                const statusMedia = MessageMedia.fromFilePath(postFilePath);
                const statusCaption = `ðŸš¨ à¤¨à¤¯à¤¾ à¤•à¤¨à¥à¤«à¤°à¥à¤® à¤²à¥‹à¤¡ à¤†à¤¯à¤¾ à¤¹à¥ˆ! ðŸšš\n\n${originalText.substring(0, 100)}...\n\nà¤—à¤¾à¤¡à¤¼à¥€ à¤²à¤—à¤¾à¤¨à¥‡ à¤•à¥‡ à¤²à¤¿à¤ à¤¤à¥à¤°à¤‚à¤¤ à¤•à¥‰à¤² à¤•à¤°à¥‡à¤‚: ðŸ“ž ${MY_NUMBER}`;
                await client.sendMessage('status@broadcast', statusMedia, { caption: statusCaption });
                console.log(`[âœ”] Cargo load successfully published to your WhatsApp Status!`);
            } catch (whatsappStatusErr) {
                console.error('[!] Failed to publish to WhatsApp Status:', whatsappStatusErr.message);
            }

            // 8. Generate professional caption using Gemini AI!
            const captionText = await generateGeminiPostCaption(originalText);

            // 9. Post to Instagram Feed & Stories (Background Auto-posts)
            // Post to Feed
            postToInstagram(postFilePath, captionText).then(() => {
                console.log(`[âœ”] Instagram background Feed auto-post finished.`);
            }).catch(e => {
                console.error(`[!] Instagram Feed auto-post error:`, e.message);
            });

            // Post to Story
            postToInstagramStory(postFilePath).then(() => {
                console.log(`[âœ”] Instagram background Story auto-post finished.`);
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
                const driverNotification = `à¤¨à¤®à¤¸à¥à¤¤à¥‡ à¤­à¤¾à¤ˆ à¤¸à¤¾à¤¹à¤¬! à¤†à¤ªà¤•à¥€ à¤—à¤¾à¤¡à¤¼à¥€ à¤•à¥‡ à¤²à¤¿à¤ à¤à¤• *à¤•à¤¨à¥à¤«à¤°à¥à¤® à¤²à¥‹à¤¡* à¤®à¤¿à¤²à¤¾ à¤¹à¥ˆ! ðŸšš\n\n*à¤²à¥‹à¤¡ à¤¡à¤¿à¤Ÿà¥‡à¤²à¥à¤¸:*\n${textToSend}\n\n*à¤¤à¥à¤°à¤‚à¤¤ à¤¸à¤‚à¤ªà¤°à¥à¤• à¤•à¤°à¥‡à¤‚ (à¤•à¥‰à¤² à¤•à¤°à¥‡à¤‚):* ðŸ“ž ${MY_NUMBER}`;
                await client.sendMessage(driver.sender, driverNotification);
                console.log(`[âœ”] Auto-match: Sent load alert to driver ${driver.sender}`);
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

// ðŸ”´ USER SETTINGS / à¤¸à¥‡à¤Ÿà¤¿à¤‚à¤—à¥à¤¸ (à¤¯à¤¹à¤¾à¤ à¤…à¤ªà¤¨à¥‡ à¤¡à¤¿à¤Ÿà¥‡à¤²à¥à¤¸ à¤­à¤°à¥‡à¤‚)
// ==========================================

// 1. Channel Names (à¤‰à¤¨ à¤¸à¤­à¥€ à¤šà¥ˆà¤¨à¤²à¥à¤¸ à¤•à¥‡ à¤¨à¤¾à¤® à¤¯à¤¹à¤¾à¤ à¤²à¤¿à¤–à¥‡à¤‚ à¤œà¤¹à¤¾à¤ à¤¸à¥‡ à¤®à¥ˆà¤¸à¥‡à¤œ à¤‰à¤ à¤¾à¤¨à¤¾ à¤¹à¥ˆ)
const DEFAULT_CHANNELS = [
    "Transport Parivar Corporation",
    "BANNA TRANSPORT & CONSTRUCTION",
    "shree shyam transport ðŸ™ðŸ»ðŸ™ðŸ»",
    "National Freight Transport (NFS)",
    "ð‘µð‘¨ð‘´ð‘¶ ð‘»ð‘¹ð‘¨ð‘µð‘ºð‘·ð‘¶ð‘¹ð‘»",
    "ð‘´ð‘¨ð‘¯ð‘¨ð‘²ð‘¨ð‘³ ð‘»ð‘¹ð‘¨ð‘µð‘ºð‘·ð‘¶ð‘¹ð‘»",
    "ats transport service",
    "Sri Velavan Transport ðŸšš",
    "Mama Sarkar Group",
    "ð—¦ð—¼ð—»ð—®ð—¹ð—¶ ð—§ð—¿ð—®ð—»ð˜€ð—½ð—¼ð—¿ð˜,ð—¦ð—²ð—¿ð˜ƒð—¶ð—°ð—² ð—¥ð—®ð—»ð—·ð—®ð—»ð—´ð—®ð—¼ð—»,ð— ð—¶ð—±ð—°",
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
    "Rolexâšœï¸ transport",
    "Siddheshwar Transport",
    "Shayam Transport",
    "SSR TRANSPORTATION ðŸ”±",
    "Abhi Transporter ðŸš›",
    "Shukla Logistic Chhattisgarh",
    "GADI WALA TRANSPORT SERVICE ðŸššðŸ›³ï¸ðŸš†âœˆï¸ðŸš›",
    "Shukla logistic Vapi ðŸ“Œ",
    "à¤†à¤°à¥à¤µà¥€ à¤¹à¥‡à¤µà¥€ à¤°à¥‹à¤¡à¤²à¤¾à¤ˆà¤¨à¥à¤¸ à¤®à¤¹à¤¾à¤°à¤¾à¤·à¥à¤Ÿà¥à¤°",
    "à¤¸à¥‹à¤²à¤¾à¤ªà¥‚à¤° 2 ðŸš›ðŸšš",
    "Shukla logistic Madhya Pradesh ðŸ“Œ",
    "~MAHAKAL TRANSPORT ~ 01"
];

const SOURCE_CHANNELS = process.env.SOURCE_CHANNELS ? process.env.SOURCE_CHANNELS.split(",") : DEFAULT_CHANNELS;

// 2. Target Group Name (à¤œà¤¿à¤¸ à¤—à¥à¤°à¥à¤ª à¤®à¥‡à¤‚ à¤®à¥ˆà¤¸à¥‡à¤œ à¤­à¥‡à¤œà¤¨à¤¾ à¤¹à¥ˆ à¤‰à¤¸à¤•à¤¾ à¤¬à¤¿à¤²à¥à¤•à¥à¤² à¤¸à¤¹à¥€ à¤¨à¤¾à¤®)
const TARGET_GROUP_NAME = process.env.TARGET_GROUP_NAME || "GlobalUnido loading requirements";

// 3. YOUR DETAILS (à¤†à¤ªà¤•à¥€ à¤¡à¤¿à¤Ÿà¥‡à¤²à¥à¤¸ à¤œà¥‹ à¤¹à¤° à¤®à¥ˆà¤¸à¥‡à¤œ à¤®à¥‡à¤‚ à¤‘à¤Ÿà¥‹à¤®à¥ˆà¤Ÿà¤¿à¤•à¤²à¥€ à¤²à¤—à¤¾à¤¨à¥€ à¤¹à¥ˆà¤‚)
const MY_NUMBER = process.env.MY_NUMBER || "8200210397";
const MY_EMAIL = process.env.MY_EMAIL || "supportglobalunido@gmail.com";
const MY_COMPANY = process.env.MY_COMPANY || "GLOBALUNIDO";

// 4. GEMINI AI API KEY (à¤¯à¤¹à¤¾à¤ à¤…à¤ªà¤¨à¥€ à¤œà¥‡à¤®à¤¿à¤¨à¥€ API à¤•à¥€ à¤¡à¤¾à¤²à¥‡à¤‚ à¤¤à¤¾à¤•à¤¿ à¤œà¥‡à¤®à¤¿à¤¨à¥€ à¤ªà¥à¤°à¥‹à¤«à¥‡à¤¶à¤¨à¤² à¤•à¥‰à¤ªà¥€à¤°à¤¾à¤‡à¤Ÿà¤¿à¤‚à¤— à¤•à¤° à¤¸à¤•à¥‡, à¤–à¤¾à¤²à¥€ à¤°à¤–à¤¨à¥‡ à¤ªà¤° à¤¡à¤¿à¤«à¥‰à¤²à¥à¤Ÿ à¤•à¥‰à¤ªà¥€à¤°à¤¾à¤‡à¤Ÿà¤¿à¤‚à¤— à¤¹à¥‹à¤—à¥€)
const GEMINI_API_KEY = process.env.GEMINI_API_KEY || "";
global.GEMINI_API_KEY = GEMINI_API_KEY;

// ==========================================
// ðŸ”µ CLOUD WEB SERVER & DYNAMIC QR DISPLAY (Render.com 24/7 Support)
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
                    <h1>ðŸ† GLOBALUNIDO Bot Is Active & Running! ðŸš€</h1>
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
                    <h1>Scan to Login GLOBALUNIDO Bot ðŸšš</h1>
                    <p>Open WhatsApp on your phone, go to Linked Devices, and scan this QR code:</p>
                    <div class="qr-container">
                        <img src="https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=${encodeURIComponent(latestQRCode)}" />
                    </div>
                    <p style="color: #94a3b8; margin-top: 20px; font-size: 14px;">This page refreshes automatically. Scan within 20 seconds of load!</p>
                </div>
                <script>
                    setInterval(() => { location.reload(); }, 20000);
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
                <h1>â³ Initializing WhatsApp Web Client...</h1>
                <p>Generating a secure QR code connection. Please wait, this page refreshes automatically...</p>
                <script>
                    setInterval(() => { location.reload(); }, 3000);
                </script>
            </body>
            </html>
        `);
    }
});

app.listen(PORT, '0.0.0.0', () => {
    console.log(`[âœ”] Cloud Web Server listening on port ${PORT} (0.0.0.0)`);
});

// ==========================================
// ðŸ”µ AUTOMATION LOGIC / à¤•à¥‹à¤¡à¤¿à¤‚à¤— (à¤¨à¥€à¤šà¥‡ à¤•à¥à¤› à¤®à¤¤ à¤¬à¤¦à¤²à¥‡à¤‚)
// ==========================================

function cleanSessionLocks() {
    const sessionDir = path.join(__dirname, '.wwebjs_auth');
    if (!fs.existsSync(sessionDir)) return;
    
    console.log('[i] Checking for and cleaning Chrome lock files in session directory...');
    function deleteLocks(dir) {
        try {
            const files = fs.readdirSync(dir);
            for (const file of files) {
                const fullPath = path.join(dir, file);
                const stat = fs.statSync(fullPath);
                if (stat.isDirectory()) {
                    deleteLocks(fullPath);
                } else if (file === 'SingletonLock' || file === 'lock' || file.includes('lock')) {
                    fs.unlinkSync(fullPath);
                    console.log(`[âœ”] Deleted active session lock file: ${fullPath}`);
                }
            }
        } catch (e) {
            // Ignore file errors
        }
    }
    deleteLocks(sessionDir);
}

// Execute clean-up on boot
cleanSessionLocks();

function getChromeExecutablePath() {
    if (process.platform === 'win32') return null; // Use default on Windows
    
    const cacheDir = process.env.PUPPETEER_CACHE_DIR || '/opt/render/.cache/puppeteer';
    console.log(`[i] Searching for Chrome executable in cache dir: ${cacheDir}`);
    
    function searchChrome(dir) {
        if (!fs.existsSync(dir)) return null;
        try {
            const files = fs.readdirSync(dir);
            for (const file of files) {
                const fullPath = path.join(dir, file);
                const stat = fs.statSync(fullPath);
                if (stat.isDirectory()) {
                    const found = searchChrome(fullPath);
                    if (found) return found;
                } else if (file === 'chrome') {
                    return fullPath;
                }
            }
        } catch (e) {
            // Ignore read errors
        }
        return null;
    }
    
    const foundPath = searchChrome(cacheDir);
    if (foundPath) {
        console.log(`[âœ”] Located Chrome executable: ${foundPath}`);
        return foundPath;
    }
    
    const fallbacks = [
        '/usr/bin/google-chrome',
        '/usr/bin/chromium',
        '/usr/bin/chromium-browser'
    ];
    for (const fb of fallbacks) {
        if (fs.existsSync(fb)) {
            console.log(`[âœ”] Using fallback Chrome path: ${fb}`);
            return fb;
        }
    }
    
    console.log('[!] Warning: Could not locate Chrome executable. Falling back to default Puppeteer launch.');
    return null;
}

const chromePath = getChromeExecutablePath();

const MONGO_URI = process.env.MONGO_URI || "mongodb+srv://globalunido:unidosecretpass2026@cluster0.3n46pbk.mongodb.net/whatsapp_sessions?retryWrites=true&w=majority";

// NOTE: client is declared here, assigned inside mongoose.connect
let client;

process.on('uncaughtException', (err) => {
    console.error('[ðŸš¨ UNCAUGHT EXCEPTION]:', err.stack || err);
});

process.on('unhandledRejection', (reason, promise) => {
    console.error('[ðŸš¨ UNHANDLED REJECTION]:', reason.stack || reason);
});

// Connect to MongoDB FIRST, then create MongoStore + Client
console.log('[i] Connecting to MongoDB Atlas database...');
mongoose.connect(MONGO_URI).then(() => {
    console.log('[âœ”] MongoDB Connected successfully!');

    const store = new MongoStore({ mongoose: mongoose });
    client = new Client({
        authStrategy: new RemoteAuth({
            store: store,
            backupSyncIntervalMs: 300000
        }),
        puppeteer: {
            headless: true,
            executablePath: chromePath || undefined,
            args: [
                '--no-sandbox',
                '--disable-setuid-sandbox',
                '--disable-dev-shm-usage',
                '--disable-accelerated-2d-canvas',
                '--no-first-run',
                '--no-zygote',
                '--disable-gpu',
                '--disable-extensions',
                '--single-process',
                '--disable-features=site-per-process',
                '--disable-features=IsolateOrigins',
                '--js-flags="--max-old-space-size=100"'
            ]
        }
    });

    client.on('remote_session_saved', () => {
        console.log('[âœ”] Remote session saved to MongoDB Atlas!');
    });

    client.on('qr', (qr) => {
        console.log('\n[!] Please SCAN the QR Code below with your WhatsApp:\n');
        qrcode.generate(qr, { small: true });
        latestQRCode = qr;
        isBotLoggedIn = false;
        console.log('[i] QR Code generated. Visit the web page to scan.');
    });

    client.on('ready', () => {
        console.log('\n[âœ”] WhatsApp Bot is READY and ACTIVE!');
        isBotLoggedIn = true;
        latestQRCode = null;
    });

    client.on('authenticated', () => {
        console.log('[âœ”] WhatsApp session authenticated successfully!');
        isBotLoggedIn = true;
    });

    client.on('auth_failure', (msg) => {
        console.error('[!] Auth failure:', msg);
        isBotLoggedIn = false;
    });

    client.on('disconnected', (reason) => {
        console.log('[!] WhatsApp client disconnected:', reason);
        isBotLoggedIn = false;
        // Auto reconnect after 5 seconds
        setTimeout(() => {
            console.log('[i] Attempting to reconnect...');
            client.initialize().catch(e => console.error('[!] Reconnect failed:', e.message));
        }, 5000);
    });

    client.on('message', async (msg) => {
        try {
            const chat = await msg.getChat();
            await processIncomingMessage(client, msg, chat, msg.body);
        } catch (err) {
            console.error('[!] Error processing message:', err.message);
        }
    });

    client.on('message_create', async (msg) => {
        try {
            if (msg.fromMe) {
                const chat = await msg.getChat();
                await processIncomingMessage(client, msg, chat, msg.body);
            }
        } catch (err) {
            console.error('[!] Error processing message_create:', err.message);
        }
    });

    console.log('[i] Starting WhatsApp Client...');
    return client.initialize();
}).then(() => {
    console.log('[âœ”] client.initialize() resolved!');
}).catch(err => {
    console.error('[âŒ] Fatal startup error:', err.message || err);
    process.exit(1);
});
