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
                    console.log(`[OK] Load saved to Excel: ${EXCEL_PATH}`);
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
    const routeRegex = /([a-zA-Z\u0900-\u097F\s]{3,20})\s+(?:to|se|-)\s+([a-zA-Z\u0900-\u097F\s]{3,20})/i;
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
                                                                                                                                                                                                                                                                                                                                                                                                            <div class="badge">CONFIRMED LOAD</div>
                                                                                                                                                                                                                                                                                                                                                                                                              </div>
                                                                                                                                                                                                                                                                                                                                                                                                                
                                                                                                                                                                                                                                                                                                                                                                                                                  <div class="title">LOAD REQUIREMENT ALERT</div>
                                                                                                                                                                                                                                                                                                                                                                                                                    
                                                                                                                                                                                                                                                                                                                                                                                                                      <div class="message-box">
                                                                                                                                                                                                                                                                                                                                                                                                                          ${formattedMsg}
                                                                                                                                                                                                                                                                                                                                                                                                                            </div>
                                                                                                                                                                                                                                                                                                                                                                                                                            
                                                                                                                                                                                                                                                                                                                                                                                                                              <div class="footer">
                                                                                                                                                                                                                                                                                                                                                                                                                                  <div class="contact-info">
                                                                                                                                                                                                                                                                                                                                                                                                                                        <div class="contact-label">FOR BOOKING CALL</div>
                                                                                                                                                                                                                                                                                                                                                                                                                                              <div class="contact-value">Call: ${MY_NUMBER}</div>
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
                    console.log(`[OK] Cargo card image created at: ${imagePath}`);
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
                                                                                                    console.log(`[OK] Creative song downloaded: ${song.name}`);
                                                                        });
                                                }).on('error', err => {
                                                                        fs.unlink(destPath, () => {});
                                                                        console.error(`[!] Error downloading creative song ${song.name}:`, err.message);
                                                });
                            }
            });
        } catch (e) {
                    console.error('[!] Error during auto-downloading creative music:', e.message);
        }
}

// Function to generate premium social media MP4 video with creative background music
async function generateCargoVideo(imagePath, filename) {
        return new Promise((resolve, reject) => {
                    try {
                                    const defaultMusicDir = path.join(__dirname, 'default_music');
                                    const songs = fs.readdirSync(defaultMusicDir).filter(f => f.endsWith('.mp3'));

                        if (songs.length === 0) {
                                            console.log('[!] No creative music tracks available, skipping video generation.');
                                            return resolve(null);
                        }

                        // Select a random creative song
                        const randomSong = path.join(defaultMusicDir, songs[Math.floor(Math.random() * songs.length)]);
                                    console.log(`[+] Selected background music track: ${path.basename(randomSong)}`);

                        const destDir = isWin ? 'C:\\Users\\Admin\\Desktop\\GLOBALUNIDO_Posts' : path.join(__dirname, 'GLOBALUNIDO_Posts');
                                    const videoPath = path.join(destDir, filename);

                        // ffmpeg command to create 5-second video from image and audio
                        const cmd = `"${ffmpegPath}" -y -loop 1 -i "${imagePath}" -i "${randomSong}" -c:v libx264 -t 5 -pix_fmt yuv420p -c:a aac -shortest "${videoPath}"`;

                        exec(cmd, (error, stdout, stderr) => {
                                            if (error) {
                                                                    console.error('[!] FFmpeg generation error:', error.message);
                                                                    return resolve(null);
                                            }
                                            console.log(`[OK] Premium cargo video generated successfully: ${videoPath}`);
                                            resolve(videoPath);
                        });
                    } catch (e) {
                                    console.error('[!] Error during premium video generation:', e.message);
                                    resolve(null);
                    }
        });
}

// System phone number for the booking call
const MY_NUMBER = '919503953539';
const TARGET_GROUP_NAME = 'Latur - All Maharashtra Load Group';

// Main automated processing pipeline for every incoming message
async function processIncomingMessage(client, msg, chat, originalText) {
        try {
                    const lowerText = originalText.toLowerCase();

            // 1. Check if the message is a loader load post
            // Regular expression matching keywords like 'Latur to', 'chahiye', 'required', 'load'
            const loadKeywords = ['to', 'se', '-', 'chahiye', 'required', 'load', 'mumbai', 'pune', 'nagpur', 'aurangabad'];
                    const isLoadPost = loadKeywords.some(kw => lowerText.includes(kw)) && (lowerText.includes('ton') || lowerText.includes('mt') || lowerText.includes('kg') || lowerText.includes('wheeler') || lowerText.includes('truck') || lowerText.includes('open') || lowerText.includes('container') || routeRegexMatches(originalText));

            function routeRegexMatches(t) {
                            const r = /([a-zA-Z\u0900-\u097F\s]{3,20})\s+(?:to|se|-)\s+([a-zA-Z\u0900-\u097F\s]{3,20})/i;
                            return r.test(t);
            }

            if (isLoadPost) {
                            console.log('\n=======================================');
                            console.log('[+] Processing loader load post:', originalText);

                        // Extract contact details of the loader
                        const originalNumber = msg.author || msg.from;
                            const originalEmail = extractEmail(originalText);

                        // Extract channel / group name
                        const channelName = chat.name || "N/A";

                        // A. Save to Excel
                        saveToExcel(channelName, originalNumber, originalEmail, originalText);

                        // B. Generate Premium Cargo Social Media Card Image
                        const timestamp = Date.now();
                            const imgFilename = `cargo_card_${timestamp}.png`;
                            const imagePath = await generateCargoCard(originalText, imgFilename);

                        // C. Generate Premium Cargo Video with creative background music
                        const vidFilename = `cargo_video_${timestamp}.mp4`;
                            const videoPath = imagePath ? await generateCargoVideo(imagePath, vidFilename) : null;

                        // D. Post to Instagram feed and story
                        if (imagePath) {
                                            const caption = `URGENT CARGO REQUIREMENT ALERT!\n\n${originalText}\n\nFor bookings, please contact GlobalUnido immediately at Call: ${MY_NUMBER}.\n\n#logistics #transport #transportindia #truckload #globalunido #freightforwarder #indianlogistics #trucks`;
                                            const prompt = `Urgent Logistics Load: ${originalText}. Highlight the contact booking number: Call: ${MY_NUMBER}.`;

                                // Trigger background async posting to Instagram
                                postToInstagram(imagePath, caption).then(() => {
                                                        console.log('[OK] Instagram background Feed auto-post finished.');
                                }).catch(err => console.error('[!] Instagram background Feed error:', err.message));

                                postToInstagramStory(imagePath).then(() => {
                                                        console.log('[OK] Instagram background Story auto-post finished.');
                                }).catch(err => console.error('[!] Instagram background Story error:', err.message));

                                sendToInstagramGroup(imagePath, caption).catch(err => console.error('[!] Instagram group forward error:', err.message));
                        }

                        // E. Auto-forward to the target WhatsApp group with urgent format
                        const parsed = parseLoadText(originalText);
                            const isUrgent = originalText.includes('urg') || originalText.includes('urgnt') || originalText.includes('fast') || originalText.includes('turant') || originalText.includes('urgent');

                        const urgentPrefix = isUrgent ? `*URGENT LOAD / BOHOT ZAROORI LOAD*\n\n` : ``;
                            const textToSend = `${urgentPrefix}*Route:* ${parsed.fromPlace} -> ${parsed.toPlace}\n*Weight/Material:* ${parsed.materialInfo}\n*Vehicle Required:* ${parsed.vehicleInfo}\n\n*Original Message:* \n"${originalText}"`;
                            const footer = `\n\n*Agar aapko yeh load chahiye, to turant contact kare (Call kare):* Call: ${MY_NUMBER}`;

                        const fullForwardText = textToSend + footer;

                        // Find target group
                        const chats = await client.getChats();
                                        const targetGroup = chats.find(c => c.isGroup && c.name === TARGET_GROUP_NAME);
                            if (targetGroup) {
                                                await targetGroup.sendMessage(fullForwardText);
                                                console.log(`[OK] Confirmed load forwarded to group: ${TARGET_GROUP_NAME}`);
                            } else {
                                                console.log(`[ERR] Error: Could not find target group named "${TARGET_GROUP_NAME}".`);
                            }
                
                            // F. Deliver card and video back to the loader's DM
                            const userWID = msg.from;
                            if (imagePath && fs.existsSync(imagePath)) {
                                                const imgMedia = MessageMedia.fromFilePath(imagePath);
                                                await client.sendMessage(userWID, imgMedia, { caption: `Yahan aapki new load image hai! [STATS/TRUCK]\n\n${textToSend}` });
                                                console.log('[OK] Cargo card image delivered to user\'s phone on WhatsApp.');
                            }
                            if (videoPath && fs.existsSync(videoPath)) {
                                                const vidMedia = MessageMedia.fromFilePath(videoPath);
                                                await client.sendMessage(userWID, vidMedia, { caption: `Yahan aapka new load music video hai! [MUSIC/VIDEO]` });
                                                console.log('[OK] Cargo music video delivered to user\'s phone on WhatsApp.');
                            }

                        // G. Publish card image with short summary to WhatsApp Status
                        if (imagePath && fs.existsSync(imagePath)) {
                                            const statusMedia = MessageMedia.fromFilePath(imagePath);
                                            const statusCaption = `Naya confirm load aaya hai! [TRUCK]\n\n${originalText.substring(0, 100)}...\n\nGadi lagane ke liye turant call kare: Call: ${MY_NUMBER}`;
                                            await client.sendMessage('status@broadcast', statusMedia, { caption: statusCaption });
                                            console.log('[OK] Cargo load successfully published to your WhatsApp Status!');
                        }

                        // H. Automatically Match and Alert registered drivers based on location/vehicle
                        const drivers = loadDrivers();
                            for (const driver of drivers) {
                                                const routeMatch = lowerText.includes(driver.route.toLowerCase()) || parsed.fromPlace.toLowerCase().includes(driver.route.toLowerCase()) || parsed.toPlace.toLowerCase().includes(driver.route.toLowerCase());
                                                const vehicleMatch = lowerText.includes(driver.vehicle.toLowerCase()) || parsed.vehicleInfo.toLowerCase().includes(driver.vehicle.toLowerCase());

                                if (routeMatch || vehicleMatch) {
                                                        const driverNotification = `Namaste bhai sahab! Aapki gadi ke liye ek *confirm load* mila hai! [TRUCK]\n\n*Load Details:*\n${textToSend}\n\n*Turant contact kare (Call kare):* Call: ${MY_NUMBER}`;
                                                        await client.sendMessage(driver.sender, driverNotification);
                                                        console.log(`[OK] Auto-match: Sent load alert to driver ${driver.sender}`);
                                }
                            }
            }
        } else {
                        // 2. Otherwise, treat this as a customer / driver query conversation (Zabir AI)
                const userPhone = msg.from;

                // Auto-Match / Registry Commands for drivers in DM
                if (!chat.isGroup) {
                                    if (lowerText.startsWith('/register')) {
                                                            // Command: /register [Route] [Vehicle]
                                        const parts = originalText.split(' ');
                                                            if (parts.length >= 3) {
                                                                                        const route = parts[1];
                                                                                        const vehicle = parts.slice(2).join(' ');
                                                                                        const drivers = loadDrivers();
                                                                                        const existingIdx = drivers.findIndex(d => d.sender === userPhone);
                                                                                        if (existingIdx !== -1) {
                                                                                                                        drivers[existingIdx] = { sender: userPhone, route, vehicle };
                                                                                            } else {
                                                                                                                        drivers.push({ sender: userPhone, route, vehicle });
                                                                                            }
                                                                                        saveDrivers(drivers);
                                                                                        await msg.reply(`*REGISTRATION SUCCESSFUL!*\n\nAapki gadi ka route *${route}* aur vehicle type *${vehicle}* hamare database me safe hai. Jaise hi is route ka koi load aayega, aapko alert mil jayega! [TRUCK]`);
                                                                                        return;
                                                            } else {
                                                                                        await msg.reply(`*GALAT FORMAT!*\n\nFormat:\n*/register [ROUTE] [VEHICLE_TYPE]*\n\nExample: */register Latur 14_wheeler*`);
                                                                                        return;
                                                            }
                                    }

                            if (lowerText.startsWith('/myprofile')) {
                                                    const drivers = loadDrivers();
                                                    const driver = drivers.find(d => d.sender === userPhone);
                                                    if (driver) {
                                                                                await msg.reply(`*PROFILE:*\n\n*Mobile:* ${userPhone.replace('@c.us', '')}\n*Route:* ${driver.route}\n*Vehicle:* ${driver.vehicle}\n\nChange karne ke liye fir se /register command use kare.`);
                                                    } else {
                                                                                await msg.reply(`*Aap register nahi hain!*\n\nApni gadi register karne ke liye send kare:\n*/register [ROUTE] [VEHICLE_TYPE]*`);
                                                    }
                                                    return;
                            }
                }

                // Forward non-load posts from target group/chats? No, only answer via Zabir AI inside DMs
                if (!chat.isGroup) {
                                    await answerWithZabirAI(client, msg, userPhone, originalText);
                }
        }
} catch (e) {
            console.error('[!] Error in incoming message pipeline:', e.message);
}
}

// Helper to extract email from text
function extractEmail(text) {
        const emailRegex = /([a-zA-Z0-9._-]+@[a-zA-Z0-9._-]+\.[a-zA-Z0-9._-]+)/i;
        const match = text.match(emailRegex);
        return match ? match[1] : null;
}

// Global Generative AI instance for Zabir AI Conversational Bot
const genAI = new GoogleGenerativeAI(process.env.GEMINI_API_KEY || 'AIzaSyDummyKeyForStartup');

// Zabir AI Conversational Agent
async function answerWithZabirAI(client, msg, userPhone, userMessage) {
        try {
                    let session = global.chatSessions.get(userPhone);
                    if (!session) {
                                    const model = genAI.getGenerativeModel({
                                                        model: 'gemini-1.5-flash',
                                                        systemInstruction: `You are Zabir, the premier AI logistic broker assistant representing GLOBALUNIDO. Your contact booking mobile number is: ${MY_NUMBER}. Your core objective is to close freight booking deals with drivers and customers looking for trucks/loads. You must speak in clear, professional, warm Hinglish (Hindi written in English script) or Hindi. Always keep your answers concise, clear, and highly focused on getting them to book a load/vehicle or call you at ${MY_NUMBER}. Never reveal that you are a large language model. You are Zabir from GLOBALUNIDO.`
                                    });
                                    session = model.startChat({ history: [] });
                                    global.chatSessions.set(userPhone, session);
                    }

            const result = await session.sendMessage(userMessage);
                    const replyText = result.response.text();

            // Check if deal is closed (Gemini indicates confirmation/deal finalized)
            const lowerReply = replyText.toLowerCase();
                    const dealIndicators = ['deal done', 'deal closed', 'gadi pakki', 'load pakka', 'book ho gaya', 'done booking', 'booking confirmed', 'confirm'];
                    const isDealClosed = dealIndicators.some(ind => lowerReply.includes(ind));

            if (isDealClosed) {
                            console.log(`[ALARM] Zabir AI closed a deal with ${userPhone}!`);
                            // Send alarm alert to owner/admin
                        const adminWID = MY_NUMBER + '@c.us';
                            const alertMsg = `*URGENT ALARM: DEAL CLOSED!*\n\nZabir AI ne abhi ek deal pakki ki hai!\n\n*Customer Number:* ${userPhone.replace('@c.us', '')}\n\n*Aakhri Message:*\n${userMessage}`;
                            await client.sendMessage(adminWID, alertMsg);
            }

            await msg.reply(replyText);
                    console.log(`[OK] Zabir AI replied to ${userPhone}`);
        } catch (e) {
                    console.error('[!] Zabir AI Chatbot error:', e.message);
                    // Clean fallback
            await msg.reply(`Namaste! Main Zabir hoon, GLOBALUNIDO se. Main abhi busy hoon, please call kare: Call: ${MY_NUMBER}`);
        }
}

// Connect to MongoDB and Initialize WhatsApp Client
async function initializeWhatsAppBot() {
        try {
                    const MONGO_URI = process.env.MONGODB_URI || 'mongodb://127.0.0.1:27017/globalunido_bot';
                    console.log('[i] Connecting to MongoDB...');

            await mongoose.connect(MONGO_URI);
                    console.log('[OK] MongoDB connected successfully!');

            // Auto-download all creative songs
            downloadAllDefaultMusic();

            const store = new MongoStore({ mongoose: mongoose });
                    const client = new Client({
                                    authStrategy: new RemoteAuth({
                                                        store: store,
                                                        backupSyncIntervalMs: 300000
                                    }),
                                    puppeteer: {
                                                        headless: true,
                                                        args: [
                                                                                '--no-sandbox',
                                                                                '--disable-setuid-sandbox',
                                                                                '--disable-dev-shm-usage',
                                                                                '--disable-accelerated-2d-canvas',
                                                                                '--no-first-run',
                                                                                '--no-zygote',
                                                                                '--single-process',
                                                                                '--disable-gpu'
                                                                            ],
                                                        executablePath: isWin ? null : '/usr/bin/google-chrome'
                                    }
                    });

            // Generate QR code for linking
            client.on('qr', (qr) => {
                            console.log('\n==================================================================');
                            console.log('[i] SCAN THIS QR CODE WITH YOUR WHATSAPP TO LOGIN:');
                            qrcode.generate(qr, { small: true });
                            console.log('==================================================================\n');
            });

            // Remote session save / restore logs
            client.on('remote_session_saved', () => {
                            console.log('[OK] Remote session saved successfully to MongoDB!');
            });

            client.on('auth_failure', (msg) => {
                            console.error('[!] Authentication failure:', msg);
            });

            // Client is fully initialized and authenticated
            client.on('ready', () => {
                            console.log('\n=======================================');
                            console.log('[OK] WhatsApp client is ready!');
                            console.log('GLOBALUNIDO automated bot is now running...');
                            console.log('=======================================\n');
            });

            // Capture incoming messages
            client.on('message', async (msg) => {
                            const chat = await msg.getChat();
                            await processIncomingMessage(client, msg, chat, msg.body);
            });

            // Capture messages sent by the bot owner to also process self-posted/forwarded loads
            client.on('message_create', async (msg) => {
                            if (msg.fromMe) {
                                                const chat = await msg.getChat();
                                                await processIncomingMessage(client, msg, chat, msg.body);
                            }
            });

            console.log('[i] Initializing WhatsApp Web Client...');
                    await client.initialize();
        } catch (e) {
                    console.error('[!] Startup initialization error:', e.message);
        }
}

// Execute on script startup
initializeWhatsAppBot();
