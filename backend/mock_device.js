const axios = require('axios');
const FormData = require('form-data');
const fs = require('fs');

const SERVER_URL = 'http://localhost:3000/device_event';   // Change if your endpoint is different

async function sendHikvisionEvent() {
    try {
        const form = new FormData();

        // === AccessControllerEvent part (exactly as Hikvision sends) ===
        const eventData = {
            AccessControllerEvent: {
                name: "John Doe",
                employeeNoString: "wE2Y7Hbvu",
                // Add any other fields you want:
                // cardNo: "123456789",
                // doorNo: 1,
                // currentVerifyMode: "face",
                // dateTime: new Date().toISOString()
            }
        };

        form.append('AccessControllerEvent', JSON.stringify(eventData), {
            contentType: 'text/json'   // Important for Hikvision compatibility
        });

        // === Picture part (uncomment if you want to send an image) ===
        const imagePath = 'test.jpg'; // Put your image in the same folder or change path
        if (fs.existsSync(imagePath)) {
            form.append('Picture_Name_Placeholder', fs.createReadStream(imagePath), {
                filename: 'snapshot.jpg',
                contentType: 'image/jpeg'
            });
        } else {
            console.log('⚠️ No image found, sending event only.');
        }

        const res = await axios.post(SERVER_URL, form, {
            headers: {
                ...form.getHeaders(),
                'Accept-Language': 'en-us',
                'Connection': 'keep-alive',
                // 'Date': new Date().toUTCString() // Axios usually handles this
            },
            maxBodyLength: Infinity,
            maxContentLength: Infinity
        });

        console.log(`✅ Event sent successfully → Status: ${res.status}`);
        console.log('Response:', res.data);
    } catch (err) {
        console.error('❌ Failed to send event:', err.message);
        if (err.response) {
            console.error('Status:', err.response.status);
            console.error('Response:', err.response.data);
        }
    }
}

// Run immediately + every 2 seconds
sendHikvisionEvent();
setInterval(sendHikvisionEvent, 2000);

console.log('🚀 Hikvision mock device started...');
console.log(`Sending to → ${SERVER_URL}`);


// const axios = require('axios');
// const FormData = require('form-data');
// const fs = require('fs');

// const SERVER_URL = 'http://localhost:3000/device_event';

// async function sendHikvisionEvent() {
//     try {
//         const form = new FormData();

//         // ==================== XML Event Payload ====================
//         const xmlEvent = `<?xml version="1.0" encoding="UTF-8"?>
// <AccessControllerEvent xmlns="http://www.hikvision.com/ver20/XMLSchema">
//     <name>John Doe</name>
//     <employeeNoString>iC3BTazvS</employeeNoString>
//     <cardNo>123456789</cardNo>
//     <doorNo>1</doorNo>
//     <currentVerifyMode>face</currentVerifyMode>
//     <verifyType>local</verifyType>
//     <dateTime>${new Date().toISOString()}</dateTime>
//     <statusValue>1</statusValue>
// </AccessControllerEvent>`;

//         // Add as form field with text/xml content type (this matches Hikvision behavior)
//         form.append('AccessControllerEvent', xmlEvent, {
//             contentType: 'text/xml',
//             filename: 'event.xml'        // optional but helps some parsers
//         });

//         // ==================== Optional Picture ====================
//         const imagePath = 'test.jpg';   // change if needed
//         if (fs.existsSync(imagePath)) {
//             form.append('Picture_Name_Placeholder', fs.createReadStream(imagePath), {
//                 filename: 'snapshot.jpg',
//                 contentType: 'image/jpeg'
//             });
//         } else {
//             console.log('⚠️  test.jpg not found → sending event only');
//         }

//         const res = await axios.post(SERVER_URL, form, {
//             headers: {
//                 ...form.getHeaders(),
//                 'Accept-Language': 'en-us',
//                 'Connection': 'keep-alive',
//             },
//             maxBodyLength: Infinity,
//             maxContentLength: Infinity
//         });

//         console.log(`✅ XML Event sent successfully → Status: ${res.status}`);
//         if (res.data) console.log('Response:', res.data);
//     } catch (err) {
//         console.error('❌ Failed to send event:', err.message);
//         if (err.response) {
//             console.error('Status:', err.response.status);
//             console.error('Data:', err.response.data);
//         }
//     }
// }

// // Start sending
// sendHikvisionEvent();
// setInterval(sendHikvisionEvent, 8000);

// console.log('🚀 Hikvision XML Mock Device running...');
// console.log(`Sending XML events every 8s to ${SERVER_URL}`);
