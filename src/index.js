const express = require('express');
const { google } = require('googleapis');
const fs = require('fs');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

// Your OAuth2 credentials from Google Cloud Console
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const REDIRECT_URI = process.env.REDIRECT_URI;

// Scopes for Google Drive (you can modify this as needed)
const SCOPES = 'https://www.googleapis.com/auth/drive';

// Create a new OAuth2 client
const oauth2Client = new google.auth.OAuth2(
    CLIENT_ID,
    CLIENT_SECRET,
    REDIRECT_URI);

try {
    const creds = fs.readFileSync("creds.json");
    oauth2Client.setCredentials(JSON.parse(creds));
} catch (error) {
    console.log("No creds found");
}

app.get('/auth/google', (req, res) => {
    //    try {
    const url = oauth2Client.generateAuthUrl({
        access_type: "offline",
        scope: ["https://www.googleapis.com/auth/userinfo.profile",
        'https://www.googleapis.com/auth/spreadsheets']
    });
    res.redirect(url);
    //    } catch (error) {
    //     return res.status(500).json({message: error.message});
    //    }
})

app.get('/google/redirect', async (req, res) => {
    try {

        const { code } = req.query;
        const { tokens } = await oauth2Client.getToken(code);
        oauth2Client.setCredentials(tokens);
        fs.writeFileSync('creds.json', JSON.stringify(tokens));
        res.send('Success');
    } catch (error) {
        return res.status(500).json({ message: error.message });
    }
})

app.get('/get-sheets', async (req, res) => {
    try {
        const spreadsheetId = '1swpwoOccWZ7KCjRb1dLyvhZtmQ-UaQUGqAtN38XmZvI';
        const sheets = google.sheets({ version: 'v4', auth: oauth2Client });
        let values = [
          ['Chris', 'Male', '1. Freshman', 'FL', 'Art', 'Baseball'],
          // Potential next row
        ];
        const resource = {
          values,
        };
        sheets.spreadsheets.values.append(
          {
            spreadsheetId,
            range: 'hue!A1',
            valueInputOption: 'RAW',
            resource: resource,
          },
          (err, result) => {
            if (err) {
              // Handle error
              console.log(err);
            } else {
              console.log(
                '%d cells updated on range: %s',
                result.data.updates.updatedCells,
                result.data.updates.updatedRange
              );
            }
          }
        );
     
    } catch (error) {
        return res.status(500).json({ message: error.message });
    }
})

app.get('/get-hue-sheets', async (req, res) => {
    try {
        const spreadsheetId = '1swpwoOccWZ7KCjRb1dLyvhZtmQ-UaQUGqAtN38XmZvI';
        const sheets = google.sheets({ version: 'v4', auth: oauth2Client });
        sheets.spreadsheets.values.get(
            {
              spreadsheetId,
            range: 'hue!A1:F',
            valueRenderOption: 'FORMULA',
            fields: '*',
            },
            (err, res) => {
              if (err) return console.log('The API returned an error: ' + err);
                const rows = res.data.values;
              console.log('after',rows);
              createData(rows, res);
            }
          );
    } catch (error) {
        return res.status(500).json({ message: error.message });
    }
})

const createData = async (rows, res) => {
  try {
    const spreadsheetId = '1nP8aV1bIpTzQmTn5QV1BEdJuTrjfOms8wTrM3vFf_10';
    const sheets = google.sheets({ version: 'v4', auth: oauth2Client });

    const values = rows;

    const resource = {
      values,
    };

  await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: 'love!A1', // Specify the target range
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      resource,
    });

    // console.log('success');

    // Apply cell borders to the inserted rows
    const batchUpdateResult = await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      resource: {
        requests: [
          {
            updateBorders: {
              range: {
                sheetId: 0,
                startRowIndex: 0, // Adjust as needed
                endRowIndex: values.length, // Apply borders to all inserted rows
                startColumnIndex: 0,
                endColumnIndex: values.length, // Adjust the column range as needed
              },
              // insertDataOption: 'INSERT_ROWS',
              top: {
                style: "SOLID",
                width: 1,
                color: {
                  red: 0,
                  green: 0,
                  blue: 0,
                },
              },
              bottom: {
                style: "SOLID",
                width: 1,
                color: {
                  red: 0,
                  green: 0,
                  blue: 0,
                },
              },
                innerHorizontal: {
                  style: "SOLID",
                  width: 1,
                  color: {
                    red: 0,
                    green: 0,
                    blue: 0,
                  },
                },
                innerVertical: { // Specify inner vertical borders
                  style: "SOLID",
                  width: 1,
                  color: {
                    red: 0,
                    green: 0,
                    blue: 0,
                  },
                },
            },
          },
        ],
      },
    });

    console.log(batchUpdateResult);
  } catch (error) {
    console.log(error);
  }
};


app.get('/create-sheets', async (req, res) => {
    try {
        const title = req.query.title;
        const service = google.sheets({version: 'v4', auth: oauth2Client});
        const resource = {
          properties: {
            title,
          },
        };
        try {
            const spreadsheet = await service.spreadsheets.create({
                resource,
                fields: 'spreadsheetId',
            });
            console.log(`Spreadsheet ID: ${spreadsheet.data.spreadsheetId}`);
            // return spreadsheet.data.spreadsheetId;
        } catch (err) {
            console.log(err);
        }
    } catch (error) {
        return res.status(500).json({ message: error.message });
    }
})

app.get('/get-all-sheets', async (req, res) => {
    try {
      const sheets = google.sheets({ version: 'v4', auth: oauth2Client });
      const response = await sheets.spreadsheets.list();
      const spreadsheets = response.data.files;
  
      // Send the list of spreadsheets as a JSON response to the client
      res.json({ spreadsheets });
    } catch (error) {
      console.error('Error listing spreadsheets:', error);
      return res.status(500).json({ message: error.message });
    }
  });


const port = process.env.PORT;

app.listen(port, (req, res) => console.log('server running on port ', port));
