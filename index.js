const {
  existsSync,
  mkdirSync,
  readFile,
  readFileSync,
  writeFile,
} = require('fs')
const readline = require('readline')
const { google } = require('googleapis')
const pdf = require('pdf-creator-node')
const { prompt } = require('inquirer')
require('colors')
const data = require('./data.json')

const SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
const TOKEN_PATH = 'token.json'

const options = {
  format: 'A4',
  orientation: 'portrait',
  border: '10mm',
}

const courseQuestion = [
  {
    type: 'list',
    name: 'course',
    message: 'Choose one of the following courses:',
    choices: ['mobile', 'oosd'],
  },
]

const { oosd, mobile } = data

let courseName = ''
let outputDirectory = ''
let spreadsheetId = ''

prompt(courseQuestion).then((answer) => {
  const { course } = answer
  if (!existsSync(course)) mkdirSync(course)
  switch (course) {
    case 'mobile':
      courseName =
        'IN721: Design and Development of Applications for Mobile Devices'
      outputDirectory = `${mobile.output_directory}/${course}`
      spreadsheetId = mobile.spreadsheet_id
      break
    case 'oosd':
      courseName = 'IN710: Object-Oriented Systems Development'
      outputDirectory = `${oosd.output_directory}/${course}`
      spreadsheetId = oosd.spreadsheet_id
      break
  }
  readFile('credentials.json', (err, content) => {
    if (err) return console.log(`Error loading client secret file: ${err}`)
    authorize(JSON.parse(content), generatePDF)
  })
})

const authorize = (credentials, callback) => {
  const { client_secret, client_id, redirect_uris } = credentials.installed
  const oAuth2Client = new google.auth.OAuth2(
    client_id,
    client_secret,
    redirect_uris[0]
  )

  readFile(TOKEN_PATH, (err, token) => {
    if (err) return getNewToken(oAuth2Client, callback)
    oAuth2Client.setCredentials(JSON.parse(token))
    callback(oAuth2Client)
  })
}

const getNewToken = (oAuth2Client, callback) => {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
  })
  console.log(`Authorize this app by visiting this url: ${authUrl}`)
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  })
  rl.question('Enter the code from that page here: ', (code) => {
    rl.close()
    oAuth2Client.getToken(code, (err, token) => {
      if (err)
        return console.error(
          `Error while trying to retrieve access token ${err}`
        )
      oAuth2Client.setCredentials(token)
      writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
        if (err) return console.error(err)
        console.log(`Token stored to ${TOKEN_PATH}`)
      })
      callback(oAuth2Client)
    })
  })
}

const generatePDF = (auth) => {
  const sheets = google.sheets({ version: 'v4', auth })
  sheets.spreadsheets.values.get(
    {
      spreadsheetId: spreadsheetId,
      range: 'overall!A2:I17',
    },
    (err, res) => {
      if (err) return console.log(`The API returned an error: ${err}`)
      const rows = res.data.values
      const studentData = []
      if (rows.length) {
        rows.map((row) => {
          studentData.push({
            first_name: row[0],
            last_name: row[1],
            email_address: row[2],
            exam_percentage: row[3],
            software_percentage: row[5],
            overall_percentage: row[7],
            overall_grade: row[8],
            course_name: courseName,
          })
        })
      } else {
        console.log('No data found.')
      }

      const html = readFileSync('./public/template.html', 'utf8')
      studentData.map((data) => {
        const firstName = data.first_name.toLowerCase()
        const lastName = data.last_name.toLowerCase()
        const filename = `./${outputDirectory}-${firstName}-${lastName}-output.pdf`
        const document = {
          html: html,
          data: {
            data: [data],
          },
          path: filename,
        }
        console.log(`Generating PDF file for ${firstName} ${lastName}.`.green)
        pdf.create(document, options)
        console.log(`PDF file generated for ${firstName} ${lastName}.`.blue)
      })
      console.log('Complete.'.green)
    }
  )
}
