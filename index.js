const {
  existsSync,
  mkdirSync,
  readFile,
  readFileSync,
  writeFile,
} = require('fs')
const readline = require('readline')
require('colors')
const { google } = require('googleapis')
const { prompt } = require('inquirer')
const nodeoutlook = require('nodejs-nodemailer-outlook')
const pdf = require('pdf-creator-node')
const data = require('./data.json')

const SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
const TOKEN_PATH = 'token.json'

const options = {
  format: 'A4',
  orientation: 'portrait',
  border: '10mm',
}

const courseQuestion = {
  type: 'list',
  name: 'course',
  message: 'Choose one of the following courses:',
  choices: ['mobile', 'oosd'],
}

const processQuestion = {
  type: 'list',
  name: 'process',
  message: 'Choose one of the following processes:',
  choices: ['generate pdf', 'email pdf', 'merge pdf'],
}

const { oosd, mobile, email, password } = data

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
    authorize(JSON.parse(content), runProcess)
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

const runProcess = (auth) => {
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

      prompt(processQuestion).then((answer) => {
        const { process } = answer
        switch (process) {
          case 'generate pdf':
            generatePDF(studentData)
            break
          case 'email pdf':
            emailPDF(studentData)
            break
          case 'merge pdf':
            mergePDF()
            break
        }
      })
    }
  )
}

const generatePDF = (studentData) => {
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
}

const emailPDF = (studentData) => {
  let interval = 7500
  studentData.map((data, idx) => {
    const firstName = data.first_name.toLowerCase()
    const lastName = data.last_name.toLowerCase()
    const filename = `./${outputDirectory}-${firstName}-${lastName}-output.pdf`
    setTimeout((_) => {
      console.log(`Emailing PDF file to ${firstName} ${lastName}.`.green)
      nodeoutlook.sendEmail({
        auth: {
          user: email,
          pass: password,
        },
        from: email,
        to: `orrgl1@student.op.ac.nz`,
        subject: 'Results',
        html: `Kia ora, <br /> <br />
        I have attached your final results. Your results will be released officially on EBS in the next day or two. I would like to personally thank you for the semester. I have thoroughly enjoyed the experience and hope you have learned something during this time. Enjoy yours holidays and take care of yourself. <br /> <br />
        Ngā mihi nui, <br /> <br />
        Grayson Orr`,
        attachments: [
          {
            path: filename,
          },
        ],
        onError: (err) => console.log(err),
        onSuccess: (_) => {
          console.log(`PDF file emailed to ${firstName} ${lastName}.`.blue)
        },
      })
    }, idx * interval)
  })
  
}

const mergePDF = () => console.log('merge')
