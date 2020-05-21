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
const { sendEmail } = require('nodejs-nodemailer-outlook')
const pdfMerge = require('pdfmerge')
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
  choices: ['prog-four', 'mobile', 'oosd'],
}

const processQuestion = {
  type: 'list',
  name: 'process',
  message: 'Choose one of the following processes:',
  choices: ['generate pdf', 'email pdf', 'merge pdf'],
}

const { prog_four, oosd, mobile, email, password } = data

let courseName = ''
let outputDirectory = ''
let spreadsheetId = ''
let template = ''

prompt(courseQuestion).then((answer) => {
  const { course } = answer
  if (!existsSync(course)) mkdirSync(course)
  switch (course) {
    case 'prog-four':
      courseName = prog_four.name
      outputDirectory = `${prog_four.output_directory}/${course}`
      spreadsheetId = prog_four.spreadsheet_id
      range = prog_four.range
      template = 'second-year'
      break
    case 'mobile':
      courseName = mobile.name
      outputDirectory = `${mobile.output_directory}/${course}`
      spreadsheetId = mobile.spreadsheet_id
      range = mobile.range
      template = 'third-year'
      break
    case 'oosd':
      courseName = oosd.name
      outputDirectory = `${oosd.output_directory}/${course}`
      spreadsheetId = oosd.spreadsheet_id
      range = oosd.range
      template = 'third-year'
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
      range: range,
    },
    (err, res) => {
      if (err) return console.log(`The API returned an error: ${err}`)
      const rows = res.data.values
      const studentData = []
      if (rows.length) {
        rows.map((row) => {
          let obj = {
            course_name: courseName,
            first_name: row[0],
            last_name: row[1],
            email_address: row[2],
            overall_percentage: row[3],
            overall_grade: row[4],
            exam_percentage: row[5],
            software_percentage: row[7],
          }
          if (courseName == 'IN628 Programming 4') {
            obj.checkpoint_percentage = row[9]
            studentData.push(obj)
          } else {
            studentData.push(obj)
          }
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
            mergePDF(studentData)
            break
        }
      })
    }
  )
}

const generatePDF = (studentData) => {
  const html = readFileSync(`./public/${template}.html`, 'utf8')
  studentData.map((data) => {
    const firstName = data.first_name.toLowerCase()
    const lastName = data.last_name.toLowerCase()
    const filename = `./${outputDirectory}-${firstName}-${lastName}-results.pdf`
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
    const filename = `./${outputDirectory}-${firstName}-${lastName}-results.pdf`
    setTimeout((_) => {
      console.log(`Emailing PDF file to ${firstName} ${lastName}.`.green)
      sendEmail({
        auth: {
          user: email,
          pass: password,
        },
        from: email,
        to: data.email_address.toLowerCase(),
        subject: 'Results',
        html: `Kia ora, <br /> <br />
        I have attached your course results for ${courseName}. <br /> <br />
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

const mergePDF = (studentData) =>  {
    let interval = 7500
    studentData.map((data, idx) => {
      const firstName = data.first_name.toLowerCase()
      const lastName = data.last_name.toLowerCase()
      setTimeout((_) => {
        console.log(`Merging PDF file for ${firstName} ${lastName}.`.green)
        pdfMerge(
          [
            `./${outputDirectory}-${firstName}-${lastName}-results.pdf`,
            `./prog-four/01-assessment-${firstName}-${lastName}.pdf`,
          ],
          `./${outputDirectory}-${firstName}-${lastName}-final.pdf`
        )
          .then((_) =>
            console.log(`PDF files merged for ${firstName} ${lastName}.`.blue)
          )
          .catch((err) => console.log(err))
      }, idx * interval)
    })
}