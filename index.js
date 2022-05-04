const {
  access,
  mkdir,
  readFile,
  readFileSync,
  writeFile,
  writeFileSync,
} = require('fs')
const { join } = require('path')
const readline = require('readline')
require('colors')
const { google } = require('googleapis')
const { prompt } = require('inquirer')
const { sendEmail } = require('nodejs-nodemailer-outlook')
const PZ = require('pizzip')
const DocxTemp = require('docxtemplater')
const data = require('./data.json')

const SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
const TOKEN_PATH = 'token.json'

const courseQuestion = {
  type: 'list',
  name: 'course',
  message: 'Choose one of the following courses:',
  choices: ['id607001', 'id721001', 'id737001'],
}

const assessmentQuestion = {
  type: 'list',
  name: 'assessment',
  message: 'Choose one of the following assessments:',
  choices: ['a1', 'a2', 'a3', 'overall'],
}

const processQuestion = {
  type: 'list',
  name: 'process',
  message: 'Choose one of the following processes:',
  choices: ['generate word docx', 'email word docx'],
}

const {
  email,
  password
} = data

let courseCode
let assessmentNum
let courseName
let spreadsheetId
let spreadsheetRange
let assessmentName
let assessmentWordDocxName

prompt([courseQuestion, assessmentQuestion]).then((answer) => {
  const { course, assessment } = answer

  courseCode = course
  assessmentNum = assessment

  const { course_name, spreadsheet_id } = data[courseCode]
  const { assessment_name, spreadsheet_range, word_docx_name } = data[courseCode][assessmentNum]

  courseName = course_name
  spreadsheetId = spreadsheet_id
  spreadsheetRange = spreadsheet_range
  assessmentName = assessment_name
  assessmentWordDocxName = word_docx_name
    
  const courseDirectory = join(__dirname, courseCode, assessmentNum)

  access(courseDirectory, (error) => {
    if (error) {
      mkdir(courseDirectory, { recursive: true }, (error) => {
        if (error) return console.error(error)
        console.log(`${courseCode}/${assessmentNum} directories successfully created`.green)
      })
    }
  })

  readFile('credentials.json', (error, content) => {
    if (error) return console.log(`Error loading client secret file: ${error}`)
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

  readFile(TOKEN_PATH, (error, token) => {
    if (error) return getNewToken(oAuth2Client, callback)
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
    oAuth2Client.getToken(code, (error, token) => {
      if (error)
        return console.error(
          `Error while trying to retrieve access token ${error}`
        )
      oAuth2Client.setCredentials(token)
      writeFile(TOKEN_PATH, JSON.stringify(token), (error) => {
        if (error) return console.error(error)
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
      range: spreadsheetRange,
    },
    (error, res) => {
      if (error) return console.log(`The API returned an error: ${error}`)
      const rows = res.data.values
      const studentData = []
      const today = new Date()
      const date = `${today.getDate()}/${today.getMonth() + 1}/${today.getFullYear()}`
      if (rows.length) {
        rows.map((row) => {
          let obj = {
            date: date,
            first_name: row[0],
            last_name: row[1],
            learner_id: row[2],
            email_address: row[3],
            points: row[4],
            percentage: row[5],
            grade: row[6],
            crit_one: row[7],
            crit_two: row[8]
          }

          if (courseCode === 'id607001') {
            obj['crit_three'] = row[9]
            obj['comment_one'] = row[10]
            obj['comment_two'] = row[11]
            obj['comment_three'] = row[12]
            if (assessmentNum === 'a1' || assessmentNum === 'a2') {
              obj['crit_one_score'] = ((row[7] * 0.40) * 10).toFixed(2)
              obj['crit_two_score'] = ((row[8] * 0.45) * 10).toFixed(2)
              obj['crit_three_score'] = ((row[9] * 0.15) * 10).toFixed(2)
            } else {
              obj['crit_one_score'] = ((row[7] * 0.60) * 10).toFixed(2)
              obj['crit_two_score'] = ((row[8] * 0.30) * 10).toFixed(2)
              obj['crit_three_score'] = ((row[9] * 0.10) * 10).toFixed(2)
            }
          }

          if (courseCode === 'id737001') {
            if (assessmentNum === 'a1') {
              obj['crit_one_score'] = ((row[7] * 0.90) * 10).toFixed(2)
              obj['crit_two_score'] = ((row[8] * 0.10) * 10).toFixed(2)
            } else {
              obj['crit_one_score'] = ((row[7] * 0.80) * 10).toFixed(2)
              obj['crit_two_score'] = ((row[8] * 0.20) * 10).toFixed(2)
            }
          }

          studentData.push(obj)
        })
      } else {
        console.log('No data found.')
      }

      prompt(processQuestion).then((answer) => {
        const { process } = answer
        switch (process) {
          case 'generate word docx':
            generateWordDocx(studentData)
            break
          case 'email word docx':
            emailWordDocx(studentData)
            break
        }
      })
    }
  )
}

const generateWordDocx = (studentData) => {
  const content = readFileSync(
    join(__dirname, 'assessments', assessmentWordDocxName),
    'binary'
  )
  
  const zip = new PZ(content)
  const doc = new DocxTemp(zip)
  doc.setData(studentData)
  studentData.map((data) => {
    const firstName = data.first_name.toLowerCase()
    const lastName = data.last_name.toLowerCase()
    doc.setData(data)
    doc.render()
    const buffer = doc.getZip().generate({ type: 'nodebuffer' })
    console.log(`Generating file for ${firstName} ${lastName}.`.green)
    writeFileSync(
      join(__dirname, courseCode, assessmentNum, `${firstName}-${lastName}.docx`),
      buffer
    )
    console.log(`File generated for ${firstName} ${lastName}.`.blue)
  })
}

const emailWordDocx = (studentData) => {
  let interval = 7500
  studentData.map((data, idx) => {
    const firstName = data.first_name.toLowerCase()
    const lastName = data.last_name.toLowerCase()
    
    setTimeout((_) => {
      console.log(`Emailing document file to ${firstName} ${lastName}.`.green)
      sendEmail({
        auth: {
          user: email,
          pass: password,
        },
        from: email,
        to: data.email_address.toLowerCase(),
        subject: `${courseName} - ${assessmentName} Results`,
        html: `Kia ora ${data.first_name}, <br /> <br />
        I have attached your ${assessmentName} assessment result. If there are any issues, please do not hesitate to ask.<br /> <br />
        Ngā mihi nui, <br /> <br />
        Grayson Orr`,
        attachments: [
          {
            path: join(__dirname, courseCode, assessmentNum, `${firstName}-${lastName}.docx`),
          },
        ],
        onError: (error) => console.log(error),
        onSuccess: (_) => {
          console.log(`File emailed to ${firstName} ${lastName}.`.blue)
        },
      })
    }, idx * interval)
  })
}
