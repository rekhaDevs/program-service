const { getQuestionSet } = require("./printDocxV1.0/dataImporter");
const JSON2CSV = require("json2csv").parse;
var cheerio = require("cheerio");
var cheerioTableparser = require("cheerio-tableparser");

const buildCSVWithCallback = async (id, callback) => {
  let error = false;
  let errorMsg = "";
  let totalMarks = 0;
  const config = {
   id: id    
  }
  getQuestionSet(config)
    .then(async (data) => {
      if (data.error) {
        callback(null, data.error, data.errorMsg, null);
      } else {
        let subject;
        let grade;
        if (data.paperData.subject && data.paperData.gradeLevel) {
          subject = data.paperData.subject[0];
          grade = data.paperData.gradeLevel[0];
        }

        const examName = data.paperData.name;
        data.sectionData.forEach((d) => {
          d.questions.forEach((element, index) => {
            const marks = parseInt(d.section.children[index].marks);
            if (!isNaN(marks)) totalMarks += marks;
          });
        });

        const questionPaperContent = [];
        let questionCounter = 0;
        for (const d of data.sectionData) {
          for (const [index, question] of d.questions.entries()) {
            questionCounter += 1;
            let questionContent;
            let blooms;
            let learningOutcome;
            let chaperName;

            if (question.qType === "MCQ") {
              if (question.learningOutcome && question.learningOutcome[0]) {
                learningOutcome = question.learningOutcome[0];
              } else {
                learningOutcome = "";
              }

              if (question.bloomsLevel) {
                blooms = question.bloomsLevel;
              } else {
                blooms = "";
              }

              if (question.topic && question.topic[0]) {
                chaperName = question.topic[0];
              } else {
                chaperName = "";
              }
              questionContent = await renderMCQ(
                question,
                questionCounter,
                grade,
                subject,
                examName,
                learningOutcome,
                blooms,
                chaperName,
                question.qType
              );
              questionPaperContent.push(questionContent);
            } else if (
              question.qType === "VSA" ||
              question.qType === "SA" ||
              question.qType === "LA" ||
              question.qType === "FTB" ||
              question.qType === "COMPREHENSION" ||
              question.qType === "CuriosityQuestion"
            ) {
              if (question.learningOutcome && question.learningOutcome[0]) {
                learningOutcome = question.learningOutcome[0];
              } else {
                learningOutcome = "";
              }
              if (question.bloomsLevel) {
                blooms = question.bloomsLevel;
              } else {
                blooms = "";
              }

              if (question.topic && question.topic[0]) {
                chaperName = question.topic[0];
              } else {
                chaperName = "";
              }
              questionContent = await renderQuestion(
                question,
                questionCounter,
                grade,
                subject,
                examName,
                learningOutcome,
                blooms,
                chaperName,
                question.qType
              );
              questionPaperContent.push(questionContent);
            }
          }
        }

        let fields = [
          "Class",
          "Subject",
          "QuestionSetName",
          "Questions",
          "Option1",
          "Option2",
          "Option3",
          "Option4",
          "CorrectAnswer(1/2/3/4)",
          "Competencies",
          "Skills",
          "QuestionImageUrl",
          "ChapterName",
          "QuestionType",
          "RightColumn",
          "LeftColumn",
        ];

        let csv = JSON2CSV(questionPaperContent, {
          fields: fields,
          withBOM: true,
        });
        let filename = grade + "_" + subject + "_" + examName;
        filename = filename.replace(/\s/g, "");
        callback(csv, error, errorMsg, filename);
      }
    })
    .catch((e) => {
      error = true;
      errorMsg = "";
      callback(null, error, errorMsg, null);
    });
};

const cleanHTML = (str, nbspAsLineBreak = false) => {
  // Remove HTML characters since we are not converting HTML to PDF.
  return str
    .replace(/<[^>]+>/g, "")
    .replace(/&nbsp;/g, nbspAsLineBreak ? "\n" : "");
};

function extractTextFromElement(elem) {
  let rollUp = "";
  if (cheerio.text(elem)) return cheerio.text(elem);
  else if (elem.name === "sup")
    return { text: elem.children[0].data, sup: true };
  else if (elem.name === "sub")
    return { text: elem.children[0].data, sub: true };
  else if (elem.type === "text" && elem.data) return elem.data;
  else {
    if (elem.children && elem.children.length) {
      for (const nestedElem of elem.children) {
        let recurse = extractTextFromElement(nestedElem);
        if (Array.isArray(rollUp)) {
          rollUp.push(recurse);
        } else {
          if (Array.isArray(recurse)) {
            rollUp = recurse;
          } else if (typeof recurse === "object") {
            rollUp = [rollUp, recurse];
          } else rollUp += recurse;
        }
      }
    }
  }
  return rollUp;
}

async function getStack(htmlString, questionCounter) {
  const stack = [];
  let count = 0;
  let p = 0;
  $ = cheerio.load(htmlString);
  const elems = $("body").children().toArray();
  for (const [index, elem] of elems.entries()) {
    let nextLine = "";
    switch (elem.name) {
      case "p":
        let extractedText = extractTextFromElement(elem);
        if (Array.isArray(extractedText)) {
          nextLine = { text: extractedText };
        } else {
          nextLine += extractedText;
        }
        nextLine = { text: nextLine };
        break;
      case "ol":
        nextLine = {
          ol: elem.children.map(
            (el) =>
              el.children[0] &&
              (el.children[0].data ||
                (el.children[0].children[0] && el.children[0].children[0].data))
          ),
        };
        break;
      case "ul":
        nextLine = {
          ul: elem.children.map(
            (el) =>
              el.children[0] &&
              (el.children[0].data ||
                (el.children[0].children[0] && el.children[0].children[0].data))
          ),
        };
        break;
      case "figure":
        if (count === 0) {
          if (elem.children && elem.children.length) {
            let { src } = elem.children[0].attribs;
            if (src !== undefined)
              if (src.startsWith("data:image/png")) {
                nextLine = "";
              } else if (src.startsWith("data:image/jpeg")) {
                nextLine = "";
              } else {
                if(src.includes("https:")){
                  nextLine = src
                } else{
                  nextLine = `${envVariables.baseURL}` + src;
                }
                count++;
  
              }
          }
          if (!nextLine)
            nextLine = "<An image of an unsupported format was scrubbed>";
        }
        break;
    }
    if (index === 0 && questionCounter) {
      if (elem.name === "p") {
        if (typeof nextLine === "object") nextLine = `${nextLine.text}`;
        else nextLine = `${nextLine}`;
      } else stack.push(`${nextLine}`);
    }
    stack.push(nextLine);
  }
  return stack;
}

const ansIndex = (arr) => {
  for (let index = 0; index < arr.length; index++) {
    const element = arr[index];
    if (element.answer === true) return index+1
  }
}
async function renderMCQ(
  question,
  questionCounter,
  grade,
  subject,
  examName,
  learningOutcome,
  blooms,
  topic,
  qType
) {
  const questionOptions = [],
    answerOptions = ["A", "B", "C", "D"];
  let questionTitle;
  let finalQuestion = "";
  if (question.editorState.options) {
    for (const [index, qo] of question.editorState.options.entries()) {
      let qoBody = qo.value.body;
      let qoData =
        qoBody.search("img") >= 0 ||
        qoBody.search("sup") >= 0 ||
        qoBody.search("sub") >= 0 ||
        (qoBody.match(/<p>/g) && qoBody.match(/<p>/g).length > 1) ||
        (qoBody.match(/<ol>/g) && qoBody.match(/<ol>/g).length >= 1)
          ? await getStack(qoBody, answerOptions[index])
          : [`${cleanHTML(qoBody)}`];
      questionOptions.push(qoData);
    }
  }
  let q = question.editorState.question;

  questionTitle =
    q.search("img") >= 0 ||
    q.search("sub") >= 0 ||
    q.search("sup") >= 0 ||
    (q.match(/<p>/g) && q.match(/<p>/g).length > 1) ||
    (q.match(/<ol>/g) && q.match(/<ol>/g).length > 1)
      ? await getStack(q, questionCounter)
      : [`${cleanHTML(q)}`];

  let questionOpt = [];

  if (questionOptions[0] !== undefined) {
    if (
      questionOptions[0][0] !== undefined &&
      typeof questionOptions[0][0] === "object"
    ) {
      if (questionOptions[0][0].text) {
        questionOpt.push(["I.", questionOptions[0][0].text[1]]);
      } else {
        questionOpt.push(["I.", questionOptions[0][0].image]);
      }
    } else {
      questionOpt.push(["I.", questionOptions[0][0]]);
    }
  }

  if (questionOptions[1] !== undefined) {
    if (
      questionOptions[1][0] !== undefined &&
      typeof questionOptions[1][0] === "object"
    ) {
      if (questionOptions[1][0].text) {
        questionOpt.push(["II.", questionOptions[1][0].text[1]]);
      } else {
        questionOpt.push(["II.", questionOptions[1][0].image]);
      }
    } else {
      questionOpt.push(["II.", questionOptions[1][0]]);
    }
  }

  if (questionOptions[2] !== undefined) {
    if (
      questionOptions[2][0] !== undefined &&
      typeof questionOptions[2][0] === "object"
    ) {
      if (questionOptions[2][0].text) {
        questionOpt.push(["III.", questionOptions[2][0].text[1]]);
      } else {
        questionOpt.push(["III.", questionOptions[2][0].image]);
      }
    } else {
      questionOpt.push(["III.", questionOptions[2][0]]);
    }
  }

  if (questionOptions[3] !== undefined) {
    if (
      questionOptions[3][0] !== undefined &&
      typeof questionOptions[3][0] === "object"
    ) {
      if (questionOptions[3][0].text) {
        questionOpt.push(["IV.", questionOptions[3][0].text[1]]);
      } else {
        questionOpt.push(["IV.", questionOptions[3][0].image]);
      }
    } else {
      questionOpt.push(["IV.", questionOptions[3][0]]);
    }
  }
  let imageurl = envVariables.QUE_IMG_URL;
  let queurl = "";

  for (let que of questionTitle) {
    if (typeof que === "object") {
      finalQuestion += que.text;
    } else {
      if (que.includes(imageurl)) {
        queurl = que;
      } else if (
        que.match("<An image of an unsupported format was scrubbed>")
      ) {
        queurl = que;
      } else {
        finalQuestion = que;
      }
    }
  }
  let answer = await ansIndex(question.editorState.options)
  let data = {
    Class: grade,
    Subject: subject,
    QuestionSetName: examName,
    Questions: finalQuestion,
    "CorrectAnswer(1/2/3/4)": answer,
    Competencies: learningOutcome,
    Skills: blooms,
    QuestionImageUrl: queurl,
    ChapterName: topic,
    QuestionType: "Multiple Choice Question",
    RightColumn: "",
    LeftColumn: "",
  };
  questionOptions.forEach(( quesOpt , i)=>{
    data[`Option${i+1}`] =   quesOpt[0]
  })
  return data;
}

async function renderQuestion(
  question,
  questionCounter,
  grade,
  subject,
  examName,
  learningOutcome,
  blooms,
  topic,
  qType
) {
  let questionTitle;
  let finalQuestion = "";

  $ = cheerio.load(question.editorState.question);
  cheerioTableparser($);
  var columns = $("table").parsetable(false, false, false);
  if (columns.length !== 0) {
    return await renderMTF(
      question,
      grade,
      subject,
      examName,
      learningOutcome,
      blooms,
      topic,
      "Match The Following"
    );
  }

  let q = question.editorState.question;
  questionTitle =
    q.search("img") >= 0 ||
    q.search("sub") >= 0 ||
    q.search("sup") >= 0 ||
    (q.match(/<p>/g) && q.match(/<p>/g).length > 1) ||
    (q.match(/<ol>/g) && q.match(/<ol>/g).length > 1)
      ? await getStack(q, questionCounter)
      : [`${cleanHTML(q)}`];

  let imageurl = envVariables.QUE_IMG_URL;
  let queurl = "";

  for (let que of questionTitle) {
    if (typeof que === "object") {
      finalQuestion += que.text;
    } else {
      if (que.includes(envVariables.baseURL)) {
        queurl = que;
      } else if (
        que.match("<An image of an unsupported format was scrubbed>")
      ) {
        queurl = que;
      } else {
        finalQuestion = que;
      }
    }
  }

  let questionType = "";

  if (qType === "VSA") questionType = "Very Short Answer";
  if (qType === "SA") questionType = "Short Answer";
  if (qType === "LA") questionType = "Long Answer";
  if (qType === "FTB") questionType = "Fill in The Blank";
  if (qType === "COMPREHENSION") questionType = "COMPREHENSION";
  if (qType === "CuriosityQuestion") questionType = "CuriosityQuestion";

  let data = {
    Class: grade,
    Subject: subject,
    QuestionSetName: examName,
    Questions: finalQuestion,
    Option1: "",
    Option2: "",
    Option3: "",
    Option4: "",
    "CorrectAnswer(1/2/3/4)": "",
    Competencies: learningOutcome,
    Skills: blooms,
    QuestionImageUrl: queurl,
    ChapterName: topic,
    QuestionType: questionType,
    RightColumn: "",
    LeftColumn: "",
  };
  return data;
}

async function renderMTF(
  question,
  grade,
  subject,
  examName,
  learningOutcome,
  blooms,
  topic,
  qType
) {
  $ = cheerio.load(question.editorState.question);

  cheerioTableparser($);
  var columns = $("table").parsetable(false, false, false);
  let transposeColumns = columns[0].map((_, colIndex) =>
    columns.map((row) => row[colIndex])
  );

  const heading = cleanHTML($("p").first().text());
  transposeColumns.shift();

  let left = [],
    right = [];
  for (const r of transposeColumns) {
    if (r[0].search("img") >= 0) {
      let leftRes = await getStack(r[0]);
      left.push(leftRes[0]);
    } else left.push(cleanHTML(r[0]));
    if (r[1].search("img") >= 0) {
      let rightRes = await getStack(r[1]);
      right.push(rightRes[0]);
    } else right.push(cleanHTML(r[1]));
  }
  let data = {
    Class: grade,
    Subject: subject,
    QuestionSetName: examName,
    Questions: heading,
    Option1: "",
    Option2: "",
    Option3: "",
    Option4: "",
    "CorrectAnswer(1/2/3/4)": "",
    Competencies: learningOutcome,
    Skills: blooms,
    QuestionImageUrl: "",
    ChapterName: topic,
    QuestionType: qType,
    RightColumn: right,
    LeftColumn: left,
  };
  return data;
}

module.exports = {
  buildCSVWithCallback,
};
