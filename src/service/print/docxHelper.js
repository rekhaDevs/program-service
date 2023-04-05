var cheerio = require("cheerio");
var cheerioTableparser = require("cheerio-tableparser");
const sizeOf = require("image-size");
const ProgramServiceHelper = require("../../helpers/programHelper");
const programServiceHelper = new ProgramServiceHelper();
const optionLabelling = {
  english: {i: 'I', ii: 'II', iii: 'III', iv: 'IV'}, 
  hindi: {i: 'क', ii: 'ख', iii: 'ग', iv: 'घ'}, 
}
const defaultLanguage = 'english'

var {
  docDefinition,
  getSectionTitle,
  getTF,
  getMTFHeader,
} = require("./utils/docDefinition");

const cleanHTML = (str, nbspAsLineBreak = false) => {
  // Remove HTML characters since we are not converting HTML to Docx.
  if (str === undefined) return "";
  else
    return str
      .replace(/<[^>]+>/g, "")
      .replace(/&nbsp;/g, nbspAsLineBreak ? "\n" : "");
};

const detectLanguage = (str) => {
  const unicodeBlocks = [
    {
      name: "Tamil",
      regex: /[\u0B80-\u0BFF]+/g,
    },
    {
      name: "Hindi",
      regex: /[\u0900-\u097F]+/g,
    },
  ];

  let language = "English";

  const langSplit = {
    Hindi: 0,
    Tamil: 0,
    English: 0,
    Undefined: 0,
  };
  if (typeof str === "string") {
    str.split("").forEach((letter) => {
      let found = false;
      unicodeBlocks.forEach((block) => {
        if (letter.match(block.regex)) {
          langSplit[block.name]++;
          found = true;
        }
      });
      if (!found) {
        langSplit.English++;
      }
    });

    let max = 0;
    for (var key of Object.keys(langSplit)) {
      if (langSplit[key] > max) {
        max = langSplit[key];
        language = key;
      }
    }

    return language;
  }
  return "English";
};

function createImageElement(src, width) {
  let imageElement = {};
  if (src.search("image/gif") >= 0) return null;
  imageElement.image = src;
  let img = Buffer.from(src.split(";base64,").pop(), "base64");
  let dimensions = sizeOf(img);
  let resizedWidth = dimensions.width * width;
  imageElement.width = resizedWidth > 200 ? 200 : resizedWidth;
  imageElement.height =
    (dimensions.height / dimensions.width) * imageElement.width;
  return imageElement;
}

function extractTextFromElement(elem) {
  let rollUp = "";
  if (cheerio.text(elem)) return cheerio.text(elem);
  else if (elem.name === "sup")
    return { text: elem.children[0].data, superScript: true };
  else if (elem.name === "sub")
    return { text: elem.children[0].data, subScript: true };
  else if (elem.name === "strong") {
    if (elem.children[0].data === undefined) {
      return getStyleEle(elem);
    } else {
      return { text: elem.children[0].data, bold: true };
    }
  } else if (elem.name === "span") {
    if (elem.children[0].data === undefined) {
      return getStyleEle(elem);
    } else {
      return elem.children[0].data;
    }
  } else if (elem.name === "i") {
    if (elem.children[0].data === undefined) {
      return getStyleEle(elem);
    } else {
      return { text: elem.children[0].data, italics: true };
    }
  } else if (elem.name === "br") return { br: "break" };
  else if (elem.name === "u") {
    if (elem.children[0].data === undefined) {
      return getStyleEle(elem);
    } else {
      return { text: elem.children[0].data, underline: true };
    }
  } else if (elem.type === "text" && elem.data) return elem.data;
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

function getStyleEle(el) {
  let value = "";
  if (
    el.children[0].name &&
    (el.children[0].name === "i" ||
      el.children[0].name === "strong" ||
      el.children[0].name === "u")
  ) {
    return getStyleEle(el.children[0]);
  } else {
    if (el.children[0].data !== undefined) {
      if (el.name && el.name === "u") {
        return {
          text:
            el.children[0] &&
            (el.children[0].data ||
              (el.children[0].children[0] && el.children[0].children[0].data)),
          underline: true,
        };
      } else if (el.name && el.name === "i") {
        return {
          text:
            el.children[0] &&
            (el.children[0].data ||
              (el.children[0].children[0] && el.children[0].children[0].data)),
          italics: true,
        };
      } else if (el.name && (el.name === "b" || el.name === "strong")) {
        return {
          text:
            el.children[0] &&
            (el.children[0].data ||
              (el.children[0].children[0] && el.children[0].children[0].data)),
          bold: true,
        };
      } else {
        return (
          el.children[0] &&
          (el.children[0].data ||
            (el.children[0].children[0] && el.children[0].children[0].data))
        );
      }
    }
  }
}

async function getStack(htmlString) {
  const stack = [];
  $ = cheerio.load(htmlString);
  const elems = $("body").children().toArray();
  for (const [index, elem] of elems.entries()) {
    let nextLine = "";
    switch (elem.name) {
      case "p":
        let extractedText = extractTextFromElement(elem);
        // Returns array if superscript/subscript inside
        if (Array.isArray(extractedText)) nextLine = { text: extractedText };
        else nextLine += extractedText;
        break;
      case "ol":
        nextLine = {
          ol: elem.children.map((el) => {
            return getStyleEle(el);
          }),
        };
        break;

      case "ul":
        nextLine = {
          ul: elem.children.map((el) => {
            return getStyleEle(el);
          }),
        };
        break;
      case "figure":
        let { style } = elem.attribs;
        let width = 1;
        if (style) {
          width = parseFloat(style.split(":").pop().slice(0, -2));
          width = width / 100;
        }
        if (elem.children && elem.children.length) {
          let { src } = elem.children[0].attribs;
          if (src) {
            switch (src.slice(0, 4)) {
              case "data":
                nextLine = createImageElement(src, width);
                break;
              default:
                let res = await programServiceHelper.getQuestionMedia(src);
                nextLine = createImageElement(res, width);
            }
          }
        }
        if (!nextLine)
          nextLine = "<An image of an unsupported format was scrubbed>";
        break;
    }
    stack.push(nextLine);
  }
  return stack;
}

async function renderMCQ(question, questionCounter, marks) {
  const printLanguage = (question.medium && question.medium.length) ? question.medium[0].toLowerCase() : defaultLanguage;
  const questionOptions = [];
  let questionTitle;
  for (const [index, qo] of question.editorState.options.entries()) {
    let qoBody = qo.value.body;
    let qoData =
      qoBody.search("img") >= 0 ||
      qoBody.search("sup") >= 0 ||
      qoBody.search("sub") >= 0 ||
      (qoBody.match(/<p>/g) && qoBody.match(/<p>/g).length >= 1) ||
      (qoBody.match(/<ol>/g) && qoBody.match(/<ol>/g).length >= 1)
        ? await getStack(qoBody)
        : [`${cleanHTML(qoBody)}`];
    questionOptions.push(qoData);
  }
  let q = question.editorState.question;

  questionTitle =
    q.search("img") >= 0 ||
    q.search("sub") >= 0 ||
    q.search("sup") >= 0 ||
    (q.match(/<p>/g) && q.match(/<p>/g).length >= 1) ||
    (q.match(/<ol>/g) && q.match(/<ol>/g).length >= 1)
      ? await getStack(q, questionCounter)
      : [`${questionCounter}. ${cleanHTML(q)}`];

  let questionOpt = [];
  let imageProperties = [];
  if (questionOptions[0] !== undefined) {
    if (
      questionOptions[0][0] !== undefined &&
      typeof questionOptions[0][0] === "object"
    ) {
      if (questionOptions[0][0].text) {
        questionOpt.push([`${getLanguageKey(printLanguage, 'i')+")"}`, questionOptions[0][0].text[1]]);
        imageProperties.push({
          width: 0,
          height: 0,
        });
      } else {
        questionOpt.push([`${getLanguageKey(printLanguage, 'i')+")"}`, questionOptions[0][0].image]);
        imageProperties.push({
          width: questionOptions[0][0].width,
          height: questionOptions[0][0].height,
        });
      }
    } else {
      questionOpt.push([`${getLanguageKey(printLanguage, 'i')+")"}`, questionOptions[0][0]]);
      imageProperties.push({
        width: 0,
        height: 0,
      });
    }
  }

  if (questionOptions[1] !== undefined) {
    if (
      questionOptions[1][0] !== undefined &&
      typeof questionOptions[1][0] === "object"
    ) {
      if (questionOptions[1][0].text) {
        questionOpt.push([`${getLanguageKey(printLanguage, 'ii')+")"}`, questionOptions[1][0].text[1]]);
        imageProperties.push({
          width: 0,
          height: 0,
        });
      } else {
        questionOpt.push([`${getLanguageKey(printLanguage, 'ii')+")"}`, questionOptions[1][0].image]);
        imageProperties.push({
          width: questionOptions[1][0].width,
          height: questionOptions[1][0].height,
        });
      }
    } else {
      questionOpt.push([`${getLanguageKey(printLanguage, 'ii')+")"}`, questionOptions[1][0]]);
      imageProperties.push({
        width: 0,
        height: 0,
      });
    }
  }

  if (questionOptions[2] !== undefined) {
    if (
      questionOptions[2][0] !== undefined &&
      typeof questionOptions[2][0] === "object"
    ) {
      if (questionOptions[2][0].text) {
        questionOpt.push([`${getLanguageKey(printLanguage, 'iii')+")"}`, questionOptions[2][0].text[1]]);
        imageProperties.push({
          width: 0,
          height: 0,
        });
      } else {
        questionOpt.push([`${getLanguageKey(printLanguage, 'iii')+")"}`, questionOptions[2][0].image]);
        imageProperties.push({
          width: questionOptions[2][0].width,
          height: questionOptions[2][0].height,
        });
      }
    } else {
      questionOpt.push([`${getLanguageKey(printLanguage, 'iii')+")"}`, questionOptions[2][0]]);
      imageProperties.push({
        width: 0,
        height: 0,
      });
    }
  }

  if (questionOptions[3] !== undefined) {
    if (
      questionOptions[3][0] !== undefined &&
      typeof questionOptions[3][0] === "object"
    ) {
      if (questionOptions[3][0].text) {
        questionOpt.push([`${getLanguageKey(printLanguage, 'iv')+")"}`, questionOptions[3][0].text[1]]);
        imageProperties.push({
          width: 0,
          height: 0,
        });
      } else {
        questionOpt.push([`${getLanguageKey(printLanguage, 'iv')+")"}`, questionOptions[3][0].image]);
        imageProperties.push({
          width: questionOptions[3][0].width,
          height: questionOptions[3][0].height,
        });
      }
    } else {
      questionOpt.push([`${getLanguageKey(printLanguage, 'iv')+")"}`, questionOptions[3][0]]);
      imageProperties.push({
        width: 0,
        height: 0,
      });
    }
  }
  let data = {
    QuestionIndex: questionCounter,
    Questions: questionTitle,
    Option1: questionOpt[0],
    Option2: questionOpt[1],
    Option3: questionOpt[2],
    Option4: questionOpt[3],
    Marks: marks,
    Language: detectLanguage(questionTitle[0]),
    type: "MCQ",
    height1: imageProperties[0] ? imageProperties[0].height : undefined,
    width1: imageProperties[0] ? imageProperties[0].width : undefined,
    height2: imageProperties[1] ? imageProperties[1].height : undefined,
    width2: imageProperties[1] ? imageProperties[1].width : undefined,
    height3: imageProperties[2] ? imageProperties[2].height : undefined,
    width3: imageProperties[2] ? imageProperties[2].width : undefined,
    height4: imageProperties[3] ? imageProperties[3].height : undefined,
    width4: imageProperties[3] ? imageProperties[3].width : undefined,
  };

  return data;
}

async function renderQuestion(question, questionCounter, marks, Type) {
  let data;
  $ = cheerio.load(question.editorState.question);
  cheerioTableparser($);
  var columns = $("table").parsetable(false, false, false);
  if (columns.length !== 0) {
    return await renderMTF(question, questionCounter, marks, "MTF");
  }
  if (
    (question.media && question.media.length) ||
    question.editorState.question.search("img") >= 0 ||
    question.editorState.question.search("sub") >= 0 ||
    question.editorState.question.search("sup") >= 0 ||
    question.editorState.question.search("ul") >= 0 ||
    (question.editorState.question.match(/<p>/g) &&
      question.editorState.question.match(/<p>/g).length >= 1) ||
    (question.editorState.question.match(/<ol>/g) &&
      question.editorState.question.match(/<ol>/g).length >= 1)
  ) {
    data = await getStack(question.editorState.question, questionCounter);
  } else {
    data = [`${questionCounter}. ${cleanHTML(question.editorState.question)}`];
  }

  let quedata = {
    QuestionIndex: questionCounter,
    Questions: data,
    Marks: marks,
    type: Type,
  };
  return quedata;
}

async function renderComprehension(question, questionCounter, marks, Type) {
  let data;
  if (
    (question.media && question.media.length) ||
    question.editorState.question.search("img") >= 0 ||
    question.editorState.question.search("sub") >= 0 ||
    question.editorState.question.search("sup") >= 0 ||
    question.editorState.question.search("ol") >= 0 ||
    question.editorState.question.search("ul") >= 0 ||
    (question.editorState.question.match(/<p>/g) &&
      question.editorState.question.match(/<p>/g).length >= 1) ||
    (question.editorState.question.match(/<ol>/g) &&
      question.editorState.question.match(/<ol>/g).length >= 1)
  ) {
    data = await getStack(question.editorState.question, questionCounter);
  } else {
    data = [`${questionCounter}. ${cleanHTML(question.editorState.question)}`];
  }
  let quedata = {
    QuestionIndex: questionCounter,
    Questions: data,
    Marks: marks,
    type: Type,
  };
  return quedata;
}

async function renderMTF(question, questionCounter, marks, Type) {
  $ = cheerio.load(question.editorState.question);

  cheerioTableparser($);
  var data = [];
  var columns = $("table").parsetable(false, false, false);
  let transposeColumns = columns[0].map((_, colIndex) =>
    columns.map((row) => row[colIndex])
  );

  const heading = cleanHTML($("p").first().text());
  data.push(heading, detectLanguage(heading), marks);
  data.push(
    getMTFHeader(
      cleanHTML(transposeColumns[0][0]),
      cleanHTML(transposeColumns[0][1]),
      detectLanguage(cleanHTML(transposeColumns[0][0]))
    )
  );

  transposeColumns.shift();

  const rows = [];
  for (const r of transposeColumns) {
    let left, right;
    if (r[0].search("img") >= 0) {
      left = await getStack(r[0]);
    } else left = [cleanHTML(r[0])];
    if (r[1].search("img") >= 0) {
      right = await getStack(r[1]);
    } else right = [cleanHTML(r[1])];
    rows.push({ left, right });
  }

  let mtfData = {
    QuestionIndex: questionCounter,
    Questions: rows,
    Marks: marks,
    type: Type,
    heading: heading,
  };

  return mtfData;
}
function getLanguageKey(lang, key) {
  return optionLabelling[lang] && optionLabelling[lang][key] ? 
  optionLabelling[lang][key] : 
  optionLabelling[defaultLanguage][key]
}

module.exports = {
  renderComprehension,
  renderMCQ,
  renderMTF,
  renderQuestion,
};
