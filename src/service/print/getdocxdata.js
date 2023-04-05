const docx = require("docx");
const { text } = require("express");
const { functions } = require("lodash");
const {
  Document,
  BorderStyle,
  Paragraph,
  TextRun,
  AlignmentType,
  ImageRun,
  Table,
  TableRow,
  TableCell,
  VerticalAlign,
  WidthType,
  HeadingLevel,
  ITableCellMarginOptions,
  convertInchesToTwip,
} = docx;
const _ = require("lodash");

const defaultLanguage = 'english';
const basicDetails={
  english : require('./lang/english.json'),
  hindi : require('./lang/hindi.json')
}

function create(data, paperData) {
  const language = (paperData && paperData.language) ? paperData.language.toLowerCase(): defaultLanguage;

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          headers(
            getLanguageKey(language, 'studentName') +  ":..................................................",
            getLanguageKey(language, 'rollNo') +  ":............................."
          ),
          new Paragraph({
            children: [], // Just newline without text
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: `${paperData.examName}`,
                heading: HeadingLevel.TITLE,
                bold: true,
              }),
            ],
          }),
          new Paragraph({
            children: [], // Just newline without text
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: getLanguageKey(language, 'class')+ `:${paperData.className}`,
                bold: true,
              }),
            ],
          }),
          new Paragraph({
            children: [], // Just newline without text
          }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: getLanguageKey(language, 'subject')+ `:${paperData.subject}`,
                bold: true,
              }),
            ],
          }),
          new Paragraph({
            children: [], // Just newline without text
          }),
          headers(getLanguageKey(language, 'time')+":" + `${paperData.maxTime}`, getLanguageKey(language, 'marks')+":"+`${paperData.maxScore}`),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: "__________________________________________________________________________________________",
                bold: true,
                thematicBreak: true,
              }),
            ],
          }),

          instructionHead(paperData.instructions, language),
          instructions(paperData.instructions),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: "__________________________________________________________________________________________",
                bold: true,
                thematicBreak: true,
              }),
            ],
          }),

          ...data
            .map((question) => {
              const arr = [];

              if (question !== undefined) {
                let page = 1;
                if (question[0].type === "COMPREHENSION") {
                  let count = 0;
                  question[0].Questions.map((item) => {
                    arr.push(
                      formatview(
                        item,
                        count,
                        question[0].QuestionIndex,
                        question[0].Marks
                      )
                    );
                    count++;
                  });

                  arr.push(
                    new Paragraph({
                      children: [], // Just newline without text
                    })
                  );
                } else if (question[0].type === "CuriosityQuestion") {
                  let count = 0;
                  question[0].Questions.map((item) => {
                    arr.push(
                      formatview(
                        item,
                        count,
                        question[0].QuestionIndex,
                        question[0].Marks
                      )
                    );
                    count++;
                  });

                  arr.push(
                    new Paragraph({
                      children: [], // Just newline without text
                    })
                  );
                } else if (question[0].type === "SA") {
                  let count = 0;
                  question[0].Questions.map((item) => {
                    arr.push(
                      formatview(
                        item,
                        count,
                        question[0].QuestionIndex,
                        question[0].Marks
                      )
                    );
                    count++;
                  });

                  arr.push(
                    new Paragraph({
                      children: [], // Just newline without text
                    })
                  );
                } else if (question[0].type === "LA") {
                  let count = 0;
                  question[0].Questions.map((item) => {
                    arr.push(
                      formatview(
                        item,
                        count,
                        question[0].QuestionIndex,
                        question[0].Marks
                      )
                    );
                    count++;
                  });
                  arr.push(
                    new Paragraph({
                      children: [], // Just newline without text
                    })
                  );
                } else if (question[0].type === "VSA") {
                  let count = 0;
                  question[0].Questions.map((item) => {
                    arr.push(
                      formatview(
                        item,
                        count,
                        question[0].QuestionIndex,
                        question[0].Marks
                      )
                    );
                    count++;
                  });
                  arr.push(
                    new Paragraph({
                      children: [], // Just newline without text
                    })
                  );
                } else if (question[0].type === "FTB") {
                  let count = 0;
                  question[0].Questions.map((item) => {
                    arr.push(
                      formatview(
                        item,
                        count,
                        question[0].QuestionIndex,
                        question[0].Marks
                      )
                    );
                    count++;
                  });
                  arr.push(
                    new Paragraph({
                      children: [], // Just newline without text
                    })
                  );
                } else if (question[0].type === "MCQ") {
                  let testimage = formatOptions(question[0]);
                  let count = 0;
                  question[0].Questions.map((item) => {
                    arr.push(
                      formatview(
                        item,
                        count,
                        question[0].QuestionIndex,
                        question[0].Marks
                      )
                    );
                    count++;
                  });
                  arr.push(
                    // new Paragraph({
                    //   children: [], // Just newline without text
                    // })
                  );
                  arr.push(optionsTabel(testimage));
                  arr.push(
                    new Paragraph({
                      children: [], // Just newline without text
                    })
                  );
                } else if (question[0].type === "MTF") {
                  let count = 0;
                  arr.push(
                    formatview(
                      question[0].heading,
                      count,
                      question[0].QuestionIndex,
                      question[0].Marks
                    )
                  );
                  arr.push(
                    new Paragraph({
                      children: [], // Just newline without text
                    })
                  );

                  arr.push(mtfTableData(question));
                  arr.push(
                    new Paragraph({
                      children: [], // Just newline without text
                    })
                  );
                } else if (question[0].type === "section") {
                  arr.push(
                    new Paragraph({
                      alignment: AlignmentType.LEFT,
                      children: [
                        new TextRun({
                          text: `${question[0].sectionHeader}`,
                          heading: HeadingLevel.TITLE,
                          bold: true,
                        }),
                      ],
                    }),
                    new Paragraph({
                      children: [], // Just newline without text
                    })
                  );
                }
              }
              return arr;
            })
            .reduce((prev, curr) => prev.concat(curr), []),
        ],
      },
    ],
  });

  return doc;
}

function instructionHead(data, language) {
  const arr = [];

  if (!_.isUndefined(data)) {
    arr.push(
      new TextRun({
        text: (getLanguageKey(language, 'instructionHead')+":"),
        bold: true,
      })
    );
    return new Paragraph({
      alignment: AlignmentType.LEFT,
      children: arr,
    });
  }
}

function instructions(data) {
  const arr = [];

  if (_.isUndefined(data)) {
    return new Paragraph({
      alignment: AlignmentType.LEFT,
      indent: {
        left: 720,
      },
      children: arr,
    });
  } else {
    data
      .map((text) => {
        arr.push(
          new TextRun({
            text: `${text}`,
            break: 1,
            bold: true,
          })
        );
      })
      .reduce((prev, curr) => prev.concat(curr), []);
    return new Paragraph({
      alignment: AlignmentType.LEFT,
      indent: {
        left: 720,
      },
      children: arr,
    });
  }
}

function displayMTFHeader(data) {
  return new TableCell({
    borders: MTFborder,
    width: {
      size: 4535,
      type: WidthType.DXA,
    },
    margins: {
      top: convertInchesToTwip(0.0693701),
      bottom: convertInchesToTwip(0.0693701),
      left: convertInchesToTwip(0.3493701),
      right: convertInchesToTwip(0.0693701),
    },
    verticalAlign: VerticalAlign.CENTER,
    children: [
      new Paragraph({
        text: data,
        bold: true,
      }),
    ],
  });
}

function displayMTFData(data) {
  return new TableCell({
    borders: MTFborder,
    width: {
      size: 4535,
      type: WidthType.DXA,
    },
    margins: {
      top: convertInchesToTwip(0.0693701),
      bottom: convertInchesToTwip(0.0693701),
      left: convertInchesToTwip(0.3493701),
      right: convertInchesToTwip(0.0693701),
    },
    verticalAlign: VerticalAlign.CENTER,
    children: [createSAObject(data, 0)],
  });
}

function MTFTabel(question) {
  const arr = [];
  const rowheading = new TableRow({
    children: [displayMTFHeader("Column1"), displayMTFHeader("Column2")],
  });
  arr.push(rowheading);
  question[0].Questions.map((item) => {
    arr.push(
      new TableRow({
        children: [displayMTFData(item.left[0]), displayMTFData(item.right[0])],
      })
    );
  });
  return new Table({
    columnWidths: [5000, 5000],
    rows: arr,
  });
}

function createFTB(data, count) {
  if (_.isUndefined(data)) {
    return new Paragraph({
      alignment: AlignmentType.LEFT,
      children: [
        new TextRun({
          text: ``,
          thematicBreak: true,
        }),
      ],
    });
  }
  if (count !== 0) {
    return new Paragraph({
      alignment: AlignmentType.LEFT,
      children: [
        new TextRun({
          text: "    " + `${data}`,
          thematicBreak: true,
        }),
      ],
    });
  } else {
    return new Paragraph({
      alignment: AlignmentType.LEFT,
      children: [
        new TextRun({
          text: `${data}`,
          thematicBreak: true,
        }),
      ],
    });
  }
}

function createSAObject(data, count) {
  const arr = [];
  if (_.isUndefined(data)) {
    return createFTB(data, count);
  }
  if (data.text) {
    data.text
      .map((text) => {
        if (typeof text === "object") {
          if (text.br === "break") {
            arr.push(
              new TextRun({
                break: 0.5,
              })
            );
          } else {
            arr.push(new TextRun(text));
          }
        } else {
          arr.push(
            new TextRun({
              text: `${text}`,
            })
          );
        }
      })
      .reduce((prev, curr) => prev.concat(curr), []);
    return new Paragraph({
      alignment: AlignmentType.LEFT,
      children: arr,
    });
  } else if (data.image) {
    if (data.image.includes("data:image/")) {
      let image = getBufferImg(data.image);

      return new Paragraph({
        children: [
          new ImageRun({
            data: image,
            transformation: {
              width: data.width,
              height: data.height,
            },
          }),
        ],
      });
    }
  } else if (data.ol) {
    let count1 = 0;
    data.ol
      .map((text) => {
        count1++;
        if (typeof text === "object") {
          text = count1 + text.text;
          arr.push(
            new TextRun({
              text: `${count1}.${text}`,
            })
          );
        } else {
          arr.push(
            new TextRun({
              text: `${count1}.${text}`,
              break: 0.5,
            })
          );
        }
      })
      .reduce((prev, curr) => prev.concat(curr), []);
    return new Paragraph({
      alignment: AlignmentType.LEFT,
      indent: {
        left: 200,
      },
      children: arr,
    });
  } else if (data.ul) {
    let count1 = 0;
    data.ul
      .map((text) => {
        if (typeof text === "object") {
          arr.push(new TextRun(text));
        } else {
          arr.push(
            new TextRun({
              text: ` ${text}`,
              break: 0.5,
            })
          );
        }
      })
      .reduce((prev, curr) => prev.concat(curr), []);
    return new Paragraph({
      alignment: AlignmentType.LEFT,
      indent: {
        left: 200,
      },
      children: arr,
    });
  } else {
    return createFTB(data, count);
  }
}

function imageData(image) {
  let bufferImage;
  if (image.includes("data:image/png;base64,")) {
    return (bufferImage = image.replace("data:image/png;base64,", ""));
  } else if (image.includes("data:image/jpg;base64")) {
    return (bufferImage = image.replace("data:image/jpg;base64,", ""));
  } else if (image.includes("data:image/jpeg;base64")) {
    return (bufferImage = image.replace("data:image/jpeg;base64,", ""));
  }
}

function getBufferImg(data) {
  let image = imageData(data);
  return image;
}
function formatOptions(data) {
  let optionArr = [];
  let image;
  let testimage = data;
  if (testimage) {
    optionArr.push(testimage.Option1);
    optionArr.push(testimage.Option2);
    optionArr.push(testimage.Option3);
    optionArr.push(testimage.Option4);
    optionArr.push(testimage.height1);
    optionArr.push(testimage.width1);
    optionArr.push(testimage.height2);
    optionArr.push(testimage.width2);
    optionArr.push(testimage.height3);
    optionArr.push(testimage.width3);
    optionArr.push(testimage.height4);
    optionArr.push(testimage.width4);
  }
  return optionArr;
}

const border = {
  left: {
    style: BorderStyle.NIL,
    size: 0,
  },
  right: {
    style: BorderStyle.NIL,
    size: 0,
  },
  top: {
    style: BorderStyle.NIL,
    size: 0,
  },
  bottom: {
    style: BorderStyle.NIL,
    size: 0,
  },
};

const MTFborder = {
  left: {
    style: BorderStyle.SINGLE,
    size: 2,
  },
  right: {
    style: BorderStyle.SINGLE,
    size: 2,
  },
  top: {
    style: BorderStyle.SINGLE,
    size: 2,
  },
  bottom: {
    style: BorderStyle.SINGLE,
    size: 2,
  },
};

const MCQborder = {
  left: {
    style: BorderStyle.NONE,
    size: 0,
    color: "ffffff",
  },
  right: {
    style: BorderStyle.NONE,
    size: 0,
    color: "ffffff",
  },
  top: {
    style: BorderStyle.NONE,
    size: 0,
    color: "ffffff",
  },
  bottom: {
    style: BorderStyle.NONE,
    size: 0,
    color: "ffffff",
  },
};

function headers(text1, text2) {
  return new Table({
    columnWidths: [4505, 4505],
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: border,
            width: {
              size: 100 / 2,
              type: WidthType.PERCENTAGE,
            },
            children: [
              new Paragraph({
                alignment: AlignmentType.LEFT,
                children: [
                  new TextRun({
                    text: text1,
                    bold: true,
                  }),
                ],
              }),
            ],
          }),
          new TableCell({
            borders: border,
            width: {
              size: 100 / 2,
              type: WidthType.PERCENTAGE,
            },
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT,
                children: [
                  new TextRun({
                    text: text2,
                    bold: true,
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

function displayNumber(data) {
  if (data !== undefined) {
    if (typeof data === "object") {
      return new TableCell({
        borders: MCQborder,
        width: {
          size: 650,
          type: WidthType.DXA,
        },
        margins: {
          top: convertInchesToTwip(0.0693701),
          bottom: convertInchesToTwip(0.0693701),
          left: convertInchesToTwip(0.3493701),
          right: convertInchesToTwip(0.0693701),
        },
        verticalAlign: VerticalAlign.CENTER,
        children: [
          new Paragraph({
            text: data[0],
          }),
        ],
      });
    } else {
      return new TableCell({
        borders: MCQborder,
        width: {
          size: 650,
          type: WidthType.DXA,
        },
        margins: {
          top: convertInchesToTwip(0.0093701),
          bottom: convertInchesToTwip(0.0093701),
          left: convertInchesToTwip(0.3493701),
          right: convertInchesToTwip(0.0693701),
        },
        verticalAlign: VerticalAlign.CENTER,
        children: [
          new Paragraph({
            text: data[0],
          }),
        ],
      });
    }
  } else {
    return new TableCell({
      borders: MCQborder,
      width: {
        size: 650,
        type: WidthType.DXA,
      },
      margins: {
        top: convertInchesToTwip(0.0693701),
        bottom: convertInchesToTwip(0.0693701),
        left: convertInchesToTwip(0.3493701),
        right: convertInchesToTwip(0.0693701),
      },
      verticalAlign: VerticalAlign.CENTER,
      children: [
        new Paragraph({
          text: "",
        }),
      ],
    });
  }
}

function displayOptionsObject(data, count) {
  const arr = [];
  if (data.text) {
    if (typeof data === "object") {
      arr.push(new TextRun(data));
    } else {
      data.text
        .map((text) => {
          if (typeof text === "object") {
            arr.push(new TextRun(text));
          }
        })
        .reduce((prev, curr) => prev.concat(curr), []);
    }

    return new TableCell({
      borders: MCQborder,
      width: {
        size: 4505,
        type: WidthType.DXA,
      },
      margins: {
        top: convertInchesToTwip(0.0693701),
        bottom: convertInchesToTwip(0.0693701),
        left: convertInchesToTwip(0.0693701),
        right: convertInchesToTwip(0.0693701),
      },
      verticalAlign: VerticalAlign.CENTER,
      children: [
        new Paragraph({
          children: arr,
        }),
      ],
    });
  }
}

function displayOptions(option, height, width) {
  if (option !== undefined) {
    if (typeof option[1] === "object") {
      return displayOptionsObject(option[1]);
    } else if (option[1].includes("data:image/")) {
      let image = getBufferImg(option[1]);
      return new TableCell({
        borders: MCQborder,
        width: {
          size: 4505,
          type: WidthType.DXA,
        },
        margins: {
          top: convertInchesToTwip(0.0693701),
          bottom: convertInchesToTwip(0.0693701),
          left: convertInchesToTwip(0.0693701),
          right: convertInchesToTwip(0.0693701),
        },
        verticalAlign: VerticalAlign.CENTER,
        children: [
          new Paragraph({
            children: [
              new ImageRun({
                data: image,
                transformation: {
                  width: width,
                  height: height,
                },
              }),
            ],
          }),
        ],
      });
    } else {
      return new TableCell({
        borders: MCQborder,
        width: {
          size: 4505,
          type: WidthType.DXA,
        },
        margins: {
          top: convertInchesToTwip(0.0693701),
          bottom: convertInchesToTwip(0.0693701),
          left: convertInchesToTwip(0.0693701),
          right: convertInchesToTwip(0.0693701),
        },
        verticalAlign: VerticalAlign.CENTER,
        children: [
          new Paragraph({
            children: [
              new TextRun({
                text: option[1],
              }),
            ],
          }),
        ],
      });
    }
  } else {
    return new TableCell({
      borders: MCQborder,
      width: {
        size: 650,
        type: WidthType.DXA,
      },
      margins: {
        top: convertInchesToTwip(0.0693701),
        bottom: convertInchesToTwip(0.0693701),
        left: convertInchesToTwip(0.3493701),
        right: convertInchesToTwip(0.0693701),
      },
      verticalAlign: VerticalAlign.CENTER,
      children: [
        new Paragraph({
          text: "",
        }),
      ],
    });
  }
}

function displayViewData(data) {
  return new TableCell({
    borders: MCQborder,
    width: {
      size: 8000,
      type: WidthType.DXA,
    },
    verticalAlign: VerticalAlign.LEFT,
    children: [createSAObject(data, 0)],
  });
}

function displayQueNum(data) {
  return new TableCell({
    borders: MCQborder,
    width: {
      size: 500,
      type: WidthType.DXA,
    },
    verticalAlign: VerticalAlign.LEFT,
    children: [createSAObject(data, 0)],
  });
}

function displayMarks(data) {
  return new TableCell({
    borders: MCQborder,
    width: {
      size: 500,
      type: WidthType.DXA,
    },
    margins: {
      left: convertInchesToTwip(0.3493701),
    },
    verticalAlign: VerticalAlign.LEFT,
    children: [createSAObject(data, 0)],
  });
}

function formatview(data, count, questionCounter, marks) {
  if (count === 0) {
    return new Table({
      borders: MCQborder,
      columnWidths: [300, 10000, 300],
      rows: [
        new TableRow({
          indent: {
            left: 100,
          },
          children: [
            displayQueNum(questionCounter),
            displayViewData(data),
            displayMarks(marks),
          ],
        }),
      ],
    });
  } else {
    return new Table({
      borders: MCQborder,
      columnWidths: [300, 10000, 300],
      rows: [
        new TableRow({
          indent: {
            left: 100,
          },
          children: [
            displayQueNum(""),
            displayViewData(data),
            displayMarks(""),
          ],
        }),
      ],
    });
  }
}

function mtfTableData(data) {
  const cell = new TableCell({
    children: [MTFTabel(data)],
  });

  return new Table({
    borders: MCQborder,
    columnWidths: [300, 10000, 300],
    rows: [
      new TableRow({
        children: [displayQueNum(""), cell, displayMarks("")],
      }),
    ],
  });
}

function optionsTabel(testimage) {
  return new Table({
    borders: MCQborder,
    columnWidths: [4505, 4505],
    rows: [
      new TableRow({
        children: [
          displayNumber(testimage[0]),

          displayOptions(testimage[0], testimage[4], testimage[5]),

          displayNumber(testimage[1]),

          displayOptions(testimage[1], testimage[6], testimage[7]),
        ],
      }),
      new TableRow({
        indent: {
          left: 600,
        },
        children: [
          displayNumber(testimage[2]),

          displayOptions(testimage[2], testimage[8], testimage[9]),

          displayNumber(testimage[3]),

          displayOptions(testimage[3], testimage[10], testimage[11]),
        ],
      }),
    ],
  });
}

function getLanguageKey(lang, key) {
  return basicDetails[lang] && basicDetails[lang][key] ? 
  basicDetails[lang][key] : 
  basicDetails[defaultLanguage][key]
}

module.exports = {
  create,
};