const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, AlignmentType } = require('docx');

const initialPath = './project-logs/';

const logsFileNames = fs
  .readdirSync(initialPath)
  .filter((fileName) => fileName.endsWith('.csv'));

const globalStyles = {
  font: 'Arial',
};
const dateStyles = {
  size: 32,
  bold: true,
  alignment: AlignmentType.CENTER,
};

const nameStyles = {
  size: 28,
  bold: true,
};

const bulletStyles = {
  size: 24,
  level: 0,
};

const dateParagraph = (date) =>
  new Paragraph({
    children: [
      new TextRun({
        text: date,
        size: dateStyles.size,
        bold: dateStyles.bold,
        font: globalStyles.font,
      }),
    ],
    alignment: dateStyles.alignment,
  });

const nameParagraph = (name) =>
  new Paragraph({
    children: [
      new TextRun({
        text: name,
        size: nameStyles.size,
        bold: nameStyles.bold,
        font: globalStyles.font,
      }),
    ],
  });

const bulletParagraph = (text) =>
  new Paragraph({
    children: [
      new TextRun({
        text,
        size: bulletStyles.size,
        font: globalStyles.font,
      }),
    ],
    bullet: {
      level: bulletStyles.level,
    },
  });

const children = [];

const months = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December',
];

const getFormattedDate = (date) =>
  new Date(date.split(',')[0]).toLocaleDateString('pk', {
    month: 'long',
    day: 'numeric',
    weekday: 'long',
  });

const getFormattedDescriptions = (description) =>
  description
    .slice(0, description.lastIndexOf('('))
    .replaceAll('[Coding] - ', '')
    .replaceAll('[Meeting] - ', '')
    .replaceAll('[Reporting/Analysis] - ', '')
    .replaceAll('[Training/Learning] - ', '')
    .split('.');

logsFileNames.sort((m1, m2) => {
  const month1 = m1.slice(m1.lastIndexOf(' ') + 1, m1.lastIndexOf('-'));
  const month2 = m2.slice(m2.lastIndexOf(' ') + 1, m2.lastIndexOf('-'));

  return months.indexOf(month1) > months.indexOf(month2) ? 1 : -1;
});
logsFileNames.forEach((fileName, index) => {
  let data = fs.readFileSync(`${initialPath}${fileName}`, 'ascii');
  data = data.split('\n"');
  const teamNames = ['BVS - Deets - Frontend', 'H-Track - Frontend'];
  data.forEach((entry) => {
    const entryData = entry.split(',"');

    const indexesToFilter = [0, 1, 3, 4];

    const [name, team, date, descriptions] = entryData
      .filter((field, index) => indexesToFilter.includes(index))
      .map((field) => field.replaceAll('"', ''));

    if (teamNames.includes(team)) {
      const formattedDate = getFormattedDate(date);

      children.push(dateParagraph(formattedDate), nameParagraph(name));

      descriptions.split('\n').forEach((description) => {
        getFormattedDescriptions(description).forEach(
          (entry) =>
            entry.trim() && children.push(bulletParagraph(entry.trim()))
        );
      });

      children.push(...['', '', ''].map(dateParagraph));
    }
  });

  if (index === logsFileNames.length - 1) {
    const doc = new Document({
      sections: [
        {
          properties: {},
          children,
        },
      ],
    });

    Packer.toBuffer(doc).then((dataBuffer) =>
      fs.writeFileSync(`logs.docx`, dataBuffer)
    );
  }
});
