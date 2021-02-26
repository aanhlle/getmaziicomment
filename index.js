const puppeteer = require('puppeteer');
const params = process.argv.slice(2);
const inputFileName = params[0];
const outputFileName = params[1];
const XLSX = require('xlsx');
const Excel = require('exceljs');
const readline = require('readline');
var URL;

const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

rl.question("Wanna find comments from Kanji or words? ", function(url) {

  if(url.toLowerCase() =="kanji"){
    URL = "https://mazii.net/#!/search?type=k&query="
    console.log("start searching kanji comment")
  } else {
    URL = "https://mazii.net/search?type=w&query="
    console.log("start searching word comment")
  }
  rl.close();
  main();
});

function main() {
  console.log("======= START ======")
  const wb = XLSX.readFile(`${inputFileName}.xlsx`);
  const sheet_name_list = wb.SheetNames;
  const data = XLSX.utils.sheet_to_json(wb.Sheets[sheet_name_list[0]]);
  doCrawl(data);
}

async function getCommentsOfKanji(URL, listKanji) {
  try {
    const browser = await puppeteer.launch();
    const pdfs = listKanji.map(async (k, i) => {
      const page = await browser.newPage();
      await page.goto(`${URL}${k}`, {
        waitUntil: 'networkidle2',
        timeout: 120000,
      });
      let content = await page.evaluate(async () => {
        let comments = document.querySelectorAll('p.mean');
        if (comments.length < 1) return 'Không có comment';
        comments = [...comments][0];
        return comments.innerHTML.trim();
      });
      await page.close();
      return {
        kanji: k,
        comment: content
      };
    });

    return Promise.all(pdfs).then((result) => {
      browser.close();
      return result;
    });
  } catch (err) {
    console.log("err", err);
  }
}

async function doCrawl(data) {
  let kanji = [];
  let count = 0;
  while (data.length > 0) {
    kanji = data.splice(0, 10).map(k => k['Kanji']);
    let num = count * 10;
    let eachRs = await getCommentsOfKanji(URL, kanji);
    await appendToExcel(eachRs,num);
    count++;
  }
  console.log("======= COMPLETED ======")
}

function appendToExcel(each10Rs,num) {
  const workbook = new Excel.Workbook();
  return workbook.xlsx.readFile(`${outputFileName}.xlsx`)
    .then(function (data) {
      var worksheet = workbook.getWorksheet(1);
      each10Rs.forEach((rs,index) => {
        let row = worksheet.getRow((num + index + 1));
        row.getCell(1).value = rs.kanji;
        row.getCell(2).value = rs.comment;
        row.commit();
      })
      console.log("DONE in " + ((num/10)+1).toString());
      return workbook.xlsx.writeFile(`${outputFileName}.xlsx`);
    })
  }
