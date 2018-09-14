let fs = require("fs");
let cheerio = require("cheerio");
let xlsx = require("node-xlsx");
let superagent = require('superagent-charset');

let url = "http://www.wzfg.com/realweb/stat/ProjectSellingList.jsp";
let data = [{
  name: "1",
  data: [
    [1, 2, 3],
    [3, 4, 5]
  ]
}];

function formatTimeNow() {
  let time = new Date();
  return time.getFullYear() + "_" + (time.getMonth() + 1) + "_" + time.getDate() + "_" + time.getHours() + "_" + time.getMinutes()
}

function get(url) {
  return new Promise((resolve, reject) => {
    superagent.get(url).charset("gbk").end((err, sres) => {
      if (err) {
        reject(err);
      } else {
        let html = sres.text;
        let $ = cheerio.load(html, {decodeEntities: false});
        resolve($)
      }
    })
  })
}

function createExcel() {
  if (!fs.existsSync("excel")) {
    fs.mkdirSync("excel")
  }
  fs.writeFileSync("./excel/" + formatTimeNow() + ".xlsx", xlsx.build(data), "binary")
}

function getProjectSellingList($) {
  let projectList = [];
  return new Promise((resolve, reject) => {
    let $trs = $("tr[onclick]");
    $trs.each((i, tr) => {
      let projectName = $(tr).find("td:nth-child(2)").html();
      let onclick = $(tr).attr("onclick");
      console.log(onclick)
    })
    resolve();
  })
}

get(url)
  .then(($) => {
    return getProjectSellingList($)
  })
  .then(createExcel)
