let fs = require("fs");
let cheerio = require("cheerio");
let xlsx = require("node-xlsx");
let superagent = require('superagent-charset');

let projectUrl = "http://www.wzfg.com/realweb/stat/ProjectSellingList.jsp?";
let infoUrl = "http://www.wzfg.com/realweb/stat/FirstHandProjectInfo.jsp?"

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

function createExcel(filename, data) {
  if (!fs.existsSync("excel")) {
    fs.mkdirSync("excel")
  }

  fs.writeFileSync("./excel/" + filename + ".xlsx", xlsx.build(data), "binary")
  console.log("已生成 " + filename + ".xlsx ......");
}

function getProjectSellingList(min, max) {
  console.log("正在读取第" + min + "页数据")
  let data = [{
    name: "项目概况",
    data: []
  }, {
    name: "分类价格",
    data: []
  }, {
    name: "销售情况",
    data: []
  }];

  function initData($, table, data) {
    table.find("> tbody").children().each((i, tr) => {
      let trArr = []
      $(tr).find("> td").each((i, td) => {
        let $td = $(td);
        if ($td.attr("rowspan") || $td.attr("colspan")) {
          return
        }
        trArr.push($td.text())
      })
      data.push(trArr)
    })
  }

  get(projectUrl + "currPage=" + (min - 1)).then(($) => {
    if (max === true) {
      if (/goPage\((\d*)\)/.test($("a[href*=goPage]").last().attr("href"))) {
        max = RegExp.$1
      }
    }
    let $trs = $("tr[onclick]");
    let promises = [];
    $trs.each((i, tr) => {
      let promise = new Promise((resolve) => {
        //项目名字
        let projectName = $(tr).find("td:nth-child(2)").html();
        let onclick = $(tr).attr("onclick");
        let query = onclick.match(/projectID=\d*/)[0];
        get(infoUrl + query).then(($) => {
          let table1 = $("table.ab")
          let table2 = $("table:not(.MsoNormalTable)").eq(5);
          let table3 = $("table:not(.MsoNormalTable)").eq(7);
          initData($, table1, data[0].data)
          initData($, table2, data[1].data)
          initData($, table3, data[2].data)
          createExcel(projectName, data)
          resolve();
        })
      })
      promises.push(promise)
    })
    Promise.all(promises).then(() => {
      if (min < max) {
        getProjectSellingList(min + 1, max)
      }
    })
  })
}

function getPage() {
  return new Promise((resolve, reject) => {
    let pages = null;
    let str = process.argv[2]
    if (str) {
      let split = null, min = 1, max = 1;
      if (/^(\d*)$/.test(str)) {
        min = max = RegExp.$1;
      } else if (split = str.split("-")) {
        if (split.length === 2) {
          min = Math.min(split[0], split[1])
          max = Math.max(split[0], split[1])
        } else {
          reject();
        }
      }
      resolve({
        min: min,
        max: max
      })
    }
    resolve(pages)


  })
}

getPage().then((pages) => {
  if (!pages) {
    console.log("开始读取所有项目")
    getProjectSellingList(1, true)
  } else {
    if (pages.min === pages.max) {
      console.log("开始读取第" + pages.min + "页项目")
    } else {
      console.log("开始读取第" + pages.min + "-" + pages.max + "页项目")
    }
    getProjectSellingList(pages.min, pages.max)
  }
}).catch(() => {
  console.log("请输入正确的页数，如\n node index.js 1-100,\n node index.js 1")
})
