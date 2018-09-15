let fs = require("fs");
let cheerio = require("cheerio");
let xlsx = require("node-xlsx");
let superagent = require('superagent-charset');

let projectUrl = "http://www.wzfg.com/realweb/stat/ProjectSellingList.jsp?";
let infoUrl = "http://www.wzfg.com/realweb/stat/FirstHandProjectInfo.jsp?"

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

function createExcel(filename, data) {
  if (!fs.existsSync("excel")) {
    fs.mkdirSync("excel")
  }

  // fs.writeFileSync("./excel/" + formatTimeNow() + ".xlsx", xlsx.build(data), "binary")
  fs.writeFileSync("./excel/" + filename + ".xlsx", xlsx.build(data), "binary")
  console.log("已生成 " + filename + ".xlsx ......");
}

function getProjectSellingList(pageIndex) {
  console.log("正在读取第" + (pageIndex + 1) + "页数据")
  let data = [{
    name: "项目概况",
    data: [
      ["开发单位", "预售许可证号", "发证日期", "所在地区", "样本区域", "项目测算面积", "项目名称", "项目地址", "开盘日期", "售楼地址", "售楼电话"],
      [3, 4, 5]
    ]
  }, {
    name: "分类价格",
    data: [
      [1, 2, 3],
      [3, 4, 5]
    ]
  }, {
    name: "销售情况",
    data: [
      [1, 2, 3],
      [3, 4, 5]
    ]
  }];
  get(projectUrl + "currPage=" + pageIndex).then(($) => {
    let $trs = $("tr[onclick]");
    $trs.each((i, tr) => {
      let projectName = $(tr).find("td:nth-child(2)").html();
      let onclick = $(tr).attr("onclick");
      let query = onclick.match(/projectID=\d*/)[0];
      get(infoUrl + query).then(($) => {
        console.log($(`td:contains('开发单位')`)).html()
        // createExcel(projectName, data)
      })
    })
  })

}

getProjectSellingList(0)
