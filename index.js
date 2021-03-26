let rawTemplate = document.querySelector("#rawTemplate");
let templateData = document.querySelector("#data");
let renderedTemplate = document.querySelector("#renderedTemplate");
let downloadBtn = document.querySelector("#download");

let sheetData = [
  {
    name: "任务1",
    address: "美国纽约",
    time: "2019.12.18-2019.12.23",
    day: 6,
    people: [
      { name: "张三1", dept: "部长1" },
      { name: "张三2", dept: "部长2" }
    ]
  },
  {
    name: "任务2",
    address: "美国纽约2",
    time: "2019.12.18-2019.12.23",
    day: 9,
    people: [
      { name: "张三3", dept: "部长3" },
      { name: "张三4", dept: "部长4" }
    ]
  }
];

templateData.innerText = JSON.stringify(sheetData, null, 2);

axios({
  method: "get",
  url: "https://cdn.jsdelivr.net/gh/GoldSubmarine/export-excel-demo@master/template.zip",
  responseType: "arraybuffer"
}).then(function(req) {
  let zip = new JSZip();
  zip.loadAsync(req.data).then(data => {
    data
      .file("xl/worksheets/sheet1.xml")
      .async("string")
      .then(sheet1 => {
        let sheetStr = ejs.render(sheet1, { data: sheetData });
        rawTemplate.innerText = sheet1
        renderedTemplate.innerText = sheetStr
        downloadBtn.onclick = function() {
          zip.file("xl/worksheets/sheet1.xml", sheetStr);
          zip
            .generateAsync({
              type: "blob",
              mimeType:
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            })
            .then(function(blob) {
              saveAs(blob, "output.xlsx");
            });
        };
      });
  });
});
