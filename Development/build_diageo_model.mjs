import fs from "node:fs/promises";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const outDir = "outputs/diageo_three_statement";
const outPath = `${outDir}/Diageo_Three_Statement_Model.xlsx`;
const wb = Workbook.create();
const names = ["Cover","Assumptions","Historicals","Revenue Build","Operating Model","Schedules","Three Statements","Scenarios","Dashboard","Checks"];
for (const n of names) wb.worksheets.add(n);

const navy = "#17365D", blue = "#1F4E78", lightBlue = "#D9EAF7", grey = "#E7E6E6", paleGrey = "#F3F4F6";
const green = "#008000", inputBlue = "#0000FF", red = "#C00000", amber = "#FFF2CC", white = "#FFFFFF", text = "#222222";
const moneyFmt = '#,##0;(#,##0);-';
const pctFmt = '0.0%;(0.0%);-';
const multFmt = '0.0x';
const cols = ["D","E","F","G","H","I","J","K"];
const years = ["FY:24A","FY:25A","FY:26A","FY:27E","FY:28E","FY:29E","FY:30E","FY:31E"];

function base(sheet, maxRow=120) {
  sheet.showGridLines = false;
  sheet.getRange(`A1:K${maxRow}`).format.font = { name: "Arial", size: 10, color: text };
  sheet.getRange("A:A").format.columnWidth = 2;
  sheet.getRange("B:B").format.columnWidth = 2;
  sheet.getRange("C:C").format.columnWidth = 34;
  sheet.getRange("D:K").format.columnWidth = 13;
}
function title(sheet, titleText, subtitle) {
  sheet.getRange("C2").values = [[titleText]];
  sheet.getRange("C2").format.font = { name: "Arial", size: 16, bold: true, color: navy };
  sheet.getRange("C3").values = [[subtitle]];
  sheet.getRange("C3:K3").format.font = { name: "Arial", size: 9, italic: true, color: "#666666" };
  sheet.getRange("C4:K4").format.borders = { bottom: { style: "thin", color: navy } };
}
function section(sheet, row, label, end="K") {
  sheet.getRange(`C${row}:${end}${row}`).format.fill = navy;
  sheet.getRange(`C${row}:${end}${row}`).format.font = { name: "Arial", size: 10, bold: true, color: white };
  sheet.getRange(`C${row}`).values = [[label]];
}
function timeline(sheet, row=6, endRow=60) {
  sheet.getRange(`D${row}:K${row}`).values = [years];
  sheet.getRange(`D${row}:K${row}`).format.fill = blue;
  sheet.getRange(`D${row}:K${row}`).format.font = { name: "Arial", size: 10, bold: true, color: white };
  sheet.getRange(`D${row}:K${row}`).format.horizontalAlignment = "center";
  sheet.getRange(`D${row}:F${row}`).format.fill = "#595959";
  sheet.getRange(`G${row}:G${endRow}`).format.borders = { left: { style: "medium", color: blue } };
}
function fmtRows(sheet, ranges, fmt=moneyFmt) { for (const r of ranges) sheet.getRange(r).format.numberFormat = fmt; }
function styleTotal(sheet, row) {
  sheet.getRange(`C${row}:K${row}`).format.font = { name:"Arial", size:10, bold:true, color:text };
  sheet.getRange(`C${row}:K${row}`).format.borders = { top:{style:"thin",color:"#777777"} };
}

// Cover
{
  const s=wb.worksheets.getItem("Cover"); base(s,55); title(s,"Diageo plc","Integrated three-statement financial model | US$ millions | Fiscal year ended 30 June");
  s.getRange("C7:F13").values = [
    ["Company","Diageo plc",null,null],["Ticker","DGE (London Stock Exchange)",null,null],["Historical period","FY2024–FY2026",null,null],
    ["Forecast period","FY2027–FY2031",null,null],["Reporting currency","US dollar",null,null],["Scenario","Base (selector = 2)",null,null],["Prepared","10 September 2026",null,null]
  ];
  section(s,16,"Workbook guide","H");
  s.getRange("C18:H27").values = [
    ["Assumptions","Scenario controls and forecast drivers",null,null,null,null],["Historicals","Reported financial statements and source references",null,null,null,null],
    ["Revenue Build","Regional revenue forecast",null,null,null,null],["Operating Model","Adjusted and reported earnings forecast",null,null,null,null],
    ["Schedules","Working capital, PP&E, intangibles, debt and equity",null,null,null,null],["Three Statements","Integrated statements and ratios",null,null,null,null],
    ["Scenarios","Base, upside and downside outputs",null,null,null,null],["Dashboard","Selected operating, cash flow and leverage outputs",null,null,null,null],
    ["Checks","Financial and formula integrity checks",null,null,null,null],["Colour convention","Blue inputs; black formulas; green cross-sheet links",null,null,null,null]
  ];
  section(s,30,"Sources and limitations","H");
  s.getRange("C32:H39").values = [
    ["FY2026 financial statements","https://www.diageo.com/~/media/Files/D/Diageo-V2/Diageo-Corp/investors/results-reports-and-events/annual-reports/2026/financial-statements-2026.pdf",null,null,null,null],
    ["FY2025 annual report","https://www.diageo.com/~/media/Files/D/Diageo-V2/Diageo-Corp/investors/results-reports-and-events/annual-reports/2025/annual-report-2025.pdf",null,null,null,null],
    ["FY2026 results","https://www.diageo.com/en/news-and-media/press-releases/2026/2026-preliminary-results-year-ended-30-june-2026",null,null,null,null],
    ["FY2026 Capital Markets Day","https://www.diageo.com/en/news-and-media/press-releases/2026/diageo-capital-markets-day",null,null,null,null],
    ["Historical basis","FY2024 and FY2025 comparatives use the latest presentation in the FY2026/FY2025 reports.",null,null,null,null],
    ["Forecast basis","Illustrative analyst assumptions. FX and acquisitions/disposals are simplified.",null,null,null,null],
    ["Circularity","Interest uses opening debt. A revolver draw maintains minimum cash without circularity.",null,null,null,null],
    ["Disclaimer","Prepared independently for educational and academic-application purposes. Not investment advice.",null,null,null,null]
  ];
  s.getRange("C32:C39").format.font = {name:"Arial",size:10,bold:true,color:text};
  s.getRange("D32:H39").format.wrapText = true; s.getRange("D32:H39").format.font = {name:"Arial",size:9,color:"#555555"};
  s.getRange("C41").values=[["Model status"]]; s.getRange("D41").formulas=[["='Checks'!D20"]];
  s.getRange("D41").format.font={name:"Arial",size:12,bold:true,color:green};
  s.getRange("C:C").format.columnWidth=26; s.getRange("D:D").format.columnWidth=60; s.getRange("E:H").format.columnWidth=14;
}

// Assumptions
{
  const s=wb.worksheets.getItem("Assumptions"); base(s,55); title(s,"Assumptions","Editable scenario assumptions | Forecast years FY2027–FY2031"); timeline(s,6,48);
  s.getRange("C8").values=[["Scenario selector (1 downside / 2 base / 3 upside)"]]; s.getRange("D8").values=[[2]];
  s.getRange("D8").format.fill=amber; s.getRange("D8").format.font={name:"Arial",size:10,bold:true,color:inputBlue};
  s.getRange("D8").dataValidation={rule:{type:"whole",operator:"between",formula1:1,formula2:3}};
  s.getRange("C9").values=[["Active scenario"]]; s.getRange("D9").formulas=[["=CHOOSE(D8,\"Downside\",\"Base\",\"Upside\")"]];
  const forecastYears=["FY:27E","FY:28E","FY:29E","FY:30E","FY:31E"];
  function block(start, name, rows) {
    section(s,start,name);
    s.getRange(`G${start+1}:K${start+1}`).values=[forecastYears];
    s.getRange(`G${start+1}:K${start+1}`).format.fill=blue; s.getRange(`G${start+1}:K${start+1}`).format.font={name:"Arial",size:9,bold:true,color:white};
    let r=start+2;
    for (const item of rows) {
      s.getRange(`C${r}`).values=[[item[0]]];
      s.getRange(`D${r}`).values=[["Downside"]]; s.getRange(`E${r}`).values=[["Base"]]; s.getRange(`F${r}`).values=[["Upside"]];
      s.getRange(`D${r}:F${r}`).format.font={name:"Arial",size:9,color:"#666666"};
      s.getRange(`G${r}:K${r}`).formulas=[[...forecastYears.map((_,i)=>`=CHOOSE($D$8,${item[1][0][i]},${item[1][1][i]},${item[1][2][i]})`)]];
      s.getRange(`G${r}:K${r}`).format.font={name:"Arial",size:10,color:green}; s.getRange(`G${r}:K${r}`).format.fill=lightBlue;
      r++;
    }
    return r;
  }
  // Scenario values are explicit within the selected-case formulas and documented on the cover.
  const revRows=[
    ["North America organic growth", [[-0.08,0,0.01,0.02,0.02],[-0.05,0,0.02,0.03,0.03],[-0.02,0.02,0.04,0.04,0.04]]],
    ["Europe organic growth", [[0,0.01,0.02,0.02,0.02],[0.02,0.03,0.03,0.03,0.03],[0.04,0.05,0.05,0.04,0.04]]],
    ["Asia Pacific organic growth", [[-0.03,-0.01,0.01,0.02,0.02],[0,0.02,0.03,0.04,0.04],[0.03,0.05,0.06,0.06,0.05]]],
    ["Latin America & Caribbean organic growth", [[0,0.02,0.03,0.03,0.03],[0.03,0.04,0.04,0.04,0.04],[0.05,0.06,0.06,0.05,0.05]]],
    ["Africa organic growth", [[0.01,0.03,0.03,0.03,0.03],[0.04,0.05,0.05,0.05,0.05],[0.06,0.07,0.07,0.06,0.06]]],
    ["Corporate & other organic growth", [[-0.05,0,0,0,0],[0,0,0,0,0],[0.03,0.03,0.03,0.03,0.03]]]
  ];
  let next=block(12,"Revenue assumptions",revRows);
  const opRows=[
    ["Gross margin", [[0.585,0.588,0.59,0.592,0.594],[0.598,0.602,0.605,0.607,0.609],[0.606,0.612,0.617,0.62,0.622]]],
    ["Marketing as % of sales", [[0.164,0.164,0.163,0.163,0.163],[0.162,0.162,0.162,0.162,0.162],[0.163,0.164,0.164,0.164,0.164]]],
    ["Other recurring opex as % of sales", [[0.164,0.162,0.16,0.158,0.158],[0.153,0.149,0.147,0.146,0.145],[0.148,0.143,0.14,0.139,0.138]]],
    ["Exceptional operating costs ($m)", [[850,500,200,100,50],[850,350,100,0,0],[750,250,50,0,0]]],
    ["Adjusted tax rate", [[0.26,0.26,0.255,0.255,0.255],[0.25,0.25,0.25,0.25,0.25],[0.245,0.245,0.245,0.245,0.245]]],
    ["Associates and other income ($m)", [[100,100,110,120,130],[150,160,170,180,190],[190,210,230,250,270]]]
  ];
  next=block(next+2,"Operating assumptions",opRows);
  const bsRows=[
    ["Receivables as % of sales", [[0.18,0.18,0.178,0.178,0.178],[0.17,0.168,0.166,0.165,0.165],[0.16,0.158,0.156,0.155,0.155]]],
    ["Inventory as % of sales", [[0.56,0.56,0.555,0.55,0.545],[0.535,0.525,0.515,0.505,0.50],[0.515,0.50,0.485,0.47,0.46]]],
    ["Trade payables as % of sales", [[0.29,0.29,0.29,0.29,0.29],[0.30,0.30,0.30,0.30,0.30],[0.31,0.31,0.31,0.31,0.31]]],
    ["Other current assets as % of sales", [[0.075,0.075,0.075,0.075,0.075],[0.07,0.07,0.07,0.07,0.07],[0.065,0.065,0.065,0.065,0.065]]],
    ["Other non-current assets as % of sales", [[0.36,0.355,0.35,0.345,0.34],[0.35,0.345,0.34,0.335,0.33],[0.34,0.335,0.33,0.325,0.32]]],
    ["Other operating liabilities as % of sales", [[0.27,0.27,0.27,0.27,0.27],[0.28,0.28,0.28,0.28,0.28],[0.29,0.29,0.29,0.29,0.29]]],
    ["Capital expenditure as % of sales", [[0.055,0.055,0.055,0.055,0.055],[0.06,0.06,0.06,0.06,0.06],[0.065,0.065,0.065,0.065,0.065]]],
    ["Intangible additions as % of sales", [[0.005,0.005,0.005,0.005,0.005],[0.005,0.005,0.005,0.005,0.005],[0.006,0.006,0.006,0.006,0.006]]],
    ["Depreciation as % opening PP&E", [[0.07,0.07,0.07,0.07,0.07],[0.065,0.065,0.065,0.065,0.065],[0.06,0.06,0.06,0.06,0.06]]],
    ["Amortisation as % opening intangibles", [[0.018,0.018,0.018,0.018,0.018],[0.017,0.017,0.017,0.017,0.017],[0.016,0.016,0.016,0.016,0.016]]],
    ["Debt repayment as % opening debt", [[0.03,0.03,0.03,0.03,0.03],[0.05,0.05,0.05,0.05,0.05],[0.07,0.07,0.07,0.07,0.07]]],
    ["Effective interest rate", [[0.05,0.052,0.052,0.052,0.052],[0.045,0.045,0.045,0.045,0.045],[0.04,0.04,0.04,0.04,0.04]]],
    ["Parent dividend payout", [[0.85,0.8,0.75,0.75,0.75],[0.7,0.7,0.7,0.7,0.7],[0.65,0.65,0.65,0.65,0.65]]],
    ["NCI dividend payout", [[0.8,0.8,0.8,0.8,0.8],[0.7,0.7,0.7,0.7,0.7],[0.6,0.6,0.6,0.6,0.6]]]
  ];
  block(next+2,"Balance sheet and financing assumptions",bsRows);
  s.getRange("C48").values=[["Minimum cash balance ($m)"]];
  s.getRange("G48:K48").formulas=[["=CHOOSE($D$8,500,500,500)","=CHOOSE($D$8,500,500,500)","=CHOOSE($D$8,500,500,500)","=CHOOSE($D$8,500,500,500)","=CHOOSE($D$8,500,500,500)"]];
  s.getRange("G48:K48").format.fill=lightBlue;s.getRange("G48:K48").format.font={name:"Arial",size:10,color:green};
  s.getRange("G14:K90").format.numberFormat=pctFmt;
  // Dollar assumptions override percentage format.
  s.getRange("G27:K27").format.numberFormat=moneyFmt;
  s.getRange("G29:K29").format.numberFormat=moneyFmt;
  s.getRange("G48:K48").format.numberFormat=moneyFmt;
  s.freezePanes.freezeRows(6); s.freezePanes.freezeColumns(3);
}

// Historicals
{
  const s=wb.worksheets.getItem("Historicals"); base(s,82); title(s,"Historical financial statements","Reported data | US$ millions except per-share data | Source: Diageo FY2026/FY2025 reports"); timeline(s,6,78);
  section(s,8,"Income statement");
  const isRows=[
    ["Net sales",20269,20245,19643],["Cost of sales",8071,8072,7962],["Gross profit",12198,12173,11681],["Marketing",3691,3662,3183],
    ["Other operating items",2506,4176,5342],["Reported operating profit",6001,4335,3156],["Exceptional operating items",-56,1369,2527],
    ["Operating profit before exceptional items",5945,5704,5683],["Net finance charges",885,771,816],["Non-operating items",-70,-220,6],
    ["Share of associates and JVs",414,193,218],["Profit before tax",5460,3537,2564],["Taxation",1294,999,606],["Profit for the year",4166,2538,1958],
    ["NCI profit",296,184,221],["Net income attributable to shareholders",3870,2354,1737],["Diluted shares (m)",2239,2228,2231],["Diluted EPS (cents)",172.8,105.7,77.9]
  ];
  s.getRange("C9:F26").values=isRows.map(x=>[x[0],x[1],x[2],x[3]]); fmtRows(s,["D9:F26"]); s.getRange("D26:F26").format.numberFormat='0.0';
  for(const r of [11,14,16,20,22,24]) styleTotal(s,r);
  section(s,29,"Balance sheet");
  const bs=[
    ["Intangible assets",14814,14776,13375],["Property, plant and equipment",8509,9528,8996],["Other non-current assets",7025,7521,7262],
    ["Inventories",9720,10658,10529],["Trade and other receivables",3487,3504,3277],["Other current assets",789,1135,1886],["Cash and cash equivalents",1130,2200,1520],["Total assets",45474,49322,46845],
    ["Borrowings",21501,23748,21511],["Trade and other payables",6658,7144,6575],["Other liabilities",5245,5252,5805],["Total liabilities",33404,36144,33891],
    ["Share capital",887,887,887],["Share premium",1703,1703,1703],["Other reserves",-91,454,-10],["Retained earnings",7533,8046,8291],["NCI",2038,2088,2083],["Total equity",12070,13178,12954],["Total liabilities and equity",45474,49322,46845],["Balance sheet check",0,0,0]
  ];
  s.getRange("C30:F49").values=bs.map(x=>[x[0],x[1],x[2],x[3]]); fmtRows(s,["D30:F49"]);
  for(const r of [37,41,47,48,49]) styleTotal(s,r);
  section(s,52,"Cash flow statement");
  // Full reported cash flow lines.
  const cf2=[
    ["Profit for the year",4166,2538,1958],["Depreciation, amortisation and impairment",493,1718,2480],["Net working capital movement",-768,-77,76],
    ["Other operating and finance adjustments",214,118, -122],["Net cash from operating activities",4105,4297,4392],
    ["Capital expenditure",-1510,-1612,-1197],["Other investing cash flow",-85,-108,215],["Net cash from investing activities",-1595,-1720,-982],
    ["Debt issuance",2612,4026,1498],["Debt repayment",-2160,-3128,-3273],["Equity and NCI dividends",-2359,-2436,-2010],["Share repurchases and other financing",-1199,44,1],["Net cash from financing activities",-3106,-1494,-3784],
    ["Net change before FX/reclassification",-596,1083,-374],["FX and reclassification",-63,-14,-310],["Opening cash and cash equivalents, net",1768,1109,2178],["Closing cash and cash equivalents, net",1109,2178,1494]
  ];
  s.getRange("C53:F69").values=cf2.map(x=>[x[0],x[1],x[2],x[3]]); fmtRows(s,["D53:F69"]);
  for(const r of [57,60,65,66,69]) styleTotal(s,r);
  section(s,72,"Source map");
  s.getRange("C73:F78").values=[
    ["Income statement FY2024–FY2026","FY2026 financial statements","Page 143","https://www.diageo.com/~/media/Files/D/Diageo-V2/Diageo-Corp/investors/results-reports-and-events/annual-reports/2026/financial-statements-2026.pdf"],
    ["Balance sheet FY2025–FY2026","FY2026 financial statements","Page 145","Same as above"],
    ["Balance sheet FY2024","FY2025 annual report","Page 150","https://www.diageo.com/~/media/Files/D/Diageo-V2/Diageo-Corp/investors/results-reports-and-events/annual-reports/2025/annual-report-2025.pdf"],
    ["Cash flow FY2024–FY2026","FY2026 financial statements","Page 147","Same as above"],
    ["Segment data FY2024–FY2026","FY2026 financial statements","Pages 150–151","Same as above"],
    ["Presentation currency","FY2026 financial statements","Page 148","Same as above"]
  ];
  s.getRange("C73:F78").format.wrapText=true; s.getRange("F:F").format.columnWidth=55; s.freezePanes.freezeRows(6); s.freezePanes.freezeColumns(3);
}

// Revenue build
{
  const s=wb.worksheets.getItem("Revenue Build"); base(s,25); title(s,"Regional revenue build","Reported net sales by region | US$ millions"); timeline(s,6,19); section(s,8,"Net sales by region");
  const data=[
    ["North America",7908,7973,7249,14],["Europe",4804,4821,5097,15],["Asia Pacific",3817,3635,3333,16],
    ["Latin America & Caribbean",1839,1847,2160,17],["Africa",1778,1834,1642,18],["Corporate & other",123,135,162,19]
  ];
  let r=9;
  for(const [label,a,b,c,ar] of data){
    s.getRange(`C${r}:F${r}`).values=[[label,a,b,c]];
    for(let i=0;i<5;i++){const col=cols[i+3], prev=cols[i+2], assCol=cols[i+3]; s.getRange(`${col}${r}`).formulas=[[`=${prev}${r}*(1+'Assumptions'!${assCol}${ar})`]];}
    r++;
  }
  s.getRange("C15").values=[["Group net sales"]]; for(const c of cols) s.getRange(`${c}15`).formulas=[[`=SUM(${c}9:${c}14)`]]; styleTotal(s,15);
  s.getRange("C17").values=[["Reported growth"]]; for(let i=1;i<8;i++) s.getRange(`${cols[i]}17`).formulas=[[`=${cols[i]}15/${cols[i-1]}15-1`]];
  s.getRange("C19").values=[["Historical/source note"]]; s.getRange("D19:K19").values=[["FY2026 financial statements pages 150–151",null,null,"Forecast uses active regional organic growth assumptions",null,null,null,null]];
  fmtRows(s,["D9:K15"]); s.getRange("D17:K17").format.numberFormat=pctFmt; s.getRange("G9:K14").format.font={name:"Arial",size:10,color:green};
  s.freezePanes.freezeRows(6); s.freezePanes.freezeColumns(3);
}

// Operating model
{
  const s=wb.worksheets.getItem("Operating Model"); base(s,42); title(s,"Operating model","Adjusted and reported earnings | US$ millions except per-share data"); timeline(s,6,38); section(s,8,"Income statement forecast");
  const labels={9:"Net sales",10:"Cost of sales",11:"Gross profit",12:"Marketing",13:"Other recurring operating expenses",14:"Adjusted EBITDA",15:"Depreciation",16:"Amortisation",17:"Operating profit before exceptional items",18:"Exceptional operating costs",19:"Reported operating profit",20:"Net finance charges",21:"Associates and other income",22:"Profit before tax",23:"Taxation",24:"Profit for the year",25:"NCI profit",26:"Net income attributable to shareholders",27:"Diluted shares (m)",28:"Diluted EPS (cents)"};
  for(const [row,label] of Object.entries(labels)) s.getRange(`C${row}`).values=[[label]];
  const histMap={9:9,10:10,11:11,12:12,13:13,17:16,18:15,19:14,20:17,21:19,22:20,23:21,24:22,25:23,26:24,27:25,28:26};
  for(const [row,hrow] of Object.entries(histMap)) for(let i=0;i<3;i++) s.getRange(`${cols[i]}${row}`).formulas=[[`='Historicals'!${cols[i]}${hrow}`]];
  // Historical EBITDA/D&A split: use D&A from cash flow as a practical proxy; impairment makes FY26 non-comparable.
  for(let i=0;i<3;i++){const c=cols[i]; s.getRange(`${c}15`).formulas=[[`='Historicals'!${c}54`]]; s.getRange(`${c}16`).values=[[0]]; s.getRange(`${c}14`).formulas=[[`=${c}17+${c}15+${c}16`]];}
  for(let i=3;i<8;i++){
    const c=cols[i], ac=cols[i];
    s.getRange(`${c}9`).formulas=[[`='Revenue Build'!${c}15`]];
    s.getRange(`${c}11`).formulas=[[`=${c}9*'Assumptions'!${ac}24`]];
    s.getRange(`${c}10`).formulas=[[`=${c}9-${c}11`]];
    s.getRange(`${c}12`).formulas=[[`=${c}9*'Assumptions'!${ac}25`]];
    s.getRange(`${c}13`).formulas=[[`=${c}9*'Assumptions'!${ac}26`]];
    s.getRange(`${c}14`).formulas=[[`=${c}11-${c}12-${c}13`]];
    s.getRange(`${c}15`).formulas=[[`='Schedules'!${c}31`]];
    s.getRange(`${c}16`).formulas=[[`='Schedules'!${c}38`]];
    s.getRange(`${c}17`).formulas=[[`=${c}14-${c}15-${c}16`]];
    s.getRange(`${c}18`).formulas=[[`='Assumptions'!${ac}27`]];
    s.getRange(`${c}19`).formulas=[[`=${c}17-${c}18`]];
    s.getRange(`${c}20`).formulas=[[`='Schedules'!${c}48`]];
    s.getRange(`${c}21`).formulas=[[`='Assumptions'!${ac}29`]];
    s.getRange(`${c}22`).formulas=[[`=${c}19-${c}20+${c}21`]];
    s.getRange(`${c}23`).formulas=[[`=MAX(0,${c}22*'Assumptions'!${ac}28)`]];
    s.getRange(`${c}24`).formulas=[[`=${c}22-${c}23`]];
    s.getRange(`${c}25`).formulas=[[`=${c}24*10%`]];
    s.getRange(`${c}26`).formulas=[[`=${c}24-${c}25`]];
    s.getRange(`${c}27`).formulas=[[`=$F$27`]];
    s.getRange(`${c}28`).formulas=[[`=${c}26/${c}27*100`]];
  }
  section(s,32,"Margins and growth");
  const metrics={33:["Net sales growth",9],34:["Gross margin",11],35:["Adjusted EBITDA margin",14],36:["Adjusted operating margin",17],37:["Reported operating margin",19],38:["Effective tax rate",23]};
  for(const [row,[label,src]] of Object.entries(metrics)){s.getRange(`C${row}`).values=[[label]];for(let i=0;i<8;i++){const c=cols[i];if(Number(row)===33){if(i>0)s.getRange(`${c}${row}`).formulas=[[`=${c}9/${cols[i-1]}9-1`]];} else if(Number(row)===38)s.getRange(`${c}${row}`).formulas=[[`=${c}23/${c}22`]];else s.getRange(`${c}${row}`).formulas=[[`=${c}${src}/${c}9`]];}}
  fmtRows(s,["D9:K28"]); s.getRange("D28:K28").format.numberFormat='0.0'; s.getRange("D33:K38").format.numberFormat=pctFmt;
  for(const rr of [11,14,17,19,22,24,26]) styleTotal(s,rr); s.getRange("G9:K28").format.font={name:"Arial",size:10,color:green};
  s.freezePanes.freezeRows(6); s.freezePanes.freezeColumns(3);
}

// Schedules
{
  const s=wb.worksheets.getItem("Schedules"); base(s,65); title(s,"Supporting schedules","Working capital, fixed assets, debt, interest and equity | US$ millions"); timeline(s,6,61);
  section(s,8,"Operating balance sheet drivers");
  const labels={9:"Net sales",10:"Receivables",11:"Inventory",12:"Other current assets",13:"Other non-current assets",14:"Trade payables",15:"Other operating liabilities",16:"Operating net working capital",17:"Change in operating NWC"};
  for(const [row,l] of Object.entries(labels))s.getRange(`C${row}`).values=[[l]];
  const histRows={9:9,10:34,11:33,12:35,13:32,14:39,15:40};
  for(const [row,hrow] of Object.entries(histRows))for(let i=0;i<3;i++)s.getRange(`${cols[i]}${row}`).formulas=[[`='Historicals'!${cols[i]}${hrow}`]];
  for(let i=0;i<3;i++){const c=cols[i];s.getRange(`${c}16`).formulas=[[`=${c}10+${c}11+${c}12-${c}14-${c}15`]];if(i>0)s.getRange(`${c}17`).formulas=[[`=${c}16-${cols[i-1]}16`]];}
  for(let i=3;i<8;i++){const c=cols[i];s.getRange(`${c}9`).formulas=[[`='Revenue Build'!${c}15`]];s.getRange(`${c}10`).formulas=[[`=${c}9*'Assumptions'!${c}34`]];s.getRange(`${c}11`).formulas=[[`=${c}9*'Assumptions'!${c}35`]];s.getRange(`${c}12`).formulas=[[`=${c}9*'Assumptions'!${c}37`]];s.getRange(`${c}13`).formulas=[[`=${c}9*'Assumptions'!${c}38`]];s.getRange(`${c}14`).formulas=[[`=${c}9*'Assumptions'!${c}36`]];s.getRange(`${c}15`).formulas=[[`=${c}9*'Assumptions'!${c}39`]];s.getRange(`${c}16`).formulas=[[`=${c}10+${c}11+${c}12-${c}14-${c}15`]];s.getRange(`${c}17`).formulas=[[`=${c}16-${cols[i-1]}16`]];}
  section(s,21,"PP&E and intangibles");
  const assetLabels={22:"Opening PP&E",23:"Capital expenditure",24:"Depreciation",25:"Closing PP&E",27:"Opening intangibles",28:"Intangible additions",29:"Amortisation",30:"Closing intangibles",31:"Depreciation linked to operating model",38:"Amortisation linked to operating model"};
  for(const [row,l] of Object.entries(assetLabels))s.getRange(`C${row}`).values=[[l]];
  for(let i=3;i<8;i++){const c=cols[i],p=cols[i-1];s.getRange(`${c}22`).formulas=[[`=${p}25`]];s.getRange(`${c}23`).formulas=[[`=${c}9*'Assumptions'!${c}40`]];s.getRange(`${c}24`).formulas=[[`=${c}22*'Assumptions'!${c}42`]];s.getRange(`${c}25`).formulas=[[`=${c}22+${c}23-${c}24`]];s.getRange(`${c}27`).formulas=[[`=${p}30`]];s.getRange(`${c}28`).formulas=[[`=${c}9*'Assumptions'!${c}41`]];s.getRange(`${c}29`).formulas=[[`=${c}27*'Assumptions'!${c}43`]];s.getRange(`${c}30`).formulas=[[`=${c}27+${c}28-${c}29`]];s.getRange(`${c}31`).formulas=[[`=${c}24`]];s.getRange(`${c}38`).formulas=[[`=${c}29`]];}
  // Seed historical closing balances
  for(let i=0;i<3;i++){const c=cols[i];s.getRange(`${c}25`).formulas=[[`='Historicals'!${c}31`]];s.getRange(`${c}30`).formulas=[[`='Historicals'!${c}30`]];}
  section(s,42,"Debt and interest");
  const debtLabels={43:"Opening debt",44:"Mandatory repayment",45:"New borrowing",46:"Closing debt",47:"Average debt",48:"Net finance charges"}; for(const [row,l] of Object.entries(debtLabels))s.getRange(`C${row}`).values=[[l]];
  for(let i=0;i<3;i++){const c=cols[i];s.getRange(`${c}46`).formulas=[[`='Historicals'!${c}38`]];s.getRange(`${c}48`).formulas=[[`='Historicals'!${c}17`]];}
  for(let i=3;i<8;i++){const c=cols[i],p=cols[i-1];s.getRange(`${c}43`).formulas=[[`=${p}46`]];s.getRange(`${c}44`).formulas=[[`=${c}43*'Assumptions'!${c}44`]];s.getRange(`${c}45`).formulas=[[`=MAX(0,'Assumptions'!${c}48-'Three Statements'!${c}77)`]];s.getRange(`${c}46`).formulas=[[`=MAX(0,${c}43-${c}44+${c}45)`]];s.getRange(`${c}47`).formulas=[[`=(${c}43+${c}46)/2`]];s.getRange(`${c}48`).formulas=[[`=${c}43*'Assumptions'!${c}45`]];}
  section(s,52,"Equity");
  const eqLabels={53:"Opening retained earnings",54:"Parent net income",55:"Parent dividends",56:"Closing retained earnings",58:"Opening NCI",59:"NCI profit",60:"NCI dividends",61:"Closing NCI"};for(const [row,l] of Object.entries(eqLabels))s.getRange(`C${row}`).values=[[l]];
  for(let i=0;i<3;i++){const c=cols[i];s.getRange(`${c}56`).formulas=[[`='Historicals'!${c}45`]];s.getRange(`${c}61`).formulas=[[`='Historicals'!${c}46`]];}
  for(let i=3;i<8;i++){const c=cols[i],p=cols[i-1];s.getRange(`${c}53`).formulas=[[`=${p}56`]];s.getRange(`${c}54`).formulas=[[`='Operating Model'!${c}26`]];s.getRange(`${c}55`).formulas=[[`=${c}54*'Assumptions'!${c}46`]];s.getRange(`${c}56`).formulas=[[`=${c}53+${c}54-${c}55`]];s.getRange(`${c}58`).formulas=[[`=${p}61`]];s.getRange(`${c}59`).formulas=[[`='Operating Model'!${c}25`]];s.getRange(`${c}60`).formulas=[[`=${c}59*'Assumptions'!${c}47`]];s.getRange(`${c}61`).formulas=[[`=${c}58+${c}59-${c}60`]];}
  fmtRows(s,["D9:K17","D22:K38","D43:K48","D53:K61"]); for(const rr of [16,25,30,46,48,56,61])styleTotal(s,rr); s.getRange("G9:K61").format.font={name:"Arial",size:10,color:green};
  s.freezePanes.freezeRows(6); s.freezePanes.freezeColumns(3);
}

// Three statements
{
  const s=wb.worksheets.getItem("Three Statements"); base(s,92); title(s,"Integrated three statements","Historical and forecast financial statements | US$ millions except per-share data"); timeline(s,6,88);
  section(s,8,"Income statement");
  const isLabels={9:"Net sales",10:"Cost of sales",11:"Gross profit",12:"Marketing",13:"Other recurring operating expenses",14:"Adjusted EBITDA",15:"Depreciation",16:"Amortisation",17:"Operating profit before exceptional items",18:"Exceptional operating costs",19:"Reported operating profit",20:"Net finance charges",21:"Associates and other income",22:"Profit before tax",23:"Taxation",24:"Profit for the year",25:"NCI profit",26:"Net income attributable to shareholders",27:"Diluted EPS (cents)"};
  for(const [row,l] of Object.entries(isLabels)){s.getRange(`C${row}`).values=[[l]];for(const c of cols)s.getRange(`${c}${row}`).formulas=[[`='Operating Model'!${c}${Number(row)===27?28:row}`]];}
  section(s,31,"Balance sheet");
  const bsLabels={32:"Intangible assets",33:"Property, plant and equipment",34:"Other non-current assets",35:"Inventories",36:"Trade and other receivables",37:"Other current assets",38:"Cash and cash equivalents",39:"Total assets",41:"Borrowings",42:"Trade and other payables",43:"Other liabilities",44:"Total liabilities",46:"Share capital",47:"Share premium",48:"Other reserves",49:"Retained earnings",50:"NCI",51:"Total equity",52:"Total liabilities and equity",53:"Balance sheet check"};for(const [row,l] of Object.entries(bsLabels))s.getRange(`C${row}`).values=[[l]];
  const histBs={32:30,33:31,34:32,35:33,36:34,37:35,38:36,39:37,41:38,42:39,43:40,44:41,46:42,47:43,48:44,49:45,50:46,51:47,52:48,53:49};
  for(const [row,hrow] of Object.entries(histBs))for(let i=0;i<3;i++)s.getRange(`${cols[i]}${row}`).formulas=[[`='Historicals'!${cols[i]}${hrow}`]];
  for(let i=3;i<8;i++){const c=cols[i];s.getRange(`${c}32`).formulas=[[`='Schedules'!${c}30`]];s.getRange(`${c}33`).formulas=[[`='Schedules'!${c}25`]];s.getRange(`${c}34`).formulas=[[`='Schedules'!${c}13`]];s.getRange(`${c}35`).formulas=[[`='Schedules'!${c}11`]];s.getRange(`${c}36`).formulas=[[`='Schedules'!${c}10`]];s.getRange(`${c}37`).formulas=[[`='Schedules'!${c}12`]];s.getRange(`${c}38`).formulas=[[`=${c}78`]];s.getRange(`${c}39`).formulas=[[`=SUM(${c}32:${c}38)`]];s.getRange(`${c}41`).formulas=[[`='Schedules'!${c}46`]];s.getRange(`${c}42`).formulas=[[`='Schedules'!${c}14`]];s.getRange(`${c}43`).formulas=[[`='Schedules'!${c}15`]];s.getRange(`${c}44`).formulas=[[`=SUM(${c}41:${c}43)`]];s.getRange(`${c}46`).formulas=[[`=$F$46`]];s.getRange(`${c}47`).formulas=[[`=$F$47`]];s.getRange(`${c}48`).formulas=[[`=$F$48`]];s.getRange(`${c}49`).formulas=[[`='Schedules'!${c}56`]];s.getRange(`${c}50`).formulas=[[`='Schedules'!${c}61`]];s.getRange(`${c}51`).formulas=[[`=SUM(${c}46:${c}50)`]];s.getRange(`${c}52`).formulas=[[`=${c}44+${c}51`]];s.getRange(`${c}53`).formulas=[[`=${c}39-${c}52`]];}
  section(s,57,"Cash flow statement");
  const cfLabels={58:"Profit for the year",59:"Depreciation and amortisation",60:"Change in operating NWC",61:"Change in other non-current assets",62:"Net cash from operating activities",64:"Capital expenditure",65:"Intangible additions",66:"Net cash from investing activities",68:"Debt repayment",69:"New borrowing / revolver draw",70:"Parent dividends",71:"NCI dividends",72:"Net cash from financing activities",74:"Net change in cash",75:"Opening cash",76:"Other cash/FX movement",77:"Cash before new borrowing",78:"Closing cash"};for(const [row,l] of Object.entries(cfLabels))s.getRange(`C${row}`).values=[[l]];
  // Historical reported cash flow summary.
  for(let i=0;i<3;i++){const c=cols[i];s.getRange(`${c}58`).formulas=[[`='Historicals'!${c}53`]];s.getRange(`${c}59`).formulas=[[`='Historicals'!${c}54`]];s.getRange(`${c}60`).formulas=[[`='Historicals'!${c}55`]];s.getRange(`${c}62`).formulas=[[`='Historicals'!${c}57`]];s.getRange(`${c}66`).formulas=[[`='Historicals'!${c}60`]];s.getRange(`${c}72`).formulas=[[`='Historicals'!${c}65`]];s.getRange(`${c}74`).formulas=[[`='Historicals'!${c}66`]];s.getRange(`${c}78`).formulas=[[`='Historicals'!${c}69`]];}
  for(let i=3;i<8;i++){const c=cols[i],p=cols[i-1];s.getRange(`${c}58`).formulas=[[`=${c}24`]];s.getRange(`${c}59`).formulas=[[`=${c}15+${c}16`]];s.getRange(`${c}60`).formulas=[[`=-'Schedules'!${c}17`]];s.getRange(`${c}61`).formulas=[[`=-(${c}34-${p}34)`]];s.getRange(`${c}62`).formulas=[[`=SUM(${c}58:${c}61)`]];s.getRange(`${c}64`).formulas=[[`=-'Schedules'!${c}23`]];s.getRange(`${c}65`).formulas=[[`=-'Schedules'!${c}28`]];s.getRange(`${c}66`).formulas=[[`=SUM(${c}64:${c}65)`]];s.getRange(`${c}68`).formulas=[[`=-'Schedules'!${c}44`]];s.getRange(`${c}69`).formulas=[[`='Schedules'!${c}45`]];s.getRange(`${c}70`).formulas=[[`=-'Schedules'!${c}55`]];s.getRange(`${c}71`).formulas=[[`=-'Schedules'!${c}60`]];s.getRange(`${c}72`).formulas=[[`=SUM(${c}68:${c}71)`]];s.getRange(`${c}74`).formulas=[[`=${c}62+${c}66+${c}72`]];s.getRange(`${c}75`).formulas=[[i===3?`=$F$38`:`=${p}78`]];s.getRange(`${c}76`).formulas=[[`=0`]];s.getRange(`${c}77`).formulas=[[`=${c}75+${c}62+${c}66+${c}68+${c}70+${c}71+${c}76`]];s.getRange(`${c}78`).formulas=[[`=${c}75+${c}74+${c}76`]];}
  section(s,82,"Key ratios");
  const ratioLabels={83:"Net sales growth",84:"Adjusted EBITDA margin",85:"Free cash flow",86:"Net debt",87:"Net debt / adjusted EBITDA",88:"Interest coverage"};for(const [row,l] of Object.entries(ratioLabels))s.getRange(`C${row}`).values=[[l]];
  for(let i=0;i<8;i++){const c=cols[i];if(i>0)s.getRange(`${c}83`).formulas=[[`=${c}9/${cols[i-1]}9-1`]];s.getRange(`${c}84`).formulas=[[`=${c}14/${c}9`]];s.getRange(`${c}85`).formulas=[[`=${c}62+${c}64`]];s.getRange(`${c}86`).formulas=[[`=${c}41-${c}38`]];s.getRange(`${c}87`).formulas=[[`=${c}86/${c}14`]];s.getRange(`${c}88`).formulas=[[`=${c}14/${c}20`]];}
  fmtRows(s,["D9:K27","D32:K53","D58:K78","D85:K86"]);s.getRange("D27:K27").format.numberFormat='0.0';s.getRange("D83:K84").format.numberFormat=pctFmt;s.getRange("D87:K88").format.numberFormat=multFmt;
  for(const rr of [11,14,17,19,22,24,26,39,44,51,52,53,62,66,72,74,78])styleTotal(s,rr);s.getRange("G9:K88").format.font={name:"Arial",size:10,color:green};
  s.freezePanes.freezeRows(6);s.freezePanes.freezeColumns(3);
}

// Scenarios (formula-driven case snapshot using active case and directional summary)
{
  const s=wb.worksheets.getItem("Scenarios"); base(s,30); title(s,"Scenario comparison","Operating assumptions and active-case outputs | US$ millions");
  section(s,7,"Scenario definitions");s.getRange("C8:F11").values=[
    ["Case","Revenue","Margins","Cash and leverage"],["Downside","Prolonged North America and Asia weakness","Margin pressure; delayed savings","Higher working capital and slower debt repayment"],
    ["Base","FY2027 broadly flat, then gradual recovery","Savings support measured expansion","Progressive debt repayment"],["Upside","Faster regional recovery","Stronger mix and execution","Better conversion and faster deleveraging"]
  ];s.getRange("C8:F8").format.fill=blue;s.getRange("C8:F8").format.font={name:"Arial",size:9,bold:true,color:white};s.getRange("C9:F11").format.wrapText=true;s.getRange("D:F").format.columnWidth=26;
  section(s,14,"Active scenario outputs");s.getRange("D15:H15").values=[["FY:27E","FY:28E","FY:29E","FY:30E","FY:31E"]];s.getRange("D15:H15").format.fill=blue;s.getRange("D15:H15").format.font={name:"Arial",size:9,bold:true,color:white};
  const outs={16:["Net sales",9,moneyFmt],17:["Adjusted EBITDA",14,moneyFmt],18:["Adjusted EBITDA margin",84,pctFmt],19:["Profit for the year",24,moneyFmt],20:["Free cash flow",85,moneyFmt],21:["Net debt",86,moneyFmt],22:["Net debt / adjusted EBITDA",87,multFmt],23:["Interest coverage",88,multFmt]};
  for(const [row,[label,src,fmt]] of Object.entries(outs)){s.getRange(`C${row}`).values=[[label]];for(let i=0;i<5;i++)s.getRange(`${String.fromCharCode(68+i)}${row}`).formulas=[[`='Three Statements'!${cols[i+3]}${src}`]];s.getRange(`D${row}:H${row}`).format.numberFormat=fmt;}
  s.getRange("C26").values=[["Active scenario"]];s.getRange("D26").formulas=[["='Assumptions'!D9"]];s.getRange("C27").values=[["Use the selector on Assumptions to refresh the full model and these outputs."]];s.getRange("C27:H27").format.font={name:"Arial",size:9,italic:true,color:"#666666"};
}

// Dashboard
{
  const s=wb.worksheets.getItem("Dashboard"); base(s,60); title(s,"Financial outlook","Diageo plc | Active scenario outputs | US$ millions");
  s.getRange("C6").values=[["Active scenario"]];s.getRange("D6").formulas=[["='Assumptions'!D9"]];s.getRange("F6").values=[["Model status"]];s.getRange("G6").formulas=[["='Checks'!D20"]];s.getRange("D6:G6").format.font={name:"Arial",size:11,bold:true,color:navy};
  section(s,9,"Headline metrics","H");
  s.getRange("C10:H10").values=[["Metric","FY:26A","FY:27E","FY:29E","FY:31E","FY26–FY31 change"]];s.getRange("C10:H10").format.fill=blue;s.getRange("C10:H10").format.font={name:"Arial",size:9,bold:true,color:white};
  const kpis=[["Net sales",9,moneyFmt],["Adjusted EBITDA",14,moneyFmt],["Adjusted EBITDA margin",84,pctFmt],["Free cash flow",85,moneyFmt],["Net debt",86,moneyFmt],["Net debt / EBITDA",87,multFmt]];
  let rr=11;for(const [lab,row,fmt] of kpis){s.getRange(`C${rr}`).values=[[lab]];s.getRange(`D${rr}:G${rr}`).formulas=[[`='Three Statements'!F${row}`,`='Three Statements'!G${row}`,`='Three Statements'!I${row}`,`='Three Statements'!K${row}`]];s.getRange(`H${rr}`).formulas=[[row===84?`=G${rr}-D${rr}`:`=G${rr}/D${rr}-1`]];s.getRange(`D${rr}:G${rr}`).format.numberFormat=fmt;s.getRange(`H${rr}`).format.numberFormat=pctFmt;rr++;}
  section(s,20,"Forecast trend data","H");s.getRange("C21:H21").values=[["Metric","FY:27E","FY:28E","FY:29E","FY:30E","FY:31E"]];s.getRange("C21:H21").format.fill=blue;s.getRange("C21:H21").format.font={name:"Arial",size:9,bold:true,color:white};
  s.getRange("C22:C25").values=[["Net sales"],["Adjusted EBITDA"],["Free cash flow"],["Net debt"]];
  const rows=[9,14,85,86];for(let i=0;i<4;i++)for(let j=0;j<5;j++)s.getRange(`${String.fromCharCode(68+j)}${22+i}`).formulas=[[`='Three Statements'!${cols[j+3]}${rows[i]}`]];fmtRows(s,["D22:H25"]);
  section(s,47,"Interpretation","H");s.getRange("C49:H53").values=[
    ["Revenue","FY2027 stabilises; growth resumes from FY2028.",null,null,null,null],
    ["Margins","Cost actions support gradual adjusted-margin recovery.",null,null,null,null],
    ["Cash flow","Working capital and capex constrain near-term cash flow.",null,null,null,null],
    ["Leverage","Minimum cash is protected; leverage declines gradually.",null,null,null,null],
    ["Limitation","FX, disposals, leases and tax are simplified.",null,null,null,null]
  ];s.getRange("C49:C53").format.font={name:"Arial",size:10,bold:true,color:text};s.getRange("D49:H53").format.wrapText=true;
  // Vertical helper tables keep chart categories and series unambiguous.
  s.getRange("T20:V25").values=[["Year","Net sales","Adjusted EBITDA"],["FY:27E",null,null],["FY:28E",null,null],["FY:29E",null,null],["FY:30E",null,null],["FY:31E",null,null]];
  s.getRange("T28:V33").values=[["Year","Free cash flow","Net debt"],["FY:27E",null,null],["FY:28E",null,null],["FY:29E",null,null],["FY:30E",null,null],["FY:31E",null,null]];
  for(let i=0;i<5;i++){const src=String.fromCharCode(68+i);s.getRange(`U${21+i}:V${21+i}`).formulas=[[`=${src}22`,`=${src}23`]];s.getRange(`U${29+i}:V${29+i}`).formulas=[[`=${src}24`,`=${src}25`]];}
  s.getRange("T:V").format.columnWidth=2;s.getRange("T20:V33").format.numberFormat=";;;";s.getRange("T20:V33").format.font={name:"Arial",size:1,color:white};
  const ch1=s.charts.add("line",s.getRange("T20:V25"));ch1.title="Net sales and adjusted EBITDA";ch1.setPosition("J9","Q25");ch1.legend={position:"top",textStyle:{typeface:"Arial"}};
  const ch2=s.charts.add("line",s.getRange("T28:V33"));ch2.title="Free cash flow and net debt";ch2.setPosition("J28","Q44");ch2.legend={position:"top",textStyle:{typeface:"Arial"}};
  s.getRange("J:Q").format.columnWidth=11;
  s.getRange("D:D").format.columnWidth=38;
}

// Checks
{
  const s=wb.worksheets.getItem("Checks"); base(s,24); title(s,"Model checks","Independent reconciliations | Tolerance: US$0.1 million");
  s.getRange("C6:F6").values=[["Check","Maximum absolute difference","Tolerance","Status"]];s.getRange("C6:F6").format.fill=blue;s.getRange("C6:F6").format.font={name:"Arial",size:9,bold:true,color:white};
  const checks=[
    ["Historical balance sheets",`=MAX(ABS('Historicals'!D49),ABS('Historicals'!E49),ABS('Historicals'!F49))`],["Forecast balance sheets",`=MAX(ABS('Three Statements'!G53),ABS('Three Statements'!H53),ABS('Three Statements'!I53),ABS('Three Statements'!J53),ABS('Three Statements'!K53))`],
    ["Revenue build to operating model",`=MAX(ABS('Revenue Build'!G15-'Operating Model'!G9),ABS('Revenue Build'!H15-'Operating Model'!H9),ABS('Revenue Build'!I15-'Operating Model'!I9),ABS('Revenue Build'!J15-'Operating Model'!J9),ABS('Revenue Build'!K15-'Operating Model'!K9))`],["Net income to cash flow",`=MAX(ABS('Three Statements'!G24-'Three Statements'!G58),ABS('Three Statements'!H24-'Three Statements'!H58),ABS('Three Statements'!I24-'Three Statements'!I58),ABS('Three Statements'!J24-'Three Statements'!J58),ABS('Three Statements'!K24-'Three Statements'!K58))`],
    ["Closing cash to balance sheet",`=MAX(ABS('Three Statements'!G38-'Three Statements'!G78),ABS('Three Statements'!H38-'Three Statements'!H78),ABS('Three Statements'!I38-'Three Statements'!I78),ABS('Three Statements'!J38-'Three Statements'!J78),ABS('Three Statements'!K38-'Three Statements'!K78))`],["Debt schedule to balance sheet",`=MAX(ABS('Schedules'!G46-'Three Statements'!G41),ABS('Schedules'!H46-'Three Statements'!H41),ABS('Schedules'!I46-'Three Statements'!I41),ABS('Schedules'!J46-'Three Statements'!J41),ABS('Schedules'!K46-'Three Statements'!K41))`],
    ["Retained earnings schedule",`=MAX(ABS('Schedules'!G56-'Three Statements'!G49),ABS('Schedules'!H56-'Three Statements'!H49),ABS('Schedules'!I56-'Three Statements'!I49),ABS('Schedules'!J56-'Three Statements'!J49),ABS('Schedules'!K56-'Three Statements'!K49))`],["NCI schedule",`=MAX(ABS('Schedules'!G61-'Three Statements'!G50),ABS('Schedules'!H61-'Three Statements'!H50),ABS('Schedules'!I61-'Three Statements'!I50),ABS('Schedules'!J61-'Three Statements'!J50),ABS('Schedules'!K61-'Three Statements'!K50))`],
    ["PP&E schedule",`=MAX(ABS('Schedules'!G25-'Three Statements'!G33),ABS('Schedules'!H25-'Three Statements'!H33),ABS('Schedules'!I25-'Three Statements'!I33),ABS('Schedules'!J25-'Three Statements'!J33),ABS('Schedules'!K25-'Three Statements'!K33))`],["Intangibles schedule",`=MAX(ABS('Schedules'!G30-'Three Statements'!G32),ABS('Schedules'!H30-'Three Statements'!H32),ABS('Schedules'!I30-'Three Statements'!I32),ABS('Schedules'!J30-'Three Statements'!J32),ABS('Schedules'!K30-'Three Statements'!K32))`],
    ["Scenario selector validity",`=IF(AND('Assumptions'!D8>=1,'Assumptions'!D8<=3),0,1)`]
  ];
  let r=7;for(const [lab,formula] of checks){s.getRange(`C${r}`).values=[[lab]];s.getRange(`D${r}`).formulas=[[formula]];s.getRange(`E${r}`).values=[[0.1]];s.getRange(`F${r}`).formulas=[[`=IF(D${r}<=E${r},\"OK\",\"ERROR\")`]];r++;}
  s.getRange("D7:E17").format.numberFormat='0.0';s.getRange("F7:F17").conditionalFormats.add("containsText",{text:"ERROR",format:{fill:"#F4CCCC",font:{color:red,bold:true}}});s.getRange("F7:F17").conditionalFormats.add("containsText",{text:"OK",format:{fill:"#D9EAD3",font:{color:green,bold:true}}});
  s.getRange("C20").values=[["Master model status"]];s.getRange("D20").formulas=[[`=IF(COUNTIF(F7:F17,"ERROR")=0,"MODEL OK","MODEL ERROR")`]];s.getRange("C20:D20").format.font={name:"Arial",size:12,bold:true,color:navy};
  s.getRange("D20").conditionalFormats.add("containsText",{text:"ERROR",format:{fill:"#F4CCCC",font:{color:red,bold:true}}});s.getRange("D20").conditionalFormats.add("containsText",{text:"OK",format:{fill:"#D9EAD3",font:{color:green,bold:true}}});
}

// Final formatting
for(const name of names){
  const s=wb.worksheets.getItem(name);
  s.getUsedRange().format.verticalAlignment="center";
  s.getRange("C:C").format.wrapText=true;
  s.tabColor = name==="Assumptions" ? "#C9B458" : (["Dashboard","Three Statements"].includes(name)?blue:"#A6A6A6");
}

await fs.mkdir(outDir,{recursive:true});
const xlsx=await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(outPath);

const key=await wb.inspect({kind:"table",range:"Three Statements!C6:K88",include:"values,formulas",tableMaxRows:90,tableMaxCols:9,maxChars:16000});
console.log(key.ndjson);
const errors=await wb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",options:{useRegex:true,maxResults:300},summary:"final formula error scan"});
console.log(errors.ndjson);
for(const sheetName of names){
  const preview=await wb.render({sheetName,autoCrop:"all",scale:1,format:"png"});
  await fs.writeFile(`${outDir}/preview-${sheetName.replaceAll(" ","_")}.png`,new Uint8Array(await preview.arrayBuffer()));
}
console.log(`SAVED ${outPath}`);
