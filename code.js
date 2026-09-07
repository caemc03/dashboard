function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('System Report Dashboard V4.2')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDashboardData(filters) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Copy of ใช้อันนี้");
    if (!sheet) sheet = ss.getActiveSheet() || ss.getSheets()[0];
    if (!sheet) return "Error: ไม่พบชีตข้อมูลในไฟล์ Google Sheets นี้";

    const data = sheet.getDataRange().getValues();
    if (data.length < 3) return "Error: ข้อมูลในชีตน้อยเกินไป (ต้องมีอย่างน้อย 3 แถว)";
    
    const rows = data.slice(2); 

    let stats = {
      totalLocations: 0,
      foundGuilty: 0,
      actCounts: { "พรบ.ยา": 0, "พรบ.อาหาร": 0, "พรบ.เครื่องสำอาง": 0, "พรบ.สมุนไพร": 0, "พรบ.เครื่องมือแพทย์": 0, "พรบ.วัตถุอันตราย": 0 },
      actSections: { "พรบ.ยา": {}, "พรบ.อาหาร": {}, "พรบ.เครื่องสำอาง": {}, "พรบ.สมุนไพร": {}, "พรบ.เครื่องมือแพทย์": {}, "พรบ.วัตถุอันตราย": {} },
      locationTypes: {}, appearances: {}, riskTypes: {},
      provinces: {}, districts: {}, hotspots: {},
      hotspotDetails: [], 
      caseStatus: { finished: 0, prosecuted: 0, investigating: 0, oyaFine: 0 }, 
      monthlyCounts: {},
      items: { seizedList: 0, seizedUnits: 0, frozenList: 0, frozenUnits: 0, totalList: 0, totalUnits: 0 },
      totalValue: 0
    };

    let startFilter = filters && filters.start ? new Date(filters.start).getTime() : null;
    let endFilter = filters && filters.end ? new Date(filters.end).setHours(23,59,59,999) : null;

    const parseNum = (v) => {
      if (!v) return 0;
      let n = parseFloat(v.toString().replace(/,/g, ''));
      return isNaN(n) ? 0 : n;
    };

    // ฟังก์ชันช่วยนับมาตราความผิด
    const countSections = (actName, rawText) => {
      if (!rawText || rawText.toString().trim() === "" || rawText.toString().trim() === "-") return;
      let sections = rawText.toString().split(/,|\n/);
      sections.forEach(s => {
        let sTrim = s.trim();
        if (sTrim) stats.actSections[actName][sTrim] = (stats.actSections[actName][sTrim] || 0) + 1;
      });
    };

    rows.forEach((row) => {
      let rawDate = row[5]; 
      if (!rawDate) return;
      let rowDate = new Date(rawDate);
      if (rowDate.getFullYear() > 2500) rowDate.setFullYear(rowDate.getFullYear() - 543);
      if (isNaN(rowDate.getTime())) return;

      const rowTime = rowDate.getTime();
      if (startFilter && rowTime < startFilter) return;
      if (endFilter && rowTime > endFilter) return;

      if (row[0] && row[0].toString().trim() !== "") {
        stats.totalLocations++;
      }

      let isGuilty = false;
      let rowActs = [];
      
      // ตรวจสอบพบ พรบ. และเก็บข้อมูลมาตรา (สมมติว่ามาตราอยู่คอลัมน์ถัดไป +1)
      if (row[43] === "พบ") { stats.actCounts["พรบ.ยา"]++; countSections("พรบ.ยา", row[44]); isGuilty = true; rowActs.push("พรบ.ยา"); }
      if (row[45] === "พบ") { stats.actCounts["พรบ.อาหาร"]++; countSections("พรบ.อาหาร", row[46]); isGuilty = true; rowActs.push("พรบ.อาหาร"); }
      if (row[47] === "พบ") { stats.actCounts["พรบ.วัตถุอันตราย"]++; countSections("พรบ.วัตถุอันตราย", row[48]); isGuilty = true; rowActs.push("พรบ.วัตถุอันตราย"); }
      if (row[49] === "พบ") { stats.actCounts["พรบ.เครื่องมือแพทย์"]++; countSections("พรบ.เครื่องมือแพทย์", row[50]); isGuilty = true; rowActs.push("พรบ.เครื่องมือแพทย์"); }
      if (row[51] === "พบ") { stats.actCounts["พรบ.เครื่องสำอาง"]++; countSections("พรบ.เครื่องสำอาง", row[52]); isGuilty = true; rowActs.push("พรบ.เครื่องสำอาง"); }
      if (row[53] === "พบ") { stats.actCounts["พรบ.สมุนไพร"]++; countSections("พรบ.สมุนไพร", row[54]); isGuilty = true; rowActs.push("พรบ.สมุนไพร"); }

      if (isGuilty || (row[42] && row[42].toString().includes("พบการกระทำผิด"))) {
        stats.foundGuilty++;
      }

      let caseResultText = row[58] ? row[58].toString().trim() : "";
      if (caseResultText.includes("สิ้นสุด") || caseResultText.includes("เปรียบเทียบปรับ")) stats.caseStatus.finished++;
      if (caseResultText.includes("ส่งฟ้อง")|| caseResultText.includes("ผลคดีชั้นศาล")) stats.caseStatus.prosecuted++;
      if (caseResultText.includes("ระหว่างสอบ") || caseResultText.includes("สอบสวน")) stats.caseStatus.investigating++;
      if (caseResultText.includes("อย.") || caseResultText.includes("เปรียบเทียบปรับ อย")) stats.caseStatus.oyaFine++;

      let prov = row[24] ? row[24].toString().trim() : "";
      let dist = row[23] ? row[23].toString().trim() : "";
      let riskRaw = row[30] ? row[30].toString().trim() : "ไม่ระบุความเสี่ยง";

      if (prov && prov !== "-") {
        stats.provinces[prov] = (stats.provinces[prov] || 0) + 1;
        if (dist && dist !== "-") {
          let hotspotKey = prov + " (เขต/อำเภอ " + dist + ")";
          stats.hotspots[hotspotKey] = (stats.hotspots[hotspotKey] || 0) + 1;
          stats.hotspotDetails.push({ hotspot: hotspotKey, province: prov, district: dist, act: rowActs.length > 0 ? rowActs.join(", ") : "พรบ.ยา/ทั่วไป", risk: riskRaw !== "" ? riskRaw : "ฝ่าฝืนกฎหมาย" });
        }
      }

      stats.items.seizedList += parseNum(row[31]); stats.items.seizedUnits += parseNum(row[32]);
      stats.items.frozenList += parseNum(row[34]); stats.items.frozenUnits += parseNum(row[35]);
      stats.totalValue += parseNum(row[39]);
      
      let loc = row[13] ? row[13].toString().trim() : "";
      if (loc && loc !== "-") stats.locationTypes[loc] = (stats.locationTypes[loc] || 0) + 1;
      
      if (riskRaw && riskRaw !== "-") {
        riskRaw.split(',').forEach(r => { let item = r.trim(); if (item) stats.riskTypes[item] = (stats.riskTypes[item] || 0) + 1; });
      }
    });

    return stats;
  } catch (e) { return "Error: " + e.toString(); }
}

function generateInsight(stats) {
  const result = {
    summary: "",
    topHotspotsList: [],
    topProvincesList: [],
    topRisksList: [],
    topLocationTypesList: [], // เพิ่ม Top 5 ประเภทสถานที่
    topSectionsPerAct: {},    // เพิ่ม Top 5 มาตราของแต่ละพรบ.
    actBreakdown: [],
    caseStatusSummary: {},
    itemsSummary: {}
  };

  if (!stats) return result;

  const getTop = (obj) => Object.entries(obj).reduce((a, b) => a[1] > b[1] ? a : b, ["ไม่ระบุ", 0]);
  const getTopN = (obj, n) => Object.entries(obj).sort((a, b) => b[1] - a[1]).slice(0, n);
  
  let [topAct, topActCount] = getTop(stats.actCounts);
  let [topLocation, topLocationCount] = getTop(stats.locationTypes);
  let [topRisk, topRiskCount] = getTop(stats.riskTypes);
  let [topProvince, topProvinceCount] = getTop(stats.provinces);

  // 1. Hotspots
  let sortedHotspots = Object.entries(stats.hotspots).sort((a, b) => b[1] - a[1]).slice(0, 5);
  sortedHotspots.forEach(([hsKey, count], index) => {
    let matchedDetail = stats.hotspotDetails.find(d => d.hotspot === hsKey);
    result.topHotspotsList.push({ rank: index + 1, location: hsKey, cases: count, act: matchedDetail ? matchedDetail.act : topAct, risk: matchedDetail ? matchedDetail.risk : topRisk });
  });

  // 2. Top 5 Lists
  result.topProvincesList = getTopN(stats.provinces, 5).map((item, index) => ({ rank: index + 1, province: item[0], cases: item[1] }));
  result.topRisksList = getTopN(stats.riskTypes, 5).map((item, index) => ({ rank: index + 1, risk: item[0], cases: item[1] }));
  result.topLocationTypesList = getTopN(stats.locationTypes, 5).map((item, index) => ({ rank: index + 1, locationType: item[0], cases: item[1] })); // ประเภทสถานที่

  // 3. Top 5 Sections Per Act (มาตราความผิด)
  for (let act in stats.actSections) {
    result.topSectionsPerAct[act] = getTopN(stats.actSections[act], 5).map((item, index) => ({
      rank: index + 1,
      section: item[0],
      cases: item[1]
    }));
  }

  // 4. Summaries
  result.actBreakdown = Object.entries(stats.actCounts).sort((a, b) => b[1] - a[1]).map(item => ({ act: item[0], cases: item[1] }));
  result.caseStatusSummary = { finished: stats.caseStatus.finished, oyaFine: stats.caseStatus.oyaFine, prosecuted: stats.caseStatus.prosecuted, investigating: stats.caseStatus.investigating };
  result.itemsSummary = { seizedList: stats.items.seizedList, seizedUnits: stats.items.seizedUnits, frozenList: stats.items.frozenList, frozenUnits: stats.items.frozenUnits, totalValue: stats.totalValue };

  result.summary = `สรุปข้อมูลเชิงสถิติ: เข้าตรวจสอบ <b>${stats.totalLocations} แห่ง</b> (พบการกระทำความผิด <b>${stats.foundGuilty} ครั้ง</b>) ` +
  `จังหวัดที่ลงพื้นที่สูงสุดคือ <b>${topProvince}</b> กฎหมายที่พบฝ่าฝืนมากที่สุดคือ <b>${topAct}</b> ` +
  `ประเภทสถานที่ที่พบมากที่สุดคือ <b>"${topLocation}"</b> และประเด็นความเสี่ยงหลักคือ <b>"${topRisk}"</b>`;

  return result;
}