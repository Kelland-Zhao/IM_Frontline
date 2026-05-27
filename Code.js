/***********************
 * 通用工具：配置、并发保护、内存缓存
 ***********************/

/**
 * 获取 Script Properties 中配置的环境变量
 * @param {string} key - 属性键名
 * @param {string|number} defVal - 属性不存在时的默认返回值
 * @returns {string|number} 找到的属性值或默认值
 */
function getProp(key, defVal) {
  try {
    var v = PropertiesService.getScriptProperties().getProperty(key);
    return (v == null || v === '') ? defVal : v;
  } catch (e) {
    return defVal;
  }
}

/**
 * 并发保护锁，防止多人同时写入导致的数据覆盖或行错乱
 * @param {Function} fn - 拿到锁后需要执行的回调函数
 * @param {number} [timeoutMs=30000] - 最长等待时间（毫秒）
 * @returns {any} 回调函数的返回值
 * @throws {Error} 若等待超时未拿到锁，抛出系统繁忙错误
 */
function withLock(fn, timeoutMs) {
  var lock = LockService.getScriptLock();
  var wait = Math.max(1000, Number(timeoutMs || 30000));
  if (!lock.tryLock(wait)) {
    throw new Error("系统繁忙，请稍后再试（获取锁超时）");
  }
  try { 
    return fn(); 
  } finally { 
    lock.releaseLock();
  }
}

/**
 * 安全地向表格末尾追加一行数据（带并发锁）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} ss - 目标工作表对象
 * @param {Array<any>} rowArr - 要写入的一维数组
 * @returns {boolean} 写入成功返回 true
 */
function safeAppendRow(ss, rowArr) { 
  return withLock(function(){ 
    ss.appendRow(rowArr); 
    return true; 
  });
}

/**
 * 安全地向表格末尾追加多行数据（带并发锁）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} ss - 目标工作表对象
 * @param {Array<Array<any>>} rows2d - 要写入的二维数组
 * @returns {boolean} 写入成功返回 true
 */
function safeAppendRows(ss, rows2d) { 
  return withLock(function(){ 
    ss.getRange(ss.getLastRow() + 1, 1, rows2d.length, rows2d[0].length).setValues(rows2d); 
    return true; 
  });
}

/**
 * 安全地向指定区域覆盖写入二维数据（带并发锁）
 * @param {GoogleAppsScript.Spreadsheet.Range} range - 目标写入区域
 * @param {Array<Array<any>>} values2d - 要写入的二维数组
 * @returns {boolean} 写入成功返回 true
 */
function safeSetValues(range, values2d) { 
  return withLock(function(){ 
    range.setValues(values2d); 
    return true; 
  }); 
}

/**
 * 安全地向指定单元格覆盖写入单一数据（带并发锁）
 * @param {GoogleAppsScript.Spreadsheet.Range} range - 目标写入区域
 * @param {any} value - 要写入的数据
 * @returns {boolean} 写入成功返回 true
 */
function safeSetValue(range, value) { 
  return withLock(function(){ 
    range.setValue(value); 
    return true; 
  });
}

// 当次执行内缓存（V8 全局内存，仅当前执行生命周期内有效）
const _SS_CACHE = {};
const _SHEET_CACHE = {};

/**
 * 获取指定的 Spreadsheet 表格对象，并缓存在当前运行内存中以加速重复获取
 * @param {string} id - Google Sheet 的文件 ID
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} 表格对象
 */
function getSpreadsheetCached(id) {
  if (!_SS_CACHE[id]) {
    _SS_CACHE[id] = SpreadsheetApp.openById(id);
  }
  return _SS_CACHE[id];
}

/**
 * 获取指定的 Sheet 子表对象，并缓存在当前运行内存中以加速重复获取
 * @param {string} id - Google Sheet 的文件 ID
 * @param {string} name - 子工作表 (Sheet) 名称
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} 子表对象
 */
function getSheetCached(id, name) {
  const key = id + "::" + name;
  if (!_SHEET_CACHE[key]) {
    const ss = getSpreadsheetCached(id);
    _SHEET_CACHE[key] = ss.getSheetByName(name);
  }
  return _SHEET_CACHE[key];
}

/***********************
 * 参数外置（Script Properties 常量表）
 ***********************/
const publishingWebsize = getProp("publishingWebsize", getReleaseWebPage());
const wsUserInfoId = getProp("wsUserInfoId", "1F7G3WOY5xM4fEYZ1s5RKulY4kJhqCZ9HefthmiVkraM");
const wsPlanId = getProp("wsPlanId", "1__w48CQA7rERW5VpxfNT12fxiKNSvTrDOr5Pt_f6aDs");
const wsTfBomId = getProp("wsTfBomId", "1Ikmgrv9jdTjBsa9-bNsMfyuZ5OY3ObBjA1mmPQjcROY");
const wsPkBomId = getProp("wsPkBomId", "1TjjDp5h-tizW37ntnY3sKsBaPWAgbUwq5pZUSO7U43o");
const wsAndonId = getProp("wsAndonId", "143IhduTBrtf5f4ixIp7Qy0rmK49qYQqBvrNbA0t0gmg");
const wsHandOverId = getProp("wsHandOverId", "1ce1lnrvlyzcc96VqOWxrGTBJb5mzZQvk6dCQevEkNqs");
const wsFirstConfirmId = getProp("wsFirstConfirmId", "1f74eVldvWq25oggyNnfKl658ty5pk6P6d5JFs_dPi7o");
const wsTwoHourSelfCheckId = getProp("wsTwoHourSelfCheckId", "1TQtxT3Xdilu4QA3wrljogDZL6zryuqF2qniYYoqn_-Y");
const wsInjParaId = getProp("wsInjParaId", "164BO94VJR6qNdJmJDwbz3w7u9QZfNQUv0U6eXSiM3kQ");
const wsTechFirstConfirmId = getProp("wsTechFirstConfirmId", "1guCZvOpKdQlhwnJh8VnsTsoj3kgd5KID0hGyGOyGLGs");
const wsTechPpmsWeekCheckId = getProp("wsTechPpmsWeekCheckId", "18hoqx_pnoRrjqijiDCOWz_hnyu3EUSWS-_sWKAwh868");
const wsFinalMoldCheckId = getProp("wsFinalMoldCheckId", "1BLtdIambdtfGCart3MQJs_K-CO62ludxKqZ03PPIpUg");
const wsNewOldChangeId = getProp("wsNewOldChangeId", "12GicROoALNx8uT7LMH4sUHDMULlUN0U6nZXgUf-8EJg");
const wsEquShiftId = getProp("wsEquShiftId", "10Fnrqc1AUiPqOi-b2UsKgR-Ww-BNdIla_HB_HjVdI0w");
const folderHandbagDatePictureId = getProp("folderHandbagDatePictureId", "1cbsHrPX5Awbt8E_9c6Zvj2K34QwHHcv7");
const wsDefectSampleId = getProp("wsDefectSampleId", "1tfpSYCr8it-AhmEPGvAPcksER3Z9KWHjbfOGTnOVgfs");
const wsInsertPlanSpecialId = getProp("wsInsertPlanSpecialId","1zg0e0sBNmh0IkUH5y3NqJ0V-m_oCDrlOH0uatfhiy30");
const wsPreBomId = getProp("wsPreBomId","1jQBJTzwfUZgX2-xn4HgQ4pvDhfqHgkRI0QHZLL0Tr-4");
const wsSarmId = getProp("wsSarmId","1RH2rmG9zf1ZzwS_25IpXWdYSGIx_EhSud-rWDumC9JA");
const wsTurnOverId = getProp("wsTurnOverId","1-5vkPzceVowy-UboacXv7b3Ist0H7Q-IS7pZ8CBEn1I");
const faviconUrl = getProp("faviconUrl", "https://images.ctfassets.net/m3056igwnpsm/4Osm79Kgn7Gee2S7ho2myx/43508e869900ac6176aba54977fa4526/INJ1.png");

// 外置邮件接收人配置
const emailTB1Ops = getProp("emailTB1Ops", "yong_ji@colpal.com,hua_li@colpal.com");
const emailTB2Ops = getProp("emailTB2Ops", "andy_lu@colpal.com,zhiwei_ding@colpal.com");
const emailEngCC  = getProp("emailEngCC", "jin_zhang@colpal.com,kelland_zhao@colpal.com,lyon_zhang@colpal.com");

// SpecialInfo 表配置
const PLAN_SHEET_ID = getProp("PLAN_SHEET_ID", '1__w48CQA7rERW5VpxfNT12fxiKNSvTrDOr5Pt_f6aDs');
const SpecialInfo_SHEET_NAME = getProp("SpecialInfo_SHEET_NAME", '特殊品种');

/***********************
 * 初始化与常量（利用缓存访问，避免死锁）
 ***********************/
const arrMachineType = ["IM-VIM","IM-HIM","IM-FCS","IM-6AIX-DEC","IM-6AIX","IM-ENL-100","IM-ENL-400","IM-ENL-160","IM-OMN","IM-DEC","IM-PRN"];
const arr5SAMLink = ["1DVZzQM4UJ7ogmrYFdTAMyyMRBCFBYbqPHVcQstvt1yY","1KWua_ljr7SsA_ROadAe754xZt1CY8Gn6dfCpk4Sj9f0","1yUOWdnU4NX3QCC5EZqnxf2hlrgPpZbPB7_OEpnMGhmc","170v7ZjnWhnKrM_uAysdvdTMQlZ6ApOmyypi2WTh5WMU","1-5SuKevxpKDx3nFGPQNIgGH-3rjHLf9BBac_kZf5UeY","1FosrQNgIvHYSQgyxZXaKdelZP120ZE_o2mNJc1R69g0","1r2G_HGhWMsVA9nLFk83gIAfKpCKPIm3Kuc2OlKyMjA4","1Xbp0xNwrbZEZGO65LfyRxCClSfltDSTTNuJnufZCmCc","1BUYeSjFQo5FshHw9sSSTGUQnCqZ95ANFq2SALj6bf0I","1ZqnCuxJxghD8SXqNFkjprRUcVn8ocP9hrfNGfAorpqY","1Wu1Bh1mBEAqOklgVeXl0DUHbPEoUkyd_dcIoWwOvgPQ"];
const userLogInformationId = getProp("userLogInformationId", "1ecEx7G_FX7DAJ_8cm1AxSaN2h95IVo8n-W4WMX-g1m4");

// 使用缓存读取一次（小机BOM、Cell Map）
const sbnSourceBom = getSheetCached(getProp("sabi_Bom", "1yS3aFWyVn3tnJBT4FTmVmtLyUzc0G2HrVZvRK0_JPTg"), "小机BOM");
const sbnSourceCellMap = getSheetCached(getProp("sabi_CellMapping", "11zyH65MhC-LuqsEXT6KeO3-GQ3jwW7z7kJjHD0TwLZc"), "2. Active Cell");

const startBomRow = 2, startBomCol = 1, endBomCol = 42;
const arrBomJsonData = getDataToArrJson(sbnSourceBom, startBomRow, startBomCol, endBomCol);
const startCellMapRow = 2, startCellMapCol = 1, endCellMapCol = 23; 
const arrCellMapJsonData = getDataToArrJson(sbnSourceCellMap, startCellMapRow, startCellMapCol, endCellMapCol);

// 机台列表提取并缓存为 HTML 下拉框片段
const ssMachineSheet = getSheetCached(wsPlanId, "Machine");
const arrAllMachine = (function(){
  const lr = ssMachineSheet.getLastRow();
  const vals = lr > 1 ? ssMachineSheet.getRange(2, 1, lr - 1, 8).getValues() : [];
  return vals.filter(x => x[0] != "" && String(x[7]).indexOf("INJ") != -1).map(x => '<option>' + x[0] + '</option>').join('');
})();

// 提取特殊品种的日期章转换要求
const sbnSpecialInfo = getSheetCached(PLAN_SHEET_ID, SpecialInfo_SHEET_NAME);
let arrSpecialInfo = (function(){
  const lr = sbnSpecialInfo.getLastRow();
  if(lr <= 1) return [];
  return sbnSpecialInfo.getRange(2, 14, lr - 1, 2).getDisplayValues().filter(v => v[0] != "");
})();

let judgeDateCo = (function(){
  const found = arrSpecialInfo.filter(v => v[0] == "日期章转换");
  return found.length ? found[0][1] : "N";
})();
console.log("judgeDateCo:", judgeDateCo);

/**
 * 从 Google Apps Script 脚本属性中安全捞取直配机台的自定义菜单配置
 * @returns {string|null} 返回序列化后的 JSON 字符串配置，若无则返回空
 */
function getDirectDeliveryConfig() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const config = scriptProperties.getProperty("DirectDeliveryConfig");
    return config ? config : null;
  } catch (e) {
    Logger.log("读取直配配置脚本属性失败: " + e.toString());
    return null;
  }
}

/***********************
 * 后端路由与页面渲染引擎
 ***********************/
const Route = {};

/**
 * 注册页面路由
 * @param {string} rou - 路由参数名
 * @param {Function} callback - 路由对应的页面处理函数
 */
Route.path = function(rou, callback){ 
  Route[rou] = callback; 
}

/**
 * Web App 入口函数，解析 URL 参数并分发至对应页面
 * @param {Object} e - HTTP GET 事件对象
 * @returns {GoogleAppsScript.HTML.HtmlOutput} 返回渲染后的网页内容
 */
function doGet(e) {
  Route.path("ChangePassword", Load_ChangePassword);
  Route.path("ViewPlan", Load_ViewPlan);
  Route.path("Approve", Load_Approve);
  Route.path("ViewBom", Load_ViewBom);
  Route.path("HandOverCheck", Load_HandOverCheck);
  Route.path("FirstCheck", Load_FirstCheck);
  Route.path("SelfCheck", Load_SelfCheck);
  Route.path("InsertPlan", Load_InsertPlan);
  Route.path("Transfer", Load_Transfer);
  Route.path("TechViewPara", Load_TechViewPara);
  Route.path("TechFirstCheck", Load_TechFirstCheck);
  Route.path("FinalMoldCollection", Load_FinalMoldCollection);
  Route.path("FinalMoldView", Load_FinalMoldView);
  Route.path("TechPpmsWeekCheck", Load_TechPpmsWeekCheck);
  Route.path("TechPpmsMonthCheck", Load_TechPpmsMonthCheck);

  if (e && e.parameter && e.parameter.v && Route[e.parameter.v]) {
    return Route[e.parameter.v]();
  } else {
    let rdate = Utilities.formatDate(new Date(), "Asia/Shanghai", "yyyy年M月d日 HH:mm:ss");
    return render("Login", {execUrl: publishingWebsize, rdate: rdate});
  }
}

// ---------------- 页面加载模块函数 ----------------

/**
 * 加载修改密码页面
 */
function Load_ChangePassword(){ 
  let rdate = Utilities.formatDate(new Date(), "Asia/Shanghai", "yyyy年M月d日 HH:mm:ss"); 
  return render("ChangePassword", {title: "用户密码修改页面", rdate: rdate});
}

/**
 * 加载计划查看与布局导航页面
 */
function Load_ViewPlan(){ 
  let rdate = Utilities.formatDate(new Date(), "Asia/Shanghai", "yyyy年M月d日 HH:mm:ss"); 
  return render("ViewPlan", {title: "高露洁注塑前线管理系统", rdate: rdate}); 
}

/**
 * 加载 BOM 审批页面
 */
function Load_Approve(){ 
  let rdate = Utilities.formatDate(new Date(), "Asia/Shanghai", "yyyy年M月d日 HH:mm:ss"); 
  return render("Approve", {title: "高露洁三笑BOM转换审批", rdate: rdate});
}

/**
 * 加载 BOM 查看页面
 */
function Load_ViewBom(){ 
  let rdate = Utilities.formatDate(new Date(), "Asia/Shanghai", "yyyy年M月d日 HH:mm:ss"); 
  return render("ViewBom", {title: "高露洁三笑BOM及工序质量标准展示", rdate: rdate}); 
}

/**
 * 加载交接班检查页面
 */
function Load_HandOverCheck(){ 
  let rdate = Utilities.formatDate(new Date(), "Asia/Shanghai", "yyyy年M月d日 HH:mm:ss"); 
  return render("HandOverCheck", {title: "高露洁三笑交接班检查", rdate: rdate});
}

/**
 * 加载注塑首件确认页面
 */
function Load_FirstCheck(){ 
  let rdate = Utilities.formatDate(new Date(), "Asia/Shanghai", "yyyy年M月d日 HH:mm:ss"); 
  return render("FirstCheck", {title: "高露洁三笑注塑机台首件确认", rdate: rdate}); 
}

/**
 * 加载两小时自检页面
 */
function Load_SelfCheck(){ 
  let rdate = Utilities.formatDate(new Date(), "Asia/Shanghai", "yyyy年M月d日 HH:mm:ss"); 
  return render("SelfCheck", {title: "高露洁三笑注塑工序两小时自检", rdate: rdate});
}

/**
 * 加载临时计划插单页面
 */
function Load_InsertPlan(){ 
  let rdate = Utilities.formatDate(new Date(), "Asia/Shanghai", "yyyy年M月d日 HH:mm:ss"); 
  return render("InsertPlan", {title: "高露洁三笑注塑工序临时计划", rdate: rdate});
}

/**
 * 加载 BOM 转换操作页面
 */
function Load_Transfer(){ 
  let rdate = Utilities.formatDate(new Date(), "Asia/Shanghai", "yyyy年M月d日 HH:mm:ss"); 
  return render("Transfer", {title: "高露洁三笑BOM转换", rdate: rdate}); 
}

/**
 * 加载设备工艺参数查询页面
 */
function Load_TechViewPara(){ 
  let rdate = Utilities.formatDate(new Date(), "Asia/Shanghai", "yyyy年M月d日 HH:mm:ss"); 
  return render("TechViewPara", {title: "高露洁三笑工艺参数展示", rdate: rdate});
}

/**
 * 加载设备工程首件页面
 */
function Load_TechFirstCheck(){ 
  let rdate = Utilities.formatDate(new Date(), "Asia/Shanghai", "yyyy年M月d日 HH:mm:ss"); 
  return render("TechFirstCheck", {title: "高露洁三笑设备首件确认", rdate: rdate}); 
}

/**
 * 加载末模收集填写页面
 */
function Load_FinalMoldCollection(){ 
  let rdate = Utilities.formatDate(new Date(), "Asia/Shanghai", "yyyy年M月d日 HH:mm:ss"); 
  return render("FinalMoldCollection", {title: "高露洁三笑注塑工序末模收集记录", rdate: rdate});
}

/**
 * 加载末模状态管理查询页面
 */
function Load_FinalMoldView(){ 
  let rdate = Utilities.formatDate(new Date(), "Asia/Shanghai", "yyyy年M月d日 HH:mm:ss"); 
  return render("FinalMoldView", {title: "高露洁三笑注塑工序末模管理", rdate: rdate}); 
}

/**
 * 加载设备周检页面
 */
function Load_TechPpmsWeekCheck(){ 
  let rdate = Utilities.formatDate(new Date(), "Asia/Shanghai", "yyyy年M月d日 HH:mm:ss"); 
  return render("TechPpmsWeekCheck", {title: "高露洁三笑设备工艺参数周检", rdate: rdate});
}

/**
 * 加载设备月检页面
 */
function Load_TechPpmsMonthCheck(){ 
  let rdate = Utilities.formatDate(new Date(), "Asia/Shanghai", "yyyy年M月d日 HH:mm:ss"); 
  return render("TechPpmsMonthCheck", {title: "高露洁三笑设备工艺参数月检", rdate: rdate}); 
}

/**
 * 基于 HTML 模板渲染网页内容，并支持传入后端变量至前端
 * @param {string} file - HTML 文件名 (不含后缀)
 * @param {Object} obj - 需要暴露给模板引擎的变量集合
 * @returns {GoogleAppsScript.HTML.HtmlOutput} 供客户端渲染的页面
 */
function render(file, obj){
  let tmp = HtmlService.createTemplateFromFile(file);
  if(obj){ 
    Object.keys(obj).forEach(function(key){ 
      tmp[key] = obj[key]; 
    });
  }
  if(file == "ChangePassword"){ return tmp.evaluate().setTitle('修改密码').setFaviconUrl(faviconUrl); }
  if(file == "Login"){ return tmp.evaluate().setTitle('登陆注塑前线管理').setFaviconUrl(faviconUrl); }
  if(file == "ViewPlan"){ 
    tmp.listMachine = arrAllMachine; 
    return tmp.evaluate().setTitle('查看注塑计划').setFaviconUrl(faviconUrl);
  }
  if(file == "Approve"){
    const ssAllInj = getSheetCached(wsPlanId, "All_INJ");
    const data = (function(){ 
      const lr = ssAllInj.getLastRow(); 
      return lr > 1 ? ssAllInj.getRange(2, 10, lr - 1, 1).getValues() : []; 
    })();
    let material = []; 
    data.forEach(r => { 
      if(material.indexOf(r[0]) === -1){ material.push(r[0]); } 
    });
    tmp.listMaterial = material.map(x => '<option>' + x + '</option>').join('');
    return tmp.evaluate().setTitle('审批BOM').setFaviconUrl(faviconUrl);
  }
  if(file == "ViewBom"){ return tmp.evaluate().setTitle('查看BOM').setFaviconUrl(faviconUrl); }
  if(file == "HandOverCheck"){ return tmp.evaluate().setTitle('交接班检查').setFaviconUrl(faviconUrl); }
  if(file == "FirstCheck"){ return tmp.evaluate().setTitle('首件确认').setFaviconUrl(faviconUrl); }
  if(file == "SelfCheck"){ return tmp.evaluate().setTitle('自检确认').setFaviconUrl(faviconUrl); }
  if(file == "InsertPlan"){
    const lrBom = sbnSourceBom.getLastRow();
    let arrSku = Array.from(new Set(sbnSourceBom.getRange(2, 4, Math.max(lrBom - 1, 0), 1).getValues().filter(x => String(x[0]).length > 4 && String(x[0]).indexOf("牙柄") == -1).map(x => '<option>' + x[0] + '</option>'))).join('');
    let arrHandSku = Array.from(new Set(sbnSourceBom.getRange(2, 9, Math.max(lrBom - 1, 0), 1).getValues().filter(x => String(x[0]).length > 4 && String(x[0]).indexOf("手包") == -1).map(x => '<option>' + x[0] + '</option>'))).join('');
    const sbnMachineMoldType = getSheetCached(getProp("machineInitId","1zNOM35TeOmaZTtaAR4izQdA37cLGSgPhJAw40KkYioQ"), "机台初始状态");
    let lrMM = sbnMachineMoldType.getLastRow();
    let arrMachineMoldType = sbnMachineMoldType.getRange(1, 1, lrMM, 3).getDisplayValues().filter((v, i) => i > 0).map(v => '<option data-tokens="' + v[0] + '|' + v[2] + '" data-subtext="' + v[0] + '" value="' + v[2] + '">' + v[2] + '</option>');
    const sbnSkuMoldType = getSheetCached(getProp("sabi_Bom", "1yS3aFWyVn3tnJBT4FTmVmtLyUzc0G2HrVZvRK0_JPTg"), "小机BOM");
    let lrSM = sbnSkuMoldType.getLastRow();
    let arrSkuMoldType = sbnSkuMoldType.getRange(3, 1, Math.max(lrSM - 2, 0), 14).getDisplayValues().filter((v, i) => i > 0).map(v => '<option data-tokens="' + v[3] + '|' + v[13] + '" data-subtext="' + v[3] + '" value="' + v[13] + '">' + v[13] + '</option>');
    tmp.listMachine = arrAllMachine; 
    tmp.listSku = arrSku; 
    tmp.listHandSku = arrHandSku; 
    tmp.listMachineMoldType = arrMachineMoldType; 
    tmp.listSkuMoldType = arrSkuMoldType;
    return tmp.evaluate().setTitle('临时计划').setFaviconUrl(faviconUrl);
  }
  if(file == "Transfer"){ return tmp.evaluate().setTitle('BOM转换').setFaviconUrl(faviconUrl); }
  if(file == "TechViewPara"){ return tmp.evaluate().setTitle('查询工艺参数').setFaviconUrl(faviconUrl); }
  if(file == "TechFirstCheck"){ return tmp.evaluate().setTitle('设备首件确认').setFaviconUrl(faviconUrl); }
  if(file == "FinalMoldCollection"){ return tmp.evaluate().setTitle('注塑末模收集').setFaviconUrl(faviconUrl); }
  if(file == "FinalMoldView"){ return tmp.evaluate().setTitle('注塑末模查询').setFaviconUrl(faviconUrl); }
  if(file == "TechPpmsWeekCheck"){ return tmp.evaluate().setTitle('设备参数周检').setFaviconUrl(faviconUrl); }
  if(file == "TechPpmsMonthCheck"){ return tmp.evaluate().setTitle('设备参数月检').setFaviconUrl(faviconUrl); }
}

/**
 * 加载 HTML 文件的内容用以被母版嵌入 (代替 PHP 的 include)
 * @param {string} filename - 需要引用的 HTML 片段名
 * @returns {string} 文件纯文本内容
 */
function include(filename){ 
  return HtmlService.createHtmlOutputFromFile(filename).getContent(); 
}

/**
 * 获取当前 Web App 发布的 URL，以便前端发起重定向
 * @returns {string} App URL 链接
 */
function getReleaseWebPage(){ 
  let webPageUrl = ScriptApp.getService().getUrl(); 
  console.log(webPageUrl); 
  return webPageUrl;
}

/***********************
 * 业务函数（开始）—— 内置缓存替换
 ***********************/
/***********************
 * 新版安全登录与单点查询引擎（带 CacheService 缓存加速与查表兜底）
 * 彻底废弃旧版全量下载密码的做法，保障极速体验与绝对安全
 ***********************/

/**
 * 根据工号查询用户的脱敏信息（供前端工号输入框实时回显专用）
 * 并将包含密码的完整信息打包装入后端 CacheService，有效期 3 分钟
 * @param {string} loginId - 前端传入的 5 位工号
 * @returns {Object} 包含查询状态及用户脱敏信息的 JSON 对象
 */
function getUserInfoById(loginId) {
  try {
    const ss = getSheetCached(wsUserInfoId, "userID");
    const lr = ss.getLastRow();
    if (lr <= 1) {
      return { status: "NO", message: "用户数据库为空" };
    }
    
    // 获取 A列(0) 到 H列(7) 的数据
    const data = ss.getRange(2, 1, lr - 1, 8).getValues();
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] == loginId) {
        // 构建完整的用户信息字典（包含密码，供后端自己下次验证用）
        const fullUserInfo = {
          id: data[i][0],
          pwd: data[i][2].toString(), // C列：密码
          name: data[i][1],           // B列：姓名
          dept: data[i][5],           // F列：部门
          auth: data[i][7]            // H列：权限标识 (如 "YT")
        };

        // 【核心优化】：存入后端的 CacheService 缓存，有效期 3 分钟 (180秒)
        const cache = CacheService.getScriptCache();
        cache.put("USER_" + loginId, JSON.stringify(fullUserInfo), 180);

        // 返回给前端的数据【绝对不能包含 pwd 密码字段】！
        return {
          status: "OK",
          data: {
            id: fullUserInfo.id,
            name: fullUserInfo.name,
            dept: fullUserInfo.dept,
            auth: fullUserInfo.auth
          }
        };
      }
    }
    
    return { status: "NO", message: "未找到该工号" };
  } catch (e) {
    return { status: "ERROR", message: e.toString() };
  }
}

/**
 * 智能双轨登录校验引擎：优先读取 CacheService 极速验证，未命中则降级查表兜底
 * @param {string} loginId - 前端传入的工号
 * @param {string} loginPassword - 前端传入的明文密码
 * @returns {Object} 包含验证结果及用户脱敏信息的 JSON 对象
 */
function verifyUserLogin(loginId, loginPassword) {
  try {
    const cache = CacheService.getScriptCache();
    const cachedStr = cache.get("USER_" + loginId);

    // ==========================================
    // 轨道一：缓存命中（极速验证，0 查表）
    // ==========================================
    if (cachedStr) {
      const cachedUser = JSON.parse(cachedStr);
      // 在内存中极速对账密码
      if (cachedUser.pwd === loginPassword) {
        try { writeUserLogInformation(loginId, cachedUser.name); } catch(logErr) {}
        
        return {
          status: "OK",
          data: {
            id: cachedUser.id,
            name: cachedUser.name,
            dept: cachedUser.dept,
            auth: cachedUser.auth
          }
        };
      } else {
        return { status: "NO", message: "密码不正确" };
      }
    }

    // ==========================================
    // 轨道二：缓存未命中（降级查表兜底）
    // 触发场景：前端直接用了本地微缓存没调后端，或后端缓存超 3 分钟过期了
    // ==========================================
    const ss = getSheetCached(wsUserInfoId, "userID");
    const lr = ss.getLastRow();
    if (lr <= 1) {
      return { status: "NO", message: "用户数据库为空" };
    }
    
    const data = ss.getRange(2, 1, lr - 1, 8).getValues();
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] == loginId) {
        if (data[i][2].toString() === loginPassword) {
          const userName = data[i][1];
          try { writeUserLogInformation(loginId, userName); } catch(logErr) {}
          
          return {
            status: "OK",
            data: {
              id: data[i][0],
              name: userName,
              dept: data[i][5],
              auth: data[i][7]
            }
          };
        } else {
          return { status: "NO", message: "密码不正确" };
        }
      }
    }
    
    return { status: "NO", message: "未找到该工号" };
  } catch (e) {
    return { status: "ERROR", message: e.toString() };
  }
}


/***********************
 * 业务函数（续）—— 使用缓存的 Sheet 访问 + 并发保护
 ***********************/

/**
 * 核心数据接口：一次性拉取计划、BOM、参数、机台状态、交接班及自检记录等所有全量数据
 * @returns {Array} 包含所有业务基础模块的数据大礼包
 */
function getAllPlan() {  
  try {
    // 1. 获取缺陷样表
    const sbnDefectSample = getSheetCached(wsDefectSampleId, "Total");
    const tlr = sbnDefectSample.getLastRow();
    const arrTemplate = tlr > 0 ? sbnDefectSample.getRange(1, 1, tlr, 2).getDisplayValues() : [];
    
    // 2. 获取小机 BOM（源表已在全局初始化加载过）
    const lrBom = sbnSourceBom.getLastRow();
    const arrBom = lrBom > 1 ? sbnSourceBom.getRange(2, 1, lrBom - 1, 15).getValues().filter(x => String(x[2]).length > 4) : [];

    // 3. 获取排产计划与 IoT 机台状态
    const ssPlanAll = getSheetCached(wsPlanId, "All_INJ");
    const ssIot = getSheetCached(wsPlanId, "IoT机台");
    const planLr = ssPlanAll.getLastRow();
    let arrData = [];
    if (planLr > 1) {
      arrData = ssPlanAll.getRange(2, 1, planLr - 1, 22).getDisplayValues()
        .map(x => [x[0], x[1], x[2], x[3], x[4], x[5], x[7], x[8], x[6], x[9], x[10], x[11], x[12], x[13], x[14], x[15], x[16], x[17], x[20], x[21]]);
    }

    let arrIotMachine = [];
    const iotLr = ssIot.getLastRow();
    if (iotLr > 1){
      arrIotMachine = ssIot.getRange(2, 1, iotLr - 1, 5).getDisplayValues();
    }

    // 4. 当计划存在时，进行后续级联数据（如参数、BOM等）的按需匹配抓取
    if(arrData.length > 0) {
      const ssMachine = getSheetCached(wsPlanId, "Machine");
      const mLr = ssMachine.getLastRow();
      const arrMachineInfo = ssMachine.getRange(1, 1, mLr, 8).getDisplayValues();

      // 提取唯一的查询关键词列表以进行批量查询，提升性能
      let arrKeyWord_Para = Array.from(new Set(arrData.map(v => [v[4].substring(0,8), v[8].substring(0,8), v[9]].join("【|】"))));
      let arrKeyWord_Bom  = Array.from(new Set(arrData.map(v => v[9])));
      let arrKeyWord_PkSku = Array.from(new Set(arrData.filter(v => v[10] != "").map(v => v[10])));
      console.log("arrKeyWord_PkSku", arrKeyWord_PkSku);

      let arrAllPlanBomData  = getAllBomData(arrKeyWord_Bom, arrKeyWord_PkSku);
      let arrAllPlanParaData = getAllParaData(arrMachineInfo, arrKeyWord_Para);

      let arrNowDateShiftSelfCheckRecord = getNowDateShiftExitSelfCheckRecord();
      let arrNowDateShiftHandOverRecord = getNowDateShiftExitHandOverRecord();

      let arrAllPlanPkBomData = getAllPkBomData(arrKeyWord_PkSku);

      return ["OK", arrData, arrBom, arrAllMachine, arrIotMachine, arrTemplate, arrAllPlanBomData, arrAllPlanParaData, arrNowDateShiftSelfCheckRecord, arrNowDateShiftHandOverRecord, arrMachineInfo, arrMachineType, arr5SAMLink, arrAllPlanPkBomData];
    } else {
      return ["NO", "未找到数据"];
    }
  } catch(e) {
    return ["NO", e.toString()];
  }
}

/**
 * 极简拉取接口：专供前端定时刷新或快速检查时使用
 * 仅拉取计划、IoT状态、自检与交接班记录，主动跳过极度耗时的 BOM 和 工艺参数查询
 * @returns {Array} 结构同 getAllPlan，但耗时数据块被置为空数组
 */
function getDynamicPlanOnly() {
  try {
    // 1. 获取主计划与 IoT 状态
    const ssPlanAll = getSheetCached(wsPlanId, "All_INJ");
    const ssIot = getSheetCached(wsPlanId, "IoT机台");
    const planLr = ssPlanAll.getLastRow();
    
    let arrData = [];
    if (planLr > 1) {
      arrData = ssPlanAll.getRange(2, 1, planLr - 1, 22).getDisplayValues()
        .map(x => [x[0], x[1], x[2], x[3], x[4], x[5], x[7], x[8], x[6], x[9], x[10], x[11], x[12], x[13], x[14], x[15], x[16], x[17], x[20], x[21]]);
    }

    let arrIotMachine = [];
    const iotLr = ssIot.getLastRow();
    if (iotLr > 1){
      arrIotMachine = ssIot.getRange(2, 1, iotLr - 1, 5).getDisplayValues();
    }

    // 2. 获取当前的动态自检与交接班记录
    let arrNowDateShiftSelfCheckRecord = getNowDateShiftExitSelfCheckRecord();
    let arrNowDateShiftHandOverRecord = getNowDateShiftExitHandOverRecord();

    if(arrData.length > 0) {
      // 按照原 getAllPlan 的格式返回，用空数组跳过庞大的静态数据块
      return [
        "OK", 
        arrData,                      // [1] 计划数据
        [],                           // [2] arrBom (跳过)
        [],                           // [3] arrAllMachine (跳过)
        arrIotMachine,                // [4] IoT机台
        [],                           // [5] arrTemplate (跳过)
        [],                           // [6] arrAllPlanBomData (极度耗时，跳过)
        [],                           // [7] arrAllPlanParaData (极度耗时，跳过)
        arrNowDateShiftSelfCheckRecord, // [8] 当班自检记录
        arrNowDateShiftHandOverRecord,  // [9] 当班交接记录
        [],                           // [10] arrMachineInfo (跳过)
        [],                           // [11] arrMachineType (跳过)
        [],                           // [12] arr5SAMLink (跳过)
        []                            // [13] arrAllPlanPkBomData (跳过)
      ];
    } else {
      return ["NO", "未找到计划数据"];
    }
  } catch(e) {
    return ["NO", e.toString()];
  }
}

/**
 * 根据提取的关键词列表拉取相关的 注塑 BOM、包装 BOM、烫印及喷墨信息
 * @param {Array} arrKeyWord_Bom - 注塑 SKU 集合
 * @param {Array} arrKeyWord_PkSku - 包装 SKU 集合
 * @returns {Array} 包含分类拆解后的标准物料数据
 */
function getAllBomData(arrKeyWord_Bom, arrKeyWord_PkSku) {
  const ssTfHeader = getSheetCached(wsTfBomId, "TF BOM masterdata Header");
  const injLr = ssTfHeader.getLastRow();
  let arrSsInj = injLr > 1 ? ssTfHeader.getRange(2, 1, injLr - 1, 9).getValues().filter(v => v[8] == "生效") : [];
  let arrSsInjStandard = filterInjByBom(arrSsInj, arrKeyWord_Bom);

  const ssTfInjRel = getSheetCached(wsTfBomId, "INJ相关");
  const tfLr = ssTfInjRel.getLastRow();
  let arrSsTf = tfLr > 1 ? ssTfInjRel.getRange(2, 1, tfLr - 1, 18).getValues()
    .filter(x => x[1] != "" && x[17] == "生效")
    .map(x => [x[0], x[1], x[3], x[4], x[6], x[9]]) : [];

  const ssPk = getSheetCached(wsPkBomId, "BOM masterdata");
  const pkLr = ssPk.getLastRow();
  let arrSsPk = pkLr > 1 ? ssPk.getRange(2, 1, pkLr - 1, 11).getDisplayValues()
    .filter(v => v[8] == "生效" && arrKeyWord_PkSku.indexOf(v[0]) != -1) : [];

  let arrSsTfStandard = arrSsTf.filter(x => arrKeyWord_Bom.indexOf(x[1]) != -1);

  const ssDecal = getSheetCached(wsTfBomId, "烫印膜");
  const dLr = ssDecal.getLastRow();
  let arrSsDecal = dLr > 1 ? ssDecal.getRange(2, 1, dLr - 1, 8).getDisplayValues()
    .filter(x => x[7] == "生效" && (arrKeyWord_Bom.indexOf(x[1]) != -1 || arrKeyWord_Bom.indexOf(x[2]) != -1))
    .map(x => [x[0], x[1], x[2], x[3], x[4], x[5], x[6]]) : [];

  const ssPrint = getSheetCached(wsTfBomId, "喷墨打印");
  const pLr = ssPrint.getLastRow();
  let arrSsDPrint = pLr > 1 ? ssPrint.getRange(2, 1, pLr - 1, 7).getDisplayValues()
    .filter(x => x[6] == "生效" && arrKeyWord_Bom.indexOf(x[1]) != -1)
    .map(x => [x[0], x[1], x[2], x[3], x[4]]) : [];

  console.log("arrSsPk", arrSsPk);
  return [arrSsInjStandard, arrSsTfStandard, arrSsPk, arrSsDecal, arrSsDPrint];
}

/**
 * 根据所需的 BOM 关键字过滤对应的注塑材料数据，内部会处理含多SKU字符串的拆分
 * @param {Array} data - 注塑基础数据
 * @param {Array} arrKeyWord_Bom - 目标 SKU 关键词集合
 * @returns {Array} 过滤并去重后的记录数组
 */
function filterInjByBom(data, arrKeyWord_Bom) {
  const norm = s => String(s == null ? "" : s).trim().toUpperCase();
  const bomSet = new Set((arrKeyWord_Bom || []).filter(Boolean).map(norm));
  const splitSkuList = skuMixed => {
    const raw = norm(skuMixed);
    if (!raw || raw === "NA" || raw === "N/A") return [];
    return raw.split(/[,\uFF0C;；/\\|、\s]+/).map(norm).filter(Boolean);
  };

  const filtered = data.filter(row => {
    const skuList = splitSkuList(row[4]);
    return skuList.length > 0 && skuList.some(sku => bomSet.has(sku));
  });

  const seen = new Set();
  const unique = [];
  for (const row of filtered) {
    const key = row[0];
    if (seen.has(key)) continue;
    seen.add(key);
    unique.push(row);
  }
  return unique;
}

/**
 * 基于机台信息和特定参数标识获取目标机台的标准工艺参数
 * @param {Array} arrMachineInfo - 机台及其分组台账
 * @param {Array} arrKeyWord_Para - 需要匹配的机台SKU参数联合键集合
 * @returns {Array} 符合条件的参数标准数据
 */
function getAllParaData(arrMachineInfo, arrKeyWord_Para) {
  const arrJsonData = getProcessParameterData(wsInjParaId, "INJ_New");
  const arrMachine = arrMachineInfo.map(v => v[0]);

  for(let i = 0; i < arrKeyWord_Para.length; i++){
    let machineTemp = arrKeyWord_Para[i].split("【|】")[0];
    let position = arrMachine.indexOf(machineTemp);
    let machineModel = position != -1 ? arrMachineInfo[position][1] : "";
    let machineType = machineTemp.substring(2, 4);
    let parts = arrKeyWord_Para[i].split("【|】");
    parts.push(machineType, machineModel);
    arrKeyWord_Para[i] = parts.join("【|】");
  }
  return filterWithSkuFallbackToNA(arrKeyWord_Para, arrJsonData);
}

/**
 * 参数过滤核心逻辑，如果专属 SKU 未找到工艺标准，则退而寻找通用 NA 记录
 * @param {Array} arrKeyWord_Para - 机台_SKU_模具等联合键名数组
 * @param {Array} arrJsonData - 原始全量参数 JSON
 * @returns {Array} 过滤出的匹配参数项
 */
function filterWithSkuFallbackToNA(arrKeyWord_Para, arrJsonData) {
  const SEP = "【|】";
  const norm = x => String(x == null ? "" : x).trim();
  const isNA = s => {
    const t = norm(s).toUpperCase();
    return t === "NA" || t === "N/A" || t === "";
  };
  const key3 = (c, m, g) => c + "||" + m + "||" + g;
  const key4 = (c, m, g, s) => c + "||" + m + "||" + g + "||" + s;
  
  const set3 = new Set();
  const set4 = new Set();

  (arrKeyWord_Para || [])
    .filter(Boolean)
    .forEach(s => {
      const parts = String(s).split(SEP).map(norm);
      if (parts.length >= 5) {
        const code  = parts[1];
        const sku   = parts[2];
        const model = parts[3];
        const group = parts[4];
        set3.add(key3(code, model, group));
        set4.add(key4(code, model, group, sku));
      }
    });

  const splitSkuList = skuMixed => {
    const raw = norm(skuMixed);
    if (!raw || isNA(raw)) return [];
    return raw.split(/[,\uFF0C;；/\\/|、\s]+/).map(norm).filter(s => s && !isNA(s));
  };

  const result = (arrJsonData || []).filter(v => {
    if (!v || v.状态 !== "生效") return false;
    const code  = norm(v.模具编码);
    const model = norm(v.机型);
    const group = norm(v.自动化分组);
    const skuList = splitSkuList(v.SKU);
    const hit4 = skuList.length > 0 && skuList.some(sku => set4.has(key4(code, model, group, sku)));
    if (hit4) return true;
    return set3.has(key3(code, model, group));
  });

  return result;
}

/**
 * 计算当下的物理时间所对应的“工厂逻辑班次”（早/中/夜）
 * @returns {string} 班次标识符，如 "2026-05-16_1"
 */
function getCurrentShiftKey() {
  const tz = "Asia/Shanghai";
  const now = new Date();
  const localStr = Utilities.formatDate(now, tz, "yyyy-MM-dd HH:mm");
  const [dStr, hmStr] = localStr.split(" ");
  const [h, m] = hmStr.split(":").map(Number);
  const minutes = h * 60 + m;

  // 班次边界（分钟）
  const M_07_45 = 7 * 60 + 45;
  const M_15_45 = 15 * 60 + 45;
  const M_23_45 = 23 * 60 + 45;

  let dateStr = dStr;
  let shift;

  if (minutes >= M_23_45) {
    // 23:45~23:59 属于第二天的 1 班 (夜班)
    const next = new Date(now.getTime() + 24 * 3600 * 1000);
    dateStr = Utilities.formatDate(next, tz, "yyyy-MM-dd");
    shift = 1;
  } else if (minutes < M_07_45) {
    shift = 1;
  } else if (minutes < M_15_45) {
    shift = 2;
  } else {
    shift = 3;
  }
  return dateStr + "_" + shift;
}

/**
 * 以指定列的机器号为依据，对数据进行分组
 * @param {number} colPosition - 机台号在子数组中的索引
 * @param {Array} rows - 需要分组的数据二维数组
 * @param {string} sheetName - 所在工作表的识别名，用以区分来源
 * @returns {Array} 按机台汇总的数据结构
 */
function groupByMachine(colPosition, rows, sheetName) {
  const map = new Map();
  for (const r of rows) {
    const machine = r[colPosition];
    if (!map.has(machine)) map.set(machine, []);
    map.get(machine).push(r);
  }
  return Array.from(map.entries()).map(([machine, list]) => [machine, list.length, sheetName, list]);
}

/**
 * 获取当前班次下各个机台已经执行过交接班记录的分组信息
 * @returns {Array} 各机台的提交数及记录体
 */
function getNowDateShiftExitHandOverRecord() {
  const shiftKey = getCurrentShiftKey();
  const shRecord = getSheetCached(wsHandOverId, "CheckRecord");
  const lrRecord = shRecord.getLastRow();

  // 一次性读取有效数据（跳过表头）
  const rowsRecordRaw = lrRecord > 1 ? shRecord.getRange(2, 1, lrRecord - 1, 15).getDisplayValues() : [];
  
  // 交接班按第13列（索引12）筛选当前班次
  const rowsRecord = rowsRecordRaw.filter(r => r[12] === shiftKey);
  const result = groupByMachine(1, rowsRecord, "CheckRecord");

  console.log("Exit handover groups:", result.length);
  if (result.length) console.log("Sample group:", result[0]);

  return result;
}

/**
 * 获取当前班次下各个机台已经执行过的【两小时自检】打卡记录
 * @returns {Array} 分组合并后的各个机台自检情况
 */
function getNowDateShiftExitSelfCheckRecord() {
  const shiftKey = getCurrentShiftKey();

  const shNonDecal = getSheetCached(wsTwoHourSelfCheckId, "NonDecalPrint");
  const shDecal = getSheetCached(wsTwoHourSelfCheckId, "DecalPrint");
  const lrNon = shNonDecal.getLastRow();
  const lrDecal = shDecal.getLastRow();

  const rowsNonRaw = lrNon > 1 ? shNonDecal.getRange(2, 1, lrNon - 1, 29).getDisplayValues() : [];
  const rowsDecalRaw = lrDecal > 1 ? shDecal.getRange(2, 1, lrDecal - 1, 26).getDisplayValues() : [];

  const rowsNon = rowsNonRaw.filter(r => r[27] === shiftKey);
  const rowsDecal = rowsDecalRaw.filter(r => r[24] === shiftKey);

  // 根据第5列 (索引4) 机器分组
  const groupedNon = groupByMachine(4, rowsNon, "NonDecalPrint");
  const groupedDecal = groupByMachine(4, rowsDecal, "DecalPrint");

  const result = [...groupedNon, ...groupedDecal];
  console.log("Exit self-check groups:", result.length);

  return result;
}

/**
 * 获取包装 SKU 所对应的材料 BOM 清单
 * @param {Array} arrPkSku - 包装 SKU 集合
 * @returns {Array} 包材基础信息二维数组
 */
function getAllPkBomData(arrPkSku) {
  const sh = getSheetCached(wsPkBomId, "BomData");
  const shLr = sh.getLastRow();
  let arrPkBomData = shLr > 1 ? sh.getRange(2, 1, shLr - 1, 24).getDisplayValues().filter(v => arrPkSku.indexOf(v[0]) != -1 && v[23] == "生效") : [];
  return arrPkBomData;
}

/**
 * 单条写入交接班记录 (弃用/兼容保留)
 * @param {Array} arrCheck - 单机台完整的交接班数据项数组
 * @returns {Array} 操作状态
 */
function writeHandOverRecord(arrCheck) {
  console.log("handOverRecord", arrCheck);
  try {
    let ss = getSheetCached(wsHandOverId, "CheckRecord");
    let arrRecord = [];
    let arrHoMachine = arrCheck[1].split(",");
    for(let i = 0; i < arrHoMachine.length; i++){
      safeAppendRow(ss, [arrCheck[0], arrHoMachine[i].toString().substring(0,8), arrCheck[2], arrCheck[3], arrCheck[4], arrCheck[5], arrCheck[6], arrCheck[7], arrCheck[8], arrCheck[9], arrCheck[10], arrCheck[11], arrCheck[12], arrCheck[13], arrCheck[14]]);
      arrRecord.push([arrCheck[0], arrHoMachine[i].toString().substring(0,8), arrCheck[2], arrCheck[3], arrCheck[4], arrCheck[5], arrCheck[6], arrCheck[7], arrCheck[8], arrCheck[9], arrCheck[10], arrCheck[11], arrCheck[12], arrCheck[13], arrCheck[14]]);
    }
    return ["OK", arrRecord];
  } catch(e) {
    return ["NO", e.toString()];
  }
}

/**
 * 根据视图及时间请求 SARM 退料管理中的原始物料数据
 * @param {boolean} judgeView - 是否全量查询的视图状态标识
 * @param {Array} viewType - 想要过滤的具体类型集合
 * @returns {Array} 符合要求的物料记录表
 */
function getSarmTotalData(judgeView, viewType) {
  try {
    let arrSarmTotalData = [];
    let now = new Date();
    let nowHm = Utilities.formatDate(now, "Asia/Shanghai", "HH:mm");
    
    if(!judgeView && ((nowHm > "01:15" && nowHm < "07:45") || (nowHm > "09:15" && nowHm < "15:45") || (nowHm > "17:15" && nowHm < "23:45"))){
      const sbn_Sarm = getSheetCached(getProp("sabiId_Sarm", "1RH2rmG9zf1ZzwS_25IpXWdYSGIx_EhSud-rWDumC9JA"), "Data");
      let lr = sbn_Sarm.getLastRow();
      let startRow = 1;
      if(lr > 200){ startRow = lr - 199; } else if(lr > 1){ startRow = 2; }
      
      if(startRow > 1){
        arrSarmTotalData = sbn_Sarm.getRange(startRow, 1, lr - startRow + 1, 15).getDisplayValues();
        arrSarmTotalData.forEach((v, index) => v.push(startRow + index));
        arrSarmTotalData.forEach(v => { 
          if(v[14] == "TRUE"){ v[3] = v[3] + "【加急】"; } 
          v[1] = Utilities.formatDate(new Date(v[1]), "Asia/Shanghai", "yyyy-MM-dd"); 
          if(v[8] == "减料"){ v[10] = 0; } 
          v.splice(12, 3); 
        });
        arrSarmTotalData = arrSarmTotalData.filter(v => viewType.indexOf(v[8]) != -1);
      }
      return ["OK", arrSarmTotalData];
    } else if(judgeView){
      const sbn_Sarm = getSheetCached(getProp("sabiId_Sarm", "1RH2rmG9zf1ZzwS_25IpXWdYSGIx_EhSud-rWDumC9JA"), "Data");
      let lr = sbn_Sarm.getLastRow();
      let startRow = 1;
      if(lr > 200){ startRow = lr - 199; } else if(lr > 1){ startRow = 2; }
      
      if(startRow > 1){
        arrSarmTotalData = sbn_Sarm.getRange(startRow, 1, lr - startRow + 1, 15).getDisplayValues();
        arrSarmTotalData.forEach((v, index) => v.push(startRow + index));
        arrSarmTotalData.forEach(v => { 
          if(v[14] == "TRUE"){ v[3] = v[3] + "【加急】"; } 
          v[1] = Utilities.formatDate(new Date(v[1]), "Asia/Shanghai", "yyyy-MM-dd"); 
          if(v[8] == "减料"){ v[10] = 0; } 
          v.splice(12, 3); 
        });
        arrSarmTotalData = arrSarmTotalData.filter(v => viewType.indexOf(v[8]) != -1);
      }
      return ["OK", arrSarmTotalData];
    }
  } catch(e) { 
    return ["NO", e.toString()]; 
  }
}

/**
 * 保存用户对 Sars 物料数据的修改请求
 * @param {Array} arrSaveSarsData - 修改请求包集合
 * @returns {Array}
 */
function saveEditSarsData(arrSaveSarsData) {
  try {
    const sbn_Sarm = getSheetCached(getProp("sabiId_Sarm", "1RH2rmG9zf1ZzwS_25IpXWdYSGIx_EhSud-rWDumC9JA"), "Data");
    let arrEditTbodyRows = [];
    for(let i = 0; i < arrSaveSarsData.length; i++){
      safeSetValues(sbn_Sarm.getRange(arrSaveSarsData[i][0], 9, 1, 3), [[arrSaveSarsData[i][1], arrSaveSarsData[i][2], arrSaveSarsData[i][3]]]);
      arrEditTbodyRows.push([arrSaveSarsData[i][4]]);
    }
    return ["OK", arrEditTbodyRows];
  } catch(e) { 
    return ["NO", e.toString()]; 
  }
}

/**
 * 提取近期的邮寄发料物流信息
 * @returns {Array} 物流发料数据
 */
function getPostageData() {
  try {
    let arrPostageData = [];
    let now = new Date();
    let next = new Date(now.getTime() + 24 * 3600 * 1000);
    let nowDate = Utilities.formatDate(now, "Asia/Shanghai", "yyyy-MM-dd");
    let nextDate = Utilities.formatDate(next, "Asia/Shanghai", "yyyy-MM-dd");
    let nowHourMinute = Utilities.formatDate(now, "Asia/Shanghai", "HH:mm");
    
    let currentDate = nowDate, nowShift = "";
    if(nowHourMinute >= "23:45" && nowHourMinute < "24:00"){ currentDate = nextDate; nowShift = "夜"; }
    else if(nowHourMinute >= "00:00" && nowHourMinute < "07:45"){ currentDate = nowDate; nowShift = "夜"; }
    else if(nowHourMinute >= "07:45" && nowHourMinute < "15:45"){ currentDate = nowDate; nowShift = "早"; }
    else if(nowHourMinute >= "15:45" && nowHourMinute < "23:45"){ currentDate = nowDate; nowShift = "中"; }

    const sbn_Postage = getSheetCached(getProp("sabiId_Postage", "1aJ1riCcR0a4GsDWSPpXW0u517sZnrsngl9gBPR7HFto"), "送料");
    const lr = sbn_Postage.getLastRow();
    let startRow = 1;
    if(lr > 500){ startRow = lr - 499; } else if(lr > 1){ startRow = 2; }
    
    if(startRow > 1){
      arrPostageData = sbn_Postage.getRange(startRow, 1, lr - startRow + 1, 10).getDisplayValues();
      arrPostageData.forEach(v => { v[0] = Utilities.formatDate(new Date(v[0]), "Asia/Shanghai", "yyyy-MM-dd"); });
      arrPostageData = arrPostageData.filter(v => v[0] == currentDate && v[1] == nowShift);
    }
    return ["OK", arrPostageData];
  } catch(e) { 
    return ["NO", e.toString()]; 
  }
}

/**
 * 用户自发修改登录密码逻辑
 * @param {string} oldPassword - 旧口令索引
 * @param {string} newPassword - 新目标口令
 * @returns {Array} 操作响应
 */
function writeChangePassword(oldPassword, newPassword) {
  try {
    const ss = getSheetCached(wsUserInfoId, "userID");
    const lr = ss.getLastRow();
    const data = lr > 0 ? ss.getRange(1, 1, lr, 3).getValues() : [];
    const ids = data.map(r => r[0]);
    const pos = ids.indexOf(oldPassword) + 1; // 1-based index
    
    safeSetValue(ss.getRange(pos, 3, 1, 1), newPassword);
    return ["OK", "更新成功"];
  } catch(e) { 
    return ["NO", e.toString()];
  }
}

/**
 * 精确查找物料与包装对应的完整 BOM 档案库
 * @param {string} materialInputId - 产品基础物料
 * @param {string} psku - 包装层 SKU
 * @returns {Array} 返回分离好的注塑和包装 BOM 数据字典
 */
function getBom(materialInputId, psku) {
  try {
    const ssTf = getSheetCached(wsTfBomId, "TF BOM masterdata Header");
    const tfLr = ssTf.getLastRow();
    const arrSsTf = tfLr > 1 ? ssTf.getRange(2, 1, tfLr - 1, 11).getValues() : [];

    const ssPk = getSheetCached(wsPkBomId, "BOM masterdata");
    const pkLr = ssPk.getLastRow();
    const arrSsPk = pkLr > 1 ? ssPk.getRange(2, 1, pkLr - 1, 11).getValues() : [];

    const arrSsTfSku = arrSsTf
      .filter(x => String(x[4]).indexOf(materialInputId.toString().split(",")[0]) > -1)
      .map(x => [x[0], x[1], x[2], x[3], x[7], x[8], x[10] ? Utilities.formatDate(new Date(x[10].toString()), "Asia/Shanghai", "yyyy-MM-dd") : "", x[5]]);
      
    let arrSsPkSku = [];
    if(psku.toString().split(",")[0].length > 4){
      arrSsPkSku = arrSsPk
        .filter(x => String(x[0]).indexOf(psku.toString().split(",")[0]) > -1 && x[8] == "生效")
        .map(x => [x[2], x[3], x[1], "", x[7], x[8], x[10] ? Utilities.formatDate(new Date(x[10].toString()), "Asia/Shanghai", "yyyy-MM-dd") : "", x[5]]);
    }
    
    return ["OK", arrSsTfSku, arrSsPkSku];
  } catch(e) { 
    return ["NO", e.toString()];
  }
}

/***********************
 * 业务函数（续）—— 缓存 + 并发保护 完整收尾
 ***********************/

/**
 * 判断是否已有设备参数周检记录
 * @param {string} nowDateShift - 日期与班次标识
 * @param {string} sku - 物料 SKU
 * @param {string} machine - 机台号
 * @returns {Array} 状态结果数组
 */
function checkWriteTechPpmsWeekCheck(nowDateShift, sku, machine) {
  try {
    const ssFcRecord = getSheetCached(wsTechPpmsWeekCheckId, "Injection");
    let arrSsFcRecord = [];
    const lr = ssFcRecord.getLastRow();
    const lc = ssFcRecord.getLastColumn();
    if(lr > 1000) {
      arrSsFcRecord = ssFcRecord.getRange(lr - 1000, 1, 1001, lc).getValues();
    } else {
      arrSsFcRecord = ssFcRecord.getRange(1, 1, lr, lc).getValues();
    }
    const arrMap = arrSsFcRecord.map(function(x) { return x[lc - 1] + x[3] + x[4]; });
    const key = nowDateShift + sku + machine;
    const pos = arrMap.indexOf(key);
    return pos == -1 ? ["OK", "开始读取数据……"] : ["NO", "已有该机台该品种的设备首件确认记录"];
  } catch(e) { 
    return ["NO", e.toString()]; 
  }
}

/**
 * 写入首件确认记录（遗留的老接口，并发保护+缓存）
 * @deprecated 建议前端统一使用 batchProcessFirstConfirm
 */
function writeFirstConfirm(userNameID, SAPCodeID, workshop, materialInputID, lineInputId, arrInjectionMaterial, arrDecalPrintMaterial, arrPkDate, arrPkMaterial, radioIf, arrInjParameter, arrKeyParameter, arrStdParameter, arrInjParaHistoryInputData, picShiftModelId, passSelectResons, reportSelectResons, exemptionResons, exemptionId, selfCheckNum, moldModel, mailBomIf, arrWarnParameter, needRemindIf, arrScanResult) {
  try {
    const wsFirstConfirm = getSpreadsheetCached(wsFirstConfirmId);
    const ssFcRecord = (lineInputId.charAt(0) != "S") ? wsFirstConfirm.getSheetByName("NonDecalPrint") : wsFirstConfirm.getSheetByName("DecalPrint");
    const now = new Date();
    const nowTimeStr = Utilities.formatDate(now, "Asia/Shanghai", "yyyy-MM-dd HH:mm:ss");
    const nowDate = Utilities.formatDate(now, "Asia/Shanghai", "yyyy-MM-dd");
    const nextDate = Utilities.formatDate(new Date(now.getTime() + 24 * 3600 * 1000), "Asia/Shanghai", "yyyy-MM-dd");
    const hm = Utilities.formatDate(now, "Asia/Shanghai", "HH:mm");
    let nowDateShift = "";
    if(hm > "23:44" && hm < "24:00") { nowDateShift = nextDate + "_1"; }
    else if(hm >= "00:00" && hm < "07:45") { nowDateShift = nowDate + "_1"; }
    else if(hm > "07:44" && hm < "15:45") { nowDateShift = nowDate + "_2"; }
    else if(hm > "15:44" && hm < "23:45") { nowDateShift = nowDate + "_3"; }

    const hasMultiSubmit = arrInjParaHistoryInputData.map(function(v) { return JSON.stringify(v[2]); }).some(function(s) { return s.indexOf("不") != -1; });
    if(hasMultiSubmit && reportSelectResons == "正常") {
      if (typeof SendGoogleChat !== "undefined" && SendGoogleChat.SendMessagePara) {
        SendGoogleChat.SendMessagePara(userNameID, workshop, materialInputID, lineInputId, arrInjParaHistoryInputData);
      }
    }
    if(arrWarnParameter.length > 0) { 
      SendNoStdMail(workshop, arrWarnParameter, mailBomIf, lineInputId, materialInputID, selfCheckNum); 
    }
    if((needRemindIf || []).join("【|】").indexOf("NeedRemind") != -1) {
      if (typeof SendGoogleChat !== "undefined" && SendGoogleChat.SendMessageRadio) {
        SendGoogleChat.SendMessageRadio(workshop, needRemindIf, lineInputId, materialInputID);
      }
    }

    const arrMachine = lineInputId.split(",");
    const arrMaterialSku = materialInputID.split(",");
    const outRows = [];
    for(let k = 0; k < arrMachine.length; k++){
      const m = arrMachine[k].substr(0, 8);
      const sku = arrMaterialSku[k];
      const row = [];
      row.push(userNameID, SAPCodeID, workshop, sku, m, nowTimeStr);
      if(lineInputId.charAt(0) != "S") {
        row.push(JSON.stringify(arrInjectionMaterial));
        for(let n = 0; n < arrPkDate.length; n++) { row.push(arrPkDate[n]); }
        row.push(JSON.stringify(arrPkMaterial));
      } else {
        row.push(JSON.stringify(arrDecalPrintMaterial));
      }
      for(let r = 0; r < radioIf.length; r++) { row.push(radioIf[r]); }
      let strParaInput = "[]";
      const ip = arrInjParameter.filter(function(v) { return v[0] == arrMachine[k]; }).map(function(v) { return v[1]; });
      if(ip.length > 0 && ip[0]) { strParaInput = JSON.stringify(ip[0]); }
      row.push(strParaInput);
      
      let cyc1 = "", cyc2 = "";
      const kp = arrKeyParameter.filter(function(v) { return v[0] == arrMachine[k]; }).map(function(v) { return v[1]; });
      if(kp.length > 0 && kp[0]) { 
        cyc1 = kp[0][0]; 
        cyc2 = kp[0][1]; 
      }
      row.push(cyc1, cyc2);
      
      let strStd = "[]";
      const sp = arrStdParameter.filter(function(v) { return v[0] == arrMachine[k]; }).map(function(v) { return v[1]; });
      if(sp.length > 0 && sp[0]) { 
        strStd = JSON.stringify(sp[0].map(function(v) { return [v[2], v[4], v[5]]; })); 
      }
      row.push(strStd);
      row.push(picShiftModelId ? ("https://drive.google.com/file/d/" + picShiftModelId + "/preview") : "");
      
      let strHis = "[]";
      const hp = arrInjParaHistoryInputData.filter(function(v) { return v[0] == arrMachine[k] || v[0] == "领班确认"; });
      if(hp.length > 0) {
        if(lineInputId.charAt(0) != "S" ? strStd.length > 2 : true) {
          strHis = JSON.stringify(hp.map(function(v) { return [v[1], v[2]]; }));
        }
      }
      row.push(strHis);
      
      const exRes = (String(exemptionResons).length < 3 ? "" : exemptionResons);
      row.push(moldModel, passSelectResons, reportSelectResons, exRes, exemptionId, nowDateShift);
      outRows.push(row);
    }
    
    for(let n = 0; n < outRows.length; n++) {
      if(outRows[n][7] != "") {
        SendGoogleChat.sendMessageInjOnline(outRows[n], arrScanResult);
      }
    }

    return withLock(function() {
      ssFcRecord.getRange(ssFcRecord.getLastRow() + 1, 1, outRows.length, outRows[0].length).setValues(outRows);
      return ["OK", "写入成功"];
    });
  } catch(e) { 
    return ["NO", e.toString()]; 
  }
}

/**
 * 写入设备参数周检停机状态（并发保护+缓存）
 */
function writeTechPpmsWeekCheckStopStatus(viewName, viewId, workshop, sku, machine, nowDateShift) {
  try {
    const ssFcRecord = getSheetCached(wsTechPpmsWeekCheckId, "Injection");
    const now = new Date();
    const nowStr = Utilities.formatDate(now, "Asia/Shanghai", "yyyy-MM-dd HH:mm:ss");
    const row = [viewName, viewId, workshop, sku, machine, nowStr, "", "", "", "", "", "", "停机", "正常", "", "", nowDateShift];
    return withLock(function() { 
      ssFcRecord.appendRow(row); 
      return ["OK", "写入成功"]; 
    });
  } catch(e) { 
    return ["NO", e.toString()]; 
  }
}

/**
 * 写入设备参数周检记录（并发保护+缓存）
 */
function writeTechPpmsWeekCheck(userNameID, SAPCodeID, workshop, materialInputID, lineInputId, arrInjParameter, passSelectResons, reportSelectResons, exemptionResons, exemptionId, emailIf, moldModel) {
  try {
    const ssFcRecord = getSheetCached(wsTechPpmsWeekCheckId, "Injection");
    const now = new Date();
    const nowTimeStr = Utilities.formatDate(now, "Asia/Shanghai", "yyyy-MM-dd HH:mm:ss");
    const nowDate = Utilities.formatDate(now, "Asia/Shanghai", "yyyy-MM-dd");
    const nextDate = Utilities.formatDate(new Date(now.getTime() + 24 * 3600 * 1000), "Asia/Shanghai", "yyyy-MM-dd");
    const hm = Utilities.formatDate(now, "Asia/Shanghai", "HH:mm");
    let nowDateShift = "";
    if(hm > "23:44" && hm < "24:00") { nowDateShift = nextDate + "_1"; }
    else if(hm >= "00:00" && hm < "07:45") { nowDateShift = nowDate + "_1"; }
    else if(hm > "07:44" && hm < "15:45") { nowDateShift = nowDate + "_2"; }
    else if(hm > "15:44" && hm < "23:45") { nowDateShift = nowDate + "_3"; }

    const row = [userNameID, SAPCodeID, workshop, materialInputID, lineInputId.substr(0, 8), nowTimeStr];
    for(let i = 0; i < arrInjParameter.length; i++) { row.push(arrInjParameter[i]); }
    row.push(passSelectResons, reportSelectResons, exemptionResons, exemptionId, nowDateShift);

    return withLock(function() {
      ssFcRecord.appendRow(row);
      if(emailIf == 1) {
        const toEmail = (workshop == "TB1") ? emailTB1Ops : emailTB2Ops;
        const ccEmail = emailEngCC;

        // 写安东（共享表）：使用缓存
        const sbnShiftToTeam = getSheetCached(getProp("shiftToTeamId", "1yF-3w6wYNhVOvicR2HmAJeVU2MdL3do5QtHSippsoCM"), "源数据");
        const lrST = sbnShiftToTeam.getLastRow();
        const arrSbnShiftToTeam = lrST > 0 ? sbnShiftToTeam.getRange(1, 1, lrST, sbnShiftToTeam.getLastColumn()).getValues() : [];
        const arrSbnShift = arrSbnShiftToTeam.map(function(x) { return x[0]; });

        const ssEquShift = getSheetCached(wsEquShiftId, "Shift_INJ_" + workshop);
        const sbnMachineToType = getSheetCached(getProp("equMasterId", "1bYKTK5a63yJWRHzM_UPP6b4hwF67eZKEM5dCKLWR59U"), "Workcenter & Mold Matrix");
        const lrMT = sbnMachineToType.getLastRow();
        const arrSbnMachineToType = lrMT > 0 ? sbnMachineToType.getRange(1, 1, lrMT, sbnMachineToType.getLastColumn()).getValues() : [];
        const arrSbnMachine = arrSbnMachineToType.map(function(x) { return x[0]; });

        const arrFollowUp = ["李华", "丁志伟"];
        const nowMs = now.getTime();
        const arrPpmsToEquShift = [
          nowMs, getTeam(nowDateShift, arrSbnShift, arrSbnShiftToTeam), lineInputId, 
          "工艺周检不在范围\n" + exemptionResons, "", "未解决", arrFollowUp[(workshop == "TB1") ? 0 : 1], 
          0, "M 其他-工艺", "", arrFollowUp[(workshop == "TB1") ? 0 : 1], 
          Utilities.formatDate(now, "Asia/Shanghai", "yyyy/MM/dd"), "FALSE", userNameID, workshop, "INJ", 0, 
          getMachineType2(lineInputId, arrSbnMachine, arrSbnMachineToType), "", "", "", 
          '=if(row()=1,"判断是否最后",if(countif(indirect("A1:A"&row()),indirect("A"&row()))=countif(indirect("A:A"),indirect("A"&row())),if(indirect("A"&row())<>"","Last",""),""))'
        ];
        ssEquShift.appendRow(arrPpmsToEquShift);

        // 邮件
        const widths = [45, 20, 35], head_backgrounds = ["unset", "unset", "unset"], aligns = ["center", "center", "center"], scale = 60, col = -1;
        const title = "设备参数周检，工艺参数不在标准范围内";
        const call = "<i>Auto send email, do not reply</i>";
        const describe1 = "以下工艺参数无法调整在标准范围内，请知悉！";
        const describe2 = "时间：" + nowTimeStr + ";机台：" + lineInputId + ";SKU：" + materialInputID + ";模具号：" + moldModel;
        const arrParaTemp = JSON.parse(exemptionResons || "[]");
        arrParaTemp.unshift(["工艺参数", "实际值", "备注"]);
        const describe3 = General_Htmltable(arrParaTemp, widths, head_backgrounds, aligns, scale, col);
        Mail_HTML_TXT(toEmail, title, call, describe1, describe2, describe3, "", "", "", ccEmail);
        return ["OK", "写入成功且豁免内容已邮给工程师"];
      } else {
        return ["OK", "写入成功"];
      }
    });
  } catch(e) { 
    return ["NO", e.toString()]; 
  }
}

/**
 * 依据排班表计算当前班次的团队名称
 * @param {string} str - 日期和班次索引的联合键
 * @param {Array} arrSbnShift - 排班第一列索引组
 * @param {Array} arrSbnShiftToTeam - 排班字典表
 * @returns {string} 班组名称，如 "白班"
 */
function getTeam(str, arrSbnShift, arrSbnShiftToTeam) {
  let team = ""; 
  if(str) { 
    const pos = arrSbnShift.indexOf(str); 
    if(pos != -1) { team = arrSbnShiftToTeam[pos][3] + "班"; } 
  }
  return team;
}

/**
 * 依据机台主数据提取机台类别
 * @param {string} str - 待查询的机台编码
 * @param {Array} arrSbnMachine - 索引列数组
 * @param {Array} arrSbnMachineToType - 机台字典
 * @returns {string} 机台组别类型
 */
function getMachineType2(str, arrSbnMachine, arrSbnMachineToType) {
  let machineType = "";
  if(str) { 
    const pos = arrSbnMachine.indexOf(str); 
    if(pos != -1) { machineType = arrSbnMachineToType[pos][6]; } 
  }
  return machineType;
}

/**
 * 计算目标机台在当前班次内已经做了几次“两小时自检”
 * @param {string} nowDateShift - 日期加班次的联合键
 * @param {string} machine - 目标机台号
 * @returns {Array} 状态结果数组与下一个次序号
 */
function countSelfCheckNum(nowDateShift, machine) {
  try {
    const wsFirstConfirm = getSpreadsheetCached(wsTwoHourSelfCheckId);
    const ss = (machine.charAt(0) != "S") ? wsFirstConfirm.getSheetByName("NonDecalPrint") : wsFirstConfirm.getSheetByName("DecalPrint");
    let arr = [];
    const lr = ss.getLastRow();
    const lc = ss.getLastColumn();
    if(lr > 1000) { arr = ss.getRange(lr - 1000, 1, 1001, lc).getValues(); } 
    else { arr = ss.getRange(1, 1, lr, lc).getValues(); }
    
    const filtered = arr.filter(function(x) { return x[lc - 2] == nowDateShift && x[4] == machine.substr(0, 8); });
    let countNum = 0;
    if(filtered.length > 0) { countNum = Number(filtered[filtered.length - 1][lc - 1]) || 0; }
    
    return (countNum < 4) ? ["OK", countNum + 1] : ["NO", "该机台已做了4次自检"];
  } catch(e) { 
    return ["NO", e.toString()]; 
  }
}

/**
 * 写入两小时自检记录（遗留的老接口，已被前端真·批量提交取代）
 * @deprecated 建议前端统一使用 batchProcessTwoHourSelfCheck
 */
function writeTwoHourSelfCheck(userNameID, SAPCodeID, workshop, materialInputID, lineInputId, arrInjectionMaterial, arrDecalPrintMaterial, arrPkDate, arrPkMaterial, radioIf, arrInjParameter, arrKeyParameter, arrStdParameter, arrInjParaHistoryInputData, picShiftModelId, passSelectResons, reportSelectResons, exemptionResons, exemptionId, selfCheckNum, moldModel, mailBomIf, arrWarnParameter, needRemindIf, arrScanResult) {
  try {
    const wsFirstConfirm = getSpreadsheetCached(wsTwoHourSelfCheckId);
    const ss = (lineInputId.charAt(0) != "S") ? wsFirstConfirm.getSheetByName("NonDecalPrint") : wsFirstConfirm.getSheetByName("DecalPrint");
    const now = new Date();
    const nowTimeStr = Utilities.formatDate(now, "Asia/Shanghai", "yyyy-MM-dd HH:mm:ss");
    const nowDate = Utilities.formatDate(now, "Asia/Shanghai", "yyyy-MM-dd");
    const nextDate = Utilities.formatDate(new Date(now.getTime() + 24 * 3600 * 1000), "Asia/Shanghai", "yyyy-MM-dd");
    const hm = Utilities.formatDate(now, "Asia/Shanghai", "HH:mm");
    let nowDateShift = "";
    if(hm > "23:44" && hm < "24:00") { nowDateShift = nextDate + "_1"; }
    else if(hm >= "00:00" && hm < "07:45") { nowDateShift = nowDate + "_1"; }
    else if(hm > "07:44" && hm < "15:45") { nowDateShift = nowDate + "_2"; }
    else if(hm > "15:44" && hm < "23:45") { nowDateShift = nowDate + "_3"; }

    const hasMultiSubmit = arrInjParaHistoryInputData.map(function(v) { return JSON.stringify(v[2]); }).some(function(s) { return s.indexOf("不") != -1; });
    if(hasMultiSubmit && reportSelectResons == "正常") {
      if (typeof SendGoogleChat !== "undefined" && SendGoogleChat.SendMessagePara) {
        SendGoogleChat.SendMessagePara(userNameID, workshop, materialInputID, lineInputId, arrInjParaHistoryInputData);
      }
    }
    if(arrWarnParameter.length > 0) { SendNoStdMail(workshop, arrWarnParameter, mailBomIf, lineInputId, materialInputID, selfCheckNum); }
    if((needRemindIf || []).join("【|】").indexOf("NeedRemind") != -1) {
      if (typeof SendGoogleChat !== "undefined" && SendGoogleChat.SendMessageRadio) {
        SendGoogleChat.SendMessageRadio(workshop, needRemindIf, lineInputId, materialInputID);
      }
    }

    const arrMachine = lineInputId.split(",");
    const arrSku = materialInputID.split(",");
    const outRows = [];
    
    for(let k = 0; k < arrMachine.length; k++) {
      const row = [];
      row.push(userNameID, SAPCodeID, workshop, arrSku[k], arrMachine[k].substr(0, 8), nowTimeStr);
      if(lineInputId.charAt(0) != "S") {
        row.push(JSON.stringify(arrInjectionMaterial));
        for(let n = 0; n < arrPkDate.length; n++) { row.push(arrPkDate[n]); }
        row.push(JSON.stringify(arrPkMaterial));
      } else {
        row.push(JSON.stringify(arrDecalPrintMaterial));
      }
      for(let r = 0; r < radioIf.length; r++) { row.push(radioIf[r]); }
      
      let ipStr = "[]";
      const ip = arrInjParameter.filter(function(v) { return v[0] == arrMachine[k]; }).map(function(v) { return v[1]; });
      if(ip.length > 0 && ip[0]) { ipStr = JSON.stringify(ip[0]); }
      row.push(ipStr);
      
      let c1 = "", c2 = "";
      const kp = arrKeyParameter.filter(function(v) { return v[0] == arrMachine[k]; }).map(function(v) { return v[1]; });
      if(kp.length > 0 && kp[0]) { c1 = kp[0][0]; c2 = kp[0][1]; }
      row.push(c1, c2);
      
      let stdStr = "[]";
      const sp = arrStdParameter.filter(function(v) { return v[0] == arrMachine[k]; }).map(function(v) { return v[1]; });
      if(sp.length > 0 && sp[0]) { stdStr = JSON.stringify(sp[0].map(function(v) { return [v[2], v[4], v[5]]; })); }
      row.push(stdStr);
      
      row.push(picShiftModelId ? ("https://drive.google.com/file/d/" + picShiftModelId + "/preview") : "");
      
      let hisStr = "[]";
      const hp = arrInjParaHistoryInputData.filter(function(v) { return v[0] == arrMachine[k] || v[0] == "领班确认"; });
      if(hp.length > 0) {
        if(lineInputId.charAt(0) != "S" ? stdStr.length > 2 : true) {
          hisStr = JSON.stringify(hp.map(function(v) { return [v[1], v[2]]; }));
        }
      }
      row.push(hisStr);
      
      const exRes = (String(exemptionResons).length < 3 ? "" : exemptionResons);
      row.push(moldModel, passSelectResons, reportSelectResons, exRes, exemptionId, nowDateShift, selfCheckNum);
      outRows.push(row);
    }
    
    for(let n = 0; n < outRows.length; n++) {
      safeAppendRow(ss, outRows[n]);
      if(outRows[n][7] != "" && outRows[n][4].indexOf("S1") == -1 && outRows[n][4].indexOf("S2") == -1) {
        SendGoogleChat.sendMessageInjOnline(outRows[n], arrScanResult);
      }
    }
    return ["OK", outRows];
  } catch(e) { 
    return ["NO", e.toString()]; 
  }
}

/**
 * 自检未找到标准工艺/BOM 邮件推送警告
 * @param {string} workshop - 车间名
 * @param {Array} arrWarnParameter - 包含缺漏信息的数组
 * @param {string} mailBomIf - 是否缺失 BOM
 * @param {string} lineInputId - 机台编码
 * @param {string} materialInputID - SKU编码
 * @param {number} selfCheckNum - 自检次数
 */
function SendNoStdMail(workshop, arrWarnParameter, mailBomIf, lineInputId, materialInputID, selfCheckNum) {
  if(selfCheckNum < 2) {
    arrWarnParameter = arrWarnParameter.filter(function(v) { return v[1] != "FT" && v[2] != "Automation_6Aixs"; });
    if(arrWarnParameter.length > 0) {
      const nowTime = Utilities.formatDate(new Date(), "Asia/Shanghai", "yyyy-MM-dd HH:mm:ss");
      const toEmail = (workshop == "TB1") ? "yong_ji@colpal.com,xingbo_chen@colpal.com" : "andy_lu@colpal.com,xingbo_chen@colpal.com";
      const ccEmail = "jin_zhang@colpal.com,fiona_zhao@colpal.com";
      const title = "自检未找到工艺参数";
      const call = "<i>Auto send email, do not reply</i>";
      const desc1 = "以下机台品种及模具未找到工艺参数，请知悉！";
      const desc2 = "时间：" + nowTime + "；机台：" + arrWarnParameter.map(function(v) { return v[0]; }).join(",") + "；机型：" + arrWarnParameter.map(function(v) { return v[1]; }).join(",") + "；自动化类型：" + arrWarnParameter.map(function(v) { return v[2]; }).join(",") + "；模具：" + arrWarnParameter.map(function(v) { return v[3]; }).join(",") + "；SKU：" + arrWarnParameter.map(function(v) { return v[4]; }).join(",");
      Mail_HTML_TXT(toEmail, title, call, desc1, desc2, "", "", "", "", ccEmail);
    }
  }
  if(mailBomIf == "Y") {
    const nowTime = Utilities.formatDate(new Date(), "Asia/Shanghai", "yyyy-MM-dd HH:mm:ss");
    const toEmail = "fiona_zhao@colpal.com";
    const ccEmail = "jin_zhang@colpal.com";
    const title = "自检未找到Bom";
    const call = "<i>Auto send email, do not reply</i>";
    const desc1 = "以下机台品种第一次自检时未找到BOM，请知悉！";
    const desc2 = "时间：" + nowTime + ";机台：" + lineInputId + ";SKU：" + materialInputID;
    Mail_HTML_TXT(toEmail, title, call, desc1, desc2, "", "", "", "", ccEmail);
  }
}

/**
 * 写入计划产出与历史PO信息回填
 * @param {string} rowNo - 更新的行标识
 * @param {string} strPo - 采购订单号
 * @param {number} tempTotalOut - 计划总产出
 * @param {Object} formValues - 表单回传的具体键值对
 * @param {Array} arrHistoryPoInfo - 包含该批次历史快照的记录
 * @returns {Array} 更新成功并返回最新表数据
 */
function writeOutputPcs(rowNo, strPo, tempTotalOut, formValues, arrHistoryPoInfo) {
  try {
    const ssPlan = getSheetCached(wsPlanId, "All_INJ");
    const lr = ssPlan.getLastRow();
    const arrPlan = ssPlan.getRange(1, 1, Math.max(lr - 1, 0), 22).getValues();
    const arrRowNoStrPo = arrPlan.map(function(v) { return String(v[0]) + String(v[1]).replace(/\n/g, ""); });
    const pos = arrRowNoStrPo.indexOf(rowNo + "" + strPo);
    if(pos != -1) {
      return withLock(function() {
        safeSetValue(ssPlan.getRange(pos + 1, 16, 1, 1), tempTotalOut);
        safeSetValue(ssPlan.getRange(pos + 1, 18, 1, 1), JSON.stringify(formValues));
        safeSetValue(ssPlan.getRange(pos + 1, 22, 1, 1), JSON.stringify(arrHistoryPoInfo));
        const arrWsPlan = arrPlan.filter(function(v) { return v[19] == "TB1"; }).map(function(x) { return [x[0], x[1], Utilities.formatDate(new Date(x[2]), "Asia/Shanghai", "yyyy-MM-dd"), x[3], x[4], x[5], x[7], x[8], x[6], x[9], x[10], x[11], x[12], x[13], x[14], x[15], x[16], x[17], x[20], x[21]]; });
        return ["OK", arrWsPlan, "保存成功"];
      });
    } else {
      return ["NO", "未找到该行记录索引信息，确认计划更新后，你是否已刷新页面？"];
    }
  } catch(e) { 
    return ["NO", e.toString()]; 
  }
}

/**
 * 获取设备专属维护的工艺参数记录集（含范围上下限及校验部门信息）
 * @param {string} machineType - 机器主类型
 * @param {string} machineModel - 机器子型号分组
 * @param {string} moldModel - 模具类型标识
 * @returns {Array} 字典化的详细参数规格要求
 * @暂时没用到，直接读DB缓存
 */
function readTechParaRecord(machineType, machineModel, moldModel) {
  try {
    const arrJsonData = getProcessParameterData(wsInjParaId, "INJ_New");
    const arrSsInjParaJson = arrJsonData.filter(function(v) { return v.模具编码 == moldModel && v.状态 == "生效"; });
    const arrData = [];
    for(let i = 0; i < arrSsInjParaJson.length; i++) {
      const arrParaJsonTemp = JSON.parse(arrSsInjParaJson[i]["工艺参数"]);
      for(let j = 0; j < arrParaJsonTemp.length; j++) {
        if(String(arrParaJsonTemp[j]["检查部门"]).length > 0 && String(arrParaJsonTemp[j]["设定下限"]) != "9999") {
          arrData.push([
            replaceBlankToStr(arrSsInjParaJson[i]["机型"]),
            replaceBlankToStr(arrSsInjParaJson[i]["自动化分组"]),
            replaceBlankToStr(arrSsInjParaJson[i]["BigBundle"]),
            replaceBlankToStr(arrSsInjParaJson[i]["SKU"]),
            replaceBlankToStr(arrSsInjParaJson[i]["模穴数"]),
            replaceBlankToStr(arrSsInjParaJson[i]["分序类别"]),
            replaceBlankToStr(arrParaJsonTemp[j]["工艺参数"]),
            replaceBlankToStr(arrParaJsonTemp[j]["单位"]),
            replaceBlankToStr(arrParaJsonTemp[j]["设定下限"]),
            replaceBlankToStr(arrParaJsonTemp[j]["设定上限"]),
            replaceBlankToStr(arrParaJsonTemp[j]["检查部门"]),
            replaceBlankToStr(arrSsInjParaJson[i]["工艺卡编号"])
          ]);
        }
      }
    }
    const arrMachineType = Array.from(new Set(arrData.map(function(x) { return x[0]; })));
    const arrMachineModel = Array.from(new Set(arrData.map(function(x) { return x[1]; })));
    const arrBigBundle = Array.from(new Set(arrData.map(function(x) { return x[2]; })));
    const arrActSku = Array.from(new Set(arrData.map(function(x) { return x[3]; })));
    const arrCavityNumber = Array.from(new Set(arrData.map(function(x) { return x[4]; })));
    const arrClassification = Array.from(new Set(arrData.map(function(x) { return x[5]; })));
    const arrCheckDept = Array.from(new Set(arrData.map(function(x) { return x[10]; })));
    const arrProcessCardNum = Array.from(new Set(arrData.map(function(x) { return x[11]; })));
    return ["OK", arrData, arrMachineType, arrMachineModel, arrBigBundle, arrActSku, arrCavityNumber, arrClassification, arrCheckDept, arrProcessCardNum];
  } catch(e) { 
    return ["NO", e.toString()]; 
  }
}

/**
 * 工具函数：如果是空字符串或 null，转化为全角空格以占位
 * @param {string} strTemp - 源字符串
 * @returns {string} 处理后字符串
 * @暂时没有用到
 */
function replaceBlankToStr(strTemp) { 
  return (strTemp != "" && strTemp != null) ? strTemp : "　"; 
}

/**
 * 末模状态留存：将最后模次的机台状态提交保存（并发保护）
 * @param {Array} arrData - 需要插入的末模记录行
 * @returns {Array} 操作响应
 */
function writeCollectionData(arrData) {
  try {
    const ss = getSheetCached(wsFinalMoldCheckId, "FinalMoldCollectionRecord");
    const lr = ss.getLastRow();
    let arr = ss.getRange(1, 1, lr, ss.getLastColumn()).getValues();
    arr = arr.map(function(r, i) { r.push(i + 1); return r; });
    const keys = arr.map(function(x) { return x[0] + x[1] + x[2] + x[3]; });
    const idx = keys.indexOf(arrData[0] + arrData[1] + arrData[2] + arrData[3]);
    
    return withLock(function() {
      if(idx == -1) {
        ss.appendRow(arrData);
        return ["OK", "插入成功"];
      } else {
        ss.getRange(arr[idx][arr[0].length - 1], 1, 1, arrData.length).setValues([arrData]);
        return ["OK", "保存成功"];
      }
    });
  } catch(e) { 
    return ["NO", e.toString()]; 
  }
}

/**
 * 获取近期未完成的末模提交收集记录字典
 * @returns {Array} 末模收集数据的集合
 */
function getCollectionData() {
  try {
    const ss = getSheetCached(wsFinalMoldCheckId, "FinalMoldCollectionRecord");
    const lr = ss.getLastRow();
    const arr = ss.getRange(1, 1, lr, ss.getLastColumn()).getValues();
    for(let i = 0; i < arr.length; i++) {
      arr[i].unshift(i + 1);
      arr[i][7] = Utilities.formatDate(new Date(arr[i][7]), "Asia/Shanghai", "yyyy-MM-dd HH:mm:ss");
      arr[i][14] = Utilities.formatDate(new Date(arr[i][14]), "Asia/Shanghai", "yyyy-MM-dd");
    }
    const out = arr.filter(function(x) { return (x[15] != "Done" && x[15] != "维修状态" && x[16] != "") || (x[2] == ""); }).reverse();
    return out.length > 0 ? ["OK", "提取成功", out] : ["NO", "未找到记录"];
  } catch(e) { 
    return ["NO", e.toString()]; 
  }
}

/**
 * 对处于异常末模修复阶段的机台提交修复报告与状态变更
 * @param {number} rowNum - 记录在表里的行号
 * @param {string} repairMoldPlate - 维修更换件信息
 * @param {string} repairDes - 维修描述详情
 * @param {string} repairStaff - 维修工程师姓名
 * @param {string} repairDate - 修复完成的日期
 * @param {string} repairStatus - 更新后的最终状态（如Done）
 * @param {string} tbodyId - 前端回显关联用的 Table Body ID
 * @param {number} trSeq - 前端回显关联用的 TR 序号
 * @returns {Array} 保存成功的通知数组
 */
function writeRepairMoldData(rowNum, repairMoldPlate, repairDes, repairStaff, repairDate, repairStatus, tbodyId, trSeq) {
  try {
    const ss = getSheetCached(wsFinalMoldCheckId, "FinalMoldCollectionRecord");
    return withLock(function() {
      safeSetValue(ss.getRange(rowNum, 2, 1, 1), repairMoldPlate);
      safeSetValues(ss.getRange(rowNum, 12, 1, 4), [[repairDes, repairStaff, repairDate, repairStatus]]);
      return ["OK", tbodyId, trSeq];
    });
  } catch(e) { 
    return ["NO", e.toString()]; 
  }
}

/**
 * 补退料数据深度查询引擎
 * @param {Array} arrSarmPlanData - 当前总览的基础计划数据
 * @param {string} machine - 选择过滤的目标机台
 * @param {string} sku - 附带指定的子料 SKU 代码
 * @returns {Array} 结构化分析好的物料记录集合（包含近三个班次与下一班次追踪）
 */
function getSarmData(arrSarmPlanData, machine, sku) {
  try {
    const sabiId_Sarm = "1RH2rmG9zf1ZzwS_25IpXWdYSGIx_EhSud-rWDumC9JA";
    const sbnName_Sarm = "Data";
    const sabi_Sarm = SpreadsheetApp.openById(sabiId_Sarm);
    const sbn_Sarm = sabi_Sarm.getSheetByName(sbnName_Sarm);
    
    const sabiId_PreBom = "1jQBJTzwfUZgX2-xn4HgQ4pvDhfqHgkRI0QHZLL0Tr-4";
    const sbnName_PreBom = "★BOM★";
    const sabi_PreBom = SpreadsheetApp.openById(sabiId_PreBom);
    const sbn_PreBom = sabi_PreBom.getSheetByName(sbnName_PreBom);
    
    let arrPreBom = sbn_PreBom.getRange(1, 1, sbn_PreBom.getLastRow(), sbn_PreBom.getLastColumn()).getDisplayValues();
    let arrPreBom_Sku = arrPreBom.filter(function(v) { return String(v[0]).indexOf(sku) != -1; });
    
    let arrSarmPlanData_Machine_Prev3Next1 = [];
    let arrWhSarm_Machine_Prev4 = [];
    let nowDateShift = "";
    let nowDateTime = "";
    
    let now = new Date();
    let nowTime = now.getTime();
    let prevDate = Utilities.formatDate(new Date(nowTime - 24 * 3600 * 1000), "Asia/Shanghai", "yyyy-MM-dd");
    let nowDate = Utilities.formatDate(now, "Asia/Shanghai", "yyyy-MM-dd");
    let nextDate = Utilities.formatDate(new Date(nowTime + 24 * 3600 * 1000), "Asia/Shanghai", "yyyy-MM-dd");
    let nowHour = now.getHours();
    let nowMinute = now.getMinutes();
    
    let arrPrevDateShift = [];
    
    if((nowHour == 7 && nowMinute >= 45) || (nowHour >= 8 && nowHour < 15) || (nowHour == 15 && nowMinute < 45)) { 
      nowDateShift = nowDate + "_2"; 
      nowDateTime = nowDate + " 08:00:00"; 
    } else if((nowHour == 15 && nowMinute >= 45) || (nowHour >= 16 && nowHour < 23) || (nowHour == 23 && nowMinute < 45)) { 
      nowDateShift = nowDate + "_3"; 
      nowDateTime = nowDate + " 16:00:00"; 
    } else if(nowHour == 23 && nowMinute >= 45) { 
      nowDateShift = nextDate + "_1"; 
      nowDateTime = nextDate + " 00:00:00"; 
    } else if((nowHour >= 0 && nowHour < 7) || (nowHour == 7 && nowMinute < 45)) { 
      nowDateShift = nowDate + "_1"; 
      nowDateTime = nowDate + " 00:00:00"; 
    }
    
    let prevDateTime = "";
    let prevDateTime_Date = "";
    let prevDateTime_Hour = 0;
    
    for(let i = 0; i < 5; i++) {
      prevDateTime = Utilities.formatDate(new Date(new Date(nowDateTime).getTime() - i * 8 * 3600 * 1000), "Asia/Shanghai", "yyyy-MM-dd HH:mm:ss");
      prevDateTime_Date = Utilities.formatDate(new Date(new Date(nowDateTime).getTime() - i * 8 * 3600 * 1000), "Asia/Shanghai", "yyyy-MM-dd");
      prevDateTime_Hour = new Date(prevDateTime).getHours();
      
      if(prevDateTime_Hour == 0) {
        arrPrevDateShift.push(prevDateTime_Date + "_1");
      } else if(prevDateTime_Hour == 8) {
        arrPrevDateShift.push(prevDateTime_Date + "_2");
      } else if(prevDateTime_Hour == 16) {
        arrPrevDateShift.push(prevDateTime_Date + "_3");
      }
    }
    console.log(arrPrevDateShift);
    
    let arrSarmPlanData_Machine = arrSarmPlanData.filter(function(v) { return v[4] == machine && v[16] > 0; }).map(function(v) { return [v[2], v[3], v[4], v[5], v[6], v[7], v[8], v[9], v[10], v[16]]; });
    
    if(arrSarmPlanData_Machine.length > 0) {
      arrSarmPlanData_Machine.forEach(function(v) {
        v[0] = Utilities.formatDate(new Date(v[0]), "Asia/Shanghai", "yyyy-MM-dd");
        if(v[1] == "早") { v.push(v[0] + "_2"); }
        else if(v[1] == "中") { v.push(v[0] + "_3"); }
        else if(v[1] == "夜") { v.push(v[0] + "_1"); }
        else { v.push(v[0] + "_0"); }
      });
      
      let arrSarmPlanData_Machine_DateShift = Array.from(new Set(arrSarmPlanData_Machine.map(function(v) { return v[10]; }))).sort(function(a, b) { if(a > b) { return 1; } else if(a == b) { return 0; } else if(a < b) { return -1; } });
      let arrSarmPlanData_Machine_Need_DateShift = [];
      let arrSarmPlanData_Machine_Prev_DateShift = arrSarmPlanData_Machine_DateShift.filter(function(v) { return v < nowDateShift; });
      
      for(let i = arrSarmPlanData_Machine_Prev_DateShift.length - 1; i > -1; i--) {
        arrSarmPlanData_Machine_Need_DateShift.push(arrSarmPlanData_Machine_Prev_DateShift[i]);
        if(arrSarmPlanData_Machine_Need_DateShift.length > 0) { break; }
      }
      
      let arrSarmPlanData_Machine_Next_DateShift = arrSarmPlanData_Machine_DateShift.filter(function(v) { return v > nowDateShift; });
      for(let i = 0; i < arrSarmPlanData_Machine_Next_DateShift.length; i++) {
        arrSarmPlanData_Machine_Need_DateShift.push(arrSarmPlanData_Machine_Next_DateShift[i]);
        if(arrSarmPlanData_Machine_Need_DateShift.length > 2) { break; }
      }
      
      arrSarmPlanData_Machine_Need_DateShift.push(nowDateShift);
      arrSarmPlanData_Machine_Prev3Next1 = arrSarmPlanData_Machine.filter(function(v) { return arrSarmPlanData_Machine_Need_DateShift.indexOf(v[10]) != -1; });
      arrSarmPlanData_Machine_Prev3Next1.forEach(function(v) { v.push(getPreNoFromSku(v[4], arrPreBom)); });
    }
    
    let arrWhSarm = [];
    let sbn_SarmLr = sbn_Sarm.getLastRow();
    let startRow = 1;
    if(sbn_SarmLr > 500) { startRow = sbn_SarmLr - 499; } else if(sbn_SarmLr > 1) { startRow = 2; }
    if(startRow > 1) {
      arrWhSarm = sbn_Sarm.getRange(startRow, 1, sbn_SarmLr - startRow + 1, 15).getDisplayValues();
    }
    arrWhSarm.forEach(function(v, index) {
      v[1] = Utilities.formatDate(new Date(v[1]), "Asia/Shanghai", "yyyy-MM-dd");
      v.push(v[1] + "_" + v[2]);
    });
    
    let arrWhSarm_Machine = arrWhSarm.filter(function(v) { return v[5] == machine; });  
    let arrWhSarm_Machine_DateShift = arrWhSarm_Machine.map(function(v) { return v[15]; }).sort(function(a, b) { if(a > b) { return 1; } else if(a == b) { return 0; } else if(a < b) { return -1; } });
    let arrWhSarm_Machine_Prev_DateShift = arrWhSarm_Machine_DateShift.filter(function(v) { return v < nowDateShift; });
    let arrWhSarm_Machine_Prev4_DateShift = [];
    
    for(let i = arrWhSarm_Machine_Prev_DateShift.length - 1; i > -1; i--) {
      arrWhSarm_Machine_Prev4_DateShift.push(arrWhSarm_Machine_Prev_DateShift[i]);
      if(arrWhSarm_Machine_Prev4_DateShift.length > 4) { break; }
    }
    
    arrWhSarm_Machine_Prev4 = arrWhSarm_Machine.filter(function(v) { return arrPrevDateShift.indexOf(v[15]) != -1; });    
    let arrPreNo = [];
    if(arrPreBom_Sku.length > 0) { arrPreNo = arrPreBom_Sku.map(function(v) { return v[5]; }); }
    let arrWhSarm_NowDateShift = arrWhSarm.filter(function(v) { return v[15] == nowDateShift; });
    
    console.log(sku, arrPreNo);
    return ["OK", nowDateShift, arrSarmPlanData_Machine_Prev3Next1, arrWhSarm_Machine_Prev4, arrPreNo, arrWhSarm_NowDateShift];
  }
  catch(e) {
    console.log(e.toString());
    return ["NO", e.toString()];
  }
}

/**
 * 辅助：根据商品 SKU 在旧版记录提取其配料码 PRENo
 * @param {string} sku - 商品 SKU
 * @param {Array} arrPreBom - 所有的子料 BOM 清单
 * @returns {string} 匹配成功的预配子码结果
 */
function getPreNoFromSku(sku, arrPreBom) {
  let result = ""; 
  const arrTemp = arrPreBom.filter(function(v) { return String(v[0]).indexOf(sku) != -1; });
  if(arrTemp.length > 0) { result = Array.from(new Set(arrTemp.map(function(v) { return v[5]; }))).join("|"); }
  return result;
}

/**
 * 补/领料信息直接推送到仓库处理报表（并发保护）
 * @param {Array} arrSarmSubmitData - 用户填写的待申报发料需求行
 * @returns {Array} 执行结果响应通知
 */
function saveSarmInfo(arrSarmSubmitData) {
  try {
    const sbn_Sarm = getSheetCached(getProp("sabiId_Sarm", "1RH2rmG9zf1ZzwS_25IpXWdYSGIx_EhSud-rWDumC9JA"), "Data");
    const lr = sbn_Sarm.getLastRow();
    let startRow = 1;
    if(lr > 200) { startRow = lr - 199; } else if(lr > 1) { startRow = 2; }
    
    let arrExisting = [];
    if(startRow > 1) { arrExisting = sbn_Sarm.getRange(startRow, 1, lr - startRow + 1, 15).getDisplayValues(); }
    
    const exists = arrExisting.filter(function(v) { return Utilities.formatDate(new Date(v[1]), "Asia/Shanghai", "yyyy-MM-dd") + v[2] + v[5] + v[7] + v[8] == arrSarmSubmitData[1] + arrSarmSubmitData[2] + arrSarmSubmitData[5] + arrSarmSubmitData[7] + arrSarmSubmitData[8]; });
    
    if(exists.length > 0) { return ["NO", "已经有当班机台该物料的【" + arrSarmSubmitData[8] + "】记录"]; }
    
    return withLock(function() { 
      sbn_Sarm.appendRow(arrSarmSubmitData); 
      return ["OK", "已成功提交"]; 
    });
  } catch(e) { 
    return ["NO", e.toString()]; 
  }
}

/**
 * 操作员提交生产退料信息保存库（并发保护与重复行阻断）
 * @param {Array} arrData - 二维数组结构，涵盖多笔同机退料行
 * @returns {Array} 将错误/成功信息归类提示返回前端
 */
function saveRetuInfo(arrData) {
  try {
    const sbn_Sarm = getSheetCached(getProp("sabiId_Sarm", "1RH2rmG9zf1ZzwS_25IpXWdYSGIx_EhSud-rWDumC9JA"), "Data");
    const lr = sbn_Sarm.getLastRow();
    let startRow = 1;
    if(lr > 200) { startRow = lr - 199; } else if(lr > 1) { startRow = 2; }
    
    let arrExisting = [];
    if(startRow > 1) { arrExisting = sbn_Sarm.getRange(startRow, 1, lr - startRow + 1, 15).getDisplayValues(); }
    
    let warn = [];
    return withLock(function() {
      for(let i = 0; i < arrData.length; i++) {
        const exists = arrExisting.filter(function(v) { return Utilities.formatDate(new Date(v[1]), "Asia/Shanghai", "yyyy-MM-dd") + v[2] + v[5] + v[7] + v[8] == arrData[i][1] + arrData[i][2] + arrData[i][5] + arrData[i][7] + arrData[i][8]; });
        if(exists.length > 0) {
          warn.push("已有" + arrData[i][5] + "　" + arrData[i][7] + "　" + arrData[i][8] + "　" + arrData[i][9] + "的退料记录");
        } else {
          sbn_Sarm.appendRow(arrData[i]);
        }
      }
      if(warn.length > 0) { return ["OK", "以下机台物料已经存在退料记录，未保存，其它已保存\n" + warn.join("\n")]; }
      return ["OK", "已成功提交"];
    });
  } catch(e) { 
    return ["NO", e.toString()]; 
  }
}

/**
 * 万能读取工艺参数工具
 * @param {string} sabiId - 要读取的标准表格 ID
 * @param {string} sbnName - 要提取的 Sheet 页名
 * @returns {Array} JSON对象格式组成的提取库
 */
function getProcessParameterData(sabiId, sbnName) {
  const sbn = getSheetCached(sabiId, sbnName);
  const lr = sbn.getLastRow();
  const lc = sbn.getLastColumn();
  if(lr <= 1) return [];
  
  const arrHead = sbn.getRange(1, 1, 1, lc).getValues();
  const arrBody = sbn.getRange(2, 1, lr - 1, lc).getValues();
  const arrJson = [];
  
  arrBody.forEach(function(row, i) {
    const obj = {"行号": i + 2};
    row.forEach(function(val, j) { obj[arrHead[0][j]] = val; });
    arrJson.push(obj);
  });
  
  return arrJson;
}

/**
 * 数组字符串序列化组装
 * @param {Array} arr - 二维数组目标体
 * @returns {string} 以回车换行和横杠分隔组装后的文本
 */
function arrToTxt(arr) {
  let txt = "";
  for(let i = 0; i < arr.length; i++) {
    if(i > 0) { txt += "\n"; }
    for(let j = 0; j < arr[i].length; j++) {
      txt += (j > 0 ? "|" : "") + arr[i][j];
    }
  }
  return txt;
}

/**
 * 将规则文本分割还原为数组阵列
 * @param {string} txt - 需分割还原的长文本
 * @returns {Array} 标准切割分离后的二维数组
 */
function txtToArr(txt) {
  return String(txt || "").split("\n").map(function(row) { return row.split("|"); });
}

/**
 * [前端UI用] 把数据数组根据 UI 设定长度，横向排开
 * @param {Array} arrSs - 原始提取出来的扁平列列数组
 * @param {number} splitCol - 最多在表格中放多少大列
 * @returns {Array} 布局好的 UI-ready 表格用结构数据
 */
function VHDeformation(arrSs, splitCol) {
  let arrTemp = arrSs.map(function(r) { return r.slice(); });
  let rowTotal = Math.ceil(arrTemp.length / splitCol);
  let arrData = [];
  if(arrTemp.length < 4) {
    let r = 0, l = 0; rowTotal = 1;
    for(let i = 0; i < arrTemp.length; i++) {
      if(l == 0) { arrData.push([arrTemp[i][0], arrTemp[i][1], ""]); r = r + 1; }
      else { arrData[r].push(arrTemp[i][0], arrTemp[i][1], ""); }
      if(r >= rowTotal) { r = 0; l = l + 1; }
    }
  } else {
    for(let n = arrTemp.length; n < rowTotal * splitCol; n++) { arrTemp.push(["", "", ""]); }
    let r = 0, l = 0;
    for(let i = 0; i < rowTotal * splitCol; i++) {
      if(l == 0) { arrData.push([arrTemp[i][0], arrTemp[i][1], ""]); r = r + 1; }
      else { arrData[r].push(arrTemp[i][0], arrTemp[i][1], ""); r = r + 1; }
      if(r >= rowTotal) { r = 0; l = l + 1; }
    }
  }
  return arrData;
}

/**
 * [前端UI用] 第二种扩展类型的并排重整法 (比如首件/自检模块特有需要带单位)
 * @param {Array} arrSs - 原始列表数据
 * @param {number} splitCol - 平分长度
 * @returns {Array} 重构组合好的二维数组
 */
function VHDeformation2(arrSs, splitCol) {
  let arrTemp = arrSs.map(function(r) { return r.slice(); });
  let rowTotal = Math.ceil(arrTemp.length / splitCol);
  let arrData = [];
  if(arrTemp.length < 4) {
    let r = 0, l = 0; rowTotal = 1;
    for(let i = 0; i < arrTemp.length; i++) {
      if(l == 0) { arrData.push([arrTemp[i][0], arrTemp[i][1], ""]); l = l + 1; }
      else { arrData[r].push(arrTemp[i][0], arrTemp[i][1], ""); l = l + 1; }
      if(l >= splitCol) { r = r + 1; l = 0; }
    }
  } else {
    for(let n = arrTemp.length; n < rowTotal * splitCol; n++) { arrTemp.push(["", "", ""]); }
    let r = 0, l = 0;
    for(let i = 0; i < rowTotal * splitCol; i++) {
      if(l == 0) { arrData.push([arrTemp[i][0], arrTemp[i][1], ""]); l = l + 1; }
      else { arrData[r].push(arrTemp[i][0], arrTemp[i][1], ""); l = l + 1; }
      if(l >= splitCol) { r = r + 1; l = 0; }
    }
  }
  return arrData;
}

/**
 * 将平铺的数据切分重构为竖排数列
 * @param {Array} arrSs - 含平铺组的二维原数
 * @param {number} splitCol - 切割深度与复现长度
 * @returns {Array} 还原重构结果
 */
function HVDeformation(arrSs, splitCol) {
  let arrData = [];
  for(let j = 0; j < arrSs[0].length; j = j + 3) {
    for(let i = 0; i < arrSs.length; i++) {
      arrData.push([arrSs[i][j + 0]]);
      for(let k = 1; k < splitCol; k++) { arrData[arrData.length - 1].push(arrSs[i][j + k]); }
    }
  }
  return arrData;
}

/**
 * 带有特殊偏置量（针对参数值）竖向强制切片提取法
 * @param {Array} arrSs - 含平铺组的二维原数
 * @param {number} splitCol - 间隔频率
 * @returns {Array} 解析复刻结果
 */
function HVDeformation2(arrSs, splitCol) {
  let arrData = [];
  for(let i = 0; i < arrSs.length; i++) {
    for(let j = 0; j < arrSs[0].length; j = j + 3) {
      arrData.push([arrSs[i][j + 0]]);
      for(let k = 1; k < splitCol; k++) { arrData[arrData.length - 1].push(arrSs[i][j + k]); }
    }
  }
  return arrData;
}

/**
 * [工具] 去除基于第一列判定的重复脏数据行
 * @param {Array} arrData - 二维原数组
 * @returns {Array} 洗净后的孤立行集合
 */
function delRepeatItem(arrData) {
  const seen = new Set(); 
  const out = [];
  for(const r of arrData) {
    const k = r[0];
    if(!seen.has(k)) { seen.add(k); out.push([k]); }
  }
  return out;
}

/**
 * 校验目标短字符是否存在于大型字典集合字符串结构内
 * @param {Array} arr - 字典库（长关键字）
 * @param {string} str - 需要对比查岗的目标小切片
 * @returns {boolean} 在内部出现则成立
 */
function checkTrue(arr, str) {
  const s = String(str || "");
  return arr.some(function(x) { return s.indexOf(String(x)) != -1; });
}

/**
 * 后端用以在邮件或警示回显内拼接渲染自定义 Table 样式的利器
 * @param {Array} data - 被嵌套的基础数据
 * @param {Array} widths - 控制每列宽度的占比设定
 * @param {Array} head_backgrounds - 头部背景控制
 * @param {Array} aligns - 排版控制 (居中、靠左等)
 * @param {number} scale - 占所在屏幕大小宽度的缩放级别比 (0-100)
 * @param {number} col - 若哪列想被隐藏剔除不显示，则填它的索引标识
 * @returns {string} 返回整段合法的 Table Html
 */
function General_Htmltable(data, widths, head_backgrounds, aligns, scale, col) {
  let htmltable = "<table border='2' style='table-layout:fixed;margin-left:30px;border-collapse:separate;font-size:12px;width:" + scale + "%;'>";
  for(let i = 0; i < data.length; i++) {
    htmltable += "<tr>";
    for(let j = 0; j < data[i].length; j++) {
      if(i == 0 && j != col) {
        htmltable += "<td style='width:" + widths[j] + "%;text-align:center;word-break:break-all;background:" + head_backgrounds[j] + "'>" + String(data[i][j]).replace(/\n/g, "<br>") + "</td>";
      } else if(j == col) {
        htmltable += "<td style='width:" + widths[j] + "%;text-align:" + aligns[j] + ";word-break:break-all;display:none'>" + String(data[i][j]).replace(/\n/g, "<br>") + "</td>";
      } else {
        htmltable += "<td style='width:" + widths[j] + "%;text-align:" + aligns[j] + ";word-break:break-all'>" + String(data[i][j]).replace(/\n/g, "<br>") + "</td>";
      }
    }
    htmltable += "</tr>";
  }
  return htmltable + "</table>";
}

/**
 * 封装好的 Gmail 投递邮件自动化机器
 * @param {string} tomail - 接收人列表（可逗号分割多个）
 * @param {string} title - 邮件大标题
 * @param {string} call - 开头称呼语（通常是免责或警告说明）
 * @param {string} describe1 - HTML 分片叙述 1 (含样式加粗)
 * @param {string} describe2 - HTML 分片叙述 2
 * @param {string} describe3 - 核心附带内容 (通常是上述生成的 table html)
 * @param {string} describe4 - 备用叙述位
 * @param {string} describe5 - 备用叙述位
 * @param {string} describe6 - 备用叙述位
 * @param {string} ccmail - 抄送人
 * @param {any} arrAttachment - 保留位置(预留扩展发送 PDF 文件)
 */
function Mail_HTML_TXT(tomail, title, call, describe1, describe2, describe3, describe4, describe5, describe6, ccmail, arrAttachment) {
  const html_txt = '<div><p style="font-size:14px;color:black">' + (call || '') + '</p>' +
                   '<div><p style="font-size:12px;color:black;text-indent:15px"><b>' + (describe1 || '') + '</b></p>' +
                   '<div><p style="font-size:12px;color:black;text-indent:15px">' + (describe2 || '') + '</p>' +
                   '<div><p style="font-size:12px;color:black;text-indent:15px">' + (describe3 || '') + '</p>' +
                   '<div><p style="font-size:12px;color:black;text-indent:15px">' + (describe4 || '') + '</p>' +
                   '<div><p style="font-size:12px;color:black;text-indent:15px">' + (describe5 || '') + '</p>' +
                   '<div><p style="font-size:12px;color:black;text-indent:15px">' + (describe6 || '') + '</p></div>';
  GmailApp.sendEmail(tomail, title, '', {htmlBody: html_txt, cc: ccmail || "", name: 'GoogleScriptAutoSend'});
}

/**
 * 读取 Sheet 页为带 Header 的 Json KV（数组内嵌对象）格式
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sbnSheet - 需要读取的 Sheet 实例缓存
 * @param {number} startRow - 表头排在第几行
 * @param {number} startCol - 读取起始列
 * @param {number} endCol - 读取总终止跨度列
 * @returns {Array} 以表头名为 Key 的对象集合数组
 */
function getDataToArrJson(sbnSheet, startRow, startCol, endCol) {
  const lr = sbnSheet.getLastRow();
  const arrHead = sbnSheet.getRange(startRow, startCol, 1, endCol - startCol + 1).getValues();
  const arrBody = (lr > startRow) ? sbnSheet.getRange(startRow + 1, startCol, lr - startRow, endCol - startCol + 1).getValues() : [];
  const arrJson = [];
  
  arrBody.forEach(function(valsRow, indexRow) {
    const obj = {"行号": indexRow + startRow + 1};
    valsRow.forEach(function(valCol, indexCol) { obj[arrHead[0][indexCol]] = valCol; });
    arrJson.push(obj);
  });
  
  return arrJson;
}

/**
 * 查阅 BOM 在源 Json Array 中的绝对阵列序列位 (根据特征匹配定位)
 * @param {Array} arrJsonData - BOM 原始数据
 * @param {string} key1 - 主要检索健属性（如 牙柄C#）
 * @param {string} val1 - 主匹配值
 * @param {string} key2 - 交叉检索健（如 模具号）
 * @param {string} val2 - 辅匹配值
 * @returns {number} 找出的索引游标，未找到返回 -1
 */
function getBomRowJsonData(arrJsonData, key1, val1, key2, val2) {
  for(let i = 0; i < arrJsonData.length; i++) { 
    if(arrJsonData[i][key1] == val1 && arrJsonData[i][key2] == val2) { return i; } 
  }
  return -1;
}

/**
 * 查阅 Cell Map 分配关系在源 Json Array 中的阵列序列位
 * @param {Array} arrJsonData - Cell 源 Json 库
 * @param {string} key1 - 检索基准特征
 * @param {string} val1 - 对应的值
 * @param {string} key2 - 附加交叉判断位
 * @param {string} val2 - 附加条件值
 * @returns {number} 返回位置游标
 */
function getCellMapRowJsonData(arrJsonData, key1, val1, key2, val2) {
  for(let i = 0; i < arrJsonData.length; i++) { 
    if(arrJsonData[i][key1] == val1 && arrJsonData[i][key2] == val2) { return i; } 
  }
  return -1;
}

/**
 * 对文件执行 OCR (识别文本中的特属内容) - 被临时备用料判定功能使用
 * @param {string} fileId - 文件位于云盘中的 HashID
 * @param {string} tfBomId - 检索对应 BOM
 * @returns {string} 包含提示结果文本
 */
function pdfToOcr(fileId, tfBomId) {
  let lotNo = "";
  try {
    const file = DriveApp.getFileById(fileId);
    const docId = Drive.Files.insert({ title: file.getName(), supportsAllDrives: true }, file.getBlob(), { ocr: true, supportsAllDrives: true }).id;
    const doc = DocumentApp.openById(docId);
    const text = doc.getBody().getText().replace(/\s/g, "");
    const matched = text.match(/根据生产需求(.*)替代/) || text.match(/备料.*/);
    if(matched) {
      const sbn = getSheetCached(getProp("planChangeId", "12GicROoALNx8uT7LMH4sUHDMULlUN0U6nZXgUf-8EJg"), "每周转换记录");
      const lr = sbn.getLastRow();
      const arr = sbn.getRange(1, 1, lr, 8).getDisplayValues();
      const found = arr.filter(function(v) { return String(v[0]).indexOf(tfBomId) != -1 && String(v[3]).indexOf("备份料") != -1; });
      lotNo = found.length > 0 ? found[0][3] : "该BOM有备份料信息，注意检查";
    }
    DriveApp.getFileById(docId).setTrashed(true);
  } catch(e) {
    if(e.message.includes("Request Too Large")) {
      const link = Drive.Files.get(fileId).thumbnailLink.replace(/=s.+/, "=s2000");
      const docId = Drive.Files.insert({ title: "thumb_" + fileId, supportsAllDrives: true }, UrlFetchApp.fetch(link).getBlob(), { ocr: true, supportsAllDrives: true }).id;
      const text = DocumentApp.openById(docId).getBody().getText().replace(/\s/g, "");
      const matched = text.match(/根据生产需求(.*)替代/) || text.match(/备料.*/);
      if(matched) {
        const sbn = getSheetCached(getProp("planChangeId", "12GicROoALNx8uT7LMH4sUHDMULlUN0U6nZXgUf-8EJg"), "每周转换记录");
        const lr = sbn.getLastRow();
        const arr = sbn.getRange(1, 1, lr, 8).getDisplayValues();
        const found = arr.filter(function(v) { return String(v[0]).indexOf(tfBomId) != -1 && String(v[3]).indexOf("备份料") != -1; });
        lotNo = found.length > 0 ? found[0][3] : "该BOM有备份料信息，注意检查";
      }
      DriveApp.getFileById(docId).setTrashed(true);
    } else { 
      console.log(e.message); 
    }
  }
  return lotNo;
}

/**
 * 带有超时熔断机制、休眠降级与垃圾回收处理的健壮版图像分析引擎 (主用于解析上传的日期章图)
 * @param {string} fileId - 被用户刚存进盘的快照ID
 * @param {string} dateTimeMachineSku - 为了文件定位给这个临时文件附带特征命名
 * @returns {Array} 【日期OCR辨别码 , 过程暂存的垃圾文件HashID (供彻底清场)】
 */
function testPictureOcr(fileId, dateTimeMachineSku) {
  const ocrFolderId = getProp("ocrFolderId", "1hkajj384kAXqNqRkDj4rugW3hdbdRtvj");
  const file = DriveApp.getFileById(fileId);
  let docTempId = "";
  let result = ["", ""]; 
    
  // 2. 指数退避重试机制
  let maxRetries = 3;
  let attempt = 0;
  let success = false;
  
  while (attempt < maxRetries && !success) {
    try {
      attempt++;
      docTempId = Drive.Files.insert(
        { title: dateTimeMachineSku, parents: [{ id: ocrFolderId }] }, 
        file.getBlob(), 
        { ocr: true, supportsAllDrives: true }
      ).id;
      result = extractInfo(docTempId);
      success = true; // 成功则结束循环
    }
    catch (error) {
      let message = error.message;
      if (message.includes("Request Too Large") && attempt === 1) {
        // 图片过大，使用缩略图重试
        const link = Drive.Files.get(file.getId()).thumbnailLink.replace(/=s.+/, "=s2000");
        docTempId = Drive.Files.insert(
          { title: file.getName() + "_thumb", parents: [{ id: ocrFolderId }] }, 
          UrlFetchApp.fetch(link).getBlob(), 
          { ocr: true, supportsAllDrives: true }
        ).id;
        result = extractInfo(docTempId);
        success = true;
      } else if (message.includes("rate limit") || message.includes("Rate Limit")) {
        // 遇到限流，强行休眠等待。第1次等5秒，第2次等10秒
        let waitTime = 5000 * attempt;
        console.warn("第 " + attempt + " 次发生限流，休眠 " + (waitTime/1000) + " 秒后重试...");
        Utilities.sleep(waitTime);
      } else {
        console.log("其他报错: " + message);
        if (attempt >= maxRetries) break;
      }
    }
    finally {
      // 3. 【防爆盘机制】每次循环结束，无论是成功还是报错，都把产生的临时 Google Doc 删掉！
      if (docTempId) {
        try {
          DriveApp.getFileById(docTempId).setTrashed(true);
          docTempId = ""; // 置空，防止下次循环重复删除报错
        } catch(e) {
          console.warn("清理临时文档失败", e.message);
        }
      }
    }
  }
  
  console.log("OCR 最终结果集:", result);
  return result;
}

/**
 * 具体解析 Google Docs 内含提取自 OCR 分析出的物理文本，用正则匹配符合格式规范的特定印章号
 * @param {string} docId - 被创建并转换好内部文本的 Google Doc 序列 ID
 * @returns {Array} 输出分析结果数组 [被正则抠出来的合格批号文本, 溯源文件的关联Id]
 */
function extractInfo(docId) {
  console.log("docId", docId);
  let docTemp = DocumentApp.openById(docId);
  
  // 1. 获取带换行的原始文本
  let fullText = docTemp.getBody().getText();
  console.log("fillText", fullText);
  
  // 2. 按行分割成数组
  let lines = fullText.split('\n');
  
  let result = {date: "" };
  
  // 3. 遍历每一行进行处理
  lines.forEach(function(line) {
    // 对当前行去除所有空格，方便正则匹配
    let cleanLine = line.replace(/\s/g, "");
    if (!cleanLine) return; // 跳过空行

    if (!result.date) {
      let d = cleanLine.match(/(\d{2}|\dO|O\d|\dI|I\d|\dB|B\d){3}S?X(A|B|C|8)(55|58)/);
      if (d) result.date = d[0];
    }
  });
  
  console.log(result.date, docId);
  return [result.date, docId];
}

/**
 * 云盘图像落地存管核心逻辑（带超高并发阻塞队列锁处理）—— 为照片盖图服务
 * @param {string} data - 前端丢入的 base64 RAW 流
 * @param {string} lineInputBoxID - 被自检拍摄机台号
 * @param {string} picType - 是拍的“靠模”，还是拍的“日期章”，或者是拍的“清洁参数”类？
 * @returns {Array} 返回存放结果响应 ["OK/NO", "照片直通共享链接或保存提示", "[附带如果触发OCR引擎后的文本识别回吐信息]"]
 */
function uploadpicFileToGoogleDrive(data, lineInputBoxID, picType) {
  // 1. 获取脚本级别的锁，所有并发请求都会在这里排队
  var lock = LockService.getScriptLock();
  try {
    // 2. 尝试获取锁，设置最长排队等待时间为 30秒 (30000毫秒)
    lock.waitLock(30000);
    
    let nowDateTime = Utilities.formatDate(new Date(), 'Asia/Shanghai', 'yyyy年M月d日 HH:mm:ss');
    let filename = lineInputBoxID + "_" + nowDateTime + "_" + picType;
    let scanResult = [];
    let folder = DriveApp.getFolderById(folderHandbagDatePictureId);
    let contentType = data.substring(5, data.indexOf(';'));
    
    if(!contentType.match('image')) {
      return ["NO", "错误：你选择的不是图片文件"];
    } else {
      let bytes = Utilities.base64Decode(data.substr(data.indexOf('base64,') + 7));
      let blob = Utilities.newBlob(bytes, contentType, filename);   
      let fileid = folder.createFile(blob).getId();
      
      if(picType == "日期章") {
        scanResult = testPictureOcr(fileid, filename);
      }
      
      // 给 Google 服务器一点喘息时间，防止刚释放锁下一个人瞬间挤进来报错
      Utilities.sleep(1500);
      return ["OK", fileid, scanResult];
    }
  } catch (error) {
    return ["NO", error.toString(), []];
  } finally {
    // 4. 【非常关键】无论上面成功还是失败，必须释放锁！让排队的下一个人进来
    lock.releaseLock();
  }
}

/**
 * 极简专用接口：仅获取指定 BOM 的新旧转换历史备注
 * @param {string} tfBomId - 需要查询的 TF BOM 号
 * @returns {Array} ["OK", 转换历史记录数组] 或 ["NO", 错误信息]
 */
function getBomChangeRemarks(tfBomId) {
  try {
    if (!tfBomId) return ["NO", "BOM号为空"];
    
    // 获取“每周转换记录”表
    let ssNewOldChange = getSheetCached(wsNewOldChangeId, "每周转换记录");
    let lr = ssNewOldChange.getLastRow();
    if (lr <= 1) return ["OK", []];
    
    // 一次性拉取数据
    let arrSsNewOldChangeAll = ssNewOldChange.getRange(2, 1, lr - 1, 7).getValues();
    arrSsNewOldChangeAll.forEach((v,index)=>v.unshift(index+2));
    let bomNum = tfBomId.slice(-4);
    
    // 过滤、映射并倒序（最新的在最前）
    let arrSsNewOldChange = arrSsNewOldChangeAll.filter(function(x) { 
      return (x[1] + x[2]).toString().indexOf(bomNum) != -1; 
    }).map(function(x) { 
      return [x[0],x[1], x[2], x[3], x[4], x[5], x[6], x[7]]; 
    }).reverse();
    
    // 格式化日期并附加 ID 序号
    for (let i = 0; i < arrSsNewOldChange.length; i++) {
      if (arrSsNewOldChange[i][6] != "") {
        arrSsNewOldChange[i][6] = Utilities.formatDate(new Date(arrSsNewOldChange[i][6].toString()), "Asia/Shanghai", "yyyy-MM-dd");
      }
      if (arrSsNewOldChange[i][7] != "") {
        arrSsNewOldChange[i][7] = Utilities.formatDate(new Date(arrSsNewOldChange[i][7].toString()), "Asia/Shanghai", "yyyy-MM-dd");
      }
    }
    
    return ["OK", arrSsNewOldChange];
  } catch(e) {
    return ["NO", e.toString()];
  }
}


/**
 * 批量判断是否已有首件确认记录 (极致性能优化版)
 * @param {string} nowDateShift - 需要查找定位的验证班期特征
 * @param {Array} tasks - 前台推送过来的包含待处理目标（机器和物料号）合集
 * @returns {Array} 字典映射，向调用端批量反馈每一台是被挡住的 NO，还是可以绿灯放行的 OK
 */
function batchCheckWriteFirstConfirm(nowDateShift, tasks) {
  try {
    const wsFirstConfirm = getSpreadsheetCached(wsFirstConfirmId);
    const shNonDecal = wsFirstConfirm.getSheetByName("NonDecalPrint");
    const shDecal = wsFirstConfirm.getSheetByName("DecalPrint");
    
    // 一次性读取近 1000 行记录到内存，避免多次读表
    let arrNon = [], arrDecal = [];
    let lrNon = shNonDecal.getLastRow(), lcNon = shNonDecal.getLastColumn();
    if(lrNon > 1) { 
      arrNon = shNonDecal.getRange(Math.max(1, lrNon - 1000), 1, Math.min(lrNon, 1001), lcNon).getValues(); 
    }
    
    let lrDec = shDecal.getLastRow(), lcDec = shDecal.getLastColumn();
    if(lrDec > 1) { 
      arrDecal = shDecal.getRange(Math.max(1, lrDec - 1000), 1, Math.min(lrDec, 1001), lcDec).getValues(); 
    }
    
    // 建立 Set 集合，查找速度为 O(1)
    const mapNon = new Set(arrNon.map(function(x) { return x[lcNon - 1] + x[3] + x[4]; }));
    const mapDec = new Set(arrDecal.map(function(x) { return x[lcDec - 1] + x[3] + x[4]; }));

    // 循环打包传进来的所有机台，瞬间秒出结果
    let results = {};
    for (let i = 0; i < tasks.length; i++) {
       let m = tasks[i].machine.split(",")[0];
       let s = tasks[i].sku.split(",")[0];
       let key = nowDateShift + s + m;
       let isDecal = m.charAt(0) === "S";
       
       let exists = isDecal ? mapDec.has(key) : mapNon.has(key);
       
       // 将每台机器的结果存入字典
       results[m] = exists ? ["NO", "已有该机台该品种的首件确认记录"] : ["OK", "开始读取数据……"];
    }
    // 一次性全部返回！
    return ["OK", results];
  } catch(e) { 
    return ["NO", e.toString()];
  }
}

/**
 * 批量处理：首件记录提交（真·批量高并发版）
 * 将前端同时选中的所有机器拼装起来统一提交并写入 Google Sheets（完全解除旧版的重复死锁卡顿）
 * @param {Array} payloads - 从前端提取包含用户交互答题结果的核心对象阵列
 * @returns {Array} 完成打包推入并下发通知结果
 */
function batchProcessFirstConfirm(payloads) {
  try {
    const wsFirstConfirm = getSpreadsheetCached(wsFirstConfirmId);
    const ssNonDecal = wsFirstConfirm.getSheetByName("NonDecalPrint");
    const ssDecal = wsFirstConfirm.getSheetByName("DecalPrint");

    let outRowsNonDecal = [];
    let outRowsDecal = [];
    let allOutRows = [];
    
    const now = new Date();
    const nowTimeStr = Utilities.formatDate(now, "Asia/Shanghai", "yyyy-MM-dd HH:mm:ss");
    const nowDate = Utilities.formatDate(now, "Asia/Shanghai", "yyyy-MM-dd");
    const nextDate = Utilities.formatDate(new Date(now.getTime() + 24 * 3600 * 1000), "Asia/Shanghai", "yyyy-MM-dd");
    const hm = Utilities.formatDate(now, "Asia/Shanghai", "HH:mm");
    
    let nowDateShift = "";
    if (hm > "23:44" && hm < "24:00") { nowDateShift = nextDate + "_1"; }
    else if (hm >= "00:00" && hm < "07:45") { nowDateShift = nowDate + "_1"; }
    else if (hm > "07:44" && hm < "15:45") { nowDateShift = nowDate + "_2"; }
    else if (hm > "15:44" && hm < "23:45") { nowDateShift = nowDate + "_3"; }

    // 1. 在纯内存中将任务拆解、拼装为二维数组 (完全无数据库 I/O 损耗)
    for (let i = 0; i < payloads.length; i++) {
      let p = payloads[i];
      
      // 通知触发 (原逻辑保持)
      const hasMultiSubmit = (p.arrInjParaHistoryInputData || []).map(function(v) { return JSON.stringify(v[2]); }).some(function(s) { return s.indexOf("不") != -1; });
      if (hasMultiSubmit && p.reportSelectResons == "正常") {
        if (typeof SendGoogleChat !== 'undefined' && SendGoogleChat.SendMessagePara) {
          SendGoogleChat.SendMessagePara(p.userNameID, p.workshop, p.materialInputID, p.lineInputId, p.arrInjParaHistoryInputData);
        }
      }
      
      if ((p.arrWarnParameter || []).length > 0) {
        SendNoStdMail(p.workshop, p.arrWarnParameter, p.mailBomIf, p.lineInputId, p.materialInputID, p.selfCheckNum);
      }
      
      if ((p.needRemindIf || []).join("【|】").indexOf("NeedRemind") != -1) {
        if (typeof SendGoogleChat !== 'undefined' && SendGoogleChat.SendMessageRadio) {
          SendGoogleChat.SendMessageRadio(p.workshop, p.needRemindIf, p.lineInputId, p.materialInputID);
        }
      }

      const arrMachine = p.lineInputId.split(",");
      const arrSku = p.materialInputID.split(",");
      
      for (let k = 0; k < arrMachine.length; k++) {
        let machine = arrMachine[k].substr(0, 8);
        let isDecal = machine.charAt(0) === "S";
        let row = [];
        
        row.push(p.userNameID, p.SAPCodeID, p.workshop, arrSku[k], machine, nowTimeStr);
        
        if (!isDecal) {
          row.push(JSON.stringify(p.arrInjectionMaterial));
          for (let n = 0; n < (p.arrPkDate || []).length; n++) { row.push(p.arrPkDate[n]); }
          row.push(JSON.stringify(p.arrPkMaterial));
        } else {
          row.push(JSON.stringify(p.arrDecalPrintMaterial));
        }
        
        for (let r = 0; r < (p.radioIf1 || []).length; r++) { row.push(p.radioIf1[r]); }
        
        let ipStr = "[]";
        const ip = (p.arrInjParameter || []).filter(function(v) { return v[0] == machine; }).map(function(v) { return v[1]; });
        if (ip.length > 0 && ip[0]) { ipStr = JSON.stringify(ip[0]); }
        row.push(ipStr);
        
        let c1 = "", c2 = "";
        const kp = (p.arrKeyParameter || []).filter(function(v) { return v[0] == machine; }).map(function(v) { return v[1]; });
        if (kp.length > 0 && kp[0]) { c1 = kp[0][0]; c2 = kp[0][1]; }
        row.push(c1, c2);
        
        let stdStr = "[]";
        const sp = (p.arrStdParameter || []).filter(function(v) { return v[0] == machine; }).map(function(v) { return v[1]; });
        if (sp.length > 0 && sp[0]) { 
          stdStr = JSON.stringify(sp[0].map(function(v) { return [v[2], v[4], v[5]]; })); 
        }
        row.push(stdStr);
        
        row.push(p.picShiftModelId ? ("https://drive.google.com/file/d/" + p.picShiftModelId + "/preview") : "");
        
        let hisStr = "[]";
        const hp = (p.arrInjParaHistoryInputData || []).filter(function(v) { return v[0] == machine || v[0] == "领班确认"; });
        if (hp.length > 0) {
          if (!isDecal ? stdStr.length > 2 : true) { 
            hisStr = JSON.stringify(hp.map(function(v) { return [v[1], v[2]]; })); 
          }
        }
        row.push(hisStr);
        
        const exRes = (String(p.exemptionResons).length < 3 ? "" : p.exemptionResons);
        row.push(p.fullMoldStr, p.passSelectResons, p.reportSelectResons, exRes, p.exemptionId, nowDateShift);
        // 注：首件没有 selfCheckNum 列

        // 依据机器类型丢进不同的缓存区
        if (!isDecal) { 
          outRowsNonDecal.push(row);
        } else { 
          outRowsDecal.push(row);
        }
        allOutRows.push(row);

        if (row[7] != "") { 
          SendGoogleChat.sendMessageInjOnline(row, p.tabScanResult);
        }
      }
    }

    // 2. 终极写入：申请 1 次锁，一次性查行，用 setValues 瞬间排入表格！
    return withLock(function() {
      if (outRowsNonDecal.length > 0) {
        ssNonDecal.getRange(ssNonDecal.getLastRow() + 1, 1, outRowsNonDecal.length, outRowsNonDecal[0].length).setValues(outRowsNonDecal);
      }
      if (outRowsDecal.length > 0) {
        ssDecal.getRange(ssDecal.getLastRow() + 1, 1, outRowsDecal.length, outRowsDecal[0].length).setValues(outRowsDecal);
      }
      return ["OK", "写入成功"];
    });
  } catch(e) {
    return ["NO", e.toString()];
  }
}

/**
 * 批量处理：两小时自检提交（真·批量高并发版）
 * 逻辑高度同上，用于解决以前多次循环排队获取 Sheet 导致 15 秒大拥堵崩溃的问题
 * @param {Array} payloads - 任务字典阵列
 * @returns {Array} 处理状态与新追加的数据全集反馈（用以刷新前端缓存）
 */
function batchProcessTwoHourSelfCheck(payloads) {
  try {
    const wsFirstConfirm = getSpreadsheetCached(wsTwoHourSelfCheckId);
    const ssNonDecal = wsFirstConfirm.getSheetByName("NonDecalPrint");
    const ssDecal = wsFirstConfirm.getSheetByName("DecalPrint");

    let outRowsNonDecal = [];
    let outRowsDecal = [];
    let allOutRows = [];
    
    const now = new Date();
    const nowTimeStr = Utilities.formatDate(now, "Asia/Shanghai", "yyyy-MM-dd HH:mm:ss");
    const nowDate = Utilities.formatDate(now, "Asia/Shanghai", "yyyy-MM-dd");
    const nextDate = Utilities.formatDate(new Date(now.getTime() + 24 * 3600 * 1000), "Asia/Shanghai", "yyyy-MM-dd");
    const hm = Utilities.formatDate(now, "Asia/Shanghai", "HH:mm");
    
    let nowDateShift = "";
    if (hm > "23:44" && hm < "24:00") { nowDateShift = nextDate + "_1"; }
    else if (hm >= "00:00" && hm < "07:45") { nowDateShift = nowDate + "_1"; }
    else if (hm > "07:44" && hm < "15:45") { nowDateShift = nowDate + "_2"; }
    else if (hm > "15:44" && hm < "23:45") { nowDateShift = nowDate + "_3"; }

    // 1. 在内存中完成所有的逻辑拆解与二维数组组装
    for (let i = 0; i < payloads.length; i++) {
      let p = payloads[i];
      
      const hasMultiSubmit = (p.arrInjParaHistoryInputData || []).map(function(v) { return JSON.stringify(v[2]); }).some(function(s) { return s.indexOf("不") != -1; });
      if (hasMultiSubmit && p.reportSelectResons == "正常") {
        if (typeof SendGoogleChat !== 'undefined' && SendGoogleChat.SendMessagePara) {
          SendGoogleChat.SendMessagePara(p.userNameID, p.workshop, p.materialInputID, p.lineInputId, p.arrInjParaHistoryInputData);
        }
      }
      
      if ((p.arrWarnParameter || []).length > 0) {
        SendNoStdMail(p.workshop, p.arrWarnParameter, p.mailBomIf, p.lineInputId, p.materialInputID, p.selfCheckNum);
      }
      
      if ((p.needRemindIf || []).join("【|】").indexOf("NeedRemind") != -1) {
        if (typeof SendGoogleChat !== 'undefined' && SendGoogleChat.SendMessageRadio) {
          SendGoogleChat.SendMessageRadio(p.workshop, p.needRemindIf, p.lineInputId, p.materialInputID);
        }
      }

      const arrMachine = p.lineInputId.split(",");
      const arrSku = p.materialInputID.split(",");
      
      for (let k = 0; k < arrMachine.length; k++) {
        let machine = arrMachine[k].substr(0, 8);
        let isDecal = machine.charAt(0) === "S";
        let row = [];
        
        row.push(p.userNameID, p.SAPCodeID, p.workshop, arrSku[k], machine, nowTimeStr);
        
        if (!isDecal) {
          row.push(JSON.stringify(p.arrInjectionMaterial));
          for (let n = 0; n < (p.arrPkDate || []).length; n++) { row.push(p.arrPkDate[n]); }
          row.push(JSON.stringify(p.arrPkMaterial));
        } else {
          row.push(JSON.stringify(p.arrDecalPrintMaterial));
        }
        
        for (let r = 0; r < (p.radioIf1 || []).length; r++) { row.push(p.radioIf1[r]); }
        
        let ipStr = "[]";
        const ip = (p.arrInjParameter || []).filter(function(v) { return v[0] == machine; }).map(function(v) { return v[1]; });
        if (ip.length > 0 && ip[0]) { ipStr = JSON.stringify(ip[0]); }
        row.push(ipStr);
        
        let c1 = "", c2 = "";
        const kp = (p.arrKeyParameter || []).filter(function(v) { return v[0] == machine; }).map(function(v) { return v[1]; });
        if (kp.length > 0 && kp[0]) { c1 = kp[0][0]; c2 = kp[0][1]; }
        row.push(c1, c2);
        
        let stdStr = "[]";
        const sp = (p.arrStdParameter || []).filter(function(v) { return v[0] == machine; }).map(function(v) { return v[1]; });
        if (sp.length > 0 && sp[0]) { 
          stdStr = JSON.stringify(sp[0].map(function(v) { return [v[2], v[4], v[5]]; })); 
        }
        row.push(stdStr);
        
        row.push(p.picShiftModelId ? ("https://drive.google.com/file/d/" + p.picShiftModelId + "/preview") : "");
        
        let hisStr = "[]";
        const hp = (p.arrInjParaHistoryInputData || []).filter(function(v) { return v[0] == machine || v[0] == "领班确认"; });
        if (hp.length > 0) {
          if (!isDecal ? stdStr.length > 2 : true) { 
            hisStr = JSON.stringify(hp.map(function(v) { return [v[1], v[2]]; })); 
          }
        }
        row.push(hisStr);
        
        const exRes = (String(p.exemptionResons).length < 3 ? "" : p.exemptionResons);
        row.push(p.fullMoldStr, p.passSelectResons, p.reportSelectResons, exRes, p.exemptionId, nowDateShift, p.selfCheckNum);
        
        if (!isDecal) { 
          outRowsNonDecal.push(row); 
        } else { 
          outRowsDecal.push(row);
        }
        allOutRows.push(row);

        if (row[7] != "" && row[4].indexOf("S1") == -1 && row[4].indexOf("S2") == -1) {
           SendGoogleChat.sendMessageInjOnline(row, p.tabScanResult);
        }
      }
    }

    // 2. 终极写入：申请 1 次锁，一次性查行，用 setValues 瞬间排入表格！
    return withLock(function() {
      if (outRowsNonDecal.length > 0) {
        ssNonDecal.getRange(ssNonDecal.getLastRow() + 1, 1, outRowsNonDecal.length, outRowsNonDecal[0].length).setValues(outRowsNonDecal);
      }
      if (outRowsDecal.length > 0) {
        ssDecal.getRange(ssDecal.getLastRow() + 1, 1, outRowsDecal.length, outRowsDecal[0].length).setValues(outRowsDecal);
      }
      return ["OK", allOutRows];
    });
  } catch(e) {
    return ["NO", e.toString()];
  }
}

/**
 * 批量处理：交接班记录（解决 N+1 网络延迟问题专属直写通道）
 * @param {Array} payloads - 来自用户填写的交接班异常或核检清单数组
 * @returns {Array} 将存入的数据原样返回，前端用于重构页面与显示状态标识
 */
function batchProcessHandOver(payloads) {
  try {
    const ss = getSheetCached(wsHandOverId, "CheckRecord");
    return withLock(function() {
      // payloads 的格式已经是二维数组：[[user, dateShift, ws, machine, ...radioRes], [user2...]]
      // 由于 safeAppendRows 需要严格对齐列，直接使用 Google API 的 batch setValues 方法写到底部
      ss.getRange(ss.getLastRow() + 1, 1, payloads.length, payloads[0].length).setValues(payloads);
      return ["OK", payloads]; 
    });
  } catch(e) {
    return ["NO", e.toString()];
  }
}

/**
 * 批量判断是否已有设备首件确认记录 (极致性能优化版)
 * @param {string} nowDateShift - 日期与班次联合键
 * @param {Array} tasks - 前端发送的任务清单
 * @returns {Array} 字典映射，返回每台机器的检测结果
 */
function batchCheckWriteTechFirstConfirm(nowDateShift, tasks) {
  try {
    const ssFcRecord = getSheetCached(wsTechFirstConfirmId, "Injection");
    let arrSsFcRecord = [];
    const lr = ssFcRecord.getLastRow();
    const lc = ssFcRecord.getLastColumn();
    
    // 一次性读取近 1000 行记录到内存
    if (lr > 1) { 
      arrSsFcRecord = ssFcRecord.getRange(Math.max(1, lr - 1000), 1, Math.min(lr, 1001), lc).getValues();
    }
    
    // 建立 Set 集合，查找速度为 O(1)
    const mapTech = new Set(arrSsFcRecord.map(x => x[lc - 1] + x[3] + x[4]));
    
    let results = {};
    for (let i = 0; i < tasks.length; i++) {
       let m = tasks[i].machine.split(",")[0];
       let s = tasks[i].sku.split(",")[0];
       let key = nowDateShift + s + m;
       let exists = mapTech.has(key);
       results[m] = exists ? ["NO", "当班已有该机台该品种的设备首件确认记录"] : ["OK", "开始读取数据……"];
    }
    return ["OK", results];
  } catch(e) { 
    return ["NO", e.toString()];
  }
}

/**
 * 批量处理：设备首件记录提交（真·批量高并发写入及发信）
 * @param {Array} payloads - 前端传递的数据包数组
 * @returns {Array} 写入结果状态
 */
function batchProcessTechFirstConfirm(payloads) {
  try {
    const ssFcRecord = getSheetCached(wsTechFirstConfirmId, "Injection");
    let allOutRows = [];
    let emailTasks = []; // 需要发邮件的任务池
    
    const now = new Date();
    const nowTimeStr = Utilities.formatDate(now, "Asia/Shanghai", "yyyy-MM-dd HH:mm:ss");
    const nowDate = Utilities.formatDate(now, "Asia/Shanghai", "yyyy-MM-dd");
    const nextDate = Utilities.formatDate(new Date(now.getTime() + 24 * 3600 * 1000), "Asia/Shanghai", "yyyy-MM-dd");
    const hm = Utilities.formatDate(now, "Asia/Shanghai", "HH:mm");
    let nowDateShift = "";
    if (hm > "23:44" && hm < "24:00") { nowDateShift = nextDate + "_1"; }
    else if (hm >= "00:00" && hm < "07:45") { nowDateShift = nowDate + "_1"; }
    else if (hm > "07:44" && hm < "15:45") { nowDateShift = nowDate + "_2"; }
    else if (hm > "15:44" && hm < "23:45") { nowDateShift = nowDate + "_3"; }

    for (let i = 0; i < payloads.length; i++) {
      let p = payloads[i];
      const arrMachine = p.lineInputId.split(",");
      const arrSku = p.materialInputID.split(",");
      
      for (let k = 0; k < arrMachine.length; k++) {
        let machine = arrMachine[k].substr(0, 8);
        let row = [p.userNameID, p.SAPCodeID, p.workshop, arrSku[k], machine, nowTimeStr];
        
        // 展平并推入所有工艺参数
        for (let r = 0; r < p.arrInjParameter.length; r++) {
          row.push(p.arrInjParameter[r]);
        }
        
        // 推入末尾的状态列
        row.push(p.passSelectResons, p.reportSelectResons, p.exemptionResons, p.exemptionId, nowDateShift);
        allOutRows.push(row);

        if (p.emailIf == 1) {
          emailTasks.push({
             workshop: p.workshop,
             lineInputId: machine,
             materialInputID: arrSku[k],
             moldModel: p.moldModel,
             exemptionResons: p.exemptionResons
          });
        }
      }
    }

    return withLock(function() {
      // 1. 唯一的一次数据库 I/O，瞬间写入所有行
      if (allOutRows.length > 0) {
        ssFcRecord.getRange(ssFcRecord.getLastRow() + 1, 1, allOutRows.length, allOutRows[0].length).setValues(allOutRows);
      }
      
      // 2. 处理邮件发送
      if (emailTasks.length > 0) {
        for(let eTask of emailTasks) {
           const toEmail = (eTask.workshop == "TB1") ? emailTB1Ops : emailTB2Ops;
           const ccEmail = emailEngCC;
           const widths = [45, 20, 35], head_backgrounds = ["unset", "unset", "unset"], aligns = ["center", "center", "center"], scale = 60, col = -1;
           const title = "设备首件确认，工艺参数不在标准范围内";
           const call = "<i>Auto send email, do not reply</i>";
           const describe1 = "以下工艺参数无法调整在标准范围内，请知悉！";
           const describe2 = "时间：" + nowTimeStr + ";机台：" + eTask.lineInputId + ";SKU：" + eTask.materialInputID + ";模具号：" + eTask.moldModel;
           const arrParaTemp = JSON.parse(eTask.exemptionResons || "[]");
           arrParaTemp.unshift(["工艺参数", "实际值", "备注"]);
           const describe3 = General_Htmltable(arrParaTemp, widths, head_backgrounds, aligns, scale, col);
           Mail_HTML_TXT(toEmail, title, call, describe1, describe2, describe3, "", "", "", ccEmail);
        }
      }
      return ["OK", "写入成功"];
    });
  } catch(e) {
    return ["NO", e.toString()];
  }
}

/**
 * ==========================================
 * 1. 布局及状态提取 (对应车间排产与机台监控视图)
 * ==========================================
 */

/**
 * 获取纯净的车间机台物理布局数据（不带检查状态颜色）
 * 主要用于自检模块选择机台时的下拉菜单或基础视图展示。
 * @param {string} workshop - 车间名 ("TB1" 或 "TB2")
 * @returns {Array} 状态标识及二维布局数组 ["OK", [[...], ...]] 或 ["NO", 错误信息]
 */
function getPureInjLayout(workshop) {
  let nameLayout = (workshop === "TB2") ? "TB2INJ布局" : "TB1INJ布局";
  let rangeLayout = (workshop === "TB2") ? "A1:H40" : "A1:G34";
  const layoutId = '1eNrl63nWe6etP2pQmyzdii3cxGvqNLkEsxh1jIy3dnI';
  try {
    const layoutSheet = SpreadsheetApp.openById(layoutId).getSheetByName(nameLayout);
    const layoutValues = layoutSheet.getRange(rangeLayout).getDisplayValues();
    return ["OK", layoutValues];
  } catch (e) {
    return ["NO", e.toString()];
  }
}

/**
 * 获取包含机台最新点检/巡检状态的车间全景布局图
 * 会交叉对比月检(Random)和周检(TF)记录，返回应当标绿、标粉或标灰的机台列表。
 * @param {string} workshop - 车间名 ("TB1" 或 "TB2")
 * @returns {Object} 包含基础布局数据和多维度检查状态数组的对象
 */
/**
 * 获取包含机台最新点检/巡检状态的车间全景布局图
 * 包含“按自然周”和“按自然月”重置的智能判定算法
 */
/**
 * 获取包含机台最新点检/巡检状态的车间全景布局图
 * 包含“按自然周”和“按自然月”重置的智能判定算法，同时输出周检和月检所需的维度数据
 */
function getInjMachineLayoutAndStatus(workshop) {
  let nameLayout = "TB1INJ布局";
  let rangeLayout = "A1:G34";
  if(workshop == "TB2"){ nameLayout = "TB2INJ布局"; rangeLayout = "A1:H40"; }
  var layoutId = '1eNrl63nWe6etP2pQmyzdii3cxGvqNLkEsxh1jIy3dnI';
  var layoutSheet = SpreadsheetApp.openById(layoutId).getSheetByName(nameLayout);
  var layoutValues = layoutSheet.getRange(rangeLayout).getDisplayValues();

  var recordId = '18hoqx_pnoRrjqijiDCOWz_hnyu3EUSWS-_sWKAwh868';
  var randomSheet = SpreadsheetApp.openById(recordId).getSheetByName('Injection_RandomInspection'); 
  var injSheet = SpreadsheetApp.openById(recordId).getSheetByName('Injection'); 

  var today = new Date();
  var nowTime = today.getTime();
  var currentMonth = today.getMonth();
  var currentYear = today.getFullYear();

  // 计算本周一的零点时间戳 (中国习惯周一为一周起点)
  var currentDay = today.getDay();
  var mondayDiff = currentDay === 0 ? -6 : 1 - currentDay; // 周日是0，需倒退6天
  var startOfThisWeek = new Date(today.getFullYear(), today.getMonth(), today.getDate() + mondayDiff);
  startOfThisWeek.setHours(0, 0, 0, 0);
  var startOfThisWeekTime = startOfThisWeek.getTime();

  // 1. 获取当月月检（Random）机器 -> 月检灰色
  function getMonthlyRandomMachines(sheet) {
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var machines = [];
    for(var i = 1; i < data.length; i++) {
      if (data[i][5]) {
        var dateVal = new Date(data[i][5]);
        if (dateVal.getFullYear() === currentYear && dateVal.getMonth() === currentMonth) {
           var m = data[i][4] ? data[i][4].toString().trim() : "";
           if(m) machines.push(m);
        }
      }
    }
    return machines;
  }

  // 2. 解析周检（Injection）机器
  function parseInjectionMachines(sheet) {
    if (!sheet) return { weekDone: [], hour24: [], hour168: [] };
    var data = sheet.getDataRange().getValues();
    
    var weekDoneMap = {};
    var latestTimeMap = {};
    
    // 倒序遍历，获取最新状态
    for(var i = data.length - 1; i >= 1; i--) {
      var m = data[i][4] ? data[i][4].toString().trim() : "";
      if (!m) continue;
      
      if (data[i][5]) {
        var recordTime = new Date(data[i][5]).getTime();
        
        // 记录该机台最新的一笔时间（用于月检的绿/粉判定）
        if (!latestTimeMap[m]) {
          latestTimeMap[m] = recordTime;
        }

        // 判定本周是否已查（用于周检的灰/黄判定）
        if (weekDoneMap[m] === undefined) {
          if (recordTime >= startOfThisWeekTime) {
            weekDoneMap[m] = true;
          } else {
            weekDoneMap[m] = false;
          }
        }
      }
    }
    
    // 组装返回数据
    var weekDoneList = Object.keys(weekDoneMap).filter(k => weekDoneMap[k]);
    var hour24List = [];
    var hour168List = [];

    // 根据最新时间计算 24h 和 168h 内的机器
    for (let m in latestTimeMap) {
      let diffHours = (nowTime - latestTimeMap[m]) / (1000 * 60 * 60);
      if (diffHours <= 24) {
        hour24List.push(m);
      } else if (diffHours <= 168) {
        hour168List.push(m);
      }
    }

    return { 
      weekDone: weekDoneList, // 周检：本周已查 (灰色)
      hour24: hour24List,     // 月检：24小时内 (绿色候选)
      hour168: hour168List    // 月检：168小时内 (粉色候选)
    };
  }

  var randomMachines = getMonthlyRandomMachines(randomSheet); // 月检灰色
  var injParsedData = parseInjectionMachines(injSheet);

  return {
    layout: layoutValues,
    // 以下用于月检判定
    monthCheckedGray: randomMachines,       
    injChecked24h: injParsedData.hour24, 
    injChecked168h: injParsedData.hour168,
    // 以下用于周检判定
    weekCheckedGray: injParsedData.weekDone
  };
}

/**
 * 提取主页仪表盘所需的精简数据，例如今天已经完成月检的机台列表
 * 用于主界面快速过滤和状态回显。
 * @returns {Object} 包含当天已检查机台列表的对象
 */
function getMainPageData() {
  var recordId = '18hoqx_pnoRrjqijiDCOWz_hnyu3EUSWS-_sWKAwh868';
  var injSheet = SpreadsheetApp.openById(recordId).getSheetByName('Injection_RandomInspection');
  var tfData = injSheet.getDataRange().getDisplayValues();
  
  var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  var todayCheckedMachines = [];

  for(var j = 1; j < tfData.length; j++) {
    if(tfData[j][5]) {
      var dStr = Utilities.formatDate(new Date(tfData[j][5]), Session.getScriptTimeZone(), "yyyy-MM-dd");
      if(dStr === todayStr) {
        var m = tfData[j][4] ? tfData[j][4].toString().trim() : "";
        if(m && todayCheckedMachines.indexOf(m) === -1) { todayCheckedMachines.push(m); }
      }
    }
  }
  return { bundleMap: machineBundleMap, todayChecked: todayCheckedMachines };
}

/**
 * ==========================================
 * 5. 工艺参数标准查询核心
 * ==========================================
 */

/**
 * 获取设备专属维护的工艺参数标准记录集（含范围上下限及校验部门信息）
 * @param {string} injSku - 当前生产的物料 SKU
 * @param {string} lineInputId - 机台编号
 * @param {string} moldModel - 模具类型标识
 * @param {string} machineModel - 机器子型号分组 (如 Automation_3Aixs)
 * @param {string} machineType - 机器主类型 (如 HT/FT)
 * @returns {Array} 状态及参数列表 ["OK", [], arrSsInjPara]
 */
function getTechParaStandardData(injSku, lineInputId, moldModel, machineModel, machineType) {
  try {
    let arrSsInjParaJson = [];
    let arrSsInjPara = [];
    let arrJsonData = getProcessParameterData(wsInjParaId, "INJ_New");
    let arrInjSku = injSku.toString().split(",");
    
    // 匹配特定机型的专属 SKU 参数
    let arrSsInjParaSku = arrJsonData.filter(v => { return v.机型 == machineType && checkTrue(arrInjSku, v.SKU) && v.SKU != "NA"; }).map(v => { return v.SKU; }).join("");
    let arrSpecialSku = [];
    let specialSku = "";
    for(let i = 0; i < arrInjSku.length; i++) {
      let optision = arrSsInjParaSku.indexOf(arrInjSku[i]);
      if(optision != -1) { arrSpecialSku.push(arrInjSku[i]); }
    }
    if(arrSpecialSku.length > 0) { specialSku = arrSpecialSku[0]; }
    
    // 第一级查找：找完全匹配 SKU 的工艺标准
    arrSsInjParaJson = arrJsonData.filter(v => { return v.模具编码 == moldModel && v.机型 + v.自动化分组 == machineType + machineModel && v.SKU.toString().indexOf(specialSku) != -1 && specialSku != "" && v.状态 == "生效"; });
    for(let i = 0; i < arrSsInjParaJson.length; i++) {
      let arrParaJsonTemp = JSON.parse(arrSsInjParaJson[i]["工艺参数"]);
      for(let j = 0; j < arrParaJsonTemp.length; j++) {
        if(arrParaJsonTemp[j]["检查部门"].toString().indexOf("设备") != -1 && arrParaJsonTemp[j]["设定下限"].toString() != "9999") {
          arrSsInjPara.push([arrSsInjParaJson[i]["分序类别"] + "【" + arrSsInjParaJson[i]["模穴数"] + "-" + arrSsInjParaJson[i]["BigBundle"] + "】", arrParaJsonTemp[j]["工艺参数"], arrParaJsonTemp[j]["单位"], arrParaJsonTemp[j]["设定下限"], arrParaJsonTemp[j]["设定上限"], arrParaJsonTemp[j]["预设值"], arrSsInjParaJson[i]["变更原因"]]);
        }
      }
    }
    
    // 第二级兜底查找：如果没有专属 SKU 参数，退回寻找通用 NA 记录
    if(arrSsInjPara.length < 1) {
      arrSsInjParaJson = arrJsonData.filter(v => { return v.模具编码 == moldModel && v.机型 + v.自动化分组 == machineType + machineModel && v.SKU == "NA" && v.状态 == "生效"; });
      for(let i = 0; i < arrSsInjParaJson.length; i++) {
        let arrParaJsonTemp = JSON.parse(arrSsInjParaJson[i]["工艺参数"]);
        for(let j = 0; j < arrParaJsonTemp.length; j++) {
          if(arrParaJsonTemp[j]["检查部门"].toString().indexOf("设备") != -1 && arrParaJsonTemp[j]["设定下限"].toString() != "9999") {
            arrSsInjPara.push([arrSsInjParaJson[i]["分序类别"] + "【" + arrSsInjParaJson[i]["模穴数"] + "-" + arrSsInjParaJson[i]["BigBundle"] + "】", arrParaJsonTemp[j]["工艺参数"], arrParaJsonTemp[j]["单位"], arrParaJsonTemp[j]["设定下限"], arrParaJsonTemp[j]["设定上限"], arrParaJsonTemp[j]["预设值"], arrSsInjParaJson[i]["变更原因"]]);
          }
        }
      }
    }
    return ["OK", [], arrSsInjPara]; // 保持原版返回兼容性
  } catch(e) { return ["NO", e.toString()]; }
}

/**
 * ==========================================
 * 6. 月检状态记录及拦截写入
 * ==========================================
 */

/**
 * 记录设备参数月检数据（带防冲突锁、领班验证、邮件通报及安东报修联动）
 * @param {string} userNameID - 填报员工号
 * @param {string} SAPCodeID - 员工 SAP 系统 ID
 * @param {string} workshop - 车间名
 * @param {string} materialInputID - 检查的物料 SKU
 * @param {string} lineInputId - 机台编号
 * @param {Array} arrInjParameter - 各项工艺参数详情的数组
 * @param {string} passSelectResons - 通用审核结果原因
 * @param {string} reportSelectResons - 报修/异常原因
 * @param {string} exemptionResons - 豁免/强行通过原因备注
 * @param {string} exemptionId - 批准豁免的领班工号
 * @param {number|string} emailIf - 是否需要触发邮件通知工程师 (1表示是)
 * @param {string} moldModel - 模具类型编码
 * @returns {Array} 写入结果信息 ["OK", "写入成功"]
 */
function writeTechPpmsMonthCheck(userNameID, SAPCodeID, workshop, materialInputID, lineInputId, arrInjParameter, passSelectResons, reportSelectResons, exemptionResons, exemptionId, emailIf, moldModel) {
  try {
    const ssFcRecord = getSheetCached(wsTechPpmsWeekCheckId, "Injection_RandomInspection");
    const now = new Date();
    const nowTimeStr = Utilities.formatDate(now, "Asia/Shanghai", "yyyy-MM-dd HH:mm:ss");
    const nowDate = Utilities.formatDate(now, "Asia/Shanghai", "yyyy-MM-dd");
    const nextDate = Utilities.formatDate(new Date(now.getTime() + 24 * 3600 * 1000), "Asia/Shanghai", "yyyy-MM-dd");
    const hm = Utilities.formatDate(now, "Asia/Shanghai", "HH:mm");
    let nowDateShift = "";
    if(hm > "23:44" && hm < "24:00"){ nowDateShift = nextDate + "_1"; } 
    else if(hm >= "00:00" && hm < "07:45"){ nowDateShift = nowDate + "_1"; } 
    else if(hm > "07:44" && hm < "15:45"){ nowDateShift = nowDate + "_2"; } 
    else if(hm > "15:44" && hm < "23:45"){ nowDateShift = nowDate + "_3"; }

    const row = [userNameID, SAPCodeID, workshop, materialInputID, lineInputId.substr(0, 8), nowTimeStr];
    for(let i = 0; i < arrInjParameter.length; i++){ row.push(arrInjParameter[i]); }
    row.push(passSelectResons, reportSelectResons, exemptionResons, exemptionId, nowDateShift);

    return withLock(function() {
      ssFcRecord.appendRow(row);
      // 若参数超差触发报警，则联动发送邮件并推入安东异常队列
      if(emailIf == 1) {
        const toEmail = (workshop == "TB1") ? emailTB1Ops : emailTB2Ops;
        const ccEmail = emailEngCC;
        const sbnShiftToTeam = getSheetCached(getProp("shiftToTeamId","1yF-3w6wYNhVOvicR2HmAJeVU2MdL3do5QtHSippsoCM"), "源数据");
        const lrST = sbnShiftToTeam.getLastRow();
        const arrSbnShiftToTeam = lrST > 0 ? sbnShiftToTeam.getRange(1, 1, lrST, sbnShiftToTeam.getLastColumn()).getValues() : [];
        const arrSbnShift = arrSbnShiftToTeam.map(x => x[0]);
        const ssEquShift = getSheetCached(wsEquShiftId, "Shift_INJ_" + workshop);
        const sbnMachineToType = getSheetCached(getProp("equMasterId","1bYKTK5a63yJWRHzM_UPP6b4hwF67eZKEM5dCKLWR59U"), "Workcenter & Mold Matrix");
        const lrMT = sbnMachineToType.getLastRow();
        const arrSbnMachineToType = lrMT > 0 ? sbnMachineToType.getRange(1, 1, lrMT, sbnMachineToType.getLastColumn()).getValues() : [];
        const arrSbnMachine = arrSbnMachineToType.map(x => x[0]);
        const arrFollowUp = ["李华", "丁志伟"];
        const nowMs = now.getTime();
        
        // 组装安东共享表的联动报错记录
        const arrPpmsToEquShift = [nowMs, getTeam(nowDateShift, arrSbnShift, arrSbnShiftToTeam), lineInputId, "工艺周检不在范围\n" + exemptionResons, "", "未解决", arrFollowUp[(workshop == "TB1") ? 0 : 1], 0, "M 其他-工艺", "", arrFollowUp[(workshop == "TB1") ? 0 : 1], Utilities.formatDate(now, "Asia/Shanghai", "yyyy/MM/dd"), "FALSE", userNameID, workshop, "INJ", 0, getMachineType2(lineInputId, arrSbnMachine, arrSbnMachineToType), "", "", "", '=if(row()=1,"判断是否最后",if(countif(indirect("A1:A"&row()),indirect("A"&row()))=countif(indirect("A:A"),indirect("A"&row())),if(indirect("A"&row())<>"","Last",""),""))'];
        ssEquShift.appendRow(arrPpmsToEquShift);
        
        // 组装发送超差警告邮件
        const widths = [45, 20, 35], head_backgrounds = ["unset", "unset", "unset"], aligns = ["center", "center", "center"], scale = 60, col = -1;
        const title = "设备参数周检，工艺参数不在标准范围内";
        const call = "<i>Auto send email,do not reply</i>";
        const describe1 = "以下工艺参数无法调整在标准范围内，请知悉！";
        const describe2 = "时间：" + nowTimeStr + ";机台：" + lineInputId + ";SKU：" + materialInputID + ";模具号：" + moldModel;
        const arrParaTemp = JSON.parse(exemptionResons || "[]");
        arrParaTemp.unshift(["工艺参数", "实际值", "备注"]);
        const describe3 = General_Htmltable(arrParaTemp, widths, head_backgrounds, aligns, scale, col);
        Mail_HTML_TXT(toEmail, title, call, describe1, describe2, describe3, "", "", "", ccEmail);
        
        return ["OK", "写入成功且豁免内容已邮给工程师"];
      } else {
        return ["OK", "写入成功"];
      }
    });
  } catch(e) { return ["NO", e.toString()]; }
}

/**
 * ==========================================
 * 7. BOM 转换申请与审批流程处理
 * ==========================================
 */

/**
 * 用户在前端向系统提交新老 BOM 转换申请
 * @param {string} bomNum - 新目标 BOM 编号
 * @param {string} usernameID - 申请者工号
 * @param {string} lineInputId - 申请在哪个机台上执行转换
 * @returns {Array} 执行结果信息
 */
function applyBom(bomNum, usernameID, lineInputId) {
  let ssTf = getSheetCached(wsTfBomId, "TF BOM masterdata Header");
  let data = ssTf.getRange(2, 1, ssTf.getLastRow() - 1, 1).getValues();
  let nowTime = Utilities.formatDate(new Date(), "Asia/Shanghai", "yyyy-MM-dd HH:mm:ss").toString();
  try {
    let bomLists = data.map(function(r){ return r[0]; });
    let position = bomLists.indexOf(bomNum);
    // 写入 Y (申请标志)、申请人、申请时间、申请机台
    if(position > -1) { ssTf.getRange(position + 2, 13, 1, 4).setValues([['Y', usernameID, nowTime, lineInputId]]); }
    return ['OK', "申请成功"];
  } catch (e) { return ["NO", e.toString()]; }
}

/**
 * 领班/工程师审核通过 BOM 的转换申请
 * 将处于老批次的 BOM 设置为“取消”停用，并将处于等待队列的新 BOM 激活为“生效”。
 * @param {string} bomNum - 需要被激活的新 BOM 编号
 * @param {string} bomMaterial - 相关物料 SKU
 * @param {string} user - 审批者的操作人工号
 * @returns {Array} 审批动作的结果标识
 */
function processBOMTransfer(bomNum, bomMaterial, user) {
  let ssTf = getSheetCached(wsTfBomId, "TF BOM masterdata Header");
  let data = ssTf.getRange(2, 1, ssTf.getLastRow() - 1, 12).getValues();
  let date = Utilities.formatDate(new Date(), "Asia/Shanghai", "yyyy-MM-dd HH:mm:ss").toString();
  try {
    let activeBOMLists = [];
    let inactiveBOMLists = [];
    let bomLists = data.map(function(r){ return r[0]; });
    let inactivePosition = bomLists.indexOf(bomNum);
    let inactiveComment = ssTf.getRange(inactivePosition + 2, 12).getValue().toString(); // 获取替代批次备注号
    
    // 找出具有相同备注的所有新旧 BOM
    for(let i = 0; i < data.length; i++) {
      if(inactiveComment && data[i][8] == '生效' && data[i][11].toString() == inactiveComment){ activeBOMLists.push(data[i][0]); }
      if(inactiveComment && data[i][8] == '待生效' && data[i][11].toString() == inactiveComment){ inactiveBOMLists.push(data[i][0]); }
    }
    // 将已生效的旧 BOM 取消
    for(let i = 0; i < activeBOMLists.length; i++) {
      let activePosition = bomLists.indexOf(activeBOMLists[i]);
      if(activePosition > -1){ ssTf.getRange(activePosition + 2, 9).setValue('取消'); }
    }
    // 将待生效的新 BOM 激活生效，并清空申请标记
    for (let i = 0, len = inactiveBOMLists.length; i < len; i++) {
      let inactivePosition = bomLists.indexOf(inactiveBOMLists[i]);
      if(inactivePosition > -1){
        ssTf.getRange(inactivePosition + 2, 17, 1, 3).setValues([['Y', user, date]]); // 记录审批人及时间
        ssTf.getRange(inactivePosition + 2, 13).setValue(''); // 清空申请标志 "Y"
        ssTf.getRange(inactivePosition + 2, 9).setValue('生效'); // 置为生效
      }
    }
    return ['OK', "批准成功"];
  } catch (e) { return ['NO', e.toString()]; }
}

/**
 * 为审批界面拉取当前所有【正在申请转换等待审批】的 BOM 列表
 * 根据访问者的权限深度（部门职级）过滤可见的数据行。
 * @param {string} material - 用于检索过滤的物料编号（可为空查全量）
 * @param {string} dept - 访问者的部门编号 (决定其审批权限)
 * @returns {Array} 符合审批要求的 BOM 数据列表
 */
function getBOMApproveDataFromSheet(material, dept) {
  try {
    let arrAuth = [];
    let auth = dept.charAt(2);
    // 简单的职级划分权限隔离
    if(auth == 1){ arrAuth = ["0", "1"]; } else if(auth == 2){ arrAuth = ["2"]; }
    let output = [];
    let ssTf = getSheetCached(wsTfBomId, "TF BOM masterdata Header");
    let data = ssTf.getRange(2, 1, ssTf.getLastRow() - 1, 16).getValues();
    
    if(material !== '') {
      for (let i = 0; i < data.length; i++) {
        // 条件：料号匹配 且 申请标志位为Y 且 在权限组内
        if(data[i][4].indexOf(material) !== -1 && data[i][12] === 'Y' && arrAuth.indexOf(data[i][15].charAt(1)) != -1) {
          try { output.push([data[i][0], data[i][1], data[i][2], data[i][3], data[i][4], data[i][7], data[i][8], data[i][10] ? Utilities.formatDate(new Date(data[i][10].toString()), "Asia/Shanghai", "yyyy-MM-dd") : "", data[i][13], data[i][14] ? Utilities.formatDate(new Date(data[i][14].toString()), "Asia/Shanghai", "yyyy-MM-dd HH:mm:ss").toString() : "", data[i][15], data[i][5]]); } catch(error) {}
        }
      }
    } else {
      for (let i = 0; i < data.length; i++) {
        if (data[i][12] === 'Y' && arrAuth.indexOf(data[i][15].charAt(1)) != -1) {
          try { output.push([data[i][0], data[i][1], data[i][2], data[i][3], data[i][4], data[i][7], data[i][8], Utilities.formatDate(new Date(data[i][10].toString()), "Asia/Shanghai", "yyyy-MM-dd"), data[i][13], Utilities.formatDate(new Date(data[i][14].toString()), "Asia/Shanghai", "yyyy-MM-dd HH:mm:ss").toString(), data[i][15], data[i][5]]); } catch(error) {}
        }
      }
    }
    if(output.length > 0){ return ["OK", output]; } else { return ["NO", "没找到申请记录"]; }
  } catch(e) { return ["NO", e.toString()]; }
}

/**
 * ==========================================
 * 8. 高级引擎：复杂临时计划插单与 SKU 分解算法
 * ==========================================
 */

/**
 * 将前端录入的临时紧急排产计划推入系统（核心插单函数）
 * 本函数是极其复杂的工厂逻辑引擎：它会读取 Cell Map(模具配比映射) 和 BOM 表，
 * 智能将一单主 PO(采购订单) 按照模孔出件比例，自动分解为多个独立注塑机的小计划并写入数据库。
 * @param {Array} arrData - 包含计划日期、班次、原PO、物料号、排产数量等元素的新增计划行
 * @param {string} planName - 目标排产表名称 (如 "All_INJ")
 * @returns {Array} 插入结果及错误提示信息
 */
function writeInsertPlan(arrData, planName) {
  try {
    let wipProcess = "牙柄";
    if(arrData[10] != "") { wipProcess = "手包"; }
    // 组装用于备份追溯的基础日志数组
    let arrDataTurnOver = [arrData].map(v => [v[1], "TB1", wipProcess, v[2], v[3], v[4], v[9], v[7], v[8].toString().substring(0, 8), Number(v[16]), v[5] + '[班插]', v[8].toString().substring(18, 100), v[10]]);
    // 对数组列进行错位拼接，对齐到 TurnOver 归档表的宽度要求
    arrDataTurnOver[0].splice(12, 0, "", "", "");
    arrDataTurnOver[0].splice(11, 0, "", "", "", "");
    arrDataTurnOver[0].splice(9, 0, "");
    arrDataTurnOver[0].splice(1, 0, "", "", "", "", "", "", "");
    arrDataTurnOver[0].unshift("", "");
    arrDataTurnOver[0][6] = Number(arrData[16]);
    
    // 获取在线破碎和额外辅助品规则
    let sbn_specialInfo = getSheetCached(wsInsertPlanSpecialId, "SpecialInfo");
    let arrSpecialInfo = sbn_specialInfo.getRange(2, 1, sbn_specialInfo.getLastRow() - 1, 12).getValues();
    let arrOnlineCrush = arrSpecialInfo.filter(v => { return v[0] != "" && v[1] != ""; }).map(v => { return v[0] + v[1]; });
    let arrAdd_E = arrSpecialInfo.filter(v => { return v[5] != ""; }).map(v => { return v[5]; });

    let arrError = [["错误提示", "SKU-模具号-机台"]];
    let cellRatio = ""; let aemNums = 0; let poNums = 0;
    let arrNewPlan = [];
    let arrSourcePlan = arrDataTurnOver;
    
    // 日期去重和格式清洗
    for(let i = 0; i < arrSourcePlan.length; i++) { arrSourcePlan[i][12] = dateToStr(arrSourcePlan[i][12]); }
    let arrSourcePlanDeduplication = [];
    for(let i = 0; i < arrSourcePlan.length; i++) {
      arrSourcePlan[i][12] = Utilities.formatDate(new Date(arrSourcePlan[i][12]), "Asia/Shanghai", "yyyy-MM-dd");
      if(arrSourcePlanDeduplication.indexOf(arrSourcePlan[i][12] + "|" + arrSourcePlan[i][13] + "|" + arrSourcePlan[i][14]) == -1) {
        arrSourcePlanDeduplication.push(arrSourcePlan[i][12] + "|" + arrSourcePlan[i][13] + "|" + arrSourcePlan[i][14]);
      }
    }
    
    let arrFilterPlan = []; let arrSubTotalSku = []; let arrDbSku = []; let arrDbSkuDeduplication = [];
    let indexBomJsonData = -1; let indexCellMapJsonData = -1;
    let strSubSku = ""; let strSubSkuDes = "";
    
    // 核心遍历：开始按需切分计划
    for(let i = 0; i < arrSourcePlanDeduplication.length; i++) {
      arrFilterPlan = arrSourcePlan.filter(v => { return v[12] + "|" + v[13] + "|" + v[14] == arrSourcePlanDeduplication[i]; });
      
      // 判断如果是特殊的 HT (混合自动线) 机台，需要切分订单
      if(arrFilterPlan[0][14].toString().substring(2, 4) == "HT" && arrFilterPlan[0][15] != "") {
        poNums = arrFilterPlan.length;
        if(poNums == 1) {
          // 单 PO 情况：查询 BOM 与 Cell 映射，按照机位（AEM数）产出平均切分订单量
          indexBomJsonData = getBomRowJsonData(arrBomJsonData, "牙柄C#", arrFilterPlan[0][15], "模具号", arrFilterPlan[0][17]);
          indexCellMapJsonData = getCellMapRowJsonData(arrCellMapJsonData, "New Formed Cell", arrFilterPlan[0][14]);
          cellRatio = ""; aemNums = 0; strSubSku = ""; strSubSkuDes = "";
          
          if(indexBomJsonData != -1 && indexCellMapJsonData != -1) {
            arrSubTotalSku = getSubTotalSku(indexBomJsonData);
            cellRatio = arrCellMapJsonData[indexCellMapJsonData]["Cell Ratio"];
            aemNums = sumAemNums(cellRatio); // 算出这套模具对应切分成几条线
            if(!isNaN(aemNums) && aemNums > 0) {
              for(let j = 0; j < aemNums; j++) {
                strSubSku = getSubSku(indexBomJsonData, j, arrBomJsonData);
                strSubSkuDes = getSubSkuDes(indexBomJsonData, j);
                if(strSubSku.toString().length > 2) {
                  // 将拆分后的新计划行注入队列
                  arrNewPlan.push([arrFilterPlan[0][10], arrFilterPlan[0][12], getMixSubSku(strSubSku, arrFilterPlan[0][3], arrFilterPlan[0][11], j, arrFilterPlan[0][15], arrOnlineCrush, arrAdd_E, getSubAem(indexCellMapJsonData, j)), arrFilterPlan[0][25], getSubAem(indexCellMapJsonData, j), getSubSkuOutPlan(strSubSku, arrSubTotalSku, arrFilterPlan[0][6]), arrFilterPlan[0][20], "", "", arrFilterPlan[0][13], strSubSkuDes, strSubSku, getHandPackSku(indexBomJsonData, arrFilterPlan[0][11], strSubSku, arrFilterPlan[0][15]), arrFilterPlan[0][17] + "-" + arrFilterPlan[0][14], getPo(strSubSku, arrFilterPlan[0][15], arrFilterPlan[0][2]), getPoInfo(strSubSku, arrFilterPlan[0][15], arrFilterPlan[0][2], arrFilterPlan[0][6], arrFilterPlan[0][11])]);
                }
              }
            } else { arrError.push(["模具配比错误", arrFilterPlan[0][15] + " " + arrFilterPlan[0][17] + " " + arrFilterPlan[0][14]]); }
          } else { arrError.push(["未找到标准数据", arrFilterPlan[0][15] + " " + arrFilterPlan[0][17] + " " + arrFilterPlan[0][14]]); }
        } else {
          // 多个连体 PO 的情况，必须保证底层底板相同方可拆分
          arrDbSku = []; arrDbSkuDeduplication = [];
          for(let k = 0; k < arrFilterPlan.length; k++) {
            indexBomJsonData = getBomRowJsonData(arrBomJsonData, "牙柄C#", arrFilterPlan[k][15], "模具号", arrFilterPlan[k][17]);
            indexCellMapJsonData = getCellMapRowJsonData(arrCellMapJsonData, "New Formed Cell", arrFilterPlan[k][14]);
            strSubSku = "";
            if(indexBomJsonData != -1 && indexCellMapJsonData != -1) {
              arrDbSku.push(getSubSku(indexBomJsonData, 0));
              strSubSku = arrFilterPlan[k][15];
              arrNewPlan.push([arrFilterPlan[k][10], arrFilterPlan[k][12], getMixSubSku(getSubSku(indexBomJsonData, 0), arrFilterPlan[k][3], arrFilterPlan[k][11], 0, arrFilterPlan[k][15], arrOnlineCrush, arrAdd_E, getSubAemFromMultiple(indexCellMapJsonData, 0)), arrFilterPlan[k][25], getSubAemFromMultiple(indexCellMapJsonData, 0), arrFilterPlan[k][6], arrFilterPlan[k][20], "", "", arrFilterPlan[k][13], getSubSkuDes(indexBomJsonData, 0), getSubSku(indexBomJsonData, 0), getHandPackSku(indexBomJsonData, arrFilterPlan[k][11], getSubSku(indexBomJsonData, 0), arrFilterPlan[k][15]), arrFilterPlan[k][17] + "-" + arrFilterPlan[k][14], getPo(getSubSku(indexBomJsonData, 0), arrFilterPlan[k][15], arrFilterPlan[k][2]), getPoInfo(getSubSku(indexBomJsonData, 0), arrFilterPlan[k][15], arrFilterPlan[k][2], arrFilterPlan[k][6], arrFilterPlan[k][11])]);
              arrNewPlan.push([arrFilterPlan[k][10], arrFilterPlan[k][12], getMixSubSku(strSubSku, arrFilterPlan[k][3], arrFilterPlan[k][11], 1, arrFilterPlan[k][15], arrOnlineCrush, arrAdd_E, getSubAemFromMultiple(indexCellMapJsonData, (k + 1))), arrFilterPlan[k][25], getSubAemFromMultiple(indexCellMapJsonData, (k + 1)), arrFilterPlan[k][6], arrFilterPlan[k][20], "", "", arrFilterPlan[k][13], arrFilterPlan[k][16], strSubSku, getHandPackSku(indexBomJsonData, arrFilterPlan[k][11], strSubSku, arrFilterPlan[k][15]), arrFilterPlan[k][17] + "-" + arrFilterPlan[k][14], getPo(strSubSku, arrFilterPlan[k][15], arrFilterPlan[k][2]), getPoInfo(strSubSku, arrFilterPlan[k][15], arrFilterPlan[k][2], arrFilterPlan[k][6], arrFilterPlan[k][11])]);
            } else { arrError.push(["未找到标准数据", arrFilterPlan[k][15] + " " + arrFilterPlan[k][17] + " " + arrFilterPlan[k][14]]); }
          }
          if(arrDbSku.length > 1) {
            arrDbSkuDeduplication = Array.from(new Set(arrDbSku));
            if(arrDbSkuDeduplication.length > 1) { arrError.push(["同一机台有不同的底板SKU", arrFilterPlan[0][15] + " " + arrFilterPlan[0][17] + " " + arrFilterPlan[0][14]]); }
          }
        }
      } else {
        // 常规单一非切分机台，不做拆解原路送回列阵
        for(let k = 0; k < arrFilterPlan.length; k++) {
          arrNewPlan.push([arrFilterPlan[k][10], arrFilterPlan[k][12], getMixSku(arrFilterPlan[k][15], arrFilterPlan[k][11], arrAdd_E, arrFilterPlan[k][14]), arrFilterPlan[k][25], arrFilterPlan[k][14], arrFilterPlan[k][6], arrFilterPlan[k][20], "", "", arrFilterPlan[k][13], arrFilterPlan[k][16], arrFilterPlan[k][15], getHandPackSkuNoSplit(arrFilterPlan[k][29], arrFilterPlan[k][11]), arrFilterPlan[k][17] + "-" + arrFilterPlan[k][14], arrFilterPlan[k][2], getPoInfo(arrFilterPlan[k][15], arrFilterPlan[k][15], arrFilterPlan[k][2], arrFilterPlan[k][6], arrFilterPlan[k][11])]);
        }
      }
    }
    
    // 若在切解中产生报错则退回，否则推入数据库
    if(arrError.length > 1) { return ["NO", arrError.join("\n")]; } 
    else {
      let resultInfo = handClearAndUpdateInjPlan(arrNewPlan, planName);
      if(resultInfo[0] == "OK") { return ["OK", "插入成功"]; } 
      else { return ["NO", resultInfo[1]]; }
    }
  } catch(e) { return ["NO", e.toString()]; }
}

/**
 * (内部协助模块) 用于最终清洗并追加已处理好的临时计划到主系统计划表中
 * @param {Array} arrPlanData - 洗排切割后的独立注塑单元任务数组
 * @param {string} planName - 目标写入的主表名 (如 "All_INJ")
 * @returns {Array} 追加状态
 */
function handClearAndUpdateInjPlan(arrPlanData, planName) {
  try {
    const sbn_PreBom = getSheetCached(wsPreBomId, "★BOM★");
    let arrPreBom = sbn_PreBom.getRange(1, 1, sbn_PreBom.getLastRow(), sbn_PreBom.getLastColumn()).getDisplayValues();
    let ssTurnOver = getSheetCached(wsTurnOverId, "All_INJ");
    // 将计划压入追溯历史库
    arrPlanData = arrPlanData.map(v => { return [v[14], dateToStr(v[1]), v[9], v[4], v[6], v[13] + "_" + v[3], v[2], v[10], v[11], v[12], v[5], "", ""]; });
    for(let i = 0; i < arrPlanData.length; i++) { ssTurnOver.appendRow(arrPlanData[i]); }
    
    // 计算并插入带动态关联公式的高阶主数据
    let sbnYsj = getSheetCached(wsPlanId, "Machine");
    let arrAemRate = sbnYsj.getRange(2, 1, sbnYsj.getLastRow() - 1, 7).getValues();
    let arrAem = arrAemRate.map(function(x){ return x[0]; });
    let ssDest = getSheetCached(wsPlanId, planName);
    if(arrPlanData.length > 0) {
      let arrNewData = arrPlanData.map(function(x){ return [x[0], dateToStr(x[1]), x[2], x[3], x[4], x[5], x[6], x[7], x[8], x[9], x[10], x[11], Number(dateToNum(x[1]) + "" + changeShiftToNum(x[2])), "TB1", x[11], x[11]]; });
      for(let k = 0; k < arrNewData.length; k++) {
        arrNewData[k].unshift("");
        arrNewData[k][8] = arrNewData[k][8] + "【|】" + getPreNoFromSku(arrNewData[k][7].replace("手包", ""), arrPreBom);
        // 压入巨量的跨表联合查询 Google Sheet 公式
        arrNewData[k].splice(11, 0, '=iferror(VLOOKUP(indirect("T"&row())&"_"&indirect("J"&row()),\'箱含量\'!A:E,5,0),"")&"【|】"&iferror(vlookup(indirect("J"&row()),\'特殊品种\'!D:E,2,0),"")&"【|】"&iferror(vlookup(indirect("J"&row()),\'特殊品种\'!D:F,3,0),"")&"【|】"&iferror(vlookup(indirect("J"&row()),\'特殊品种\'!A:B,2,0),"")&"【|】"&iferror(vlookup(left(indirect("E"&row()),8),CurrentShiftPointCheckInfo!C:N,12,0),"")&"【|】"&iferror(join("|",ArrayFormula(vlookup(split(indirect("B"&row()),char(10)),PlanPoCoois!A:H,8,0))),"")&"【|】"&iferror(vlookup(indirect("J"&row()),filter(indirect("特殊品种!Q:R"),indirect("特殊品种!R:R")<>""),2,0),"")', '=iferror(vlookup(left(indirect("E"&row()),8),Machine!A:C,2,0),"")', '=iferror(vlookup(left(indirect("E"&row()),8),Machine!A:D,3,0),"")', '', arrNewData[k][11]);
        // 算出由于插单造成的早中夜分拆量
        arrNewData[k][14] = calcNewShiftOut(arrNewData[k][4], arrAem, arrAemRate, arrNewData[k][16]);
        ssDest.appendRow(arrNewData[k]);
      }
      SpreadsheetApp.flush();
    }
    return ["OK", "插入成功"];
  } catch(e) { return ["No", e.toString()]; }
}

/**
 * 辅助方法：读取复合式 Cell 模具映射代号中的产线总数（如 "1V2" 计算得出为 3）
 * @param {string} cellRatio - 描述子线模具比例的字符标识
 * @returns {number} 解析后的线体总数量
 */
function sumAemNums(cellRatio) {
  let result = 0; let arrTemp = cellRatio.split("V");
  for(let i = 0; i < arrTemp.length; i++) { result = result + Number(arrTemp[i]); }
  return result;
}

/**
 * 辅助方法：从单一产品拆解中定位所属的物理机头编码
 * @param {number} indexCellMapJsonData - 该 Cell Map 行的缓存位置
 * @param {number} index - 属于第几根子线
 * @returns {string} 精确物理机台名字
 */
function getSubAem(indexCellMapJsonData, index) {
  let result = ""; let cellType = ""; let judgeCellType = ""; let cellMachine = "";
  cellType = arrCellMapJsonData[indexCellMapJsonData]["Cell Type"];
  cellMachine = arrCellMapJsonData[indexCellMapJsonData]["New Formed Cell"];
  if(cellType.slice(-2) == "6X" || cellMachine == "H2HTB363") { judgeCellType = "6X"; cellMachine = "(" + cellMachine + ")"; }
  if(index == 0) { result = arrCellMapJsonData[indexCellMapJsonData]["HIM/Auto"]; } 
  else {
    if(judgeCellType == "6X") { result = arrCellMapJsonData[indexCellMapJsonData]["VIM-" + index] + cellMachine; } 
    else { result = arrCellMapJsonData[indexCellMapJsonData]["VIM-" + index]; }
  }
  return result;
}

/**
 * 辅助方法：从多订单联合拆解中定位所属的物理机头编码
 * @param {number} indexCellMapJsonData - 该 Cell Map 行的缓存位置
 * @param {number} poIndex - 第几个级联的订单序列
 * @returns {string} 精确物理机台名字
 */
function getSubAemFromMultiple(indexCellMapJsonData, poIndex) {
  let result = ""; let cellType = ""; let judgeCellType = ""; let cellMachine = "";
  cellType = arrCellMapJsonData[indexCellMapJsonData]["Cell Type"];
  cellMachine = arrCellMapJsonData[indexCellMapJsonData]["New Formed Cell"];
  if(cellType.slice(-2) == "6X" || cellMachine == "H2HTB363") { judgeCellType = "6X"; cellMachine = "(" + cellMachine + ")"; }
  if(poIndex == 0) { result = arrCellMapJsonData[indexCellMapJsonData]["HIM/Auto"]; } 
  else {
    if(judgeCellType == "6X") { result = arrCellMapJsonData[indexCellMapJsonData]["VIM-" + (poIndex)] + cellMachine; } 
    else { result = arrCellMapJsonData[indexCellMapJsonData]["VIM-" + (poIndex)]; }
  }
  return result;
}

/**
 * 辅助方法：提取多色/混合注塑模块对应的特定材料 SKU
 * @param {number} indexBomJsonData - 内存 Json 数据定位列
 * @param {number} index - 所需查询的腔位号（0是主件，其余是附件件）
 * @returns {string} 获取到的具体 SKU 物料名
 */
function getSubSku(indexBomJsonData, index) {
  let result = "";
  if(index == 0) { result = arrBomJsonData[indexBomJsonData]["H编码C#"]; } 
  else { result = arrBomJsonData[indexBomJsonData]["V" + index + "编码C#"]; }
  return result;
}

/**
 * 辅助方法：提取上述对应混合物料的纯中文描述
 * @param {number} indexBomJsonData - 内存 Json 数据定位列
 * @param {number} index - 腔位号
 * @returns {string} 提取描述
 */
function getSubSkuDes(indexBomJsonData, index) {
  let result = "";
  if(index == 0) { result = arrBomJsonData[indexBomJsonData]["H描述"]; } 
  else { result = arrBomJsonData[indexBomJsonData]["V" + index + "描述"]; }
  return result;
}

/**
 * 辅助方法：提取手包业务专用成品物料 SKU (当存在打包指令时)
 * @param {number} indexBomJsonData - BOM行
 * @param {string} judgePk - 标识是否带有包材判定 ("手包")
 * @param {string} subSku - 当前对比底层基板子型号
 * @param {string} finalSku - 最终订单需要落地的总型号
 * @returns {string} 包材专用识别码
 */
function getHandPackSku(indexBomJsonData, judgePk, subSku, finalSku) {
  let result = "";
  if(judgePk == "手包" && subSku == finalSku) { result = arrBomJsonData[indexBomJsonData]["手包成品SKU"]; }
  return result;
}

/**
 * 辅助方法：非混打产线纯提取手包标记
 * @param {string} handPkSku - 原生包材主号
 * @param {string} judgePk - 标识位
 * @returns {string} 剥离出来的手包型号
 */
function getHandPackSkuNoSplit(handPkSku, judgePk) {
  let result = "";
  if(judgePk == "手包") { result = handPkSku; }
  return result;
}

/**
 * 辅助方法：一键捞取该组合模具下涉及到的所有合法子料件号
 * @param {number} indexBomJsonData - BOM 检索列
 * @returns {Array} 排除了空白的纯净 SKU 数组集合
 */
function getSubTotalSku(indexBomJsonData) {
  let arrResult = [];
  for(let i = 0; i < 5; i++) {
    if(i == 0) { arrResult.push(arrBomJsonData[indexBomJsonData]["H编码C#"]); } 
    else { arrResult.push(arrBomJsonData[indexBomJsonData]["V" + i + "编码C#"]); }
  }
  arrResult = arrResult.filter(v => { return v != ""; });
  return arrResult;
}

/**
 * 辅助方法：按照模具总共出现的对应孔数权重，平分主订单量，计算得到该子线的标准派单量
 * @param {string} strSubSku - 该切片线生产的具体子型号
 * @param {Array} arrSubTotalSku - 上面提取的全量产线出孔明细
 * @param {number} totalOutPlan - 这张 PO 初始的总派单数量
 * @returns {number} 属于这台物理机头被分解指派的数量
 */
function getSubSkuOutPlan(strSubSku, arrSubTotalSku, totalOutPlan) {
  let result = totalOutPlan; let n = 0;
  for(let i = 0; i < arrSubTotalSku.length; i++) { if(arrSubTotalSku[i] == strSubSku) { n = n + 1; } }
  result = Math.round(totalOutPlan / n);
  return result;
}

/**
 * 辅助方法：根据特殊加工类型（诸如在线破损回收等情况），加缀相应的追踪后缀如 "_H" 或 "_E"
 * @param {string} sku - 基准 sku
 * @param {string} moldType - 模具种类标识符
 * @param {string} judgePk - 包装标识符
 * @param {number} index - 线号索引
 * @param {string} finalSku - 总成目标sku
 * @param {Array} arrOnlineCrush - 在线破碎识别表
 * @param {Array} arrAdd_E - 特定增强修饰符组
 * @param {string} subAem - 具体加工机器码
 * @returns {string} 智能组装后的专属物料跟踪短码
 */
function getMixSubSku(sku, moldType, judgePk, index, finalSku, arrOnlineCrush, arrAdd_E, subAem) {
  let result = sku;
  if(index == 0) {
    if(moldType.substring(0, 1) == "H" || arrOnlineCrush.indexOf(subAem + sku) != -1) { result = sku + "_H"; }
    if(arrAdd_E.indexOf(subAem) != -1) { result = sku + "_E"; }
  } else {
    if(moldType.substring(1, 2) == "H") { result = sku + "_H"; }
  }
  if(judgePk == "手包" && sku == finalSku) { result = result + "手包"; }
  return result;
}

/**
 * 辅助方法：针对不需要做子线打散拆分的大机台，直接施加特有的后处理跟踪标志位
 * @param {string} sku - 基准物料号
 * @param {string} judgePk - 手包状态位
 * @param {Array} arrAdd_E - 增强物料名单表
 * @param {string} subAem - 生产代号
 * @returns {string}
 */
function getMixSku(sku, judgePk, arrAdd_E, subAem) {
  let result = sku;
  if(arrAdd_E.indexOf(subAem) != -1) { result = sku + "_E"; }
  if(judgePk == "手包") { result = result + "手包"; }
  return result;
}

/**
 * 辅助方法：通过 SKU 防伪对比验证合法并放行 PO 单号
 * @param {string} sku - 本线 SKU
 * @param {string} cellSku - 表征的主件源 SKU
 * @param {string} po - 采购单号
 * @returns {string} 合法则返回原 PO 字符串，否则为空
 */
function getPo(sku, cellSku, po) {
  let result = ""; if(sku == cellSku) { result = po; } return result;
}

/**
 * 辅助方法：将匹配放行的 PO 数据转置为其后关联模块必需依赖的 JSON 对象格式数组
 * @param {string} sku - 比对参考 SKU
 * @param {string} cellSku - 基础源线 SKU
 * @param {string} po - 原指令单号
 * @param {string|number} lineNum - 线路号/项号
 * @param {string} judgePk - 验证组装特征
 * @returns {string} 包含指令详参的被 stringify 后的阵列包
 */
function getPoInfo(sku, cellSku, po, lineNum, judgePk) {
  let result = "";
  if(sku == cellSku) { result = JSON.stringify([[po, lineNum, 0, lineNum, 0, judgePk, ""]]); }
  return result;
}

/**
 * 辅助方法：安全将对象转回为不含时间与时区的标准日期字串格式 YYYY-MM-DD
 * @param {Date|string} dateTemp - 需要解析的标准或非标准时间格式数据
 * @returns {string}
 */
function dateToStr(dateTemp) {
  let strTemp = "0000-00-00";
  if(dateTemp != "") { strTemp = Utilities.formatDate(new Date(dateTemp.toString()), "Asia/Shanghai", "yyyy-MM-dd").toString(); }
  return strTemp;
}

/**
 * 辅助方法：将时间快速提取转换为利于关系型数字排序与查重的连体印记 YYYYMMDD
 * @param {Date|string} dateTemp - 被压入源数据
 * @returns {string}
 */
function dateToNum(dateTemp) {
  let strTemp = "00000000";
  if(dateTemp != "") { strTemp = Utilities.formatDate(new Date(dateTemp.toString()), "Asia/Shanghai", "yyyyMMdd").toString(); }
  return strTemp;
}

/**
 * 辅助方法：插单落库前，按班别的基础产能分配权重，把目标总产出打散分摊给当天的三班以提供进度指示参考
 * @param {string} machine - 目标产线机台号
 * @param {Array} arrAem - 设备字典序索引集合
 * @param {Array} arrAemRate - 具体不同设备在早/中/夜拥有的产能权重比例字典
 * @param {number} oldShiftOut - 当前插单预估总要求产量
 * @returns {string} 一句呈现供前端阅览的切分比例如 "100(早) 150(中) 50(夜)"
 */
function calcNewShiftOut(machine, arrAem, arrAemRate, oldShiftOut) {
  let strNewShiftOut = "";
  let position = arrAem.indexOf(machine);
  if(position != -1) { strNewShiftOut = Math.round(oldShiftOut * arrAemRate[position][4]) + "(早)　" + Math.round(oldShiftOut * arrAemRate[position][5]) + "(中)　" + Math.round(oldShiftOut * arrAemRate[position][6]) + "(夜)"; } 
  else { strNewShiftOut = Math.round(oldShiftOut * 1) + "(早)　" + Math.round(oldShiftOut * 1) + "(中)　" + Math.round(oldShiftOut * 1) + "(夜)"; }
  return strNewShiftOut;
}

/**
 * 辅助方法：物理班别代号向系统查询权重的强制量化映射转换
 * @param {string} str - "早", "中", "夜"
 * @returns {number} 排序权：1代表夜班、2早班、3中班 (用于按时间线推演)
 */
function changeShiftToNum(str) {
  let shiftNum = 0;
  if(str == "夜"){ shiftNum = 1; } 
  else if(str == "早"){ shiftNum = 2; } 
  else if(str == "中"){ shiftNum = 3; }
  return shiftNum;
}