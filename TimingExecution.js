function emailOneDayInAdavance() {
  let ws=SpreadsheetApp.openById(wsFinalMoldCheckId);
  let ss=ws.getSheetByName("FinalMoldCollectionRecord");
  let lastTime=new Date().getTime()-24*3600*1000;
  let arrFinalMoldCheckRecord=ss.getRange(1,1,ss.getLastRow(),15).getValues();
  let arrFinalMoldCheckRecordFilter=arrFinalMoldCheckRecord.filter(function(x){return x[14]=="Ongoing"}).map(function(x){return [x[0],x[3],x[4],x[5],Utilities.formatDate(new Date(x[6].toString()),"Asia/Shanghai","yyyy-MM-dd HH:mm:ss"),x[7],JSON.parse(x[8]).join("\n"),JSON.parse(x[9]).join("\n"),x[10],x[11],x[12],Utilities.formatDate(new Date(x[13].toString()),"Asia/Shanghai","yyyy-MM-dd"),x[14]]});
  if(arrFinalMoldCheckRecordFilter.length>0){
    let followMail="andy_lu@colpal.com,fei_lin@colpal.com,yong_ji@colpal.com";
    let widths=[6,6,6,6,6,6,13,13,7,13,6,6,6];
    let head_backgrounds=["unset","unset","unset","unset","unset","unset","unset","unset","unset","unset","unset","unset","unset"];
    let aligns=["center","center","center","center","center","center","center","center","center","center","center","center","center"];
    let scale=90;
    let col=-1;
    let title="有末模问题需要跟进";
    let call="<i>Auto send email,do not reply</i>";
    let describe1="<b>登陆入口：</b>"+publishingWebsize;
    let describe2="";
    let describe3="打开界面后，输入工号密码，点击【末模】，填写措施及日期后，点击保存";
    let describe4="";
    let describe5="确认以下末模问题已有措施：";
    arrFinalMoldCheckRecordFilter.unshift(["模具号","机台","SKU","描述","收集时间","收集人","外观","生产过程","确认人","维修描述","维修人","维修日期","维修状态"]);
    let tomail=followMail;
    let describe6=General_Htmltable(arrFinalMoldCheckRecordFilter,widths,head_backgrounds,aligns,scale,col)
    let ccmail="jin_zhang@colpal.com,fiona_zhao@colpal.com,ping_li@colpal.com,andy_wang@colpal.com";
    Mail_HTML_TXT(tomail,title,call,describe1,describe2,describe3,describe4,describe5,describe6,ccmail)
  }
}