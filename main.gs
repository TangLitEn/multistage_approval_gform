function checkQuota(email_count){
  var emailQuotaRemaining = MailApp.getRemainingDailyQuota();
  Logger.log("Remaining email quota: " + emailQuotaRemaining);
  if(email_count <= emailQuotaRemaining){
    return true
  }
  else {
    return false
  }
}

function testQuota(){
  checkQuota(0);
}
