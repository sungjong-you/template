const URL_OF_THE_LATEST_DEPLOYMENT = "https://script.google.com/macros/s/AKfycbyraDFl8DZGbz_4tZjS9HSCeGINfwL0cnzFexrp36rm4uA58lpHHJH-6FYf8syehcwnHg/exec"

function checkCashPaymentInDateTime_deployment() {
  const url = URL_OF_THE_LATEST_DEPLOYMENT + "?func=" + "checkCashPaymentInDateTime"
  UrlFetchApp.fetch(url)
}

  
function doGet(e) {
  try {
    if (e.parameter.func === "checkCashPaymentInDateTime") {
      checkCashPaymentInDateTime();
    }
  }catch (error) {
  }
}
