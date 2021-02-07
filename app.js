var ui = SpreadsheetApp.getUi();
function onOpen(e) {

  ui.createMenu("Orders Manager").addItem("Get orders", "getOrdersEmails").addToUi();

}

function getOrdersEmails() {
  // var input = ui.prompt('Label Name', 'Enter the label name that is assigned to your emails:', Browser.Buttons.OK_CANCEL);

  // if (input.getSelectedButton() == ui.Button.CANCEL){
  //   return;
  // }
  var activeData = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange("B2:B9999").getValues();
  var label = GmailApp.getUserLabelByName("orders");
  var threads = label.getThreads();
  var count = 0;
  for (var i = threads.length - 1; i >= 0; i--) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var emailBody = message.getPlainBody();
      var keyWord = "Your order number is";
      var regExp = new RegExp("(?<=" + keyWord + ").*\\w", 'g');
      var id = (emailBody.match(regExp)) ? emailBody.match(regExp).toString().trim() : "";
      var resultObject = activeData.findIndex(orderId => {
        return orderId == id;
      });
      if (resultObject == -1) {
        extractDetails(message);
        count++;
      }
    }
    threads[i].removeLabel(label);
  }
  if (count != 0) {
    ui.alert("Đã có "+ count + " đơn mới được thêm vào !");
  }else{
    ui.alert("Chưa có đơn được thêm vào !");
  }
}

function extractDetails(message) {

  var transactions = []
  var emailBody = message.getPlainBody();
  var regExp;

  var emailKeywords = {
    orderId: "Your order number is",
    productName: "Item:",
    style: "Style:",
    color: "Color:",
    size: "Size:",
    quantity: "Quantity:",
    personalisation: "Personalization:",
    shop: "Shop:",
    shippingName: "class='name'>",
    shippingAddress1: "class='first-line'>",
    shippingAddress2: "class='second-line'>",
    shippingZipcode: "class='zip'>",
    shippingCity: "class='city'>",
    shippingState: "class='state'>",
    shippingCountry: "class='country-name'>",
    // shippingPhone: "Null",
    // shippingEmail: "Null",
    // shippingPack: "Null",
    shippingCost: "Shipping:",
    delivery: "Delivery:",
    transactionId: "Transaction ID:",
  }

  regExp = new RegExp("(?<=" + emailKeywords.transactionId + ")(.*\\r\\n)?(.*\\r\\n)?(.*\\r\\n)?(.*\\r\\n)?(.*\\r\\n)?", 'gm');
  if (emailBody.match(regExp)) {
    Logger.log('Transactions: ' + emailBody.match(regExp));
    var transactionIds = emailBody.match(regExp);
    for (let index_one = 0; index_one < transactionIds.length; index_one++) {
      var transaction = {
        // transactionId: "",
        date: "",
        orderId: "",
        productName: "",
        sku: "",
        style: "",
        color: "",
        size: "",
        side: "",
        quantity: "",
        personalisation: "",
        shop: "",
        shippingName: "",
        shippingAddress1: "",
        shippingAddress2: "",
        shippingCity: "",
        shippingState: "",
        shippingZipcode: "",
        shippingCountry: "",
        shippingPhone: "",
        shippingEmail: "",
        shippingPack: "",
        shippingCost: "",
        subject: "",
      }

      var stringA = transactionIds[index_one].split(/\r?\n/);
      for (let index_two = 0; index_two < stringA.length; index_two++) {
        // transaction.transactionId = stringA[0].toString().trim();
        if (stringA[index_two].includes("Item:")) {
          var item = stringA[index_two].split(":")
          transaction.productName = item[1].toString().trim();
        }
        if (stringA[index_two].includes("Style")) {
          var item = stringA[index_two].split(":")
          transaction.style = item[1].toString().trim();
        }
        if (stringA[index_two].includes("Color")) {
          var item = stringA[index_two].split(":")
          transaction.color = item[1].toString().trim();
        }
        if (stringA[index_two].includes("Size")) {
          var item = stringA[index_two].split(":")
          transaction.size = item[1].toString().trim();
        }
        if (stringA[index_two].includes("Quantity")) {
          var item = stringA[index_two].split(":")
          transaction.quantity = item[1].toString().trim();
        }
      }

      transaction.date = message.getDate();

      regExp = new RegExp("(?<=" + emailKeywords.orderId + ").*\\w", 'g');
      transaction.orderId = (emailBody.match(regExp)) ? emailBody.match(regExp).toString().trim() : "";

      regExp = new RegExp("(?<=" + emailKeywords.personalisation + ").*", 'g');
      transaction.personalisation = (emailBody.match(regExp)) ? emailBody.match(regExp).toString().trim() : "";

      regExp = new RegExp("(?<=" + emailKeywords.shop + ").*", 'g');
      transaction.shop = (emailBody.match(regExp)) ? emailBody.match(regExp).toString().trim() : "";

      var regExpShippingCost = new RegExp("(?<=" + emailKeywords.shippingCost + ").*\\w", 'g');
      if (emailBody.match(regExpShippingCost)) {
        transaction.shippingCost = emailBody.match(regExpShippingCost).toString().trim();
      }
      var regExpDelivery = new RegExp("(?<=" + emailKeywords.delivery + ").*\\w", 'g');
      if (emailBody.match(regExpDelivery)) {
        transaction.shippingCost = emailBody.match(regExpDelivery).toString().trim();
      }

      regExp = new RegExp("(?<=" + emailKeywords.shippingName + ")([\\w\\s]*)", 'g');
      transaction.shippingName = (emailBody.match(regExp)) ? emailBody.match(regExp).toString().trim() : "";

      regExp = new RegExp("(?<=" + emailKeywords.shippingAddress1 + ")[^<]*", 'g');
      transaction.shippingAddress1 = (emailBody.match(regExp)) ? emailBody.match(regExp).toString().trim() : "";

      regExp = new RegExp("(?<=" + emailKeywords.shippingAddress2 + ")[^<]*");
      transaction.shippingAddress2 = (emailBody.match(regExp)) ? emailBody.match(regExp).toString().trim() : "";

      regExp = new RegExp("(?<=" + emailKeywords.shippingZipcode + ")([\\w\\s]*-?.\\w+)", 'g');
      transaction.shippingZipcode = (emailBody.match(regExp)) ? emailBody.match(regExp).toString().trim() : "";

      regExp = new RegExp("(?<=" + emailKeywords.shippingCity + ")([\\w\\s]*)", 'g');
      transaction.shippingCity = (emailBody.match(regExp)) ? emailBody.match(regExp).toString().trim() : "";

      regExp = new RegExp("(?<=" + emailKeywords.shippingState + ")([\\w\\s]*)", 'g');
      transaction.shippingState = (emailBody.match(regExp)) ? emailBody.match(regExp).toString().trim() : "";

      regExp = new RegExp("(?<=" + emailKeywords.shippingCountry + ")([\\w\\s]*)", 'g');
      transaction.shippingCountry = (emailBody.match(regExp)) ? emailBody.match(regExp).toString().trim() : "";

      transactions.push(transaction)
    }
  }

  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  var emailDataArr = [];

  if (transactions[0].orderId) {
    transactions.forEach(element => {
      var row = []
      for (var propName in element) {
        row.push(element[propName]);
      }
      emailDataArr.push(row);
    });
    for (let index = 0; index < emailDataArr.length; index++) {
      activeSheet.appendRow(emailDataArr[index]);
    }
  }
}


