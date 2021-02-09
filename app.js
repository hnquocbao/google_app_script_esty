var iban_countries = SpreadsheetApp.openById("1Iek1JLB8IEBbKbizvn2DbZmEsjT84zPBeojUNkh_Ing").getDataRange().getValues();
var mainFile = SpreadsheetApp.openById("1bO6gsAO-c9JSiEsR5PlOOptTJZpTvBqPWELaH1QKnTs"); // Bỏ Id của Sheet vào đây
var mainSheet = mainFile.getSheets()[0];

function getOrdersEmails() {
  var activeData = mainSheet.getRange("B2:B9999").getValues();
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
      var resultObject;
      if(emailBody.match(regExp)){
        var id = (emailBody.match(regExp)) ? emailBody.match(regExp).toString().trim() : "";
        if(id && id != ""){
          resultObject = activeData.findIndex(orderId => {return orderId == id;});
        }
      }
      if (resultObject && resultObject == -1) {
        extractDetails(message);
        count++;
      }
    }
    threads[i].removeLabel(label);
  }
  if (count != 0) {
    Logger.log("Đã có "+ count + " đơn mới được thêm vào !");
  }else{
    Logger.log("Chưa có đơn mới được thêm vào !");
  }
}

function extractDetails(message) {

  var transactions = []
  var emailBody = message.getPlainBody();
  var regExp;
  var emailDataArr = [];

  var emailKeywords = {
    orderId: "Your order number is",
    productName: "Item:",
    style: "Style:",
    color: "Color:",
    size: "Size:",
    side: "Side:",
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
    shippingPhone: "",
    shippingEmail: "",
    shippingPack: "",
    shippingCost: "Shipping:",
    delivery: "Delivery:",
    transactionId: "(?<=Shop:.*\\r\\n\\r\\n.*?\\r\\n\\r\\n)(.|\\r\\n)*(?=\\r\\n-*?\\r\\nItem total:)",
    noteFromBuyer: "(?<=Note from.*:\\n.*\\n)(.|\\r\\n)*(?=\\r\\n.*\\r\\nOrder Details)",
    giftMessage:"(?<=Gift message\\n)(.|\\r\\n)*(?=\\r\\nDelivery Address)",
    shippingService: "(?<=Delivery:.*\\n)(.|\\r\\n)*(?=\\r\\nOrder Total:)",
    price: "Item price:",
  }

  regExp = new RegExp(emailKeywords.transactionId,'g');
  if (emailBody.match(regExp)) {
    var transactionIds = emailBody.match(regExp)[0].split(/(?:\r\n){2,}/);;
    for (let index_one = 0; index_one < transactionIds.length; index_one++) {
      var transaction = {
        date: "",
        noteFromBuyer: "",
        giftMessage: "",
        personalisation: "",
        sku: "",
        shop: "",
        orderId: "",
        shippingName: "",
        shippingAddress1: "",
        shippingAddress2: "",
        shippingCity: "",
        shippingState: "",
        shippingZipcode: "",
        shippingCountry: "",
        shippingPhone: "",
        shippingEmail: "",
        productName: "",
        style: "",
        color: "",
        size: "",
        side: "",
        quantity: "",
        designLinkFront: "",
        designLinkBack: "",
        shippingService: "",
        shippingCost: "",
        price: "",
      }

      var stringA = transactionIds[index_one].split(/\r?\n/);
      for (let index_two = 0; index_two < stringA.length; index_two++) {
        if (stringA[index_two].includes("Item:")) {
          transaction.productName = stringA[index_two].split(":")[1].toString().trim();
          continue;
        }
        if (stringA[index_two].includes("Style")) {
          transaction.style = stringA[index_two].split(":")[1].toString().trim();
          continue;
        }
        if (stringA[index_two].includes("Personalisation:")) {
          transaction.personalisation = stringA[index_two].split(":")[1].toString().trim();
          continue;
        }
        if (stringA[index_two].includes("Color")) {
          transaction.color = stringA[index_two].split(":")[1].toString().trim();
          continue;
        }
        if (stringA[index_two].includes("Size")) {
          transaction.size = stringA[index_two].split(":")[1].toString().trim();
          continue;
        }
        if (stringA[index_two].includes("Side")) {
          transaction.side = stringA[index_two].split(":")[1].toString().trim();
          continue;
        }
        if (stringA[index_two].includes("Quantity")) {
          transaction.quantity = stringA[index_two].split(":")[1].toString().trim();
          continue;
        }
        if (stringA[index_two].includes("Item price:")) {
          transaction.price = stringA[index_two].split(":")[1].toString().trim();
          continue;
        }
      }

      transaction.date = message.getDate();

      regExp = new RegExp("(?<=Note from.*:\\r\\n.*\\r\\n)(.*\\r\\n)?(.*\\r\\n)?(.*\\r\\n)?(.*\\r\\n)?", 'gm');
      var buyerNote = (emailBody.match(regExp)) ? emailBody.match(regExp).toString().trim() : "";
      if(buyerNote.includes("The buyer did not leave a note")){
        transaction.noteFromBuyer = "";
      }else{
        transaction.noteFromBuyer = buyerNote;
      }

      regExp = new RegExp("(?<=Gift message\\r\\n)(.*\\r\\n)?(.*\\r\\n)?(.*\\r\\n)?(.*\\r\\n)?(.*\\r\\n)?(?=Delivery Address:)", 'gm');
      transaction.giftMessage = (emailBody.match(regExp)) ? emailBody.match(regExp).toString().trim() : "";

      regExp = new RegExp("(?<=" + emailKeywords.orderId + ").*\\w", 'g');
      transaction.orderId = (emailBody.match(regExp)) ? emailBody.match(regExp).toString().trim() : "";

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

      regExp = new RegExp("(?<=" + emailKeywords.shippingName + ")[^<]*", 'gu');
      transaction.shippingName = (emailBody.match(regExp)) ? emailBody.match(regExp).toString().trim() : "";

      regExp = new RegExp("(?<=" + emailKeywords.shippingAddress1 + ")[^<]*", 'g');
      transaction.shippingAddress1 = (emailBody.match(regExp)) ? emailBody.match(regExp).toString().trim() : "";

      regExp = new RegExp("(?<=" + emailKeywords.shippingAddress2 + ")[^<]*");
      transaction.shippingAddress2 = (emailBody.match(regExp)) ? emailBody.match(regExp).toString().trim() : "";

      regExp = new RegExp("(?<=" + emailKeywords.shippingZipcode + ")[^<]*", 'g');
      transaction.shippingZipcode = (emailBody.match(regExp)) ? emailBody.match(regExp).toString().trim() : "";

      regExp = new RegExp("(?<=" + emailKeywords.shippingCity + ")[^<]*", 'g');
      transaction.shippingCity = (emailBody.match(regExp)) ? emailBody.match(regExp).toString().trim() : "";

      regExp = new RegExp("(?<=" + emailKeywords.shippingState + ")[^<]*", 'g');
      transaction.shippingState = (emailBody.match(regExp)) ? emailBody.match(regExp).toString().trim() : "";

      regExp = new RegExp("(?<=" + emailKeywords.shippingCountry + ")[^<]*", 'g');
      var shippingCountry = (emailBody.match(regExp)) ? emailBody.match(regExp).toString().trim() : "";

      var result = iban_countries.findIndex(country => {
        return country[0].includes(shippingCountry);
      });
      if (result != -1) {
        transaction.shippingCountry = iban_countries[result][1];
      }

      regExp = new RegExp(emailKeywords.shippingService, 'gm');
      var paragraph = (emailBody.match(regExp)) ? emailBody.match(regExp).toString().trim() : "";
      const express = /Express/g;
      paragraph.search(express)
      if (paragraph.search(express) == -1) {
        transaction.shippingService = "Standard";
      }else{
        transaction.shippingService = "Express";
      }

      transactions.push(transaction)
    }
  }

  if (transactions[0] && transactions[0].orderId) {
    transactions.forEach(element => {
      var row = []
      for (var propName in element) {
        row.push(element[propName]);
      }
      emailDataArr.push(row);
    });
    for (let index = 0; index < emailDataArr.length; index++) {
      mainSheet.appendRow(emailDataArr[index]);
    }
  }
}


