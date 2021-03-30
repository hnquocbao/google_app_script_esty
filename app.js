// var ui = SpreadsheetApp.getUi();
var ibanGoogleSheetId = "1Iek1JLB8IEBbKbizvn2DbZmEsjT84zPBeojUNkh_Ing";
var mainGoogleSheetFile = "1bO6gsAO-c9JSiEsR5PlOOptTJZpTvBqPWELaH1QKnTs";
var designGoogleSheetFile = "1fodieDlNhhMuHvPVTVQBf6_aYocnwzXvkH55ykjV054";
var mainSheetName = "Trang tính1";
var orderIdColumn = "J2:J9999";
var gmailOrderTag = "orders";
var imageShareFolderId = "1NSGPV6DfOpEKiQgok5Sc3lXy_PTJ4LtZ";// bỏ Id của folder share ảnh vô đây nhé


var iban_countries = SpreadsheetApp.openById(ibanGoogleSheetId).getDataRange().getValues();
var mainFile = SpreadsheetApp.openById(mainGoogleSheetFile);
var mainSheet = mainFile.getSheetByName(mainSheetName);
var activeData = mainSheet.getRange(orderIdColumn).getValues();
var label = GmailApp.getUserLabelByName(gmailOrderTag);
var designSpreadsheet = SpreadsheetApp.openById(designGoogleSheetFile);
var imgShareFolder = DriveApp.getFolderById(imageShareFolderId);



function getOrdersEmails() {
  createHeadersEstyFull();
  createDesignSheet();
  var settingSheet = mainFile.getSheetByName("Setting");
  var fromDate;
  var toDate;
  var settingCondition;
  if (!settingSheet) {
    settingSheet = mainFile.insertSheet();
    settingSheet.setName("Setting");
  }
  if (settingSheet.getRange("A2")) {
    settingCondition = settingSheet.getRange("A2").getValue();
  }
  if (settingSheet.getRange("B2")) {
    fromDate = new Date(settingSheet.getRange("B2").getValue());
  }

  if (settingSheet.getRange("C2")) {
    toDate = new Date(settingSheet.getRange("C2").getValue());
  }


  var threads = label.getThreads();
  var count = 0;
  for (var i = threads.length - 1; i >= 0; i--) {
    var messages = threads[i].getMessages();
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      // var emailBody = message.getPlainBody();
      var emailBody = message.getBody();
      var keyWord = "Your order number is";
      var regExp = new RegExp("(?<=" + keyWord + ").*\\w", 'g');
      var resultObject;
      if (emailBody.match(regExp)) {
        var id = (emailBody.match(regExp)) ? emailBody.match(regExp).toString().trim() : "";
        if (id && id != "") {
          resultObject = activeData.findIndex(orderId => { return orderId == id; });
        }
      }
      if (resultObject && resultObject == -1) {
        if (settingCondition.toString().trim().startsWith("bigger") && message.getDate().valueOf() >= fromDate.valueOf()) {
          extractDetails(message);
          count++;
        }
        if (settingCondition.toString().trim().startsWith("smaller") && message.getDate().valueOf() <= fromDate.valueOf()) {
          extractDetails(message);
          count++;
        }
        // if (settingCondition.toString().trim().startsWith("between") && message.getDate().valueOf() >= fromDate.valueOf() && message.getDate().valueOf() <= toDate.valueOf()) {
        //   extractDetails(message);
        //   count++;
        // }
      }
    }
    threads[i].removeLabel(label);
  }
  if (count != 0) {
    Logger.log("Đã có " + count + " đơn mới được thêm vào !");
  } else {
    Logger.log("Chưa có đơn mới được thêm vào !");
  }
}

function extractDetails(message) {
  // var file = DriveApp.getFileById("1IqzE0jSNRb3svHzWJKYNWuePbhBoB4Ek");
  // var content = file.getBlob();
  // message = content.getDataAsString();

  var transactions = []
  var transactionsDesign = []
  var emailBody = message.getPlainBody();
  var regExp;

  var emailKeywords = {
    orderId: "Your order number is",
    productName: "Item:",
    style: "Style:",
    color: "Color:",
    size: "Size:",
    side: "Side:",
    quantity: "Quantity:",
    personalization: "Personalization:",
    shop: "Shop:",
    shippingName: "name'>",
    shippingAddress1: "first-line'>",
    shippingAddress2: "second-line'>",
    shippingZipcode: "zip'>",
    shippingCity: "city'>",
    shippingState: "state'>",
    shippingCountry: "country-name'>",
    shippingCost: "Shipping:",
    delivery: "Delivery:",
    transactionId: "(?<=Shop:.*\\r\\n\\r\\n.*?\\r\\n\\r\\n)(.|\\r\\n)*(?=\\r\\n-*?\\r\\nItem total:)",
    noteFromBuyer: "(?<=Note from.*:\\n.*\\n)(.|\\r\\n)*(?=\\r\\n.*\\r\\nOrder Details)",
    giftMessage: "(?<=Gift message\\n)(.|\\r\\n)*(?=\\r\\nDelivery Address)",
    shippingService: "(?<=Delivery:)(.|\\r\\n)*(?=\\r\\nOrder Total:)",
    price: "Item price:",
    img: "\\bhttps?:\/\/i.etsystatic.com\/\\d+[^)''" + '"' + "\\s]+[^" + '"' + "]*"
  }


  regExp = new RegExp(emailKeywords.img, 'g');
  var emailBodyFull = message.getBody();
  var imgs_temp = emailBodyFull.match(regExp);
  var imgs = [];
  var img_urls = [];
  if (imgs_temp && imgs_temp.length > 0) {
    for (let index = 0; index < imgs_temp.length; index++) {
      var url = imgs_temp[index].split("?");
      var ulr_0 = url[0].replace("75x75", "300x300").replace("=", "");
      var image = "=image(" + '"' + ulr_0 + '"' + ")";
      imgs.push(image);
      img_urls.push(ulr_0);
    }
  }

  regExp = new RegExp(emailKeywords.transactionId, 'g');
  if (emailBody.match(regExp)) {
    var regExpSlit = new RegExp("(?<=Item price.*$)", 'gm');
    var transactionIds = emailBody.match(regExp)[0].split(regExpSlit);
    for (let index_one = 0; index_one < transactionIds.length; index_one++) {
      var regExpSearch = new RegExp("^\s*$", 'gm');
      var item = transactionIds[index_one].trim();
      if (item.length) {
        var transaction = {
          id: "",
          img_url: "",
          img: "",
          date: "",
          noteFromBuyer: "",
          giftMessage: "",
          personalization: "",
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
          option: "",
          color: "",
          size: "",
          side: "",
          quantity: "",
          designLinkFront: "",
          designLinkBack: "",
          shippingService: "",
          processingTime: "",
          shippingCost: "",
          price: "",
          discountCode: "",
          subtotal: ""
        }
        var transactionDesign = {
          id: "",
          img: "",
          date: "",
          noteFromBuyer: "",
          giftMessage: "",
          personalization: "",
          sku: "",
          shop: "",
          orderId: "",
          productName: "",
          option: "",
          color: "",
          size: "",
          side: "",
          quantity: "",
        }

        regExp = new RegExp("(?<=Transaction ID:)([^:]*)(?=\\r\\n)", 'gm');
        transaction.id = (item.match(regExp)) ? item.match(regExp).toString().trim() : "";
        transactionDesign.id = transaction.id;

        regExp = new RegExp("(?<=Item:)([^:]*)(?=\\r\\n)", 'gm');
        transaction.productName = (item.match(regExp)) ? item.match(regExp).toString().trim() : "";
        transactionDesign.productName = transaction.productName;

        regExp = new RegExp("(?<=Face mask size:|Face Mask Size:|face mask size:)([^:]*)(?=\\r\\n)", 'gm');
        var faceMaskSize = (item.match(regExp)) ? item.match(regExp).toString().trim() : "";
        if (faceMaskSize) {
          transaction.size = faceMaskSize;
          transactionDesign.size = faceMaskSize;
        }

        regExp = new RegExp("(?<=Size:|size:|SIZE:)([^:]*)(?=\\r\\n)", 'gm');
        var size = (item.match(regExp)) ? item.match(regExp).toString().trim() : "";
        if (size) {
          transaction.size = size;
          transactionDesign.size = size;
        }

        regExp = new RegExp("(?<=Capacity:|capacity:|CAPACITY:)([^:]*)(?=\\r\\n)", 'gm');
        var capacity = (item.match(regExp)) ? item.match(regExp).toString().trim() : "";
        if (capacity) {
          transaction.size = capacity;
          transactionDesign.size = capacity;
        }

        regExp = new RegExp("(?<=Volume:|volume:|VOLUME:)([^:]*)(?=\\r\\n)", 'gm');
        var volume = (item.match(regExp)) ? item.match(regExp).toString().trim() : "";
        if (volume) {
          transaction.size = volume;
          transactionDesign.size = volume;
        }

        regExp = new RegExp("(?<=Option:|option:|OPTION:)([^:]*)(?=\\r\\n)", 'gm');
        transaction.option = (item.match(regExp)) ? item.match(regExp).toString().trim() : "";
        transactionDesign.option = transaction.option;

        regExp = new RegExp("(?<=Style:|style:|STYLE:)([^:]*)(?=\\r\\n)", 'gm');
        var style = (item.match(regExp)) ? item.match(regExp).toString().trim() : "";
        if (style) {
          transaction.option = style;
          transactionDesign.option = style;
        }

        regExp = new RegExp("(?<=Pack:|pack:|PACK:)([^:]*)(?=\\r\\n)", 'gm');
        var pack = (item.match(regExp)) ? item.match(regExp).toString().trim() : "";
        if (pack) {
          transaction.option = pack;
          transactionDesign.option = pack;
        }

        regExp = new RegExp("(?<=Shape:|shape:|SHAPE:)([^:]*)(?=\\r\\n)", 'gm');
        var shape = (item.match(regExp)) ? item.match(regExp).toString().trim() : "";
        if (shape) {
          transaction.option = shape;
          transactionDesign.option = shape;
        }

        regExp = new RegExp("(?<=Includes:|includes:|INCLUDES:)([^:]*)(?=\\r\\n)", 'gm');
        var includes = (item.match(regExp)) ? item.match(regExp).toString().trim() : "";
        if (includes) {
          transaction.option = includes;
          transactionDesign.option = includes;
        }

        regExp = new RegExp("(?<=Design:|design:|DESIGN:)([^:]*)(?=\\r\\n)", 'gm');
        var design = (item.match(regExp)) ? item.match(regExp).toString().trim() : "";
        if (design) {
          transaction.option = design;
          transactionDesign.option = design;
        }

        regExp = new RegExp("(?<=Personalization:|Personalisation:|personalisation:)(.|\\r\\n)*(?=\\r\\nQuantity)", 'gm');
        transaction.personalization = (item.match(regExp)) ? item.match(regExp).toString().trim() : "";
        transactionDesign.personalization = transaction.personalization;

        regExp = new RegExp("(?<=Color:|Colour:|color:|COLOR:|COLOUR:)([^:]*)(?=\\r\\n)", 'gm');
        transaction.color = (item.match(regExp)) ? item.match(regExp).toString().trim() : "";
        transactionDesign.color = transaction.color;

        regExp = new RegExp("(?<=Side:|side:)([^:]*)(?=\\r\\n)", 'gm');
        transaction.side = (item.match(regExp)) ? item.match(regExp).toString().trim() : "";
        transactionDesign.side = transaction.side;

        regExp = new RegExp("(?<=Quantity:|quantity:)([^:]*)(?=\\r\\n)", 'gm');
        transaction.quantity = (item.match(regExp)) ? item.match(regExp).toString().trim() : "";
        transactionDesign.quantity = transaction.quantity;


        var files = imgShareFolder.searchFiles('title contains "' + transaction.id +"-"+transaction.orderId + '"');
        while (files.hasNext()) {
          var designFile = files.next();
          transaction.designLinkFront = designFile.getUrl();
        }


        regExp = new RegExp("(?<=Item price:|Price:).*", 'gm');
        transaction.price = (item.match(regExp)) ? item.match(regExp).toString().trim() : "";


        regExp = new RegExp("(?<=Subtotal:|subtotal:)([^:]*)(?=\\r\\n)", 'gm');
        transaction.subtotal = (emailBody.match(regExp)) ? emailBody.match(regExp)[0].toString().trim() : "";

        regExp = new RegExp("(?<=Applied discounts)([^:]*)(?=\\r\\n)", 'gm');
        transaction.discountCode = (emailBody.match(regExp)) ? emailBody.match(regExp).toString().trim().replace("-", "") : "";

        transaction.date = message.getDate();
        transactionDesign.date = transaction.date;

        regExp = new RegExp("(?<=Note from.*:\\r\\n.*\\r\\n)(.*\\r\\n)?(.*\\r\\n)?(.*\\r\\n)?(.*\\r\\n)?", 'gm');
        var buyerNote = (emailBody.match(regExp)) ? emailBody.match(regExp).toString().trim() : "";
        if (buyerNote.includes("The buyer did not leave a note")) {
          transaction.noteFromBuyer = "";
        } else {
          transaction.noteFromBuyer = buyerNote;
        }
        transactionDesign.noteFromBuyer = transaction.noteFromBuyer;

        regExp = new RegExp("(?<=Gift message\\r\\n)(.|\\r\\n)*(?=\\r\\n\\w*\\s*Address)", 'gm');
        transaction.giftMessage = (emailBody.match(regExp)) ? emailBody.match(regExp).toString().trim() : "";
        transactionDesign.giftMessage = transaction.giftMessage;

        regExp = new RegExp("(?<=" + emailKeywords.orderId + ").*\\w", 'g');
        transaction.orderId = (emailBody.match(regExp)) ? emailBody.match(regExp)[0].toString().trim() : "";
        transactionDesign.orderId = transaction.orderId;


        transactionDesign.sku = transaction.sku;

        regExp = new RegExp("(?<=" + emailKeywords.shop + ").*", 'g');
        transaction.shop = (emailBody.match(regExp)) ? emailBody.match(regExp)[0].toString().trim() : "";
        transactionDesign.shop = transaction.shop;

        var regExpShippingCost = new RegExp("(?<=Shipping:)([^:].*)(?=\\s\\S\\S\\r\\n)", 'gm');
        if (emailBody.match(regExpShippingCost)) {
          transaction.shippingCost = emailBody.match(regExpShippingCost)[0].toString().trim();
        }
        var regExpDelivery = new RegExp("(?<=" + emailKeywords.delivery + ").*\\w", 'g');
        if (emailBody.match(regExpDelivery)) {
          transaction.shippingCost = emailBody.match(regExpDelivery)[0].toString().trim();
        }

        regExp = new RegExp("(?<=" + emailKeywords.shippingName + ")[^<]*", 'gu');
        transaction.shippingName = (emailBody.match(regExp)) ? emailBody.match(regExp)[0].toString().trim() : "";

        regExp = new RegExp("(?<=" + emailKeywords.shippingAddress1 + ")[^<]*", 'g');
        transaction.shippingAddress1 = (emailBody.match(regExp)) ? emailBody.match(regExp)[0].toString().trim().replace("=", "") : "";

        regExp = new RegExp("(?<=" + emailKeywords.shippingAddress2 + ")[^<]*");
        transaction.shippingAddress2 = (emailBody.match(regExp)) ? emailBody.match(regExp)[0].toString().trim().replace("=", "") : "";

        regExp = new RegExp("(?<=" + emailKeywords.shippingZipcode + ")[^<]*", 'g');
        transaction.shippingZipcode = (emailBody.match(regExp)) ? "'" + emailBody.match(regExp)[0].toString().trim() : "";

        regExp = new RegExp("(?<=" + emailKeywords.shippingCity + ")[^<]*", 'g');
        transaction.shippingCity = (emailBody.match(regExp)) ? emailBody.match(regExp)[0].toString().trim().replace("=", "") : "";

        regExp = new RegExp("(?<=" + emailKeywords.shippingState + ")[^<]*", 'g');
        transaction.shippingState = (emailBody.match(regExp)) ? emailBody.match(regExp)[0].toString().trim().replace("=", "") : "";

        regExp = new RegExp("(?<=" + emailKeywords.shippingCountry + ")[^<]*", 'g');
        var shippingCountry = (emailBody.match(regExp)) ? emailBody.match(regExp)[0].toString().trim().replace("=", "") : "";

        var result = iban_countries.findIndex(country => {
          return country[0].includes(shippingCountry);
        });

        if (result !== -1) {
          transaction.shippingCountry = iban_countries[result][1];
        } else {
          transaction.shippingCountry = shippingCountry;
        }

        regExp = new RegExp(emailKeywords.shippingService, 'gm');
        var paragraph = (emailBody.match(regExp)) ? emailBody.match(regExp).toString().trim() : "";
        const express = /Express/g;
        paragraph.search(express)
        if (paragraph.search(express) == -1) {
          transaction.shippingService = "Standard";
        } else {
          transaction.shippingService = "Express";
        }


        regExp = new RegExp("(?<=Processing time:)[^<]*", 'gm');
        var rawEmail = message.getRawContent();
        transaction.processingTime = (rawEmail.match(regExp)) ? rawEmail.match(regExp).toString().trim().replace("&ndash;", "-").replace("=", "") : "";

        if (imgs && imgs.length > 0) {
          transaction.img = imgs[index_one];
          transaction.img_url = img_urls[index_one];
          transactionDesign.img = transaction.img;
        }
        transactions.push(transaction);
        transactionsDesign.push(transactionDesign);
      }
    }
  }

  var emailDataArr = [];
  var emailDataArrDesign = [];

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

  if (transactionsDesign[0] && transactionsDesign[0].orderId) {
    transactionsDesign.forEach(element => {
      var row = []
      for (var propName in element) {
        row.push(element[propName]);
      }
      emailDataArrDesign.push(row);
    });
    for (let index = 0; index < emailDataArrDesign.length; index++) {
      designSpreadsheet.appendRow(emailDataArrDesign[index]);
    }
  }
}

function createHeadersEstyFull() {
  // Set the values we want for headers
  var values = ["id", "img_url", "img", "date", "noteFromBuyer", "giftMessage", "personalization", "sku",
    "shop", "orderId", "shippingName", "shippingAddress1", "shippingAddress2", "shippingCity", "shippingState",
    "shippingZipcode", "shippingCountry", "shippingPhone", "shippingEmail", "productName", "option", "color",
    "size", "side", "quantity", "designLinkFront", "designLinkBack", "shippingService",
    "processingTime", "shippingCost", "price", "discountCode", "subtotal"];

  // Set the range of cells
  var range = mainSheet.getRange("A1:AE1");
  var header_content = range.getValues();
  var hearder_index_0 = header_content[0][0];
  if (!hearder_index_0) {
    mainSheet.appendRow(values);
  }

  //Hide columns
  // mainSheet.hideColumns(1, 2);

  // Freezes the first row
  mainSheet.setFrozenRows(1);
}

function createHeadersEstyDesign() {
  // Set the values we want for headers
  var values = ["id", "img", "date", "noteFromBuyer", "giftMessage", "personalization", "sku", "shop", "orderId",
    "productName", "option", "color", "size", "side", "quantity"];

  // Set the range of cells
  var range = designSpreadsheet.getRange("A1:AE1");
  var header_content = range.getValues();
  var hearder_index_0 = header_content[0][0];
  if (!hearder_index_0) {
    designSpreadsheet.appendRow(values);
  }

  //Hide columns
  // mainSheet.hideColumns(1, 2);

  // Freezes the first row
  designSpreadsheet.setFrozenRows(1);
}

// Tạo sheet thiết kế
function createDesignSheet() {
  if (!designSpreadsheet) {
    designSpreadsheet = SpreadsheetApp.create("Esty Design")
  }
  createHeadersEstyDesign();
}



