// Create UI element
function onOpen() { 
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Script")
    .addItem("Run Now","GetEmail")
    .addItem("Clean Inbox", "cleanUpInbox")
    .addItem("Create Project Triggers","MakeTrigger")
    .addItem("Clear Project Triggers","ClearTriggers")
    .addToUi();
}


// Makes a trigger to run program every hour
function MakeTrigger() {
  ScriptApp.newTrigger('GetEmail')
      .timeBased()
      .everyHours(1)
      .create();
}


function ClearTriggers() {
  Logger.log('Current project has ' + ScriptApp.getProjectTriggers().length + ' triggers.');
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}


function GetEmail() {
  // Get all threads in inbox
  var threads = GmailApp.getInboxThreads();
  var msgArr = getMessages(threads);
  var orderArr = getOrders(msgArr);
  var orderElements = getElements(orderArr);
  var orderInfo = parseText(orderElements);
  addRow(orderInfo);
  cleanUpInbox();
}


function getMessages(threads) { // Gets the value of each thread and puts it into an array of objects
  var msgArr = []
  var count = 0;
  for (var i = 0; i < threads.length; i++) {
    var msg = threads[i].getMessages();
    var id = msg[0].getId();
    var sub = threads[i].getFirstMessageSubject();
    var unread = threads[i].isUnread();
    var msgObj = {
      Message: msg,
      ID: id,
      Subject: sub,
      Unread: unread,
    };
    if (unread == true) {
      msgArr[count] = msgObj;
      count++;
    }
  }
  return  msgArr;
}


function getOrders(msgArr) { // Takes the array of message objects and processes them into an array of orders
  var count = 0;
  var orderArr = [];
  for (var i = 0; i < msgArr.length; i++) {
    Logger.log(GmailApp.getMessageById(msgArr[i].ID).getFrom());
    if (GmailApp.getMessageById(msgArr[i].ID).getFrom() == 'confirmation@quill.com') {
      var orderObject = {
        Vendor: "Quill",
        Date: GmailApp.getMessageById(msgArr[i].ID).getDate(),
        ID: GmailApp.getMessageById(msgArr[i].ID).getId(),
        Subject: GmailApp.getMessageById(msgArr[i].ID).getSubject(),
        Body: GmailApp.getMessageById(msgArr[i].ID).getPlainBody(),
      };
      GmailApp.getThreadById(msgArr[i].ID).markRead();
      GmailApp.getThreadById(msgArr[i].ID).addLabel(GmailApp.getUserLabelByName("Orders"));
      GmailApp.getMessageById(msgArr[i].ID).star();
      orderArr[count] = orderObject;
      count++;
    }
    else if (GmailApp.getMessageById(msgArr[i].ID).getFrom() == '"Quill" <orders@quill.com>') {
      GmailApp.getThreadById(msgArr[i].ID).markRead();
      GmailApp.getThreadById(msgArr[i].ID).addLabel(GmailApp.getUserLabelByName("Orders"));
    }
    else if (GmailApp.getMessageById(msgArr[i].ID).getFrom() == '"Amazon.com" <auto-confirm@amazon.com>') {
      var orderObject = {
        Vendor: "Amazon",
        Date: GmailApp.getMessageById(msgArr[i].ID).getDate(),
        ID: GmailApp.getMessageById(msgArr[i].ID).getId(),
        Subject: GmailApp.getMessageById(msgArr[i].ID).getSubject(),
        Body: GmailApp.getMessageById(msgArr[i].ID).getPlainBody(),
      };
      GmailApp.getThreadById(msgArr[i].ID).markRead();
      GmailApp.getThreadById(msgArr[i].ID).addLabel(GmailApp.getUserLabelByName("Orders"));
      GmailApp.getMessageById(msgArr[i].ID).star();
      orderArr[count] = orderObject;
      count++;
    }
    else if (GmailApp.getMessageById(msgArr[i].ID).getFrom() == 'Amazon Business <no-reply@amazon.com>') {
      GmailApp.getThreadById(msgArr[i].ID).markRead();
      GmailApp.getThreadById(msgArr[i].ID).addLabel(GmailApp.getUserLabelByName("Orders"));
    }
    else if (GmailApp.getMessageById(msgArr[i].ID).getFrom() == '"Amazon.com" <shipment-tracking@amazon.com>') {
      GmailApp.getThreadById(msgArr[i].ID).markRead();
      GmailApp.getThreadById(msgArr[i].ID).addLabel(GmailApp.getUserLabelByName("Orders"));
    }
    else if (GmailApp.getMessageById(msgArr[i].ID).getFrom() == '"Amazon.com" <ship-confirm@amazon.com>') {
      GmailApp.getThreadById(msgArr[i].ID).markRead();
      GmailApp.getThreadById(msgArr[i].ID).addLabel(GmailApp.getUserLabelByName("Orders"));
    }
    else if (GmailApp.getMessageById(msgArr[i].ID).getFrom() == '"DocuCopies.com" <info@docucopies.com>') {
      GmailApp.getThreadById(msgArr[i].ID).markRead();
      GmailApp.getThreadById(msgArr[i].ID).addLabel(GmailApp.getUserLabelByName("Orders"));
      var subj = GmailApp.getMessageById(msgArr[i].ID).getSubject();
      if (subj.search("Has Shipped") == -1 && subj.search("[DocuCopies.com] Order #") != -1) {
        GmailApp.getMessageById(msgArr[i].ID).star();
      }
    }
    else if (GmailApp.getMessageById(msgArr[i].ID).getFrom() == 'DocuCopies.com <statements@docucopies.com>>') {
      GmailApp.getThreadById(msgArr[i].ID).markRead();
      GmailApp.getThreadById(msgArr[i].ID).addLabel(GmailApp.getUserLabelByName("Orders"));
    }
    else if (GmailApp.getMessageById(msgArr[i].ID).getFrom() == 'UPS Quantum View <pkginfo@ups.com>') {
      GmailApp.getThreadById(msgArr[i].ID).markRead();
      GmailApp.getThreadById(msgArr[i].ID).addLabel(GmailApp.getUserLabelByName("Orders"));
    }
    else if (GmailApp.getMessageById(msgArr[i].ID).getFrom() == 'XCOMWEB@xerox.com') {
      GmailApp.getThreadById(msgArr[i].ID).markRead();
      GmailApp.getThreadById(msgArr[i].ID).addLabel(GmailApp.getUserLabelByName("Orders"));
      var subj = GmailApp.getMessageById(msgArr[i].ID).getSubject();
      if (subj.search("Xerox Metered Supply Auto Replenishment Order Confirmation") != -1) {
        GmailApp.getMessageById(msgArr[i].ID).star();
      }
      else if (subj.search("Xerox Metered Supply Order Confirmation") != -1) {
        GmailApp.getMessageById(msgArr[i].ID).star();
      }
    }
  } 
  return orderArr;
}


function getElements(orderArr) {
  var orderElements = [];
  for (var i = 0; i < orderArr.length; i++) {
    var date = orderArr[i].Date;
    var vendor = orderArr[i].Vendor;
    var text = orderArr[i].Body;
    if (vendor == "Quill") { // If Quill Invoice
      Logger.log("Quill");
      var start = 0;
      var end = (text.search("Order Number: ") + 14);
      text = text.slice(end); // Clear [Start] to [Order Number]
      end = text.search("Order Date: ");
      var orderId = text.slice(start, end); // Get [orderId]
      end = (text.search("Expected Delivery: ") + 19);
      text = text.slice(end); // Clear [Order Date] to [Expected Delivery]
      end = text.search("by ");
      var expDelivery = text.slice(start, end); // Get [expDelivery]
      end = text.search("------------------------------") - 1;
      text = text.slice(end); // Clear [expDelivery] to [---]
      var count = 0;
      var items = [];
      while (text.search("------------------------------") != -1) {
        end = (text.search("------------------------------") + 30);
        text = text.slice(end); // Clear [---] before [items]
        if (text.search("------------------------------") < text.search("Subtotal: ") && text.search("------------------------------") != -1) {
          end = text.search("------------------------------") - 1;
          items[count] = text.slice(start, end); // Get [items]
          count++;
        }
        else {
          end = text.search("Subtotal: ") - 1;
          items[count] = text.slice(start, end); // Get [items]
          count++;
          if (text.search("------------------------------") == -1) {
            break;
          }
          else {
            end = text.search("------------------------------") - 1;
            text = text.slice(end); // Clear [Subtotal] to [---]
          }
        }
      }
    }
    
    else if (vendor == "Amazon") { // If Amazon Invoice
      Logger.log("Amazon");
      var items = [];
      var count = 0;
      var start = 0;
      var end = (text.search("Order #") + 7);
      text = text.slice(end); // Clear [Start] to [Order Number]
      end = text.search("www.amazon.com/ref=TE_tex_h");
      
      if (end > (text.search("Order #") + 7)) { // Alt path for weirdly formatted orders
        Logger.log("Weird Format Variant");
        end = (text.search("Order Details") + 13);
        text = text.slice(end);
        end = text.search("To learn more about ordering, go to Ordering from ");
        text = text.slice(start, end);
        while (text.indexOf("$") != -1) {
          if (text.indexOf("$") > (text.search("Order #") + 7) && text.search("Order #") != -1) {
            end = (text.search("Order #") + 7);
            text = text.slice(end); // Clear [Start] to [Order Number]
            end = text.search("Placed on ");
            if (count == 0) {
              var orderId = text.slice(start, end).trim(); // Get [orderId]
            }
            else {
              orderId = orderId.concat(" || ", text.slice(start, end).trim()); // Append additional ID to [orderId]
            }
            end = (text.search(" delivery date is:") + 18);
            text = text.slice(end); // Clear [Order Date] to [Expected Delivery]
            end = text.search("Your shipping speed:");
            if (count == 0) {
              var expDelivery = text.slice(start, end).trim(); // Get [expDelivery]
            }
            else {
              expDelivery = expDelivery.concat(" || ", text.slice(start, end).trim()); // Append additional delivery date to [expDelivery]
            }
            end = text.search("United States") + 13;
            text = text.slice(end); // Clear [expDelivery] to [Items]
            end = (text.indexOf("$") - 1);
            items[count] = text.slice(start, end); // Get [Items]
            count++;
            end = (text.indexOf("$") + 1);
            text = text.slice(end); // Remove following "$"
          }
          else {
            end = (text.indexOf("$") - 1);
            items[count] = text.slice(start, end); // Get [Items]
            count++;
            end = (text.indexOf("$") + 1);
            text = text.slice(end); // Remove following "$"
          }
          var con = (text.search("Condition: New") + 14);
          var sld = (text.search("Sold by: Amazon.com Services, Inc") + 34);
          if (con < sld &&  con < text.indexOf("$")) {
            end = con;
            text = text.slice(end); // Clear to next [Items] following "Condition:"
            if (sld < text.indexOf("$")) {
              end = (text.search("Sold by: Amazon.com Services, Inc") + 34);
              text = text.slice(end); // Clear to next [Items] following "Sold by:" if sold by Amazon
            }
          }
          else if (sld < con && sld < text.indexOf("$")) {
            end = sld;
            text = text.slice(end); // Clear to next [Items] following "Sold by:"
            if (con < text.indexOf("$")) {
              end = (text.search("Condition: New") + 14);
              text = text.slice(end); // Clear to next [Items] following "Condition:" if sold by Amazon
            }
          }
          else {
            break;
          }
          if (text.indexOf("_______________________________________________________________________________________") < text.indexOf("$")) {
            end = (text.search("=======================================================================================") + 87);
            text = text.slice(end);
          }
        }
      }
      
      else {
        var orderId = text.slice(start, end); // Get [orderId]
        end = (text.search(" delivery date is:") + 18);
        text = text.slice(end); // Clear [Order Date] to [Expected Delivery]
        end = text.search("Your shipping speed:");
        var expDelivery = text.slice(start, end).trim(); // Get [expDelivery]
        end = text.search("_______________________________________________________________________________________");
        text = text.slice(start, end); // Get rid of extraneous text at the end of the email
        
        if (text.search("Shipment 2 of ") != -1) { // Multiple shipment orders
          Logger.log("Multiple Shipments");
          end = text.search("United States") + 13;
          text = text.slice(end); // Clear [expDelivery] to [Items]
          while (text.indexOf("$") != - 1) {
            if (text.search("Shipment ") < (text.indexOf("$") - 1) && text.search("Shipment ") != -1) {
              end = (text.search(" delivery date is:") + 18);
              text = text.slice(end); // Clear to [expDelivery]
              end = (text.search("Your shipping speed:") - 1);
              expDelivery = expDelivery.concat(" || ", text.slice(start, end).trim()); // Append additional delivery date to [expDelivery]
              end = text.search("United States") + 13;
              text = text.slice(end); // Clear [expDelivery] to [Items]
              end = (text.indexOf("$") - 1);
              items[count] = text.slice(start, end); // Get [Items]
              count++;
              end = (text.indexOf("$") + 1);
              text = text.slice(end); // Remove following "$"
            }
            else {
              end = (text.indexOf("$") - 1);
              items[count] = text.slice(start, end); // Get [Items]
              count++;
              end = (text.indexOf("$") + 1);
              text = text.slice(end); // Remove following "$"
            }
            var con = (text.search("Condition: New") + 14);
            var sld = (text.search("Sold by: Amazon.com Services, Inc") + 34);
            if (con < sld &&  con < text.indexOf("$")) {
              end = con;
              text = text.slice(end); // Clear to next [Items] following "Condition:"
            }
            else if (sld < con && sld < text.indexOf("$")) {
              end = sld;
              text = text.slice(end); // Clear to next [Items] following "Sold by:"
              if (con < text.indexOf("$")) {
                end = (text.search("Condition: New") + 14);
                text = text.slice(end); // Clear to next [Items] following "Condition:" if sold by Amazon
              }
            }
            else {
              break;
            }
          }
        }
        else { // Single shipment orders
          Logger.log("Single Shipment");
          end = text.search("Placed on ") + 10;
          text = text.slice(end); // Clear [expDelivery] to [Items1]
          end = text.search(", 20") + 6;
          text = text.slice(end); // Clear [expDelivery] to [Items2]
          while (text.indexOf("$") != - 1) {
            end = (text.indexOf("$") - 1);
            items[count] = text.slice(start, end);
            count++;
            end = (text.indexOf("$") + 1);
            text = text.slice(end); // Remove following "$"
            var con = (text.indexOf("Condition: ") + 14);
            Logger.log("Condition: " + con);
            var sld = (text.indexOf("Sold by: Amazon.com Services, Inc") + 34);
            Logger.log("Sold by: " + sld);
            if (con < sld &&  con < text.indexOf("$") && con != 13) {
              end = con;
              text = text.slice(end); // Clear to next [Items] following "Condition:"
            }
            else if (sld < con && sld < text.indexOf("$") && sld != 33) {
              end = sld;
              text = text.slice(end); // Clear to next [Items] following "Sold by:"
              if (con < text.indexOf("$")) {
                end = (text.search("Condition: New") + 14);
                text = text.slice(end); // Clear to next [Items] following "Condition:" if sold by Amazon
              }
            }
            else if (sld < con && sld < text.indexOf("$") && sld == 33) {
              end = con;
              text = text.slice(end); // Clear to next [Items] following "Condition:"
            }
            else if (sld < text.indexOf("$") && sld != 33 && con == 13) {
              end = sld;
              text = text.slice(end); // Clear to next [Items] following "Sold by:"
              if (con < text.indexOf("$")) {
                end = (text.search("Condition: New") + 14);
                text = text.slice(end); // Clear to next [Items] following "Condition:" if sold by Amazon
              }
            }
            else {
              break;
            }
          }
        }
      }
    }
    
    var element = {
      Date: date,
      Vendor: vendor,
      ID: orderId,
      Delivery: expDelivery,
      Items: items,
    };
    orderElements[i] = element;
  }
  return orderElements;
}




function parseText(orderElements) {
  var orderInfo = [];
  var splitItems = [];
  var items = [];
  for(var i = 0; i < orderElements.length; i++) {
    var date = orderElements[i].Date;
    var vendor = orderElements[i].Vendor;
    var orderId = orderElements[i].ID.trim();
    var delivery = orderElements[i].Delivery.trim();
    if (vendor == "Quill") { // If Quill Invoice
      for (var j = 0; j < orderElements[i].Items.length; j++) {
        splitItems = orderElements[i].Items[j].split(" "); // Split items into indiviudal words
        var itemNumber = splitItems[0].trim(); // Get [itemNumber]
        var itemName = splitItems[1].trim(); // Get first word of [itemName]
        for (var k = 2; k < splitItems.length; k++) {
          var temp = splitItems[k];
          var rge = /\$/;
          var result = rge.test(temp);
          if (result == true) {
            break;
          }
        }
        var pos = k; // Find where the prices start to get relative position of pieces
        var q1 = (pos - 2);
        var q2 = (pos - 1);
        var itemQuantity = splitItems[q1].concat(" ", splitItems[q2].trim()); // Get [itemQuantity]
        for (var n = 2; n < q1; n++) { // Get the remaining words for [itemName]
          var sf = splitItems[n].trim();
          itemName = itemName + " " + sf;
        }
        var itemInst = { // Create array to hold each items info
          ItemID: itemNumber,
          ItemName: itemName,
          Quantity: itemQuantity,
        };
      items[j] = itemInst;
      }
      var order = { // Create array to hold order info
        Date: date,
        Vendor: vendor,
        ID: orderId,
        Delivery: delivery,
        Items: items,
      };
    }
    
    else if (vendor == "Amazon") { // If Amazon Invoice
      for (var j = 0; j < orderElements[i].Items.length; j++) {
        splitItems = orderElements[i].Items[j].trim();
        splitItems = splitItems.split(" "); // Split items into indiviudal words
        var itemNumber = " --- "
        if (splitItems[1] == "x") {
          var itemQuantity = splitItems[0].trim(); // Get [itemQuantity]
          var itemName = splitItems[2].trim(); // Get first word of [itemName]
          for (var k = 3; k < splitItems.length; k++) {
            itemName = itemName + " " + splitItems[k].trim();
          }
        }
        else {
          var itemQuantity = "1";
          var itemName = splitItems[0].trim(); // Get first word of [itemName]
          for (var k = 1; k < splitItems.length; k++) {
            itemName = itemName + " " + splitItems[k].trim(); // Add the remaining words to [itemName]
          }
        }
        var itemInst = { // Create array to hold each items info
          ItemID: itemNumber,
          ItemName: itemName,
          Quantity: itemQuantity,
        };
        items[j] = itemInst;
      }
      var order = { // Create array to hold order info
        Date: date,
        Vendor: vendor,
        ID: orderId,
        Delivery: delivery,
        Items: items,
      };
    }
    orderInfo[i] = order; // Add this array to an array of orders
    items = []; // Clears out items after each iteration
  }
  return orderInfo;
}


function addRow(orderInfo) {
  var sheet = SpreadsheetApp.getActiveSheet();
  for (var i = 0; i < orderInfo.length; i++) { 
    var date = orderInfo[i].Date.toLocaleDateString(); // Get Order Date
    var vendor = orderInfo[i].Vendor; // Get Order Vendor
    var orderNum = orderInfo[i].ID; // Get Order ID
    var expDelivery = orderInfo[i].Delivery;
    for (var j = 0; j < orderInfo[i].Items.length; j++) {
      var itemNum = orderInfo[i].Items[j].ItemID; // Get Item Number if listed
      var itemDescript = orderInfo[i].Items[j].ItemName; // Get Item Name
      var quantity = orderInfo[i].Items[j].Quantity; // Get Item Quantity if listed
      var rowVals = [
        [date, vendor, orderNum, expDelivery, itemNum, itemDescript, quantity]
      ];
      var nextRow = (sheet.getLastRow() + 1);  // Get next empty row
      var range = sheet.getRange(nextRow, 1, 1, 7); // Get the range to be filled
      range.setValues(rowVals); // Fill range with data
      if (j == (orderInfo[i].Items.length - 1)) {
        var borderRange = sheet.getRange(nextRow, 1, 1, 9);
        borderRange.setBorder(null, null, true, null, null, null);
      }
    }
  }
}


function cleanUpInbox() {
  var threads = GmailApp.getInboxThreads();
  var msgArr = getMessages(threads);
  for (var i = 0; i < msgArr.length; i++) {
    if (GmailApp.getMessageById(msgArr[i].ID).getFrom() == '<veilegv@gmail.com>') {
      GmailApp.getThreadById(msgArr[i].ID).markRead();
      GmailApp.getThreadById(msgArr[i].ID).addLabel(GmailApp.getUserLabelByName("VM & Fax"));
    }
    else if (GmailApp.getMessageById(msgArr[i].ID).getFrom() == 'email@e.quill.com') {
      GmailApp.getThreadById(msgArr[i].ID).markRead();
      GmailApp.getThreadById(msgArr[i].ID).moveToTrash();
    }
    else if (GmailApp.getMessageById(msgArr[i].ID).getFrom() == 'Quill <reviews@quill.com>') {
      GmailApp.getThreadById(msgArr[i].ID).markRead();
      GmailApp.getThreadById(msgArr[i].ID).moveToTrash();
    }
    else if (GmailApp.getMessageById(msgArr[i].ID).getFrom() == 'NAME <NAME@townandstyle.com>') {
      GmailApp.getThreadById(msgArr[i].ID).markRead();
      GmailApp.getThreadById(msgArr[i].ID).moveToTrash();
    }
    else if (GmailApp.getMessageById(msgArr[i].ID).getFrom() == 'DP Community <noreply@softerware.com>') {
      GmailApp.getThreadById(msgArr[i].ID).markRead();
      GmailApp.getThreadById(msgArr[i].ID).moveToTrash();
    }
    else if (GmailApp.getMessageById(msgArr[i].ID).getFrom() == 'Adobe Creative Cloud <mail@rt.adobesystems.com>') {
      GmailApp.getThreadById(msgArr[i].ID).markRead();
      GmailApp.getThreadById(msgArr[i].ID).moveToTrash();
    }
    else if (GmailApp.getMessageById(msgArr[i].ID).getFrom() == 'HumanResources <HumanResources@sccmo.org>') {
      GmailApp.getThreadById(msgArr[i].ID).markRead();
      GmailApp.getThreadById(msgArr[i].ID).addLabel(GmailApp.getUserLabelByName("Archived"));
      GmailApp.getMessageById(msgArr[i].ID).forward("NAME@DOMAIN.org");
    }
    else if (GmailApp.getMessageById(msgArr[i].ID).getSubject() == "Confirm Schedule") {
      GmailApp.getThreadById(msgArr[i].ID).markRead();
      GmailApp.getThreadById(msgArr[i].ID).addLabel(GmailApp.getUserLabelByName("DFS"));
    }
    
  }
}












