
function updateReadmeV2() {
  var ss = SpreadsheetApp.openById('1zLYAcnX-Mvg8BB59ZM5M0LQGTFDhaFuP-vUZm7NlBpk');
  var readme = ss.getSheetByName('README');
  if (!readme) throw new Error('README not found');
  var lines = [
    'دليل استخدام قاعدة البيانات الأساسية - Future Designers',
    'الهدف: تجميع كل الكيانات الأساسية في مكان واحد منظم، مع رصيد افتتاحي لكل جهة، عشان يبقى نقطة البداية لبناء الفورمات في النظام.',
    'الشيتات:',
    '1) Customers - كل عميل بياناته + رصيده الافتتاحي (الرصيد الافتتاحي بيشمل أي مبلغ لسه معاه شيك بيه ولسه ما اتحصلش، لأن الفلوس لسه ما دخلتش كاش فعليًا).',
    '2) Customer_Payments - شيت حركات: كل تحصيلة فعلية من عميل (كاش أو شيك اتودع البنك فعلا) بتتسجل هنا سطر بسطر، وده اللي بيقلل رصيده الافتتاحي مع الوقت.',
    '3) Suppliers - كل مورد بياناته + رصيده الافتتاحي (ممكن يكون بالجنيه أو بالدولار أو الاتنين).',
    '4) Supplier_Payments - شيت حركات: كل دفعة بتتدفع لمورد بتتسجل هنا سطر بسطر. لو رصيد المورد الافتتاحي بالدولار وانتي بتدفعي بالجنيه، سجلي المبلغ بالجنيه + سعر الصرف وقت الدفع عشان يتحول صح لدولار وينخصم من رصيده.',
    '5) Owners - كل أونر/شريك + رصيد الكاش الافتتاحي بتاعه بس.',
    '6) Cheques - شيت حركات: كل شيك بياخده أونر من عميل بيتسجل هنا (رقم الشيك - مين العميل - المبلغ - تاريخ الاستحقاق - حالته من الدروب داون: In Hand / Deposited / Spent with Customer/Supplier / Bounced).',
    '7) Products_Stock - كل صنف (زيبر) موجود عندك فعليًا في المخزون بمواصفاته + كمية المخزون الافتتاحية.',
    '8) Items - كتالوج/قائمة مرجعية لكل أنواع الأصناف اللي بتتعامل بيها (زي شيت Items القديم) - مش شرط تكون موجودة في المخزون دلوقتي. الهدف الأساسي منها: تبقى مصدر الدروب داون لاختيار الصنف لما نبني فورمات طلب الأوردرات من المورد والبيع للعميل.',
    '9) Assets - قائمة بكل أصل لوحده (اسمه - نوعه/تصنيفه - صاحبه من الأونرز - قيمته).',
    '10) Capital_Movements - شيت حركات: كل مرة أونر يدفع فلوس في التريدينج بيتسجل سطر.',
    '11) Services - قائمة أنواع الخدمات اللي بتقدميها وسعرها الافتراضي.',
    '12) Service_Income - شيت حركات: كل مرة يدخل فلوس من خدمة بيتسجل سطر.',
    'ملاحظات:',
    '- الـ ID في كل شيت اختياري لو حابة الكود يولده تلقائي، أو تكتبيه انتي بنفسك.',
    '- شيتات "الحركات" (Customer_Payments, Supplier_Payments, Cheques, Capital_Movements, Service_Income) بتتزود سطر بسطر مع الوقت.',
    '- شيتات "القوائم" (Items, Assets, Services, Products_Stock, Customers, Suppliers, Owners) هي البيانات الأساسية/الثابتة.',
    '- أي شيك لسه معاكي دلوقتي من عميل (ضمن رصيده الافتتاحي) يتسجل في Cheques بتاريخ النهاردة وحالة "In Hand"، وبرضه يتحسب جوه Opening Balance بتاع العميل عشان الرصيد يكون شامل. لما الشيك يتودع البنك فعلا، غيري حالته في Cheques لـ "Deposited" وسجلي نفس المبلغ كسطر جديد في Customer_Payments - وده اللي بيقلل رصيد العميل الفعلي في الداشبورد.',
    '- لو دفعتي لمورد بعملة مختلفة عن عملة رصيده الافتتاحي، سجلي المبلغ اللي فعلا دفعتيه بعملته + سعر الصرف وقتها في Supplier_Payments عشان يتحسب صح.',
    '- الفرق بين Products_Stock و Items: Products_Stock فيها بس الأصناف اللي عندك مخزون منها فعلا دلوقتي، أما Items فهي الكتالوج الشامل لكل الأصناف اللي بتتعامل بيها (تشتري/تبيع فيها) حتى لو مش موجودة في المخزون حاليًا.',
    '- لو فيه عمود ناقص أو مش محتاجاه، قوليلي وأعدل الهيكل قبل ما تبدأي تعبي بيانات كتير.'
  ];
  readme.getRange(1,1,80,1).clearContent();
  readme.getRange(1,1,lines.length,1).setValues(lines.map(function(l){return [l];}));
  return 'README updated, ' + lines.length + ' lines';
}


var MASTER_DB_ID = '1zLYAcnX-Mvg8BB59ZM5M0LQGTFDhaFuP-vUZm7NlBpk';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🔧 أدوات النظام')
    .addItem('تحديث المخزون من الأوردرات الجديدة', 'processNewOrders')
    .addToUi();
}

function setupOrderSystem() {
  var ss = SpreadsheetApp.openById(MASTER_DB_ID);
  var result = {};

  function ensureSheet(name, headers, statusListCol, statusListValues) {
    var sh = ss.getSheetByName(name);
    var created = false;
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1,1,1,headers.length).setValues([headers]);
      sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
      sh.setFrozenRows(1);
      created = true;
    }
    if (statusListCol && statusListValues) {
      var rule = SpreadsheetApp.newDataValidation().requireValueInList(statusListValues, true).setAllowInvalid(true).build();
      sh.getRange(2, statusListCol, 300, 1).setDataValidation(rule);
    }
    return {sheet: sh, created: created};
  }

  // Customer_Orders
  var co = ensureSheet('Customer_Orders', ['Order ID','Date','Customer','Status','Total EGP','Notes'], 4, ['Draft','Confirmed','Delivered','Cancelled']);
  result.customerOrdersCreated = co.created;
  // Customer dropdown (live-linked to Customers!B)
  var custRule = SpreadsheetApp.newDataValidation().requireValueInRange(ss.getSheetByName('Customers').getRange('B2:B301'), true).setAllowInvalid(true).build();
  co.sheet.getRange(2,3,300,1).setDataValidation(custRule);
  // Total EGP formula
  var coTotalFormulas = [];
  for (var i=2;i<=301;i++){
    coTotalFormulas.push(['=IF(A'+i+'="","",SUMPRODUCT((Customer_Order_Items!$A$2:$A$301=A'+i+')*(Customer_Order_Items!$E$2:$E$301)*(Customer_Order_Items!$F$2:$F$301)))']);
  }
  co.sheet.getRange(2,5,300,1).setFormulas(coTotalFormulas);

  // Customer_Order_Items
  var coi = ensureSheet('Customer_Order_Items', ['Order ID','Item Name','Lenght','Color','Quantity','Unit Price EGP','Stock Status','Notes']);
  result.customerOrderItemsCreated = coi.created;
  var itemRule = SpreadsheetApp.newDataValidation().requireValueInRange(ss.getSheetByName('Items').getRange('C2:C301'), true).setAllowInvalid(true).build();
  coi.sheet.getRange(2,2,300,1).setDataValidation(itemRule);

  // Supplier_Orders
  var so = ensureSheet('Supplier_Orders', ['Order ID','Date','Supplier','Status','Total EGP','Notes'], 4, ['Draft','Confirmed','Received','Cancelled']);
  result.supplierOrdersCreated = so.created;
  var suppRule = SpreadsheetApp.newDataValidation().requireValueInRange(ss.getSheetByName('Suppliers').getRange('B2:B301'), true).setAllowInvalid(true).build();
  so.sheet.getRange(2,3,300,1).setDataValidation(suppRule);
  var soTotalFormulas = [];
  for (var i=2;i<=301;i++){
    soTotalFormulas.push(['=IF(A'+i+'="","",SUMPRODUCT((Supplier_Order_Items!$A$2:$A$301=A'+i+')*(Supplier_Order_Items!$E$2:$E$301)*(Supplier_Order_Items!$F$2:$F$301)))']);
  }
  so.sheet.getRange(2,5,300,1).setFormulas(soTotalFormulas);

  // Supplier_Order_Items
  var soi = ensureSheet('Supplier_Order_Items', ['Order ID','Item Name','Lenght','Color','Quantity','Unit Cost EGP','Stock Status','Notes']);
  result.supplierOrderItemsCreated = soi.created;
  soi.sheet.getRange(2,2,300,1).setDataValidation(itemRule);

  // Reorder sheets
  var desiredOrder = ['README','Dashboard','Customers','Customer_Payments','Customer_Orders','Customer_Order_Items','Suppliers','Supplier_Payments','Supplier_Orders','Supplier_Order_Items','Owners','Cheques','Products_Stock','Items','Assets','Capital_Movements','Services','Service_Income'];
  desiredOrder.forEach(function(name, idx){
    var sh = ss.getSheetByName(name);
    if (sh) {
      ss.setActiveSheet(sh);
      ss.moveActiveSheet(idx + 1);
    }
  });
  result.finalSheetOrder = ss.getSheets().map(function(s){return s.getName();});

  return JSON.stringify(result);
}

function inferItemUnit_(itemName) {
  var n = String(itemName||'').trim().toLowerCase();
  if (n.indexOf('chain') === 0) return 'Mt.';
  return 'PCs';
}
function processNewOrders() {
  var ss = SpreadsheetApp.openById(MASTER_DB_ID);
  var stock = ss.getSheetByName('Products_Stock');
  var stockData = stock.getDataRange().getValues();
  var stockHeaders = stockData[0];
  var iName = stockHeaders.indexOf('Item Name');
  var iLen = stockHeaders.indexOf('Lenght');
  var iColor = stockHeaders.indexOf('Color');
  var iQty = stockHeaders.indexOf('Opening Stock Qty');
  var iUnitStock = stockHeaders.indexOf('Unit');
  var iDateStock = stockHeaders.indexOf('Opening Stock Date');
  var iCostStock = stockHeaders.indexOf('Cost Price EGP');
  var iNotesStock = stockHeaders.indexOf('Notes');

  function findStockRow(name, len, color) {
    for (var r=1;r<stockData.length;r++){
      if (String(stockData[r][iName]).trim() === String(name).trim() &&
          String(stockData[r][iLen]).trim() === String(len).trim() &&
          String(stockData[r][iColor]).trim() === String(color).trim()) {
        return r;
      }
    }
    return -1;
  }

  function createStockRow(name, len, color, cost) {
    var newRow = new Array(stockHeaders.length).fill('');
    newRow[iName] = name;
    newRow[iLen] = len;
    newRow[iColor] = color;
    newRow[iQty] = 0;
    if (iUnitStock > -1) newRow[iUnitStock] = inferItemUnit_(name);
    if (iDateStock > -1) newRow[iDateStock] = new Date();
    if (iCostStock > -1 && cost !== '' && cost !== null && cost !== undefined) newRow[iCostStock] = toNum_(cost);
    if (iNotesStock > -1) newRow[iNotesStock] = 'Auto-added from supplier order';
    stock.appendRow(newRow);
    stockData.push(newRow);
    return stockData.length - 1;
  }

  // Build a map of Order ID -> Status from a header sheet (Customer_Orders / Supplier_Orders)
  function buildStatusMap(headerSheetName) {
    var sh = ss.getSheetByName(headerSheetName);
    var data = sh.getDataRange().getValues();
    var headers = data[0];
    var iOrder = headers.indexOf('Order ID');
    var iStatus = headers.indexOf('Status');
    var map = {};
    for (var r=1;r<data.length;r++){
      var oid = data[r][iOrder];
      if (oid === '' || oid === null) continue;
      map[String(oid).trim()] = data[r][iStatus];
    }
    return map;
  }

  function processSheet(sheetName, sign, headerSheetName, readyStatus, autoCreateStock, costHeader) {
    var statusMap = buildStatusMap(headerSheetName);
    var sh = ss.getSheetByName(sheetName);
    var data = sh.getDataRange().getValues();
    var headers = data[0];
    var iOrder = headers.indexOf('Order ID');
    var iItem = headers.indexOf('Item Name');
    var iLen2 = headers.indexOf('Lenght');
    var iColor2 = headers.indexOf('Color');
    var iQty2 = headers.indexOf('Quantity');
    var iCost2 = costHeader ? headers.indexOf(costHeader) : -1;
    var iStatus = headers.indexOf('Stock Status');
    var processed = 0, errors = 0, waiting = 0, created = 0;
    for (var r=1;r<data.length;r++){
      var row = data[r];
      if (!row[iItem] || !row[iQty2]) continue;
      if (row[iStatus] === 'Processed') continue;
      var orderId = String(row[iOrder]).trim();
      var orderStatus = statusMap[orderId];
      if (orderStatus === undefined) {
        sh.getRange(r+1, iStatus+1).setValue('ERROR: Order #' + orderId + ' not found in ' + headerSheetName);
        errors++;
        continue;
      }
      if (orderStatus !== readyStatus) {
        // order not yet in the status that should affect stock - leave Stock Status blank, skip silently
        waiting++;
        continue;
      }
      var stockRowIdx = findStockRow(row[iItem], row[iLen2], row[iColor2]);
      if (stockRowIdx === -1) {
        if (autoCreateStock) {
          stockRowIdx = createStockRow(row[iItem], row[iLen2], row[iColor2], iCost2 > -1 ? row[iCost2] : '');
          created++;
        } else {
          sh.getRange(r+1, iStatus+1).setValue('ERROR: Item/Length/Color not found in Products_Stock');
          errors++;
          continue;
        }
      }
      var qty = toNum_(row[iQty2]);
      var currentQty = toNum_(stockData[stockRowIdx][iQty]);
      if (sign === -1 && currentQty < qty) {
        sh.getRange(r+1, iStatus+1).setValue('ERROR: Insufficient stock (available ' + currentQty + ', requested ' + qty + ')');
        errors++;
        continue;
      }
      var newQty = currentQty + sign * qty;
      stock.getRange(stockRowIdx+1, iQty+1).setValue(newQty);
      stockData[stockRowIdx][iQty] = newQty;
      sh.getRange(r+1, iStatus+1).setValue('Processed');
      processed++;
    }
    return {processed: processed, errors: errors, waiting: waiting, created: created};
  }

  var custResult = processSheet('Customer_Order_Items', -1, 'Customer_Orders', 'Delivered', false, null);
  var suppResult = processSheet('Supplier_Order_Items', 1, 'Supplier_Orders', 'Received', true, 'Unit Cost EGP');

  var msg = 'Stock Update:\n\n' +
    'Customer Orders (Delivered only): ' + custResult.processed + ' updated, ' + custResult.errors + ' with issues, ' + custResult.waiting + ' waiting to be delivered\n' +
    'Supplier Orders (Received only): ' + suppResult.processed + ' updated' + (suppResult.created ? ' (' + suppResult.created + ' new item(s) auto-added to stock)' : '') + ', ' + suppResult.errors + ' with issues, ' + suppResult.waiting + ' waiting to be received';

  var important = (custResult.processed > 0 || custResult.errors > 0 || suppResult.processed > 0 || suppResult.errors > 0 || suppResult.created > 0);
  var hasIssues = (custResult.errors > 0 || suppResult.errors > 0);

  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch(e) {
    Logger.log(msg);
  }
  return {message: msg, important: important, hasIssues: hasIssues};
}


function addOrdersToDashboardAndReadme() {
  var ss = SpreadsheetApp.openById(MASTER_DB_ID);
  var dash = ss.getSheetByName('Dashboard');
  var lastRow = dash.getLastRow();
  var colA = dash.getRange(1,1,lastRow,1).getValues().map(function(r){return r[0];});
  var already = colA.indexOf('عدد أوردرات العملاء المسجلة');
  var dashResult;
  if (already === -1) {
    // append at the end of the sheet
    var startRow = lastRow + 2;
    dash.getRange(startRow,1).setValue('الأوردرات (Orders)');
    dash.getRange(startRow,1).copyFormatToRange(dash, 1, 1, dash.getRange(colA.indexOf('الشيكات (Cheques)')+1,1).getRow(), dash.getRange(colA.indexOf('الشيكات (Cheques)')+1,1).getRow());
    dash.getRange(startRow+1,1).setValue('عدد أوردرات العملاء المسجلة');
    dash.getRange(startRow+1,2).setFormula('=COUNTA(Customer_Orders!A2:A301)');
    dash.getRange(startRow+2,1).setValue('عدد أوردرات الموردين المسجلة');
    dash.getRange(startRow+2,2).setFormula('=COUNTA(Supplier_Orders!A2:A301)');
    dashResult = 'DONE rows ' + startRow + '-' + (startRow+2);
  } else {
    dashResult = 'ALREADY_EXISTS';
  }

  var readme = ss.getSheetByName('README');
  var lines = [
    'دليل استخدام قاعدة البيانات الأساسية - Future Designers',
    'الهدف: تجميع كل الكيانات الأساسية في مكان واحد منظم، مع رصيد افتتاحي لكل جهة، عشان يبقى نقطة البداية لبناء الفورمات في النظام.',
    'الشيتات:',
    '1) Customers - كل عميل بياناته + رصيده الافتتاحي (شامل أي مبلغ لسه معاه شيك بيه ولسه ما اتحصلش).',
    '2) Customer_Payments - شيت حركات: كل تحصيلة فعلية من عميل (كاش أو شيك اتودع البنك فعلا) بتتسجل هنا سطر بسطر، وده اللي بيقلل رصيده الافتتاحي مع الوقت.',
    '3) Customer_Orders - رأس كل أوردر بيع لعميل (رقم الأوردر - التاريخ - العميل - الحالة - الإجمالي بيتحسب أوتوماتيك من Customer_Order_Items).',
    '4) Customer_Order_Items - بنود كل أوردر عميل (رقم الأوردر - الصنف - الطول - اللون - الكمية - سعر البيع - حالة تحديث المخزون).',
    '5) Suppliers - كل مورد بياناته + رصيده الافتتاحي (ممكن يكون بالجنيه أو بالدولار أو الاتنين).',
    '6) Supplier_Payments - شيت حركات: كل دفعة بتتدفع لمورد بتتسجل هنا سطر بسطر. لو رصيد المورد الافتتاحي بالدولار وانتي بتدفعي بالجنيه، سجلي المبلغ بالجنيه + سعر الصرف وقت الدفع عشان يتحول صح لدولار وينخصم من رصيده.',
    '7) Supplier_Orders - رأس كل أوردر شراء من مورد (رقم الأوردر - التاريخ - المورد - الحالة - الإجمالي بيتحسب أوتوماتيك من Supplier_Order_Items).',
    '8) Supplier_Order_Items - بنود كل أوردر مورد (رقم الأوردر - الصنف - الطول - اللون - الكمية - سعر الشراء - حالة تحديث المخزون).',
    '9) Owners - كل أونر/شريك + رصيد الكاش الافتتاحي بتاعه بس.',
    '10) Cheques - شيت حركات: كل شيك بياخده أونر من عميل بيتسجل هنا (رقم الشيك - مين العميل - المبلغ - تاريخ الاستحقاق - حالته من الدروب داون: In Hand / Deposited / Spent with Customer/Supplier / Bounced).',
    '11) Products_Stock - كل صنف (زيبر) موجود عندك فعليًا في المخزون بمواصفاته + كمية المخزون الحالية (بتتحدث تلقائي من الأوردرات).',
    '12) Items - كتالوج/قائمة مرجعية لكل أنواع الأصناف اللي بتتعامل بيها - مصدر الدروب داون لاختيار الصنف في الأوردرات.',
    '13) Assets - قائمة بكل أصل لوحده (اسمه - نوعه/تصنيفه - صاحبه من الأونرز - قيمته).',
    '14) Capital_Movements - شيت حركات: كل مرة أونر يدفع فلوس في التريدينج بيتسجل سطر.',
    '15) Services - قائمة أنواع الخدمات اللي بتقدميها وسعرها الافتراضي.',
    '16) Service_Income - شيت حركات: كل مرة يدخل فلوس من خدمة بيتسجل سطر.',
    'الأوردرات وتحديث المخزون:',
    '- سجلي الأوردر الأول في Customer_Orders أو Supplier_Orders (رقم أوردر + تاريخ + عميل/مورد + حالة)، وبعدين سجلي كل صنف في الأوردر كسطر منفصل في Customer_Order_Items أو Supplier_Order_Items بنفس رقم الأوردر.',
    '- لازم اسم الصنف + الطول + اللون في سطر البند يكونوا مطابقين بالظبط لصف موجود في Products_Stock عشان الكود يعرف يحدث الكمية الصح.',
    '- لما تخلصي تسجيل الأوردرات الجديدة، افتحي من فوق قايمة "🔧 أدوات النظام" واضغطي "تحديث المخزون من الأوردرات الجديدة" - هيخصم الكمية من المخزون لأوردرات العملاء (بيع) ويزودها لأوردرات الموردين (شراء)، ويحط "Processed" في عمود Stock Status. لو في صنف مش لاقيه في Products_Stock هيكتبلك رسالة خطأ في نفس السطر عشان تصلحيه وتجربي تاني.',
    '- القايمة دي بتظهر بس لما تقفلي وتفتحي الشيت تاني بعد أي تحديث في الكود.',
    'ملاحظات:',
    '- الـ ID في كل شيت اختياري لو حابة الكود يولده تلقائي، أو تكتبيه انتي بنفسك. رقم الأوردر مهم تكتبيه بنفسك وتستخدميه في كل بنوده.',
    '- شيتات "الحركات" (Customer_Payments, Supplier_Payments, Cheques, Capital_Movements, Service_Income, Customer_Order_Items, Supplier_Order_Items) بتتزود سطر بسطر مع الوقت.',
    '- شيتات "القوائم" (Items, Assets, Services, Products_Stock, Customers, Suppliers, Owners, Customer_Orders, Supplier_Orders) هي البيانات الأساسية/رؤوس الأوردرات.',
    '- أي شيك لسه معاكي دلوقتي من عميل (ضمن رصيده الافتتاحي) يتسجل في Cheques بتاريخ النهاردة وحالة "In Hand"، وبرضه يتحسب جوه Opening Balance بتاع العميل عشان الرصيد يكون شامل. لما الشيك يتودع البنك فعلا، غيري حالته في Cheques لـ "Deposited" وسجلي نفس المبلغ كسطر جديد في Customer_Payments.',
    '- لو دفعتي لمورد بعملة مختلفة عن عملة رصيده الافتتاحي، سجلي المبلغ اللي فعلا دفعتيه بعملته + سعر الصرف وقتها في Supplier_Payments عشان يتحسب صح.',
    '- الفرق بين Products_Stock و Items: Products_Stock فيها بس الأصناف اللي عندك مخزون منها فعلا دلوقتي، أما Items فهي الكتالوج الشامل لكل الأصناف اللي بتتعامل بيها حتى لو مش موجودة في المخزون حاليًا.',
    '- لو فيه عمود ناقص أو مش محتاجاه، قوليلي وأعدل الهيكل قبل ما تبدأي تعبي بيانات كتير.'
  ];
  readme.getRange(1,1,100,1).clearContent();
  readme.getRange(1,1,lines.length,1).setValues(lines.map(function(l){return [l];}));

  return JSON.stringify({dashResult: dashResult, readmeLines: lines.length});
}


function addDeliveredOrderTotalsToDashboardAndReadme() {
  var ss = SpreadsheetApp.openById(MASTER_DB_ID);
  var dash = ss.getSheetByName('Dashboard');
  var lastRow = dash.getLastRow();
  var colA = dash.getRange(1,1,lastRow,1).getValues().map(function(r){return r[0];});
  var already = colA.indexOf('إجمالي مبيعات العملاء (أوردرات Delivered) EGP');
  var dashResult;
  if (already === -1) {
    var ordersHeaderRow = colA.indexOf('الأوردرات (Orders)') + 1;
    var custOrdersCountRow = ordersHeaderRow + 1; // 'عدد أوردرات العملاء المسجلة'
    var suppOrdersCountRow = ordersHeaderRow + 2; // 'عدد أوردرات الموردين المسجلة'
    dash.insertRowsAfter(suppOrdersCountRow, 2);
    var r1 = suppOrdersCountRow + 1, r2 = suppOrdersCountRow + 2;
    dash.getRange(r1,1).setValue('إجمالي مبيعات العملاء (أوردرات Delivered) EGP');
    dash.getRange(r1,2).setFormula('=SUMIFS(Customer_Orders!E2:E301,Customer_Orders!D2:D301,"Delivered")');
    dash.getRange(r2,1).setValue('إجمالي مشتريات من الموردين (أوردرات Received) EGP');
    dash.getRange(r2,2).setFormula('=SUMIFS(Supplier_Orders!E2:E301,Supplier_Orders!D2:D301,"Received")');
    dashResult = 'DONE rows ' + r1 + ',' + r2;
  } else {
    dashResult = 'ALREADY_EXISTS';
  }

  var readme = ss.getSheetByName('README');
  var lines = [
    'دليل استخدام قاعدة البيانات الأساسية - Future Designers',
    'الهدف: تجميع كل الكيانات الأساسية في مكان واحد منظم، مع رصيد افتتاحي لكل جهة، عشان يبقى نقطة البداية لبناء الفورمات في النظام.',
    'الشيتات:',
    '1) Customers - كل عميل بياناته + رصيده الافتتاحي (شامل أي مبلغ لسه معاه شيك بيه ولسه ما اتحصلش).',
    '2) Customer_Payments - شيت حركات: كل تحصيلة فعلية من عميل (كاش أو شيك اتودع البنك فعلا) بتتسجل هنا سطر بسطر - بتقلل رصيد العميل.',
    '3) Customer_Orders - رأس كل أوردر بيع لعميل (رقم الأوردر - التاريخ - العميل - الحالة - الإجمالي بيتحسب أوتوماتيك). أوردر Delivered بيزود رصيد العميل (دين جديد عليه).',
    '4) Customer_Order_Items - بنود كل أوردر عميل (رقم الأوردر - الصنف - الطول - اللون - الكمية - سعر البيع - حالة تحديث المخزون).',
    '5) Suppliers - كل مورد بياناته + رصيده الافتتاحي (ممكن يكون بالجنيه أو بالدولار أو الاتنين).',
    '6) Supplier_Payments - شيت حركات: كل دفعة بتتدفع لمورد بتتسجل هنا سطر بسطر - بتقلل رصيد المورد. لو رصيد المورد الافتتاحي بالدولار وانتي بتدفعي بالجنيه، سجلي المبلغ بالجنيه + سعر الصرف وقت الدفع.',
    '7) Supplier_Orders - رأس كل أوردر شراء من مورد (رقم الأوردر - التاريخ - المورد - الحالة - الإجمالي بيتحسب أوتوماتيك). أوردر Received بيزود رصيد المورد (دين جديد عليكي).',
    '8) Supplier_Order_Items - بنود كل أوردر مورد (رقم الأوردر - الصنف - الطول - اللون - الكمية - سعر الشراء - حالة تحديث المخزون).',
    '9) Owners - كل أونر/شريك + رصيد الكاش الافتتاحي بتاعه بس.',
    '10) Cheques - شيت حركات: كل شيك بياخده أونر من عميل بيتسجل هنا (رقم الشيك - مين العميل - المبلغ - تاريخ الاستحقاق - حالته: In Hand / Deposited / Spent with Customer/Supplier / Bounced).',
    '11) Products_Stock - كل صنف (زيبر) موجود عندك فعليًا في المخزون بمواصفاته + كمية المخزون الحالية.',
    '12) Items - كتالوج/قائمة مرجعية لكل أنواع الأصناف اللي بتتعامل بيها - مصدر الدروب داون لاختيار الصنف في الأوردرات.',
    '13) Assets - قائمة بكل أصل لوحده (اسمه - نوعه/تصنيفه - صاحبه من الأونرز - قيمته).',
    '14) Capital_Movements - شيت حركات: كل مرة أونر يدفع فلوس في التريدينج بيتسجل سطر.',
    '15) Services - قائمة أنواع الخدمات اللي بتقدميها وسعرها الافتراضي.',
    '16) Service_Income - شيت حركات: كل مرة يدخل فلوس من خدمة بيتسجل سطر.',
    'الأوردرات، المخزون، ورصيد العميل/المورد:',
    '- سجلي رأس الأوردر في Customer_Orders أو Supplier_Orders بحالة "Draft" في الأول، وسجلي بنوده في Customer_Order_Items أو Supplier_Order_Items بنفس رقم الأوردر.',
    '- المخزون ورصيد العميل/المورد ملهمش دعوة بالأوردر خالص لحد ما تغيري حالته لـ "Delivered" (للعميل) أو "Received" (للمورد). يعني تقدري تعدلي في الأوردر وانتي لسه بتحضريه من غير ما تأثري على حاجة.',
    '- لما الأوردر يتسلم فعلا: غيري حالته لـ Delivered/Received، وبعدين افتحي من فوق قايمة "🔧 أدوات النظام" واضغطي "تحديث المخزون من الأوردرات الجديدة". وقتها بس المخزون هيتحدث (يقل للعميل، يزيد للمورد)، وإجمالي الأوردر هيدخل في حساب رصيد العميل/المورد في الداشبورد.',
    '- رصيد العميل الكامل تقريبًا = Opening Balance + إجمالي أوردراته الـ Delivered - إجمالي تحصيلاته في Customer_Payments. ورصيد المورد بنفس المنطق: Opening Balance + إجمالي أوردراته الـ Received - إجمالي مدفوعاته في Supplier_Payments.',
    '- لازم اسم الصنف + الطول + اللون في سطر البند يكونوا مطابقين بالظبط لصف موجود في Products_Stock عشان الكود يعرف يحدث الكمية الصح. لو في خطأ هيكتبلك رسالة في نفس السطر في عمود Stock Status.',
    '- القايمة دي بتظهر بس لما تقفلي وتفتحي الشيت تاني بعد أي تحديث في الكود.',
    'ملاحظات:',
    '- الـ ID في كل شيت اختياري لو حابة الكود يولده تلقائي، أو تكتبيه انتي بنفسك. رقم الأوردر مهم تكتبيه بنفسك وتستخدميه في كل بنوده.',
    '- شيتات "الحركات" (Customer_Payments, Supplier_Payments, Cheques, Capital_Movements, Service_Income, Customer_Order_Items, Supplier_Order_Items) بتتزود سطر بسطر مع الوقت.',
    '- شيتات "القوائم" (Items, Assets, Services, Products_Stock, Customers, Suppliers, Owners, Customer_Orders, Supplier_Orders) هي البيانات الأساسية/رؤوس الأوردرات.',
    '- أي شيك لسه معاكي دلوقتي من عميل (ضمن رصيده الافتتاحي) يتسجل في Cheques بتاريخ النهاردة وحالة "In Hand". لما يتودع البنك فعلا، غيري حالته لـ "Deposited" وسجلي نفس المبلغ كسطر جديد في Customer_Payments.',
    '- الفرق بين Products_Stock و Items: Products_Stock فيها بس الأصناف اللي عندك مخزون منها فعلا دلوقتي، أما Items فهي الكتالوج الشامل لكل الأصناف اللي بتتعامل بيها حتى لو مش موجودة في المخزون حاليًا.',
    '- لو فيه عمود ناقص أو مش محتاجاه، قوليلي وأعدل الهيكل قبل ما تبدأي تعبي بيانات كتير.'
  ];
  readme.getRange(1,1,100,1).clearContent();
  readme.getRange(1,1,lines.length,1).setValues(lines.map(function(l){return [l];}));

  return JSON.stringify({dashResult: dashResult, readmeLines: lines.length});
}


function inspectOpeningSheets() {
  var ss = SpreadsheetApp.openById('1zLYAcnX-Mvg8BB59ZM5M0LQGTFDhaFuP-vUZm7NlBpk');
  var names = ['Owners','Customers','Suppliers'];
  var out = {};
  names.forEach(function(n){
    var sh = ss.getSheetByName(n);
    if (!sh) { out[n] = 'MISSING'; return; }
    var lastCol = sh.getLastColumn();
    var lastRow = sh.getLastRow();
    var headers = lastCol > 0 ? sh.getRange(1,1,1,lastCol).getValues()[0] : [];
    out[n] = {headers: headers, lastRow: lastRow};
  });
  return JSON.stringify(out);
}

// ===================== WEB APP =====================
function doGet(e) {
  var tmpl = HtmlService.createTemplateFromFile('Index');
  return tmpl.evaluate()
    .setTitle('Future Designers - نظام الإدارة')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function ss_() { return SpreadsheetApp.openById(MASTER_DB_ID); }



function sheetToObjects_(sheetName) {
  var sh = ss_().getSheetByName(sheetName);
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var rows = [];
  for (var r=1;r<data.length;r++){
    var obj = {};
    var empty = true;
    for (var c=0;c<headers.length;c++){
      var v = data[r][c];
      if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      obj[headers[c]] = v;
      if (v !== '' && v !== null && v !== undefined) empty = false;
    }
    if (!empty) { obj.__row = r+1; rows.push(obj); }
  }
  return rows;
}

function getNextId_(sh, headers, idHeader) {
  var idIdx = headers.indexOf(idHeader);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return 1;
  var vals = sh.getRange(2, idIdx+1, lastRow-1, 1).getValues();
  var maxId = 0;
  vals.forEach(function(row){ var n = Number(row[0]); if (!isNaN(n) && n > maxId) maxId = n; });
  return maxId + 1;
}

// ---- Dashboard ----
function sumCol_(sheetName, colHeader) {
  var sh = ss_().getSheetByName(sheetName);
  var data = sh.getDataRange().getValues();
  var idx = data[0].indexOf(colHeader);
  if (idx === -1) return 0;
  var total = 0;
  for (var r=1;r<data.length;r++){ total += toNum_(data[r][idx]); }
  return total;
}
function sumIfCol_(sheetName, sumHeader, critHeader, critValue) {
  var sh = ss_().getSheetByName(sheetName);
  var data = sh.getDataRange().getValues();
  var sIdx = data[0].indexOf(sumHeader), cIdx = data[0].indexOf(critHeader);
  if (sIdx === -1 || cIdx === -1) return 0;
  var total = 0;
  for (var r=1;r<data.length;r++){ if (data[r][cIdx] === critValue) { total += toNum_(data[r][sIdx]); } }
  return total;
}
function countCol_(sheetName, colHeader) {
  var sh = ss_().getSheetByName(sheetName);
  var data = sh.getDataRange().getValues();
  var idx = data[0].indexOf(colHeader);
  if (idx === -1) return 0;
  var c = 0;
  for (var r=1;r<data.length;r++){ if (data[r][idx] !== '' && data[r][idx] !== null) c++; }
  return c;
}
function countByStatus_(sheetName, statusHeader) {
  var sh = ss_().getSheetByName(sheetName);
  var data = sh.getDataRange().getValues();
  var idx = data[0].indexOf(statusHeader);
  if (idx === -1) return {};
  var counts = {};
  for (var r=1;r<data.length;r++){
    var v = data[r][idx];
    if (v === '' || v === null || v === undefined) continue;
    counts[v] = (counts[v]||0) + 1;
  }
  return counts;
}
function sumProductCol_(sheetName, headerA, headerB) {
  var sh = ss_().getSheetByName(sheetName);
  var data = sh.getDataRange().getValues();
  var idxA = data[0].indexOf(headerA), idxB = data[0].indexOf(headerB);
  if (idxA === -1 || idxB === -1) return 0;
  var total = 0;
  for (var r=1;r<data.length;r++){ var a = toNum_(data[r][idxA]), b = toNum_(data[r][idxB]); total += a*b; }
  return total;
}
function fmt_(n) { return Number(n||0).toLocaleString('en-US', {minimumFractionDigits:2, maximumFractionDigits:2}); }

function api_getDashboard() {
  var sections = [];
  var custStatusCounts = countByStatus_('Customer_Orders', 'Status');
  var suppStatusCounts = countByStatus_('Supplier_Orders', 'Status');
  var pendingCustomerCount = 0, pendingSupplierCount = 0;
  Object.keys(custStatusCounts).forEach(function(k){ if (k !== 'Delivered' && k !== 'Cancelled') pendingCustomerCount += custStatusCounts[k]; });
  Object.keys(suppStatusCounts).forEach(function(k){ if (k !== 'Received' && k !== 'Cancelled') pendingSupplierCount += suppStatusCounts[k]; });
  sections.push({title: 'Orders', items: [
    {label: 'Customer Orders Recorded', value: countCol_('Customer_Orders','Order ID'), currency: null},
    {label: 'Supplier Orders Recorded', value: countCol_('Supplier_Orders','Order ID'), currency: null},
    {label: 'Pending Customer Orders', value: pendingCustomerCount, currency: null, action: 'customer'},
    {label: 'Pending Supplier Orders', value: pendingSupplierCount, currency: null, action: 'supplier'}
  ]});
  sections.push({title: 'Customer Orders — Status Breakdown', items: Object.keys(custStatusCounts).map(function(k){ return {label: k, value: custStatusCounts[k], currency: null}; })});
  sections.push({title: 'Supplier Orders — Status Breakdown', items: Object.keys(suppStatusCounts).map(function(k){ return {label: k, value: suppStatusCounts[k], currency: null}; })});
  sections.push({title: 'Item Quantities', items: [
    {label: 'Total Qty Ordered by Customers', value: sumCol_('Customer_Order_Items','Quantity'), currency: null},
    {label: 'Total Qty Ordered from Suppliers', value: sumCol_('Supplier_Order_Items','Quantity'), currency: null}
  ]});
  sections.push({title: 'Stock', items: [
    {label: 'Items in Stock', value: countCol_('Products_Stock','Item Name'), currency: null}
  ]});
  sections.push({title: 'Items Catalog', items: [
    {label: 'Items Registered', value: countCol_('Items','Item Name'), currency: null}
  ]});
  return sections;
}

function computeBalances_(baseSheet, nameField, ordersSheet, orderPartyField, paymentsSheet, paymentPartyField) {
  var rows = sheetToObjects_(baseSheet);
  var orders = sheetToObjects_(ordersSheet);
  var payments = sheetToObjects_(paymentsSheet);
  var orderTotals = {};
  orders.forEach(function(o){
    var key = String(o[orderPartyField]||'').trim();
    if (!key) return;
    orderTotals[key] = (orderTotals[key]||0) + toNum_(o['Total EGP']);
  });
  var paidEGP = {}, paidUSD = {};
  payments.forEach(function(p){
    var key = String(p[paymentPartyField]||'').trim();
    if (!key) return;
    paidEGP[key] = (paidEGP[key]||0) + toNum_(p['Amount EGP']);
    paidUSD[key] = (paidUSD[key]||0) + toNum_(p['Amount USD']);
  });
  rows.forEach(function(r){
    var key = String(r[nameField]||'').trim();
    var openEGP = toNum_(r['Opening Balance EGP']);
    var openUSD = toNum_(r['Opening Balance USD']);
    r['Balance EGP'] = openEGP + (orderTotals[key]||0) - (paidEGP[key]||0);
    r['Balance USD'] = openUSD - (paidUSD[key]||0);
  });
  return rows;
}

// ---- Customers ----
function api_getCustomers() { return computeBalances_('Customers', 'Customer Name', 'Customer_Orders', 'Customer', 'Customer_Payments', 'Customer'); }
function api_getChequesSummary() {
  return [
    {label: 'Cheques In Hand EGP', value: sumIfCol_('Cheques','Amount EGP','Status','In Hand'), currency: 'EGP'},
    {label: 'Cheques Deposited EGP', value: sumIfCol_('Cheques','Amount EGP','Status','Deposited'), currency: 'EGP'},
    {label: 'Total Cheques Recorded', value: countCol_('Cheques','Cheque Number'), currency: null}
  ];
}
function api_getCustomersSummary() {
  var rows = computeBalances_('Customers', 'Customer Name', 'Customer_Orders', 'Customer', 'Customer_Payments', 'Customer');
  var outEGP = 0, outUSD = 0;
  rows.forEach(function(r){ outEGP += toNum_(r['Balance EGP']); outUSD += toNum_(r['Balance USD']); });
  return [
    {label: 'Total Opening Balance EGP', value: sumCol_('Customers','Opening Balance EGP'), currency: 'EGP'},
    {label: 'Total Opening Balance USD', value: sumCol_('Customers','Opening Balance USD'), currency: 'USD'},
    {label: 'Total Collected from Customers EGP', value: sumCol_('Customer_Payments','Amount EGP'), currency: 'EGP'},
    {label: 'Total Collected from Customers USD', value: sumCol_('Customer_Payments','Amount USD'), currency: 'USD'},
    {label: 'Sales (Delivered Orders) EGP', value: sumIfCol_('Customer_Orders','Total EGP','Status','Delivered'), currency: 'EGP'},
    {label: 'Total Outstanding Now EGP', value: outEGP, currency: 'EGP'},
    {label: 'Total Outstanding Now USD', value: outUSD, currency: 'USD'}
  ].concat(api_getChequesSummary());
}
function api_addCustomer(data, user) {
  var sh = ss_().getSheetByName('Customers');
  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var nextId = getNextId_(sh, headers, 'Customer ID');
  var row = headers.map(function(h){
    if (h === 'Customer ID') return nextId;
    return data[h] !== undefined ? data[h] : '';
  });
  sh.appendRow(row);
  logAction_(user, 'Customer Added', 'Customer #' + nextId + ': ' + (data['Customer Name']||''));
  return nextId;
}
function api_prepareCustomerOrderFromSupplier(supplierOrderId) {
  var ss = ss_();
  var siSh = ss.getSheetByName('Supplier_Order_Items');
  var siData = siSh.getDataRange().getValues();
  var siH = siData[0];
  var siOrder = siH.indexOf('Order ID');
  var siItem = siH.indexOf('Item Name');
  var siLen = siH.indexOf('Lenght');
  var siColor = siH.indexOf('Color');
  var siQty = siH.indexOf('Quantity');
  var siCost = siH.indexOf('Unit Cost EGP');
  var fullItems = [];
  for (var r = 1; r < siData.length; r++) {
    if (String(siData[r][siOrder]).trim() === String(supplierOrderId).trim()) {
      fullItems.push({
        itemName: siData[r][siItem],
        length: siData[r][siLen],
        color: siData[r][siColor],
        qty: toNum_(siData[r][siQty]),
        price: siData[r][siCost]
      });
    }
  }

  var soSh = ss.getSheetByName('Supplier_Orders');
  var soData = soSh.getDataRange().getValues();
  var soH = soData[0];
  var soIdIdx = soH.indexOf('Order ID');
  var soNumIdx = soH.indexOf('Order Number');
  var supplierOrderNumber = '';
  for (var r2 = 1; r2 < soData.length; r2++) {
    if (String(soData[r2][soIdIdx]).trim() === String(supplierOrderId).trim()) {
      supplierOrderNumber = soData[r2][soNumIdx];
      break;
    }
  }

  var coSh = ss.getSheetByName('Customer_Orders');
  var coData = coSh.getDataRange().getValues();
  var coH = coData[0];
  var coIdIdx = coH.indexOf('Order ID');
  var coNotesIdx = coH.indexOf('Notes');
  var coStatusIdx = coH.indexOf('Status');
  var markerRe = new RegExp('Created from Supplier Order #' + supplierOrderId + '(\\D|$)');
  var priorOrderIds = [];
  for (var r3 = 1; r3 < coData.length; r3++) {
    var notes = String(coData[r3][coNotesIdx] || '');
    if (markerRe.test(notes) && coData[r3][coStatusIdx] !== 'Cancelled') {
      priorOrderIds.push(coData[r3][coIdIdx]);
    }
  }

  var alreadyConverted = {};
  if (priorOrderIds.length) {
    var ciSh = ss.getSheetByName('Customer_Order_Items');
    var ciData = ciSh.getDataRange().getValues();
    var ciH = ciData[0];
    var ciOrder = ciH.indexOf('Order ID');
    var ciItem = ciH.indexOf('Item Name');
    var ciLen = ciH.indexOf('Lenght');
    var ciColor = ciH.indexOf('Color');
    var ciQty = ciH.indexOf('Quantity');
    var priorSet = {};
    priorOrderIds.forEach(function(id){ priorSet[String(id)] = true; });
    for (var r4 = 1; r4 < ciData.length; r4++) {
      if (priorSet[String(ciData[r4][ciOrder]).trim()]) {
        var key = String(ciData[r4][ciItem]).trim() + '|' + String(ciData[r4][ciLen]).trim() + '|' + String(ciData[r4][ciColor]).trim();
        alreadyConverted[key] = (alreadyConverted[key] || 0) + toNum_(ciData[r4][ciQty]);
      }
    }
  }

  var remainingItems = [];
  var anyRemaining = false;
  fullItems.forEach(function(it){
    var key = String(it.itemName).trim() + '|' + String(it.length).trim() + '|' + String(it.color).trim();
    var used = alreadyConverted[key] || 0;
    var remaining = it.qty - used;
    if (remaining > 0) { anyRemaining = true; }
    remainingItems.push({itemName: it.itemName, length: it.length, color: it.color, qty: Math.max(0, remaining), price: it.price});
  });

  return {
    supplierOrderNumber: supplierOrderNumber,
    priorOrderIds: priorOrderIds,
    fullItems: fullItems,
    remainingItems: remainingItems,
    fullyConverted: priorOrderIds.length > 0 && !anyRemaining
  };
}
function api_updateCustomerDetails(customerId, data, user) {
  var sh = ss_().getSheetByName('Customers');
  var sheetData = sh.getDataRange().getValues();
  var headers = sheetData[0];
  var iId = headers.indexOf('Customer ID');
  for (var r = 1; r < sheetData.length; r++) {
    if (String(sheetData[r][iId]) === String(customerId)) {
      ['Phone','Address','Location URL','Notes'].forEach(function(field){
        var iCol = headers.indexOf(field);
        if (iCol > -1 && data[field] !== undefined) {
          sh.getRange(r+1, iCol+1).setValue(data[field]);
        }
      });
      logAction_(user, 'Customer Details Updated', 'Customer #' + customerId);
      return 'OK';
    }
  }
  throw new Error('Customer not found');
}

// ---- Suppliers ----
function api_getSuppliers() { return computeBalances_('Suppliers', 'Supplier Name', 'Supplier_Orders', 'Supplier', 'Supplier_Payments', 'Supplier'); }
function api_getSuppliersSummary() {
  var rows = computeBalances_('Suppliers', 'Supplier Name', 'Supplier_Orders', 'Supplier', 'Supplier_Payments', 'Supplier');
  var outEGP = 0, outUSD = 0;
  rows.forEach(function(r){ outEGP += toNum_(r['Balance EGP']); outUSD += toNum_(r['Balance USD']); });
  return [
    {label: 'Total Opening Balance EGP', value: sumCol_('Suppliers','Opening Balance EGP'), currency: 'EGP'},
    {label: 'Total Opening Balance USD', value: sumCol_('Suppliers','Opening Balance USD'), currency: 'USD'},
    {label: 'Total Paid to Suppliers EGP', value: sumCol_('Supplier_Payments','Amount EGP'), currency: 'EGP'},
    {label: 'Total Paid to Suppliers USD', value: sumCol_('Supplier_Payments','Amount USD'), currency: 'USD'},
    {label: 'Purchases (Received Orders) EGP', value: sumIfCol_('Supplier_Orders','Total EGP','Status','Received'), currency: 'EGP'},
    {label: 'Total Outstanding Now EGP', value: outEGP, currency: 'EGP'},
    {label: 'Total Outstanding Now USD', value: outUSD, currency: 'USD'}
  ];
}
function api_addSupplier(data, user) {
  var sh = ss_().getSheetByName('Suppliers');
  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var nextId = getNextId_(sh, headers, 'Supplier ID');
  var row = headers.map(function(h){
    if (h === 'Supplier ID') return nextId;
    return data[h] !== undefined ? data[h] : '';
  });
  sh.appendRow(row);
  logAction_(user, 'Supplier Added', 'Supplier #' + nextId + ': ' + (data['Supplier Name']||''));
  return nextId;
}
function api_updateSupplierDetails(supplierId, data, user) {
  var sh = ss_().getSheetByName('Suppliers');
  var sheetData = sh.getDataRange().getValues();
  var headers = sheetData[0];
  var iId = headers.indexOf('Supplier ID');
  for (var r = 1; r < sheetData.length; r++) {
    if (String(sheetData[r][iId]) === String(supplierId)) {
      ['Phone','Address','Location URL','Notes'].forEach(function(field){
        var iCol = headers.indexOf(field);
        if (iCol > -1 && data[field] !== undefined) {
          sh.getRange(r+1, iCol+1).setValue(data[field]);
        }
      });
      logAction_(user, 'Supplier Details Updated', 'Supplier #' + supplierId);
      return 'OK';
    }
  }
  throw new Error('Supplier not found');
}
function api_getCheques() { return sheetToObjects_('Cheques'); }
function api_addCheque(data, user) {
  var sh = ss_().getSheetByName('Cheques');
  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var nextId = getNextId_(sh, headers, 'Cheque ID');
  var row = headers.map(function(h){
    if (h === 'Cheque ID') return nextId;
    if (h === 'Status') return data[h] || 'In Hand';
    return data[h] !== undefined ? data[h] : '';
  });
  sh.appendRow(row);
  logAction_(user, 'Cheque Added', 'Cheque #' + (data['Cheque Number']||'') + ' from ' + (data['Customer']||'') + ': ' + (data['Amount EGP']||0) + ' EGP');
  return nextId;
}
function api_updateChequeStatus(chequeId, newStatus, spentOn, user) {
  var sh = ss_().getSheetByName('Cheques');
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var iId = headers.indexOf('Cheque ID');
  var iStatus = headers.indexOf('Status');
  var iSpent = headers.indexOf('Spent On / Reference');
  var iCustomer = headers.indexOf('Customer');
  var iAmount = headers.indexOf('Amount EGP');
  var iNumber = headers.indexOf('Cheque Number');
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][iId]) === String(chequeId)) {
      var prevStatus = data[r][iStatus] || 'In Hand';
      sh.getRange(r+1, iStatus+1).setValue(newStatus);
      if (iSpent > -1 && spentOn !== undefined && spentOn !== null) sh.getRange(r+1, iSpent+1).setValue(spentOn);
      if (newStatus === 'Deposited' && prevStatus !== 'Deposited') {
        var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
        api_addCustomerPayment({
          date: todayStr,
          customer: data[r][iCustomer],
          amountEGP: data[r][iAmount],
          amountUSD: '',
          rate: '',
          notes: 'Cheque #' + (data[r][iNumber]||'') + ' deposited' + (spentOn ? ' (' + spentOn + ')' : '')
        }, user);
      }
      logAction_(user, 'Cheque Status Updated', 'Cheque #' + chequeId + ' -> ' + newStatus);
      return 'OK';
    }
  }
  throw new Error('Cheque not found');
}
function api_deleteCheque(chequeId, user) {
  requireAdmin_(user);
  var sh = ss_().getSheetByName('Cheques');
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var iId = headers.indexOf('Cheque ID');
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][iId]) === String(chequeId)) {
      sh.deleteRow(r+1);
      logAction_(user, 'Cheque Deleted', 'Cheque #' + chequeId);
      return 'OK';
    }
  }
  throw new Error('Cheque not found');
}
function api_getLedger(type, partyName) {
  var isCustomer = type === 'customer';
  var baseSheet = isCustomer ? 'Customers' : 'Suppliers';
  var nameField = isCustomer ? 'Customer Name' : 'Supplier Name';
  var ordersSheet = isCustomer ? 'Customer_Orders' : 'Supplier_Orders';
  var partyField = isCustomer ? 'Customer' : 'Supplier';
  var paymentsSheet = isCustomer ? 'Customer_Payments' : 'Supplier_Payments';

  var baseRows = sheetToObjects_(baseSheet);
  var partyRow = baseRows.find(function(r){ return String(r[nameField]||'').trim() === String(partyName).trim(); });
  if (!partyRow) throw new Error((isCustomer?'Customer':'Supplier') + ' not found');

  var orders = sheetToObjects_(ordersSheet).filter(function(o){ return String(o[partyField]||'').trim() === String(partyName).trim(); });
  var payments = sheetToObjects_(paymentsSheet).filter(function(p){ return String(p[partyField]||'').trim() === String(partyName).trim(); });

  var entries = [];
  orders.forEach(function(o){
    entries.push({ date: o['Date'], type: 'order', orderId: o['Order ID'], label: 'Order' + (o['Order Number'] ? ' #' + o['Order Number'] : '') + ' (Order ID ' + o['Order ID'] + ')', amount: toNum_(o['Total EGP']), status: o['Status'] || 'Draft' });
  });
  payments.forEach(function(p){
    entries.push({ date: p['Date'], type: 'payment', label: 'Payment' + (p['Notes'] ? ' \u2014 ' + p['Notes'] : ''), amount: -toNum_(p['Amount EGP']) });
  });
  entries.sort(function(a,b){ return new Date(a.date||0) - new Date(b.date||0); });

  var runningBal = toNum_(partyRow['Opening Balance EGP']);
  entries.forEach(function(e){ runningBal += e.amount; e.runningBalance = runningBal; });

  var cheques = [];
  if (isCustomer) {
    cheques = sheetToObjects_('Cheques').filter(function(c){ return String(c['Customer']||'').trim() === String(partyName).trim(); })
      .sort(function(a,b){ return new Date(a['Cheque Due Date']||0) - new Date(b['Cheque Due Date']||0); });
  }

  return {
    party: partyName,
    openingBalance: toNum_(partyRow['Opening Balance EGP']),
    openingDate: partyRow['Opening Balance Date'],
    entries: entries,
    finalBalance: runningBal,
    cheques: cheques
  };
}

// ---- Dropdown helpers ----
function api_getItemNames() {
  var sh = ss_().getSheetByName('Items');
  var data = sh.getDataRange().getValues();
  var idx = data[0].indexOf('Item Name');
  var names = [];
  for (var r=1;r<data.length;r++){ if (data[r][idx]) names.push(data[r][idx]); }
  return names;
}
function api_getItems() {
  var sh = ss_().getSheetByName('Items');
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var rows = [];
  for (var r = 1; r < data.length; r++) {
    var obj = {};
    headers.forEach(function(h, i){ obj[h] = data[r][i]; });
    rows.push(obj);
  }
  return rows;
}
function api_addItem(data, user) {
  var sh = ss_().getSheetByName('Items');
  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var name = (data['Item Name']||'').toString().trim();
  if (!name) throw new Error('Item Name is required');
  var existing = sh.getDataRange().getValues();
  var iName = headers.indexOf('Item Name');
  for (var r = 1; r < existing.length; r++) {
    if (String(existing[r][iName]).trim().toLowerCase() === name.toLowerCase()) {
      throw new Error('Item "' + name + '" already exists');
    }
  }
  var row = headers.map(function(h){
    if (h === 'Item ID') return name;
    if (h === 'Item Name') return name;
    return data[h] !== undefined ? data[h] : '';
  });
  sh.appendRow(row);
  logAction_(user, 'Item Added', name);
  return 'OK';
}
function api_updateItem(oldName, data, user) {
  oldName = (oldName||'').toString().trim();
  var newName = (data['Item Name']||'').toString().trim();
  if (!newName) throw new Error('Item Name is required');
  var ss = ss_();
  var sh = ss.getSheetByName('Items');
  var values = sh.getDataRange().getValues();
  var headers = values[0];
  var iName = headers.indexOf('Item Name');
  var rowIdx = -1;
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][iName]).trim() === oldName) { rowIdx = r; break; }
  }
  if (rowIdx === -1) throw new Error('Item not found: ' + oldName);
  if (newName.toLowerCase() !== oldName.toLowerCase()) {
    for (var r2 = 1; r2 < values.length; r2++) {
      if (r2 !== rowIdx && String(values[r2][iName]).trim().toLowerCase() === newName.toLowerCase()) {
        throw new Error('Another item is already named "' + newName + '"');
      }
    }
  }
  headers.forEach(function(h, i){
    if (h === 'Item ID') { sh.getRange(rowIdx+1, i+1).setValue(newName); return; }
    if (data[h] !== undefined) sh.getRange(rowIdx+1, i+1).setValue(data[h]);
  });
  if (newName !== oldName) {
    cascadeItemRename_(ss, oldName, newName);
  }
  if (data['Item Type'] !== undefined) {
    cascadeItemCategorySync_(ss, newName, data['Item Type']);
  }
  logAction_(user, 'Item Updated', oldName + (newName !== oldName ? (' -> ' + newName) : ''));
  return 'OK';
}
function cascadeItemRename_(ss, oldName, newName) {
  ['Products_Stock','Customer_Order_Items','Supplier_Order_Items'].forEach(function(sheetName){
    var sh = ss.getSheetByName(sheetName);
    if (!sh) return;
    var values = sh.getDataRange().getValues();
    var headers = values[0];
    var iName = headers.indexOf('Item Name');
    var iId = headers.indexOf('Item ID');
    if (iName === -1) return;
    for (var r = 1; r < values.length; r++) {
      if (String(values[r][iName]).trim() === oldName) {
        sh.getRange(r+1, iName+1).setValue(newName);
        if (iId !== -1) sh.getRange(r+1, iId+1).setValue(newName);
      }
    }
  });
}
function cascadeItemCategorySync_(ss, itemName, itemType) {
  var sh = ss.getSheetByName('Products_Stock');
  if (!sh) return;
  var values = sh.getDataRange().getValues();
  var headers = values[0];
  var iName = headers.indexOf('Item Name');
  var iCat = headers.indexOf('Category');
  if (iName === -1 || iCat === -1) return;
  for (var r = 1; r < values.length; r++) {
    if (String(values[r][iName]).trim() === itemName && values[r][iCat] !== itemType) {
      sh.getRange(r+1, iCat+1).setValue(itemType);
    }
  }
}

function api_getCustomerNames() {
  var sh = ss_().getSheetByName('Customers');
  var data = sh.getDataRange().getValues();
  var idx = data[0].indexOf('Customer Name');
  var names = [];
  for (var r=1;r<data.length;r++){ if (data[r][idx]) names.push(data[r][idx]); }
  return names;
}
function api_getSupplierNames() {
  var sh = ss_().getSheetByName('Suppliers');
  var data = sh.getDataRange().getValues();
  var idx = data[0].indexOf('Supplier Name');
  var names = [];
  for (var r=1;r<data.length;r++){ if (data[r][idx]) names.push(data[r][idx]); }
  return names;
}

// ---- Orders ----
function api_getCustomerOrders() { return sheetToObjects_('Customer_Orders'); }
function api_getSupplierOrders() { return sheetToObjects_('Supplier_Orders'); }

function api_createCustomerOrder(header, items, user) {
  var ss = ss_();
  var ordersSh = ss.getSheetByName('Customer_Orders');
  var itemsSh = ss.getSheetByName('Customer_Order_Items');
  var headers1 = ordersSh.getRange(1,1,1,ordersSh.getLastColumn()).getValues()[0];
  var nextId = getNextId_(ordersSh, headers1, 'Order ID');
  var discount = toNum_(header.discount) || 0;
  var total = computeOrderTotal_(items) - discount;
  var newRow = [nextId, header.orderNumber||'', header.date, header.customer, 'Draft', total, header.notes||''];
  if (headers1.indexOf('Discount EGP') > -1) newRow.push(discount);
  ordersSh.appendRow(newRow);
  items.forEach(function(it){
    itemsSh.appendRow([nextId, it.itemName, it.length, it.color, it.qty, it.price, '', it.notes||'']);
  });
  logAction_(user, 'Customer Order Created', 'Order #' + nextId + ' for ' + header.customer);
  return nextId;
}
function api_createSupplierOrder(header, items, user) {
  var ss = ss_();
  var ordersSh = ss.getSheetByName('Supplier_Orders');
  var itemsSh = ss.getSheetByName('Supplier_Order_Items');
  var headers1 = ordersSh.getRange(1,1,1,ordersSh.getLastColumn()).getValues()[0];
  var nextId = getNextId_(ordersSh, headers1, 'Order ID');
  var total = computeOrderTotal_(items);
  ordersSh.appendRow([nextId, header.orderNumber||'', header.date, header.supplier, 'Draft', total, header.notes||'']);
  items.forEach(function(it){
    itemsSh.appendRow([nextId, it.itemName, it.length, it.color, it.qty, it.cost, '', it.notes||'']);
  });
  logAction_(user, 'Supplier Order Created', 'Order #' + nextId + ' for ' + header.supplier);
  return nextId;
}

function checkStockSufficiencyForOrder_(orderId) {
  var ss = ss_();
  var itemsSh = ss.getSheetByName('Customer_Order_Items');
  var idata = itemsSh.getDataRange().getValues();
  var iheaders = idata[0];
  var iOrder = iheaders.indexOf('Order ID');
  var iItem = iheaders.indexOf('Item Name');
  var iLen = iheaders.indexOf('Lenght');
  var iColor = iheaders.indexOf('Color');
  var iQty = iheaders.indexOf('Quantity');
  var iStockStatus = iheaders.indexOf('Stock Status');
  var stockSh = ss.getSheetByName('Products_Stock');
  var stockData = stockSh.getDataRange().getValues();
  var sHeaders = stockData[0];
  var sName = sHeaders.indexOf('Item Name'), sLen = sHeaders.indexOf('Lenght'), sColor = sHeaders.indexOf('Color'), sQty = sHeaders.indexOf('Opening Stock Qty');
  var shortages = [];
  for (var r = 1; r < idata.length; r++) {
    if (String(idata[r][iOrder]).trim() !== String(orderId).trim()) continue;
    if (idata[r][iStockStatus] === 'Processed') continue;
    var qty = toNum_(idata[r][iQty]);
    if (!qty) continue;
    var name = idata[r][iItem], len = idata[r][iLen], color = idata[r][iColor];
    var available = 0;
    for (var s = 1; s < stockData.length; s++) {
      if (String(stockData[s][sName]).trim() === String(name).trim() &&
          String(stockData[s][sLen]).trim() === String(len).trim() &&
          String(stockData[s][sColor]).trim() === String(color).trim()) {
        available = toNum_(stockData[s][sQty]);
        break;
      }
    }
    if (available < qty) {
      shortages.push({itemName: name, length: len, color: color, requested: qty, available: available});
    }
  }
  return shortages;
}

function api_updateOrderStatus(type, orderId, newStatus, user) {
  var sheetName = type === 'customer' ? 'Customer_Orders' : 'Supplier_Orders';
  if (type === 'customer' && newStatus === 'Delivered') {
    var shortages = checkStockSufficiencyForOrder_(orderId);
    if (shortages.length) {
      var lines = shortages.map(function(s){
        return '- ' + s.itemName + ' ' + s.length + ' ' + s.color + ': requested ' + s.requested + ', only ' + s.available + ' in stock';
      });
      throw new Error('Cannot mark as Delivered \u2014 not enough stock for:\n' + lines.join('\n'));
    }
  }
  var sh = ss_().getSheetByName(sheetName);
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var idIdx = headers.indexOf('Order ID');
  var statusIdx = headers.indexOf('Status');
  for (var r=1;r<data.length;r++){
    if (String(data[r][idIdx]) === String(orderId)) {
      sh.getRange(r+1, statusIdx+1).setValue(newStatus);
      break;
    }
  }
  var stockMsg = '', stockImportant = false, stockHasIssues = false;
  if (newStatus === 'Delivered' || newStatus === 'Received') {
    var stockResult = processNewOrders();
    stockMsg = stockResult.message;
    stockImportant = stockResult.important;
    stockHasIssues = stockResult.hasIssues;
  }
  logAction_(user, 'Order Status Changed', (type === 'customer' ? 'Customer' : 'Supplier') + ' order #' + orderId + ' -> ' + newStatus);
  return {status: 'OK', stockMsg: stockMsg, stockImportant: stockImportant, stockHasIssues: stockHasIssues};
}

// ---- Payments ----
function api_getCustomerPayments() { return sheetToObjects_('Customer_Payments'); }
function api_getSupplierPayments() { return sheetToObjects_('Supplier_Payments'); }
function api_addCustomerPayment(data, user) {
  var sh = ss_().getSheetByName('Customer_Payments');
  sh.appendRow([data.date, data.customer, data.amountEGP||'', data.amountUSD||'', data.rate||'', data.notes||'']);
  logAction_(user, 'Customer Payment Recorded', data.customer + ': ' + (data.amountEGP||0) + ' EGP / ' + (data.amountUSD||0) + ' USD');
  return 'OK';
}
function api_addSupplierPayment(data, user) {
  var sh = ss_().getSheetByName('Supplier_Payments');
  sh.appendRow([data.date, data.supplier, data.amountEGP||'', data.amountUSD||'', data.rate||'', data.notes||'']);
  logAction_(user, 'Supplier Payment Recorded', data.supplier + ': ' + (data.amountEGP||0) + ' EGP / ' + (data.amountUSD||0) + ' USD');
  return 'OK';
}


// ===================== USERS / AUTH =====================
function hashPw_(pw) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(pw), Utilities.Charset.UTF_8);
  var hex = '';
  for (var i=0;i<raw.length;i++){
    var b = raw[i];
    if (b < 0) b += 256;
    var h = b.toString(16);
    if (h.length === 1) h = '0' + h;
    hex += h;
  }
  return hex;
}

function toNum_(v) {
  if (typeof v === 'number') return v;
  if (v === null || v === undefined || v === '') return 0;
  var s = String(v).replace(/[^0-9.\-]/g, '');
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function computeOrderTotal_(items) {
  var t = 0;
  items.forEach(function(it){ t += toNum_(it.qty) * toNum_(it.price); });
  return t;
}

function getUserRole_(username) {
  var sh = ss_().getSheetByName('Users');
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var iUser = headers.indexOf('Username');
  var iRole = headers.indexOf('Role');
  for (var r=1;r<data.length;r++){
    if (String(data[r][iUser]).trim().toLowerCase() === String(username).trim().toLowerCase()) return data[r][iRole];
  }
  return null;
}

function requireAdmin_(username) {
  if (getUserRole_(username) !== 'Admin') throw new Error('This page is for Admin only');
}

function api_login(username, password) {
  var sh = ss_().getSheetByName('Users');
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var iUser = headers.indexOf('Username');
  var iPass = headers.indexOf('Password Hash');
  var iRole = headers.indexOf('Role');
  var iName = headers.indexOf('Full Name');
  var hash = hashPw_(password);
  for (var r=1;r<data.length;r++){
    if (String(data[r][iUser]).trim().toLowerCase() === String(username).trim().toLowerCase()) {
      if (String(data[r][iPass]) === hash) {
        logAction_(data[r][iUser], 'Login', 'Successful login');
        return {success:true, user:{username:data[r][iUser], role:data[r][iRole], fullName:data[r][iName]}};
      }
      return {success:false, message:'Incorrect password'};
    }
  }
  return {success:false, message:'Username not found'};
}

function api_getUsers(currentUser) {
  requireAdmin_(currentUser);
  var rows = sheetToObjects_('Users');
  rows.forEach(function(r){ delete r['Password Hash']; });
  return rows;
}

function api_addUser(currentUser, data) {
  requireAdmin_(currentUser);
  var sh = ss_().getSheetByName('Users');
  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var nextId = getNextId_(sh, headers, 'User ID');
  var pwHash = hashPw_(data.password);
  var row = headers.map(function(h){
    if (h === 'User ID') return nextId;
    if (h === 'Password Hash') return pwHash;
    return data[h] !== undefined ? data[h] : '';
  });
  sh.appendRow(row);
  logAction_(currentUser, 'User Added', 'New user: ' + data['Username'] + ' (' + data['Role'] + ')');
  return nextId;
}

function api_resetPassword(currentUser, targetUsername, newPassword) {
  requireAdmin_(currentUser);
  var sh = ss_().getSheetByName('Users');
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var iUser = headers.indexOf('Username');
  var iPass = headers.indexOf('Password Hash');
  for (var r=1;r<data.length;r++){
    if (String(data[r][iUser]).trim().toLowerCase() === String(targetUsername).trim().toLowerCase()) {
      sh.getRange(r+1, iPass+1).setValue(hashPw_(newPassword));
      logAction_(currentUser, 'Password Reset', 'Password changed: ' + targetUsername);
      return 'OK';
    }
  }
  throw new Error('User not found');
}

// ===================== LOGS =====================
function logAction_(username, action, details) {
  try {
    var sh = ss_().getSheetByName('Logs');
    if (!sh) return;
    sh.appendRow([new Date(), username||'', action, details||'']);
  } catch(e) {
    // never let logging break the main action
  }
}

function api_getLogs(currentUser) {
  requireAdmin_(currentUser);
  var sh = ss_().getSheetByName('Logs');
  var data = sh.getDataRange().getValues();
  var rows = [];
  for (var r=1;r<data.length;r++){
    var row = data[r];
    if (!row[0]) continue;
    var ts = row[0];
    var tsStr = ts instanceof Date ? Utilities.formatDate(ts, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : String(ts);
    rows.push({timestamp: tsStr, user: row[1], action: row[2], details: row[3]});
  }
  return rows.reverse();
}

// ===================== STOCK =====================
function api_updateStockCost(itemName, length, color, newCost, user) {
  var sh = ss_().getSheetByName('Products_Stock');
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var iName = headers.indexOf('Item Name');
  var iLen = headers.indexOf('Lenght');
  var iColor = headers.indexOf('Color');
  var iCost = headers.indexOf('Cost Price EGP');
  for (var r=1;r<data.length;r++){
    if (String(data[r][iName]).trim() === String(itemName).trim() &&
        String(data[r][iLen]).trim() === String(length).trim() &&
        String(data[r][iColor]).trim() === String(color).trim()) {
      sh.getRange(r+1, iCost+1).setValue(toNum_(newCost));
      logAction_(user, 'Stock Price Updated', itemName + ' ' + length + ' ' + color + ' -> ' + toNum_(newCost) + ' EGP');
      return 'OK';
    }
  }
  throw new Error('Stock item not found');
}

function ensureOverheadCategoriesSheet_() {
  var ss = ss_();
  var sh = ss.getSheetByName('Overhead_Categories');
  if (!sh) {
    sh = ss.insertSheet('Overhead_Categories');
    sh.getRange(1,1,1,2).setValues([['ID','Name']]);
    sh.getRange(1,1,1,2).setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    var defaults = ['Salaries','Rent','Stationery','Utilities','Transport','Maintenance','Other'];
    defaults.forEach(function(name, i){ sh.appendRow([i+1, name]); });
  }
  return sh;
}

function api_getOverheadCategories() {
  var sh = ensureOverheadCategoriesSheet_();
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var iId = headers.indexOf('ID'), iName = headers.indexOf('Name');
  var rows = [];
  for (var r=1;r<data.length;r++){
    if (data[r][iName]) rows.push({id: data[r][iId], name: data[r][iName]});
  }
  return rows;
}

function api_addOverheadCategory(name, user) {
  name = String(name||'').trim();
  if (!name) throw new Error('Category name is required');
  var sh = ensureOverheadCategoriesSheet_();
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var iName = headers.indexOf('Name');
  for (var r=1;r<data.length;r++){
    if (String(data[r][iName]).trim().toLowerCase() === name.toLowerCase()) throw new Error('Category already exists');
  }
  var nextId = getNextId_(sh, headers, 'ID');
  sh.appendRow([nextId, name]);
  logAction_(user, 'Overhead Category Added', name);
  return nextId;
}

function api_editOverheadCategory(id, newName, user) {
  newName = String(newName||'').trim();
  if (!newName) throw new Error('Category name is required');
  var sh = ensureOverheadCategoriesSheet_();
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var iId = headers.indexOf('ID'), iName = headers.indexOf('Name');
  var oldName = null, rowIdx = -1;
  for (var r=1;r<data.length;r++){
    if (String(data[r][iId]) === String(id)) { oldName = data[r][iName]; rowIdx = r; break; }
  }
  if (rowIdx === -1) throw new Error('Category not found');
  sh.getRange(rowIdx+1, iName+1).setValue(newName);
  if (oldName && oldName !== newName) {
    var ohSh = ensureOverheadSheet_();
    var ohData = ohSh.getDataRange().getValues();
    var ohHeaders = ohData[0];
    var ohCat = ohHeaders.indexOf('Category');
    for (var r2=1;r2<ohData.length;r2++){
      if (String(ohData[r2][ohCat]).trim() === String(oldName).trim()) {
        ohSh.getRange(r2+1, ohCat+1).setValue(newName);
      }
    }
  }
  logAction_(user, 'Overhead Category Renamed', oldName + ' -> ' + newName);
  return 'OK';
}

function api_deleteOverheadCategory(id, user) {
  requireAdmin_(user);
  var sh = ensureOverheadCategoriesSheet_();
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var iId = headers.indexOf('ID');
  for (var r=1;r<data.length;r++){
    if (String(data[r][iId]) === String(id)) {
      sh.deleteRow(r+1);
      logAction_(user, 'Overhead Category Deleted', 'Category #' + id);
      return 'OK';
    }
  }
  throw new Error('Category not found');
}

function migrateOverheadOwnerEmployeeColsRUN() {
  var sh = ensureOverheadSheet_();
  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  if (headers.indexOf('Owner') === -1) {
    var col = sh.getLastColumn() + 1;
    sh.getRange(1, col).setValue('Owner').setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
  }
  headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  if (headers.indexOf('Employee') === -1) {
    var col2 = sh.getLastColumn() + 1;
    sh.getRange(1, col2).setValue('Employee').setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
  }
  return 'Overhead headers now: ' + sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].join(', ');
}

function ensureOverheadSheet_() {
  var ss = ss_();
  var sh = ss.getSheetByName('Overhead');
  if (!sh) {
    sh = ss.insertSheet('Overhead');
    var headers = ['ID','Date','Category','Description','Amount EGP','Added By','Owner','Employee'];
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function api_getOverheads() {
  ensureOverheadSheet_();
  return sheetToObjects_('Overhead');
}

function api_addOverhead(data, user) {
  var sh = ensureOverheadSheet_();
  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var nextId = getNextId_(sh, headers, 'ID');
  var row = headers.map(function(h){
    if (h === 'ID') return nextId;
    if (h === 'Date') return data.date||'';
    if (h === 'Category') return data.category||'';
    if (h === 'Description') return data.description||'';
    if (h === 'Amount EGP') return toNum_(data.amount)||0;
    if (h === 'Added By') return user||'';
    if (h === 'Owner') return data.owner||'';
    if (h === 'Employee') return data.employee||'';
    return '';
  });
  sh.appendRow(row);
  logAction_(user, 'Overhead Expense Added', (data.category||'') + ' - ' + (data.description||'') + ': ' + (toNum_(data.amount)||0) + ' EGP');
  return nextId;
}

function api_deleteOverhead(id, user) {
  requireAdmin_(user);
  var sh = ensureOverheadSheet_();
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var idIdx = headers.indexOf('ID');
  for (var r=1;r<data.length;r++){
    if (String(data[r][idIdx]) === String(id)) {
      sh.deleteRow(r+1);
      logAction_(user, 'Overhead Expense Deleted', 'Expense #' + id);
      return 'OK';
    }
  }
  throw new Error('Expense not found');
}

// ---- Owners / Assets ----
function api_getOwnersExtraCards() {
  var overheadTotal = sumCol_('Overhead', 'Amount EGP');
  var customersOut = 0;
  computeBalances_('Customers', 'Customer Name', 'Customer_Orders', 'Customer', 'Customer_Payments', 'Customer').forEach(function(r){ customersOut += toNum_(r['Balance EGP']); });
  var suppliersOut = 0;
  computeBalances_('Suppliers', 'Supplier Name', 'Supplier_Orders', 'Supplier', 'Supplier_Payments', 'Supplier').forEach(function(r){ suppliersOut += toNum_(r['Balance EGP']); });
  return [
    {label: 'Total Overhead EGP', value: overheadTotal, currency: 'EGP'},
    {label: 'Customers Balance EGP', value: customersOut, currency: 'EGP', valueColor: 'good'},
    {label: 'Suppliers Balance EGP', value: suppliersOut, currency: 'EGP', valueColor: 'bad'}
  ];
}
function api_getOwnerNames() {
  var sh = ss_().getSheetByName('Owners');
  var data = sh.getDataRange().getValues();
  var idx = data[0].indexOf('Owner Name');
  var names = [];
  for (var r=1;r<data.length;r++){ if (data[r][idx]) names.push(data[r][idx]); }
  return names;
}
function api_getUsernames() {
  var sh = ss_().getSheetByName('Users');
  var data = sh.getDataRange().getValues();
  var idx = data[0].indexOf('Username');
  var names = [];
  for (var r=1;r<data.length;r++){ if (data[r][idx]) names.push(data[r][idx]); }
  return names;
}
function api_getOwners() {
  var rows = sheetToObjects_('Owners');
  var movements = sheetToObjects_('Capital_Movements');
  var contribEGP = {}, contribUSD = {};
  movements.forEach(function(m){
    var key = String(m['Owner']||'').trim();
    if (!key) return;
    contribEGP[key] = (contribEGP[key]||0) + toNum_(m['Amount EGP']);
    contribUSD[key] = (contribUSD[key]||0) + toNum_(m['Amount USD']);
  });
  var assets = sheetToObjects_('Assets');
  var assetEGP = {}, assetUSD = {};
  assets.forEach(function(a){
    var key = String(a['Owner']||'').trim();
    if (!key) return;
    assetEGP[key] = (assetEGP[key]||0) + toNum_(a['Value EGP']);
    assetUSD[key] = (assetUSD[key]||0) + toNum_(a['Value USD']);
  });
  var cheques = sheetToObjects_('Cheques');
  var chequesInHandEGP = {};
  cheques.forEach(function(ch){
    var key = String(ch['Owner']||'').trim();
    if (!key) return;
    if (String(ch['Status']||'').trim() !== 'In Hand') return;
    chequesInHandEGP[key] = (chequesInHandEGP[key]||0) + toNum_(ch['Amount EGP']);
  });
  rows.forEach(function(r){
    var key = String(r['Owner Name']||'').trim();
    r['Capital Contributed EGP'] = contribEGP[key] || 0;
    r['Capital Contributed USD'] = contribUSD[key] || 0;
    r['Assets Value EGP'] = assetEGP[key] || 0;
    r['Assets Value USD'] = assetUSD[key] || 0;
    r['Cheques In Hand EGP'] = chequesInHandEGP[key] || 0;
  });
  return rows;
}
function api_getAssets() { return sheetToObjects_('Assets'); }
function api_addAsset(data, user) {
  var sh = ss_().getSheetByName('Assets');
  var headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var nextId = getNextId_(sh, headers, 'Asset ID');
  var row = headers.map(function(h){
    if (h === 'Asset ID') return nextId;
    return data[h] !== undefined ? data[h] : '';
  });
  sh.appendRow(row);
  logAction_(user, 'Asset Added', (data['Asset Name']||'') + ' (' + (data['Owner']||'') + ')');
  return nextId;
}
function api_updateAsset(assetId, data, user) {
  var sh = ss_().getSheetByName('Assets');
  var sheetData = sh.getDataRange().getValues();
  var headers = sheetData[0];
  var iId = headers.indexOf('Asset ID');
  for (var r = 1; r < sheetData.length; r++) {
    if (String(sheetData[r][iId]) === String(assetId)) {
      headers.forEach(function(h, c){
        if (h === 'Asset ID') return;
        if (data[h] !== undefined) sh.getRange(r+1, c+1).setValue(data[h]);
      });
      logAction_(user, 'Asset Updated', 'Asset #' + assetId);
      return 'OK';
    }
  }
  throw new Error('Asset not found');
}
function api_deleteAsset(assetId, user) {
  requireAdmin_(user);
  var sh = ss_().getSheetByName('Assets');
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var iId = headers.indexOf('Asset ID');
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][iId]) === String(assetId)) {
      sh.deleteRow(r+1);
      logAction_(user, 'Asset Deleted', 'Asset #' + assetId);
      return 'OK';
    }
  }
  throw new Error('Asset not found');
}

// ---- Valuation ----
function ensureValuationSheet_() {
  var ss = ss_();
  var sh = ss.getSheetByName('Valuation_Snapshots');
  if (!sh) {
    sh = ss.insertSheet('Valuation_Snapshots');
    var headers = ['ID','Date','Valuation EGP','Notes','Added By'];
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  }
  return sh;
}
function computeCurrentValuationEGP_() {
  var assetsTotal = sumCol_('Assets', 'Value EGP');
  var ownersCapital = sumCol_('Owners', 'Opening Balance EGP') + sumCol_('Capital_Movements', 'Amount EGP');
  var stockValue = 0;
  try { stockValue = api_getStock().summary.totalValue || 0; } catch(e) { stockValue = 0; }
  var customersOut = 0;
  computeBalances_('Customers', 'Customer Name', 'Customer_Orders', 'Customer', 'Customer_Payments', 'Customer').forEach(function(r){ customersOut += toNum_(r['Balance EGP']); });
  var suppliersOut = 0;
  computeBalances_('Suppliers', 'Supplier Name', 'Supplier_Orders', 'Supplier', 'Supplier_Payments', 'Supplier').forEach(function(r){ suppliersOut += toNum_(r['Balance EGP']); });
  return assetsTotal + ownersCapital + stockValue + customersOut - suppliersOut;
}
function recordValuationSnapshot_(user) {
  var sh = ensureValuationSheet_();
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var iDate = headers.indexOf('Date');
  var iVal = headers.indexOf('Valuation EGP');
  var value = computeCurrentValuationEGP_();
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][iDate]) === today) {
      sh.getRange(r+1, iVal+1).setValue(value);
      return;
    }
  }
  var headers2 = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  var nextId = getNextId_(sh, headers2, 'ID');
  sh.appendRow([nextId, today, value, 'Auto', user||'System']);
}
function api_getValuationHistory(user) {
  recordValuationSnapshot_(user);
  var rows = sheetToObjects_('Valuation_Snapshots');
  rows.sort(function(a,b){ return String(a['Date']).localeCompare(String(b['Date'])); });
  return rows;
}
function api_addValuationPoint(data, user) {
  var sh = ensureValuationSheet_();
  var sheetData = sh.getDataRange().getValues();
  var headers = sheetData[0];
  var iDate = headers.indexOf('Date');
  var dateStr = data.date;
  for (var r = 1; r < sheetData.length; r++) {
    if (String(sheetData[r][iDate]) === String(dateStr)) {
      sh.getRange(r+1, headers.indexOf('Valuation EGP')+1).setValue(toNum_(data.value));
      sh.getRange(r+1, headers.indexOf('Notes')+1).setValue(data.notes||'');
      sh.getRange(r+1, headers.indexOf('Added By')+1).setValue(user||'');
      logAction_(user, 'Valuation Point Updated', dateStr + ': ' + toNum_(data.value) + ' EGP');
      return 'OK';
    }
  }
  var nextId = getNextId_(sh, headers, 'ID');
  sh.appendRow([nextId, dateStr, toNum_(data.value), data.notes||'', user||'']);
  logAction_(user, 'Valuation Point Added', dateStr + ': ' + toNum_(data.value) + ' EGP');
  return 'OK';
}
function api_deleteValuationPoint(id, user) {
  requireAdmin_(user);
  var sh = ensureValuationSheet_();
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var iId = headers.indexOf('ID');
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][iId]) === String(id)) {
      sh.deleteRow(r+1);
      logAction_(user, 'Valuation Point Deleted', 'Point #' + id);
      return 'OK';
    }
  }
  throw new Error('Point not found');
}

function api_getStock() {
  var sh = ss_().getSheetByName('Products_Stock');
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var iName = headers.indexOf('Item Name');
  var iLen = headers.indexOf('Lenght');
  var iColor = headers.indexOf('Color');
  var iCat = headers.indexOf('Category');
  var iUnit = headers.indexOf('Unit');
  var iQty = headers.indexOf('Opening Stock Qty');
  var iCost = headers.indexOf('Cost Price EGP');
  var iReorder = headers.indexOf('Reorder Level');
  var rows = [];
  var totalValue = 0, lowCount = 0, outCount = 0;
  for (var r=1;r<data.length;r++){
    var row = data[r];
    if (!row[iName]) continue;
    var qty = toNum_(row[iQty]);
    var cost = toNum_(row[iCost]);
    var reorder = iReorder > -1 ? Number(row[iReorder]) : NaN;
    if (isNaN(reorder)) reorder = 5;
    var low = qty <= reorder;
    if (low) lowCount++;
    if (qty <= 0) outCount++;
    totalValue += qty*cost;
    rows.push({
      itemName: row[iName], length: row[iLen], color: row[iColor], category: row[iCat],
      unit: row[iUnit], qty: qty, cost: cost, reorderLevel: reorder, low: low
    });
  }
  return {
    summary: { totalItems: rows.length, totalValue: totalValue, lowCount: lowCount, outCount: outCount },
    rows: rows
  };
}

// ===================== SETUP: Users / Logs / Stock reorder level =====================
function setupUsersLogsStock() {
  var ss = ss_();
  var result = {};

  function ensureSheet(name, headers) {
    var sh = ss.getSheetByName(name);
    var created = false;
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1,1,1,headers.length).setValues([headers]);
      sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
      sh.setFrozenRows(1);
      created = true;
    }
    return {sheet: sh, created: created};
  }

  var users = ensureSheet('Users', ['User ID','Username','Password Hash','Role','Full Name','Notes']);
  result.usersCreated = users.created;
  if (users.created) {
    var roleRule = SpreadsheetApp.newDataValidation().requireValueInList(['Admin','Staff'], true).setAllowInvalid(true).build();
    users.sheet.getRange(2,4,300,1).setDataValidation(roleRule);
    users.sheet.appendRow([1, 'admin', hashPw_('Future@2026'), 'Admin', 'Gemmy', 'Default admin account - change the password from the Users page after first login']);
  }

  var logs = ensureSheet('Logs', ['Timestamp','User','Action','Details']);
  result.logsCreated = logs.created;

  var stock = ss.getSheetByName('Products_Stock');
  var stockHeaders = stock.getRange(1,1,1,stock.getLastColumn()).getValues()[0];
  if (stockHeaders.indexOf('Reorder Level') === -1) {
    var newCol = stock.getLastColumn() + 1;
    stock.getRange(1, newCol).setValue('Reorder Level').setFontWeight('bold').setBackground('#4a86e8').setFontColor('#ffffff');
    var lastRow = stock.getLastRow();
    if (lastRow > 1) {
      var defaults = [];
      for (var i=2;i<=lastRow;i++) defaults.push([5]);
      stock.getRange(2, newCol, defaults.length, 1).setValues(defaults);
    }
    result.reorderLevelAdded = true;
  } else {
    result.reorderLevelAdded = false;
  }

  var desiredOrder = ['README','Dashboard','Customers','Customer_Payments','Customer_Orders','Customer_Order_Items','Suppliers','Supplier_Payments','Supplier_Orders','Supplier_Order_Items','Owners','Users','Cheques','Products_Stock','Items','Assets','Capital_Movements','Services','Service_Income','Logs'];
  desiredOrder.forEach(function(name, idx){
    var sh = ss.getSheetByName(name);
    if (sh) {
      ss.setActiveSheet(sh);
      ss.moveActiveSheet(idx + 1);
    }
  });
  result.finalSheetOrder = ss.getSheets().map(function(s){return s.getName();});

  return JSON.stringify(result);
}


function api_getOrderDetails(type, orderId) {
  var sheetName = type === 'customer' ? 'Customer_Orders' : 'Supplier_Orders';
  var itemsSheetName = type === 'customer' ? 'Customer_Order_Items' : 'Supplier_Order_Items';
  var priceCol = type === 'customer' ? 'Unit Price EGP' : 'Unit Cost EGP';
  var partyCol = type === 'customer' ? 'Customer' : 'Supplier';
  var ss = ss_();
  var sh = ss.getSheetByName(sheetName);
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var idIdx = headers.indexOf('Order ID');
  var header = null;
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][idIdx]) === String(orderId)) {
      header = {
        date: data[r][headers.indexOf('Date')],
        party: data[r][headers.indexOf(partyCol)],
        notes: data[r][headers.indexOf('Notes')],
        status: data[r][headers.indexOf('Status')],
        orderNumber: data[r][headers.indexOf('Order Number')],
        discount: (type === 'customer' && headers.indexOf('Discount EGP') > -1) ? toNum_(data[r][headers.indexOf('Discount EGP')]) : 0
      };
      break;
    }
  }
  if (!header) throw new Error('Order not found');
  if (header.date instanceof Date) {
    header.date = Utilities.formatDate(header.date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var itemsSh = ss.getSheetByName(itemsSheetName);
  var idata = itemsSh.getDataRange().getValues();
  var iheaders = idata[0];
  var iOrder = iheaders.indexOf('Order ID');
  var iItem = iheaders.indexOf('Item Name');
  var iLen = iheaders.indexOf('Lenght');
  var iColor = iheaders.indexOf('Color');
  var iQty = iheaders.indexOf('Quantity');
  var iPrice = iheaders.indexOf(priceCol);
  var iStockStatus2 = iheaders.indexOf('Stock Status');
  var items = [];
  for (var r2 = 1; r2 < idata.length; r2++) {
    if (String(idata[r2][iOrder]) === String(orderId)) {
      items.push({
        itemName: idata[r2][iItem],
        length: idata[r2][iLen],
        color: idata[r2][iColor],
        qty: idata[r2][iQty],
        price: idata[r2][iPrice],
        stockStatus: idata[r2][iStockStatus2]
      });
    }
  }
  return {header: header, items: items};
}

function api_updateOrderDetails(type, orderId, header, items, user) {
  requireAdmin_(user);
  var sheetName = type === 'customer' ? 'Customer_Orders' : 'Supplier_Orders';
  var itemsSheetName = type === 'customer' ? 'Customer_Order_Items' : 'Supplier_Order_Items';
  var priceCol = type === 'customer' ? 'Unit Price EGP' : 'Unit Cost EGP';
  var partyCol = type === 'customer' ? 'Customer' : 'Supplier';
  var readyStatus = type === 'customer' ? 'Delivered' : 'Received';
  var stockSign = type === 'customer' ? -1 : 1;
  var ss = ss_();
  var sh = ss.getSheetByName(sheetName);
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var idIdx = headers.indexOf('Order ID');
  var rowIdx = -1, currentStatus = '';
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][idIdx]) === String(orderId)) { rowIdx = r; currentStatus = data[r][headers.indexOf('Status')]; break; }
  }
  if (rowIdx === -1) throw new Error('Order not found');

  var itemsSh = ss.getSheetByName(itemsSheetName);
  var idata = itemsSh.getDataRange().getValues();
  var iheaders = idata[0];
  var iOrder = iheaders.indexOf('Order ID');
  var iItem = iheaders.indexOf('Item Name');
  var iLen = iheaders.indexOf('Lenght');
  var iColor = iheaders.indexOf('Color');
  var iQty = iheaders.indexOf('Quantity');
  var iStockStatus = iheaders.indexOf('Stock Status');

  var stockSh = ss.getSheetByName('Products_Stock');
  var stockData = stockSh.getDataRange().getValues();
  var sHeaders = stockData[0];
  var sName = sHeaders.indexOf('Item Name'), sLen = sHeaders.indexOf('Lenght'), sColor = sHeaders.indexOf('Color'), sQty = sHeaders.indexOf('Opening Stock Qty');
  function findStockRow_(name, len, color) {
    for (var rr = 1; rr < stockData.length; rr++) {
      if (String(stockData[rr][sName]).trim() === String(name).trim() &&
          String(stockData[rr][sLen]).trim() === String(len).trim() &&
          String(stockData[rr][sColor]).trim() === String(color).trim()) return rr;
    }
    return -1;
  }

  var rowsToDelete = [];
  for (var r2 = 1; r2 < idata.length; r2++) {
    if (String(idata[r2][iOrder]) === String(orderId)) {
      if (idata[r2][iStockStatus] === 'Processed') {
        var sIdx = findStockRow_(idata[r2][iItem], idata[r2][iLen], idata[r2][iColor]);
        if (sIdx !== -1) {
          var qtyOld = toNum_(idata[r2][iQty]);
          var cur = toNum_(stockData[sIdx][sQty]);
          var reversed = cur - stockSign * qtyOld;
          stockSh.getRange(sIdx + 1, sQty + 1).setValue(reversed);
          stockData[sIdx][sQty] = reversed;
        }
      }
      rowsToDelete.push(r2 + 1);
    }
  }
  rowsToDelete.sort(function (a, b) { return b - a; });
  rowsToDelete.forEach(function (rowNum) { itemsSh.deleteRow(rowNum); });

  items.forEach(function (it) {
    itemsSh.appendRow([orderId, it.itemName, it.length, it.color, it.qty, it.price, '', it.notes || '']);
  });

  sh.getRange(rowIdx + 1, headers.indexOf('Date') + 1).setValue(header.date);
  sh.getRange(rowIdx + 1, headers.indexOf(partyCol) + 1).setValue(header.party);
  sh.getRange(rowIdx + 1, headers.indexOf('Notes') + 1).setValue(header.notes || '');
  var discount = 0;
  if (type === 'customer' && headers.indexOf('Discount EGP') > -1) {
    discount = toNum_(header.discount) || 0;
    sh.getRange(rowIdx + 1, headers.indexOf('Discount EGP') + 1).setValue(discount);
  }
  sh.getRange(rowIdx + 1, headers.indexOf('Total EGP') + 1).setValue(computeOrderTotal_(items) - discount);
  if (currentStatus !== readyStatus) {
    sh.getRange(rowIdx + 1, headers.indexOf('Order Number') + 1).setValue(header.orderNumber || '');
  }

  var stockMsg = '', stockImportant = false, stockHasIssues = false;
  if (currentStatus === readyStatus) {
    var stockResult = processNewOrders();
    stockMsg = stockResult.message;
    stockImportant = stockResult.important;
    stockHasIssues = stockResult.hasIssues;
  }

  logAction_(user, 'Order Edited', (type === 'customer' ? 'Customer' : 'Supplier') + ' order #' + orderId);
  return {status: 'OK', stockMsg: stockMsg, stockImportant: stockImportant, stockHasIssues: stockHasIssues};
}

function api_deleteOrder(type, orderId, user) {
  requireAdmin_(user);
  var sheetName = type === 'customer' ? 'Customer_Orders' : 'Supplier_Orders';
  var itemsSheetName = type === 'customer' ? 'Customer_Order_Items' : 'Supplier_Order_Items';
  var stockSign = type === 'customer' ? -1 : 1;
  var ss = ss_();
  var sh = ss.getSheetByName(sheetName);
  var data = sh.getDataRange().getValues();
  var headers = data[0];
  var idIdx = headers.indexOf('Order ID');
  var rowIdx = -1;
  for (var r = 1; r < data.length; r++) {
    if (String(data[r][idIdx]) === String(orderId)) { rowIdx = r; break; }
  }
  if (rowIdx === -1) throw new Error('Order not found');

  var itemsSh = ss.getSheetByName(itemsSheetName);
  var idata = itemsSh.getDataRange().getValues();
  var iheaders = idata[0];
  var iOrder = iheaders.indexOf('Order ID');
  var iItem = iheaders.indexOf('Item Name');
  var iLen = iheaders.indexOf('Lenght');
  var iColor = iheaders.indexOf('Color');
  var iQty = iheaders.indexOf('Quantity');
  var iStockStatus = iheaders.indexOf('Stock Status');

  var stockSh = ss.getSheetByName('Products_Stock');
  var stockData = stockSh.getDataRange().getValues();
  var sHeaders = stockData[0];
  var sName = sHeaders.indexOf('Item Name'), sLen = sHeaders.indexOf('Lenght'), sColor = sHeaders.indexOf('Color'), sQty = sHeaders.indexOf('Opening Stock Qty');
  function findStockRowDel_(name, len, color) {
    for (var rr = 1; rr < stockData.length; rr++) {
      if (String(stockData[rr][sName]).trim() === String(name).trim() &&
          String(stockData[rr][sLen]).trim() === String(len).trim() &&
          String(stockData[rr][sColor]).trim() === String(color).trim()) return rr;
    }
    return -1;
  }

  var rowsToDelete = [];
  for (var r2 = 1; r2 < idata.length; r2++) {
    if (String(idata[r2][iOrder]) === String(orderId)) {
      if (idata[r2][iStockStatus] === 'Processed') {
        var sIdx = findStockRowDel_(idata[r2][iItem], idata[r2][iLen], idata[r2][iColor]);
        if (sIdx !== -1) {
          var qtyOld = toNum_(idata[r2][iQty]);
          var cur = toNum_(stockData[sIdx][sQty]);
          var reversed = cur - stockSign * qtyOld;
          stockSh.getRange(sIdx + 1, sQty + 1).setValue(reversed);
          stockData[sIdx][sQty] = reversed;
        }
      }
      rowsToDelete.push(r2 + 1);
    }
  }
  rowsToDelete.sort(function (a, b) { return b - a; });
  rowsToDelete.forEach(function (rowNum) { itemsSh.deleteRow(rowNum); });

  sh.deleteRow(rowIdx + 1);

  logAction_(user, 'Order Deleted', (type === 'customer' ? 'Customer' : 'Supplier') + ' order #' + orderId);
  return {status: 'OK'};
}
