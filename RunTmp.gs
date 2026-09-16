function checkHeaders_RUN() {
  var ss = ss_();
  ['Customer_Orders','Supplier_Orders','Customer_Order_Items','Supplier_Order_Items','Products_Stock'].forEach(function(name){
    var sh = ss.getSheetByName(name);
    if (sh) console.log(name + ': ' + JSON.stringify(sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]));
  });
}
