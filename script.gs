function findContainingTable(element) {
  if (element.getType() == 'TABLE') {
    return element;
  }
  var parent = element.getParent()
  if (parent) {
    Logger.log(parent);
    return findContainingTable(parent);
  }
}

function findContainingTableRow(element) {
  if (element.getType() == 'TABLE_ROW') {
    return element;
  }
  var parent = element.getParent()
  if (parent) {
    Logger.log(parent.getType());
    return findContainingTableRow(parent);
  }
}

function findPlaceholder(element, placeholder) {
  if (element.getNumChildren !== undefined) {
    for (var i=0;i<element.getNumChildren();i++) {
      var child = element.getChild(i);
      //Logger.log(child.getType());
      if (child.getType() == 'PARAGRAPH') {
        //Logger.log(child.getText());
        if (child.getText().indexOf(placeholder) > -1) {
          return child;
        }
      }
      var res = findPlaceholder(child, placeholder);
      if (res) {
        return res;
      }
    }
  }
  return null;
}

function to2decimal(num) {
  return Math.round(num * 100) / 100;
}

function num2str(value) {
   var decimals = 2;
   return value.toFixed(decimals);
}

function insertData(documentId, invoiceNumber, date, noVat, client, lines, currency, paymentDays) {
  var body = DocumentApp.openById(documentId).getBody();
  body.replaceText('#{N}', invoiceNumber);
  body.replaceText('#{DD}', date.day);
  body.replaceText('#{MM}', date.month);
  body.replaceText('#{YY}', date.year);

  for (clientKey in client) {
      var capitalizedKey = clientKey[0].toUpperCase() + clientKey.substring(1);
      body.replaceText('#{client' + capitalizedKey + '}', client[clientKey]);
  }

  Logger.log("START lines="+lines);
  var placeholder = findPlaceholder(body, '#{lineDescription}');
  Logger.log("res="+placeholder);
  var table = findContainingTable(placeholder);
  Logger.log("table="+table);
  var totalAmount = 0.0;
  var totalVAT = 0.0;
  var totalTotal = 0.0;
  for (var i=lines.length;i>0;i--) {
    var tableRow = findContainingTableRow(placeholder);
    if (i!=1) {
      Logger.log("inserting at "+(lines.length-i+1));
      tableRow = table.insertTableRow(lines.length-i+1, tableRow.copy());
    }
    var line = lines[lines.length - i];
    tableRow.replaceText('#{lineDescription}', line.description);
    tableRow.replaceText('#{lineAmount}', num2str(line.amount));
    var vat = to2decimal(line.amount * (line.vatRate/100.0));
    tableRow.replaceText('#{lineVAT}', num2str(vat));
    tableRow.replaceText('#{lineTotal}', num2str(line.amount + vat));

    totalAmount += line.amount;
    totalVAT += vat;
    totalTotal += line.amount + vat;
  }
  body.replaceText('#{totalAmount}', num2str(to2decimal(totalAmount)));
  body.replaceText('#{totalVAT}', num2str(to2decimal(totalVAT)));
  body.replaceText('#{totalTotal}', num2str(to2decimal(totalTotal)));

  body.replaceText('#{currency}', currency || 'Euro');
  body.replaceText('#{paymentDays}', paymentDays || '15');

  if (!noVat) {
    var par = findPlaceholder(body, 'Value added tax levied');
    par.clear();
  }
}

function main() {
  insertData('1QnLMAXiSQg8ut_RiK66BDJKke_N6AD8A38e2IIeehTE', 11,
             {'day':11, 'month':'April', 'year':2016},
             false,
             {'name':'Pinco', 'address':'Via dei pinchi palli','vatID':'FOO1234','contact':'Mr. Pallo'},
             [{'description':'Stuff done', 'amount':128.34, 'vatRate':20.0},
              {'description':'Other Stuff', 'amount':80.0, 'vatRate':20.0},
              {'description':'Third line', 'amount':85.0, 'vatRate':20.0}]);
}