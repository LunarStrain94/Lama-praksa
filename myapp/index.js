const express = require('express');
const app = express();
const port = 3000;
const data = require('./data.json');
const fs = require('fs');
const file = "./data.json";
const exceljs = require('exceljs');
app.use(express.json());
const letters = {
  0 : 'A',
  1 : 'B',
  2 : 'C',
  3 : 'D',
  4 : 'E',
  5 : 'F',
  6 : 'G',
  7 : 'H',
  8 : 'I',
  9 : 'J',
  10 : 'K',
  11 : 'L',
  12 : 'M',
  13 : 'N'
};

async function write(workbook, filename) {
  await workbook.xlsx.writeFile(filename);
};

app.get('/excel', (req, res) => {
  const workbook = new exceljs.Workbook();
  workbook.creator = 'Me';
  workbook.lastModifiedBy = 'Her';
  workbook.created = new Date(1985, 8, 30);
  workbook.modified = new Date();
  workbook.properties.date1904 = true;
  workbook.calcProperties.fullCalcOnLoad = true;
  workbook.views = [{
      x: 0, y: 0, width: 10000, height: 20000,

      firstSheet: 0, activeTab: 1, visibility: 'visible'
    }]
  const sheet = workbook.addWorksheet('My Sheet');

  var ind = 2;
  var nastavnici = [];
  var statusi = [];
  var predavanja = [];
  var seminari = [];
  var vjezbe = [];
  var normaPredavanja = [];
  var normaSeminari = [];
  var normaVjezbe = [];

  var databook = new exceljs.Workbook();
  databook.xlsx.readFile("data.xlsx")
    .then(function(ind, nastavnici, statusi, predavanja, seminari, vjezbe, normaPredavanja, normaSeminari, normaVjezbe) {
      var datasheet = databook.getWorksheet('List1');

      while (true) { // get data
        var nas = "D" + ind.toString();
        var stat = "E" + ind.toString();
        var pred = "N" + ind.toString();
        var sem = "O" + ind.toString();
        var vj = "P" + ind.toString();
        var nPred = "" + ind.toString();
        var nSem = "O" + ind.toString();
        var nVj = "P" + ind.toString();
        if (datasheet.getCell(nas).value == "") {
          ind -= 2;
          break;
        }
        nastavnici.push(datasheet.getCell(nas).value);
        statusi.push(datasheet.getCell(stat).value);
        predavanja.push(datasheet.getCell(pred).value);
        seminari.push(datasheet.getCell(sem).value);
        vjezbe.push(datasheet.getCell(vj).value);
        normaPredavanja.push(datasheet.getCell(nPred).value);
        normaSeminari.push(datasheet.getCell(nSem).value);
        normaVjezbe.push(datasheet.getCell(nVj).value);
        ind++;
      }
        
    });

  var len = ind;
  var ind = 0;
  var cells = 'A' + (18 + len).toString() + ":C" + (18 + len).toString();
  sheet.mergeCells(cells);
  for (i = 0; i < 14; i++) { // A17:N25 border + first row numbers + data
    for (j = 17; j < 19 + len; j++) {
      if (j == (18 + len)) {
        var str = letters[i] + j.toString();
        sheet.getCell(str).border = { top: {style:'medium'}, left: {style:'medium'}, bottom: {style:'medium'}, right: {style:'medium'}};
      }
      else if (i < 13) {
        var str = letters[i] + j.toString();
        sheet.getCell(str).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'}};
        if (str[0] == 'B' && ind <= len) {
          sheet.getCell(str).value = nastavnici[ind];
        }
        else if (str[0] == 'D' && ind <= len) {
          sheet.getCell(str).value = statusi[ind];
        }
        else if (str[0] == 'E' && ind <= len) {
          sheet.getCell(str).value = predavanja[ind];
        }
        else if (str[0] == 'F' && ind <= len) {
          sheet.getCell(str).value = seminari[ind];
        }
        else if (str[0] == 'G' && ind <= len) {
          sheet.getCell(str).value = vjezbe[ind];
        }
      }
      else if (i == 13) {
        var str = letters[i] + j.toString();
        sheet.getCell(str).border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'medium'}};
      }
      if (i == 0 && j < 18 + len) {
        var str = letters[i] + j.toString();
        sheet.getCell(str).value = j - 16;
        sheet.getCell(str).alignment = { horizontal: 'center', wrapText: true };
      }
    }
    ind++;
  }

  sheet.getColumn('A').width = 7;
  sheet.getColumn('B').width = 18.5;
  sheet.getColumn('C').width = 22;
  sheet.getColumn('D').width = 22;
  sheet.getColumn('E').width = 7;
  sheet.getColumn('F').width = 8.5;
  sheet.getColumn('G').width = 7.5;
  sheet.getColumn('H').width = 11.5;
  sheet.getColumn('I').width = 11.5;
  sheet.getColumn('J').width = 11;
  sheet.getRow('12').height = 65;
  sheet.getRow('16').height = 90;
  
  sheet.mergeCells('A6:I7');
  sheet.mergeCells('A8:I11');
  sheet.mergeCells('A12:B12');
  sheet.mergeCells('H12:I12');
  sheet.mergeCells('H13:I13');
  sheet.mergeCells('A13:B13');
  sheet.mergeCells('A15:A16');
  sheet.mergeCells('B15:B16');
  sheet.mergeCells('C15:C16');
  sheet.mergeCells('D15:D16');
  sheet.mergeCells('E15:G15');
  sheet.mergeCells('H15:H16');
  sheet.mergeCells('I15:I16');
  sheet.mergeCells('J15:J16');
  sheet.mergeCells('K15:M15');
  sheet.mergeCells('N15:N16');

  const image1 = workbook.addImage({
    filename: 'Logo.png',
    extension: 'png',
  });
  sheet.addImage(image1, 'A1:B4');

  for (i = 0; i < 8; i++) { // A12:H12 color and border
    var str = letters[i] + "12";
    sheet.getCell(str).fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FFE5E5FF'}};
    sheet.getCell(str).border = { top: {style:'medium'}, left: {style:'medium'}, bottom: {style:'medium'}, right: {style:'medium'}};
    sheet.getCell(str).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  }
  for (i = 0; i < 14; i++) { // A15:N15 color and border
    var str = letters[i] + "15";
    sheet.getCell(str).fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FFE5E5FF'}};
    sheet.getCell(str).border = { top: {style:'medium'}, left: {style:'medium'}, bottom: {style:'medium'}, right: {style:'medium'}};
    sheet.getCell(str).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  }
  for (i = 4; i < 13; i++) { // E16:M16 color and border
    var str = letters[i] + "16";
    sheet.getCell(str).fill = { type: 'pattern', pattern:'solid', fgColor:{argb:'FFE5E5FF'}};
    sheet.getCell(str).border = { top: {style:'medium'}, left: {style:'medium'}, bottom: {style:'medium'}, right: {style:'medium'}};
    sheet.getCell(str).alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  }
  for (i = 0; i < 7; i++) { // A13:H13 border
    var str = letters[i] + "13";
    sheet.getCell(str).border = { top: {style:'medium'}, left: {style:'thin'}, bottom: {style:'medium'}, right: {style:'thin'}};
  }

  sheet.getCell('A13').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  sheet.getCell('H13').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  sheet.getCell('H13').border = { top: {style:'medium'}, left: {style:'thin'}, bottom: {style:'medium'}, right: {style:'medium'}};

  sheet.getCell('A5').value = {
    'richText': [
      {'font': {'size': 11, 'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'Predmet: '},
      {'font': {'size': 11,'color': {'argb': 'FFFF0000'},'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'naziv predmeta sa sifrom'}
    ]
  };
  sheet.getCell('A6').value = { 'richText': [ {'font': {'bold': true,'size': 14,'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'NALOG ZA ISPLATU'}]};
  sheet.getCell('A6').alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
  sheet.getCell('A6').border = { bottom: {style:'medium', color: {argb:'FFFFFFFF'}}};
  sheet.getCell('A8').value = 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.'
  sheet.getCell('A8').alignment = { vertical: 'top', wrapText: true };
  sheet.getCell('A12').value = { 'richText': [{'font': {'bold': true,'size': 11,'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'Katedra'}]};
  sheet.getCell('C12').value = { 'richText': [{'font': {'bold': true,'size': 11,'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'Studij'}]};
  sheet.getCell('D12').value = { 'richText': [{'font': {'bold': true,'size': 11,'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'ak. god.'}]};
  sheet.getCell('E12').value = { 'richText': [{'font': {'bold': true,'size': 11,'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'stud. god.'}]};
  sheet.getCell('F12').value = { 'richText': [{'font': {'bold': true,'size': 11,'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'početak turnusa'}]};
  sheet.getCell('G12').value = { 'richText': [{'font': {'bold': true,'size': 11,'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'kraj turnusa'}]};
  sheet.getCell('H12').value = { 'richText': [{'font': {'bold': true,'size': 11,'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'br sati predviđen programom'}]};
  sheet.getCell('A13').value = { 'richText': [{'font': {'size': 11,'color': {'argb': 'FFFF0000'},'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'Katedra'}]};
  sheet.getCell('C13').value = { 'richText': [{'font': {'size': 11,'color': {'argb': 'FFFF0000'},'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'Studiji'}]};
  sheet.getCell('D13').value = { 'richText': [{'font': {'size': 11,'color': {'argb': 'FFFF0000'},'name': 'Calibri','family': 2,'scheme': 'minor'},'text': '2022./23.'}]};
  sheet.getCell('E13').value = { 'richText': [{'font': {'size': 11,'color': {'argb': 'FFFF0000'},'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'npr. 1'}]};
  sheet.getCell('F13').value = { 'richText': [{'font': {'size': 11,'color': {'argb': 'FFFF0000'},'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'datum'}]};
  sheet.getCell('G13').value = { 'richText': [{'font': {'size': 11,'color': {'argb': 'FFFF0000'},'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'datum'}]};
  sheet.getCell('H13').value = {
    'richText': [
      {'font': {'size': 11,'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'P: '},
      {'font': {'size': 11,'color': {'argb': 'FFFF0000'}, 'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'XY '},
      {'font': {'size': 11,'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'S: '},
      {'font': {'size': 11,'color': {'argb': 'FFFF0000'}, 'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'XY '},
      {'font': {'size': 11,'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'V: '},
      {'font': {'size': 11,'color': {'argb': 'FFFF0000'},'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'XY '}
    ]
  };
  sheet.getCell('A15').value = { 'richText': [{'font': {'bold': true,'size': 11,'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'Redni broj'}]};
  sheet.getCell('B15').value = { 'richText': [{'font': {'bold': true,'size': 11,'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'Nastavnik/Suradnik'}]};
  sheet.getCell('C15').value = { 'richText': [{'font': {'bold': true,'size': 11,'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'Zvanje'}]};
  sheet.getCell('D15').value = { 'richText': [{'font': {'bold': true,'size': 11,'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'Status'}]};
  sheet.getCell('E15').value = { 'richText': [{'font': {'bold': true,'size': 11,'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'Sati nastave'}]};
  sheet.getCell('H15').value = { 'richText': [{'font': {'bold': true,'size': 11,'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'Bruto satnica predavanja (EUR)'}]};
  sheet.getCell('I15').value = { 'richText': [{'font': {'bold': true,'size': 11,'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'Bruto satnica seminari (EUR)'}]};
  sheet.getCell('J15').value = { 'richText': [{'font': {'bold': true,'size': 11,'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'Bruto satnica vježbe (EUR)'}]};
  sheet.getCell('K15').value = { 'richText': [{'font': {'bold': true,'size': 11,'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'Bruto iznos'}]};
  sheet.getCell('N15').value = { 'richText': [{'font': {'bold': true,'size': 11,'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'Ukupno za isplatu (EUR)'}]};
  sheet.getCell('E16').value = { 'richText': [{'font': {'bold': true,'size': 11,'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'pred'}]};
  sheet.getCell('F16').value = { 'richText': [{'font': {'bold': true,'size': 11,'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'sem'}]};
  sheet.getCell('G16').value = { 'richText': [{'font': {'bold': true,'size': 11,'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'vjež'}]};
  sheet.getCell('K16').value = { 'richText': [{'font': {'bold': true,'size': 11,'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'pred'}]};
  sheet.getCell('L16').value = { 'richText': [{'font': {'bold': true,'size': 11,'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'sem'}]};
  sheet.getCell('M16').value = { 'richText': [{'font': {'bold': true,'size': 11,'name': 'Calibri','family': 2,'scheme': 'minor'},'text': 'vjež'}]};
  write(workbook, "Sheet.xlsx");
  res.send("Placeholder");
})

app.get('/userID/:userId', (req, res) => {
  res.send(data["users"][req.params.userId-1]);
})


app.get('/postID/:postId', (req, res) => {
  res.send(data["posts"][req.params.postId-1]);
})


app.get('/date/:fromDate/:toDate', (req, res) => {
  let d1 = new Date(req.params.fromDate);
  let d2 = new Date(req.params.toDate);
  let dates = data["posts"];
  let list = [];
  dates.forEach(element => {
    let date = element["last_update"];
    date = date.replace(" ", "T");
    date += "Z";
    let d3 = new Date(date);
    if (d3 > d1 && d3 < d2) {
      list.push(element);
    }
  })
  let e = d1 < d2;
  res.send(list);
})


app.put('/user', (req, res) => {
  let users = data["users"];
  let posts = data["posts"];
  let changed = false;
  var { email, user_id } = req.body;
  users.forEach(element => {
    if (element["id"] == user_id) {
      element["email"] = email;
      changed = true;
    }
  })

  if (changed) {
    let content = JSON.stringify({users, posts}, undefined, 2);
    fs.writeFile(file, content, err => {
      if (err) {
        res.send(err);
      }
      res.send("Email updated successfully");
    });
  }
  else {
    res.status(404).send('User not found');
  }
});


app.post('/post', (req, res) => {
  let posts = data["posts"];
  let users = data["users"];
  let post = {"id" : (posts.length + 1)};
  let {user_id, title, body} = req.body;
  post["user_id"] = user_id;
  post["title"] = title;
  post["body"] = body;
  post["last_update"] = new Date();
  posts.push(post);

  let content = JSON.stringify({users, posts}, undefined, 2);
  fs.writeFile(file, content, err => {
    if (err) {
      res.send(err);
    }
    res.send("Post added successfully");
  });
});


app.listen(port, () => {
  console.log(`Example app listening on port ${port}`)
})