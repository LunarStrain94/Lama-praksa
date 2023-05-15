const express = require('express');
const app = express();
const port = 3000;
const data = require('./data.json');
const fs = require('fs');
const file = "./data.json";

app.use(express.json());

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

app.put('/user/:userId/:email', (req, res) => {
  let users = data["users"];
  let posts = data["posts"];
  let changed = false;
  users.forEach(element => {
    if (element["id"] == req.params.userId) {
      element["email"] = req.params.email;
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

app.post('/post/:userId/:title/:body', (req, res) => {
  let posts = data["posts"];
  let users = data["users"];
  let post = {"id" : (posts.length + 1).toString()};
  post["title"] = req.params.title;
  post["body"] = req.params.body;
  post["user_id"] = req.params.userId;
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