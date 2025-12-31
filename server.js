const express = require('express');
const bodyParser = require('body-parser');
const path = require('path');

const app = express();
app.use(bodyParser.json());
app.use(express.static(path.join(__dirname))); // serves index.html, style.css, app.js

// Local demo storage (in memory)
let ideas = [];

// API route to accept submissions
app.post('/api/submit', (req, res) => {
  const idea = req.body;
  idea.submittedAt = new Date().toISOString();
  ideas.push(idea);
  console.log("New idea:", idea);

  // TODO: Push to Microsoft Graph (Excel/SharePoint)
  res.json({ success: true, message: "Idea received!" });
});

// API route to list all ideas
app.get('/api/ideas', (req, res) => {
  res.json(ideas);
});

app.listen(3000, () => {
  console.log("Server running at http://localhost:3000");
});