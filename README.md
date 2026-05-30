# Campus Ideas Hub 🎓💡

> A Microsoft-ecosystem-powered web application for collecting, organizing, visualizing, and summarizing student ideas on **Sustainability & Innovation** — built for colleges and academic institutions.

---

## Table of Contents

1. [Overview](#overview)
2. [Features](#features)
3. [Tech Stack](#tech-stack)
4. [System Design — High-Level Design (HLD)](#system-design--high-level-design-hld)
   - [Architecture Diagram](#architecture-diagram)
   - [Component Overview](#component-overview)
   - [Data Flow](#data-flow)
5. [Low-Level Design (LLD)](#low-level-design-lld)
   - [Frontend Architecture](#frontend-architecture)
   - [Backend Architecture](#backend-architecture)
   - [Data Model](#data-model)
   - [Module Breakdown](#module-breakdown)
6. [API Reference](#api-reference)
7. [Microsoft Integrations (Planned)](#microsoft-integrations-planned)
8. [Project Structure](#project-structure)
9. [Getting Started](#getting-started)
10. [Configuration](#configuration)
11. [Deployment](#deployment)
12. [Roadmap](#roadmap)
13. [Contributing](#contributing)

---

## Overview

**Campus Ideas Hub** is a lightweight, full-stack web application that enables students to submit ideas categorized under themes like Energy, Waste, Innovation, and Social Impact. Administrators can view, analyze, and export submissions through a built-in dashboard with canvas-rendered visualizations.

The application is designed to integrate with the **Microsoft 365 ecosystem** (Forms, SharePoint, Excel, Power BI, Copilot) while providing a fully functional offline-first demo mode using browser `localStorage` and an in-memory Node.js server.

---

## Features

| Feature | Status |
|---|---|
| Idea submission form (local fallback) | ✅ Implemented |
| Microsoft Forms embed | ✅ Iframe placeholder |
| In-memory REST API (`/api/submit`, `/api/ideas`) | ✅ Implemented |
| Browser `localStorage` persistence | ✅ Implemented |
| Admin dashboard with tabular view | ✅ Implemented |
| CSV export of local ideas | ✅ Implemented |
| Category bar chart (native Canvas) | ✅ Implemented |
| Word cloud from idea text (native Canvas) | ✅ Implemented |
| Impact vs Feasibility matrix (native Canvas) | ✅ Implemented |
| Microsoft MSAL sign-in integration | 🔲 Placeholder |
| Microsoft Graph API (Excel/SharePoint write) | 🔲 Placeholder |
| Microsoft Graph API (Excel/SharePoint read/sync) | 🔲 Placeholder |
| Power BI embed dashboard | 🔲 Placeholder |
| File upload via SharePoint DriveItem | 🔲 Placeholder |
| Copilot AI summaries | 🔲 Planned |

---

## Tech Stack

### Frontend
| Technology | Role |
|---|---|
| HTML5 | Page structure and SPA views |
| CSS3 (custom variables) | Dark-themed responsive styling |
| Vanilla JavaScript (ES2020) | SPA navigation, form handling, canvas charts |
| HTML5 Canvas API | Bar chart, word cloud, matrix visualization |
| `localStorage` API | Client-side data persistence (demo mode) |
| `crypto.randomUUID()` | Unique idea ID generation |

### Backend
| Technology | Role |
|---|---|
| Node.js | Server runtime |
| Express 5 | HTTP framework + static file serving |
| `body-parser` | JSON request body parsing |
| `nodemon` | Development auto-reload |

### Planned / Pluggable
| Technology | Role |
|---|---|
| MSAL.js | Microsoft identity authentication |
| Microsoft Graph API | Excel/SharePoint read & write |
| Power BI Embedded | Advanced analytics dashboards |
| Microsoft Forms | Structured student submission front-end |
| Microsoft Copilot | AI-powered idea summaries |

---

## System Design — High-Level Design (HLD)

### Architecture Diagram

```
┌─────────────────────────────────────────────────────────────────┐
│                         CLIENT (Browser)                        │
│                                                                 │
│  ┌──────────┐  ┌─────────────┐  ┌───────────┐  ┌───────────┐  │
│  │  Home    │  │  Submit     │  │ Dashboard │  │  Summary  │  │
│  │  View    │  │  View       │  │  View     │  │  View     │  │
│  └──────────┘  └──────┬──────┘  └─────┬─────┘  └─────┬─────┘  │
│                       │               │               │         │
│              ┌────────▼───────────────▼───────────────▼──────┐  │
│              │              app.js (SPA Controller)           │  │
│              │  - SPA navigation                              │  │
│              │  - Form submission logic                       │  │
│              │  - localStorage CRUD                          │  │
│              │  - Canvas chart rendering                     │  │
│              │  - MSAL auth stub                             │  │
│              └────────────────────┬───────────────────────────┘  │
│                                   │                              │
│              ┌────────────────────▼───────────────────────────┐  │
│              │           localStorage (Demo Storage)          │  │
│              └────────────────────────────────────────────────┘  │
└──────────────────────────────┬──────────────────────────────────┘
                               │ HTTP (fetch / REST)
                               │
┌──────────────────────────────▼──────────────────────────────────┐
│                      NODE.JS SERVER (server.js)                 │
│                                                                 │
│   Express App                                                   │
│   ├── Static file serving  (Index.html, style.css, app.js)     │
│   ├── POST /api/submit  →  In-memory ideas[] array             │
│   └── GET  /api/ideas   →  Returns all in-memory ideas         │
│                                                                 │
│   In-Memory Store: ideas[]  (resets on server restart)         │
└──────────────────────────────┬──────────────────────────────────┘
                               │ (Planned: Graph API calls)
                               │
┌──────────────────────────────▼──────────────────────────────────┐
│                   MICROSOFT 365 ECOSYSTEM                       │
│                                                                 │
│  ┌──────────────────┐   ┌──────────────────┐                   │
│  │  Microsoft Forms │   │  SharePoint /    │                   │
│  │  (Submission UI) │   │  OneDrive Excel  │                   │
│  └──────────────────┘   └──────────────────┘                   │
│                                                                 │
│  ┌──────────────────┐   ┌──────────────────┐                   │
│  │  Power BI Embed  │   │  Microsoft       │                   │
│  │  (Dashboards)    │   │  Copilot (AI)    │                   │
│  └──────────────────┘   └──────────────────┘                   │
└─────────────────────────────────────────────────────────────────┘
```

### Component Overview

| Layer | Components | Responsibility |
|---|---|---|
| **Presentation** | `Index.html`, `style.css` | Renders 4 SPA views (Home, Submit, Dashboard, Summary); responsive dark theme |
| **Application Logic** | `app.js` | SPA routing, form handling, data read/write, chart rendering, auth stubs |
| **API Server** | `server.js` | Serves static files, exposes REST endpoints, holds in-memory idea store |
| **Client Storage** | Browser `localStorage` | Offline-first demo storage; key: `campus_ideas` |
| **Cloud Storage** (planned) | Microsoft SharePoint / Excel | Persistent, multi-user, organization-wide idea storage via Graph API |
| **Authentication** (planned) | MSAL.js + Azure AD | Microsoft 365 identity sign-in for Graph-connected features |
| **Visualization** (planned) | Power BI Embedded | Rich dashboards embedded inside the Summary view |

### Data Flow

#### Demo / Offline Mode (No sign-in)
```
Student fills form
      │
      ▼
app.js validates + creates idea object
      │
      ▼
idea saved to localStorage["campus_ideas"]
      │
      ▼
renderTable() + renderSummary() re-render from localStorage
```

#### REST API Mode (Server running)
```
Client calls POST /api/submit  →  JSON body
      │
      ▼
server.js adds submittedAt timestamp
      │
      ▼
Pushed to in-memory ideas[] array
      │
      ▼
GET /api/ideas returns full list
```

#### Microsoft Graph Mode (Planned)
```
User clicks "Sign in with Microsoft"
      │
      ▼
MSAL.js acquires Azure AD token
      │
      ▼
sendIdeaToGraph(idea) called on submit
      │
      ▼
Graph API: POST to Excel table row  OR  SharePoint List item
      │
      ▼
"Sync from Graph" reads data back and populates UI
```

---

## Low-Level Design (LLD)

### Frontend Architecture

The frontend is a **Single-Page Application (SPA)** with four named views toggled via CSS (`display: none` / `display: block`).

```
Index.html
├── <header>         — App title + tagline
├── <nav>            — View switcher buttons + Auth buttons
│     ├── #home      — Home/About card
│     ├── #submit    — Forms iframe + Local fallback form
│     ├── #dashboard — Admin table + action buttons
│     └── #summary   — Charts + Power BI container
└── <footer>         — Branding
```

**SPA Navigation Logic (`app.js`)**
```
document.querySelectorAll(".nav-btn")
  → on click:
      remove "active" from all .view elements
      add "active" to target view by data-target attribute
```

---

### Backend Architecture

**`server.js`** is a minimal Express 5 server with three responsibilities:

1. **Static File Server** — serves the entire project directory (HTML, CSS, JS)
2. **POST `/api/submit`** — accepts a JSON idea object, stamps `submittedAt`, appends to `ideas[]`
3. **GET `/api/ideas`** — returns the full in-memory `ideas[]` array as JSON

```
app.js (Express)
│
├── Middleware: bodyParser.json()
├── Middleware: express.static(__dirname)
│
├── POST /api/submit
│     ├── Read req.body (idea fields)
│     ├── Stamp idea.submittedAt = new Date().toISOString()
│     ├── ideas.push(idea)
│     └── res.json({ success: true, message: "Idea received!" })
│
└── GET /api/ideas
      └── res.json(ideas)
```

> ⚠️ The server-side `ideas[]` array is ephemeral — it resets on every server restart. Persistent storage via Microsoft Graph (Excel/SharePoint) is the intended production path.

---

### Data Model

#### Idea Object

```json
{
  "id":           "uuid-v4-string",
  "name":         "Student full name",
  "email":        "student@college.edu",
  "title":        "Short idea title",
  "desc":         "Full description of the idea",
  "category":     "Energy | Waste | Innovation | Social Impact",
  "impact":       3,
  "feasibility":  4,
  "submittedAt":  "2026-05-30T10:00:00.000Z",
  "fileName":     "optional-attachment.pdf"
}
```

| Field | Type | Validation | Description |
|---|---|---|---|
| `id` | `string` (UUID v4) | Auto-generated | Unique identifier per idea |
| `name` | `string` | Required | Submitter's full name |
| `email` | `string` | Required, email format | Submitter's email |
| `title` | `string` | Required | Brief idea headline |
| `desc` | `string` | Required | Detailed idea description |
| `category` | `enum` | Required | One of: Energy, Waste, Innovation, Social Impact |
| `impact` | `number` | 1–5 | Self-rated impact score |
| `feasibility` | `number` | 1–5 | Self-rated feasibility score |
| `submittedAt` | `string` (ISO 8601) | Auto-generated | Timestamp at submission |
| `fileName` | `string \| null` | Optional | Name of attached file (if any) |

#### LocalStorage Schema

```
Key:   "campus_ideas"
Value: JSON.stringify(Idea[])   — array of Idea objects
```

---

### Module Breakdown

#### `app.js` — Client-Side Modules

| Function / Block | Responsibility |
|---|---|
| SPA nav listener | Toggles `.active` class on `.view` sections |
| `loadLocalIdeas()` | Reads and parses `campus_ideas` from `localStorage` |
| `saveLocalIdeas(list)` | Serializes and writes idea list to `localStorage` |
| `msAuth` object | Stub for MSAL sign-in/sign-out state |
| `ideaForm` submit handler | Validates form, builds idea object, routes to Graph or local storage |
| `sendIdeaToGraph(idea)` | Placeholder for Microsoft Graph API POST; falls back to `localStorage` |
| `btnClearLocal` listener | Clears `localStorage` and re-renders |
| `btnExportLocal` listener | Generates CSV from local ideas, triggers browser download |
| `btnSyncFromGraph` listener | Placeholder for Microsoft Graph read/sync |
| `renderTable()` | Populates `#ideasTable tbody` from `localStorage` with XSS-safe HTML |
| `renderSummary()` | Calls all chart renderers; aggregates category counts |
| `drawBarChart(id, counts)` | Canvas: renders vertical bar chart of idea counts by category |
| `wordFrequency(text)` | Tokenizes text, removes stop words, returns `{word: count}` map |
| `drawWordCloud(id, freq)` | Canvas: scatter-renders top-50 words sized by frequency |
| `drawMatrix(id, list)` | Canvas: plots ideas on Impact (X) vs Feasibility (Y) matrix with quadrant labels |
| `clamp(n, min, max)` | Utility: numeric clamping |
| `escapeHTML(s)` | Utility: XSS prevention — escapes `& < > " '` |
| `toCSV(list)` | Utility: converts idea array to RFC-4180 CSV string |
| `embedPowerBI()` | Placeholder scaffold for Power BI embedded report |

#### `server.js` — Server-Side Modules

| Component | Responsibility |
|---|---|
| `express()` | HTTP server and middleware pipeline |
| `bodyParser.json()` | Parse incoming JSON request bodies |
| `express.static()` | Serve `Index.html`, `style.css`, `app.js` from project root |
| `ideas[]` | In-memory store for submitted ideas (non-persistent) |
| `POST /api/submit` | Accept, timestamp, and store a new idea |
| `GET /api/ideas` | Return all stored ideas as JSON |
| `app.listen(3000)` | Start server on port 3000 |

---

## API Reference

### `POST /api/submit`

Submit a new idea to the server-side in-memory store.

**Request**
```
Content-Type: application/json
```

```json
{
  "name": "Jane Doe",
  "email": "jane@college.edu",
  "title": "Solar-powered bike stations",
  "desc": "Install solar panels over campus bike racks to power USB charging.",
  "category": "Energy",
  "impact": 4,
  "feasibility": 3,
  "fileName": null
}
```

**Response — 200 OK**
```json
{
  "success": true,
  "message": "Idea received!"
}
```

---

### `GET /api/ideas`

Retrieve all ideas from the in-memory store.

**Response — 200 OK**
```json
[
  {
    "name": "Jane Doe",
    "email": "jane@college.edu",
    "title": "Solar-powered bike stations",
    "desc": "Install solar panels over campus bike racks to power USB charging.",
    "category": "Energy",
    "impact": 4,
    "feasibility": 3,
    "fileName": null,
    "submittedAt": "2026-05-30T10:00:00.000Z"
  }
]
```

> **Note:** The server resets on restart. All data is lost unless Graph/SharePoint integration is configured.

---

## Microsoft Integrations (Planned)

### 1. Microsoft Forms
- Replace the `<iframe src="...">` in `Index.html` with your organization's Microsoft Forms embed URL.
- Responses collected in Forms flow automatically to the linked Excel spreadsheet in OneDrive.

### 2. MSAL.js Authentication
```
1. Register an app in Azure Active Directory (App Registration)
2. Note the Client ID and Tenant ID
3. Add MSAL.js (msal.min.js) to the project
4. Fill in the msAuth object in app.js:
   - clientId, authority, redirectUri
5. Call msAuth.signIn() to acquire an access token for Graph scopes
```

### 3. Microsoft Graph API (Excel/SharePoint)
```
Scope required: Files.ReadWrite (Excel) or Sites.ReadWrite.All (SharePoint)

Write: POST https://graph.microsoft.com/v1.0/me/drive/items/{item-id}/workbook/tables/{table}/rows/add
       Body: { values: [[name, email, title, desc, category, impact, feasibility, submittedAt, fileName]] }

Read:  GET  https://graph.microsoft.com/v1.0/me/drive/items/{item-id}/workbook/tables/{table}/rows
       Map returned values back to Idea objects and save to localStorage
```

### 4. Power BI Embedded
```
1. Install powerbi-client: <script src="powerbi.min.js">
2. Create a Power BI workspace and report connected to your Excel/SharePoint source
3. Generate an embed token (Power BI REST API or Azure AD)
4. Configure embedPowerBI() in app.js:
   - type: "report"
   - id: <your Report ID>
   - embedUrl: <your Embed URL>
   - accessToken: <your embed token>
5. Call powerbi.embed(document.getElementById("powerBiContainer"), config)
```

### 5. Microsoft Copilot
- Connect Copilot Studio to your SharePoint list or Excel data source.
- Build a Copilot plugin that summarizes the top ideas by category, highest impact/feasibility scores, and keyword trends.

---

## Project Structure

```
idea-collector/
│
├── Index.html          # Main HTML — 4-view SPA shell
├── style.css           # Dark-theme stylesheet with CSS variables
├── app.js              # Client-side SPA logic, charts, storage, auth stubs
├── server.js           # Express REST API server + static file server
├── package.json        # npm manifest (Express 5, nodemon)
├── package-lock.json   # Lockfile
├── .gitignore          # Ignored files
└── README.md           # This file
```

---

## Getting Started

### Prerequisites

- [Node.js](https://nodejs.org/) v18 or higher
- npm v9 or higher

### Installation

```bash
# Clone the repository
git clone https://github.com/Sayan-dev731/idea-collector.git
cd idea-collector

# Install dependencies
npm install
```

### Running in Development

```bash
npm run dev
# Server starts with nodemon (auto-reloads on file changes)
# Open: http://localhost:3000
```

### Running in Production

```bash
npm start
# Open: http://localhost:3000
```

---

## Configuration

All Microsoft integration settings are currently **placeholder stubs** inside `app.js` and `Index.html`. To enable them:

| Setting | File | Location |
|---|---|---|
| Microsoft Forms embed URL | `Index.html` | `<iframe src="YOUR_FORM_EMBED_URL">` |
| MSAL Client ID / Tenant | `app.js` | `msAuth` object (TODO comment) |
| Excel workbook item ID | `app.js` | `sendIdeaToGraph()` function |
| SharePoint site / list ID | `app.js` | `sendIdeaToGraph()` function |
| Power BI Report ID | `app.js` | `embedPowerBI()` config object |
| Power BI Embed URL | `app.js` | `embedPowerBI()` config object |
| Power BI Access Token | `app.js` | `embedPowerBI()` config object |

---

## Deployment

### Local / On-Premise
```bash
npm start   # Runs on port 3000
```

### Azure App Service
```bash
# Set start command in App Service configuration:
node server.js

# Or push via GitHub Actions to Azure Web Apps
```

### Docker
```dockerfile
FROM node:20-alpine
WORKDIR /app
COPY package*.json ./
RUN npm install --production
COPY . .
EXPOSE 3000
CMD ["node", "server.js"]
```

---

## Roadmap

- [ ] Implement full MSAL.js sign-in flow
- [ ] Implement Microsoft Graph write (Excel row append / SharePoint list item)
- [ ] Implement Microsoft Graph read (sync data from SharePoint to UI)
- [ ] Add file upload via Graph DriveItem API
- [ ] Add server-side persistent database (PostgreSQL / Azure Cosmos DB) as an alternative to in-memory store
- [ ] Integrate Power BI Embedded with live data
- [ ] Add Microsoft Copilot plugin for AI-generated idea summaries
- [ ] Add pagination and search/filter to the admin dashboard table
- [ ] Add user role support (student vs. admin) via Azure AD groups
- [ ] Add email notification on idea submission via Microsoft Graph Mail API

---

## Contributing

1. Fork the repository
2. Create a feature branch: `git checkout -b feature/your-feature`
3. Commit your changes: `git commit -m "Add your feature"`
4. Push the branch: `git push origin feature/your-feature`
5. Open a Pull Request

---

*Built for colleges using Microsoft apps: Forms, SharePoint/Excel, Power BI, Copilot, Teams.*
