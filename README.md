# RestoPulse — Restaurant Sales Analytics & BI

A React dashboard for Indian restaurant sales analysis. This project was reconstructed from the supplied single-file Gemini AI Studio export (`restaurant_sales_analytics_bi_platform.tsx`). The original dashboard and its analytics logic are retained, with a local Vite/React project, local Excel parsing, and a small Express backend for secure Gemini API access.

## What's included

- Existing dark-mode React dashboard and navigation.
- INR-first formatting, Indian sample restaurant data, GST selector and DD-MM-YYYY display.
- CSV, XLSX and XLS upload with schema detection and column mapping.
- Customer and menu-item analytics, chart views, customer segmentation, menu-engineering categories and downloadable Markdown/CSV reports.
- A local deterministic chatbot fallback that can answer supported questions from calculated analytics without an API key.
- Optional Gemini chatbot integration through a server-side API route. The API key is not embedded in frontend code.
- Vite development server and Express production server.

## Requirements (macOS / Apple Silicon)

- macOS on Apple Silicon (M1/M2/M3/M4).
- Node.js 22 LTS or newer (20.19+ is required by this project's Vite setup).
- npm, installed with Node.js.
- Visual Studio Code (recommended).

Check your installation in the VS Code terminal:

```bash
node --version
npm --version
```

## Run locally in VS Code

1. Extract the ZIP file.
2. Open the extracted `restaurant-analytics-app` folder in Visual Studio Code.
3. Open **Terminal → New Terminal**.
4. Install dependencies:

```bash
npm install
```

5. (Optional, for Gemini AI responses) Create your local environment file:

```bash
cp .env.example .env
```

Open `.env` and paste your Gemini API key after `GEMINI_API_KEY=`. You can create/manage a key in [Google AI Studio](https://aistudio.google.com/apikey). Keep the key private. The `.env` file is ignored by Git and must never be placed in `src/`.

6. Start the frontend and backend together:

```bash
npm run dev
```

7. Open the Vite URL printed in the terminal (normally **http://localhost:5173**).

The Vite frontend proxies `/api` requests to the local Express server on port 8787. Leave that terminal running while using the app. To check the backend, visit **http://localhost:8787/api/health**.

### Using the chatbot without an API key

The dashboard's deterministic local answer engine remains available without Gemini. It covers a set of common questions by calculating answers from the current dataset. Add a valid `GEMINI_API_KEY` to `.env` to enable open-ended Gemini responses.

## Build and run the production version locally

```bash
npm run build
npm start
```

Then open **http://localhost:8787**. The Express server serves the built React frontend and the `/api` route from the same origin.

## Dataset notes and limits

The dashboard supports two kinds of source data:

1. **Transaction-level data**: ideally includes transaction date, customer ID, item name, quantity, unit price and/or line total. This is needed for daily, weekly and monthly trend analysis.
2. **Customer-summary data**: includes each customer's top ordered items, frequencies, prices and optional customer revenue. This can support cumulative customer/menu summaries but does not contain enough information to reconstruct monthly trends without transaction dates.

The app keeps raw numeric values as supplied and defaults the display currency to INR. Changing the currency selector changes formatting; it does not perform foreign-exchange conversion. Select the correct currency for the uploaded file before interpreting its amounts.

The demo data is synthetic and is clearly labelled as demo data in the dashboard. Do not use demo figures as real business results.

## Data privacy

- File parsing and dashboard calculations run in your browser.
- The app does not persist uploaded source files to a database or disk in this starter project; the active data is held in browser memory for the current session.
- When the chatbot is used, the app sends a compact analytics summary (KPIs, leading/low-performing items, selected customer summary fields and trends) to your Express backend, which forwards it to the Gemini API. It does not send the complete raw upload or customer email addresses in that context.
- Gemini interactions are submitted with `store: false` in the API request. Review Google's current API terms and privacy documentation before using real customer/business data.
- Add authentication, access control, retention rules and a proper database before using this as a public multi-user service.

## Deploy to Render

1. Push this folder to a private or public GitHub repository. Do not commit `.env`.
2. In Render, create a **New → Blueprint** service and select the repository, or create a Node web service manually.
3. The included `render.yaml` specifies the build and start commands.
4. Add `GEMINI_API_KEY` as a secret environment variable in Render. Keep `GEMINI_MODEL` as `gemini-3.8-flash` unless you intentionally change it.
5. Deploy. Render uses the `PORT` environment variable automatically; the Express server listens on it.

## Deploy to Google Cloud Run

The included `Dockerfile` can be used by Cloud Run's source deployment. First create a secret named `GEMINI_API_KEY` in Google Secret Manager and grant the Cloud Run service account permission to access it. Then run:

```bash
gcloud run deploy restopulse \
  --source . \
  --region asia-south1 \
  --allow-unauthenticated \
  --set-secrets GEMINI_API_KEY=GEMINI_API_KEY:latest
```

For a production deployment, add appropriate authentication and rate limiting before exposing the chatbot endpoint publicly.

## Project structure

```text
restaurant-analytics-app/
├── src/
│   ├── App.jsx             # Original dashboard, analytics and UI
│   ├── main.jsx            # React entry point
│   └── index.css           # Tailwind and global styles
├── server/
│   └── index.js            # Express API and secure Gemini proxy
├── index.html
├── package.json
├── vite.config.js
├── tailwind.config.js
├── postcss.config.js
├── .env.example
├── render.yaml
├── Dockerfile
└── README.md
```

## Notes

This package is a local-development and deployment starter, not a complete SaaS platform. It does not include user registration, a database, multi-restaurant tenancy, persistent uploads, automated backups, or role-based access control. The original analytics are retained, but business decisions should be checked against the source data, especially where customer-summary rows only contain top-three-item information.
