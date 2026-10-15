# Upland Test Plan / Cases Generator

An AI-assisted web application for creating test plans and test cases from Jira requirements. It can send generated cases to TestRail and export plans for review.

## Features

- Fetch Jira issues and acceptance criteria into organized test-plan sections.
- Generate test cases with GitHub Models or Amazon Bedrock.
- Upload generated cases to TestRail.
- Export test plans and cases to Word or Excel.
- Manage user accounts, integration settings, and AI usage from the application.

## Requirements

- Node.js 22 LTS recommended
- npm
- Access credentials for the integrations you plan to use

## Run Locally

From the repository directory:

```bash
npm install
node server.js
```

Open [http://localhost:3377](http://localhost:3377). The server listens on `127.0.0.1` by default. To use a different port, set the `PORT` environment variable before starting the server.

On first launch, create the initial administrator account when prompted. Configure Jira, TestRail, and AI provider credentials from **Settings**. GitHub Models uses a GitHub token; Amazon Bedrock settings are managed by an administrator.

## Run the Regression Check

```bash
npm run test:ac-parser
```

## Data and Security

User accounts, integration credentials, and usage data are stored locally by the application. Treat the data directory and its backups as sensitive, and do not commit credentials, tokens, customer data, or local data files to GitHub. Environment files are excluded by `.gitignore`.

The local server is intended for local use. For shared or production use, follow the HTTPS reverse-proxy, access-control, and service-account guidance in [deployment/README-QA.md](deployment/README-QA.md). Do not expose the Node.js listener directly to the public internet.
