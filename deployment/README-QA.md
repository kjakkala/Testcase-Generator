# QA Dashboard Deployment

## Package Contents

The deployment archive contains the Node.js server, public app files, production dependency manifests, regression script, and these deployment instructions. It intentionally excludes `config.json`, `data/local-store.json`, `token-history.json`, `node_modules`, and SQL files.

## Requirements

- Windows Server with Node.js 22 LTS (64-bit) and npm
- An internal HTTPS reverse proxy (IIS ARR, nginx, or equivalent)
- A dedicated, non-administrator Windows service account
- Outbound HTTPS access to the configured Jira, TestRail, GitHub Models, and/or AWS Bedrock endpoints

## Install

1. Copy and extract the archive to a versioned application directory, for example `C:\Apps\UplandQADashboard\current`.
2. From that directory, install the locked production dependencies:

   ```powershell
   npm ci --omit=dev
   ```

3. Run the included parser regression check:

   ```powershell
   npm run test:ac-parser
   ```

4. Create the persistent data directory and grant Modify access only to the service account and QA administrators:

   ```text
   C:\ProgramData\Upland\TestPlanGenerator
   ```

   This directory contains password hashes, per-user integration secrets, shared Bedrock keys, and usage history. Those values are stored locally, not encrypted by the app. Restrict its Windows ACLs and protect backups accordingly.

## HTTPS and Network Access

The app listens on `127.0.0.1:3377` by default. Keep that loopback binding and configure the reverse proxy to terminate HTTPS and forward requests to `http://127.0.0.1:3377`. Forward the original `Host` header and `X-Forwarded-Proto: https`. Set an HTTPS redirect at the proxy and allow inbound traffic only to the proxy listener.

The provided launcher sets `COOKIE_SECURE=true`, so the session cookie is marked Secure. Do not expose the Node listener directly to the network or run the app over public HTTP.

Before initial setup, restrict the proxy route to the Admin/operator network. The first account created through the setup page becomes the Admin. After that account exists, allow the intended QA users through the proxy and add their accounts from the Users page.

## Start and Stop

Run `deployment\start-qa.ps1` as the dedicated service account. It starts Node in the background, writes logs and a PID file under the data directory, and refuses to kill or replace a process already using the configured port.

Run `deployment\stop-qa.ps1` as the same account (or an administrator) to stop only the process recorded by the launcher.

For automatic startup, register `start-qa.ps1` as a Windows Task Scheduler task that runs at system startup under the dedicated service account, whether or not that account is logged on. Configure the task to restart on failure and record its action/working directory.

Optional environment variables supported by the launcher include:

- `QA_DASHBOARD_PORT` (default `3377`)
- `QA_DASHBOARD_DATA_DIR` (default `%ProgramData%\Upland\TestPlanGenerator`)

The app data paths are kept outside the release directory so replacing application files does not replace user accounts or credentials.

## First Login and Configuration

1. Start the app and open its HTTPS proxy URL from the restricted Admin network.
2. Create the first Admin account. Do not use a shared/default password.
3. From **Users**, create individual QA accounts.
4. Each user enters their own Jira, TestRail, and GitHub settings in **Settings**.
5. An Admin enters shared Bedrock region/model and AWS credentials in **Settings**. Grant the AWS identity only the required Bedrock invocation permissions and enable the chosen model/inference profile in that region.
6. Verify one Jira fetch, one small test-case generation, personal token history, and (as Admin) the **All users** weekly/monthly usage view. Do not upload test cases during smoke testing.

Usage written before account attribution was added is shown to Admins as unattributed legacy usage. New AI calls record user, model, purpose, ticket/story key, and token counts. Bedrock cost is not estimated by the app; use AWS billing for charges.

## Upgrade and Rollback

1. Stop the current process with `stop-qa.ps1`.
2. Back up the persistent data directory to an access-restricted location.
3. Extract the new version to a new versioned application directory and run `npm ci --omit=dev` plus the regression test.
4. Update the scheduled task or deployment pointer to the new directory, then run `start-qa.ps1`.
5. Verify `/ping`, login, user settings, and generation before reopening the proxy to QA users.

Rollback by stopping the new version, restoring the previous application directory, and restarting it with the same persistent data directory. Do not restore an older `local-store.json` over newer user changes unless the QA owner approves it.

## Operational Notes

- Sessions are held in server memory and expire after 24 hours. Restarting the app signs everyone out.
- The deployment scripts do not kill arbitrary processes or change firewall/proxy/ACL settings.
- Do not copy the development `config.json`, local-store, or token history into the package.