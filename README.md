# Household Finances

Stage 1: Microsoft sign-in and a synthetic OneDrive save/load test.

Published address: https://gibsonjs27.github.io/household-finances/

This is not yet a budgeting application. No workbook, CSV, balance, or account history is bundled or imported. The public browser application uses a personal Microsoft sign-in and delegated `Files.ReadWrite.AppFolder` access to the signed-in user's app folder. It requests no client secret and no whole-drive access. A different visitor cannot access the owner's OneDrive through this application.

## Connection test

1. Select **Sign in with Microsoft** and complete sign-in and consent yourself.
2. Select **Save sample to OneDrive**. Note the generated test code.
3. Select **Load latest sample** and verify the same code appears.
4. Open the published address on a second device. Sign in with the same account and load the latest sample. Do not save a second sample until you have compared the code.
5. Sign out, then verify the sample and account are no longer displayed.

Each save creates a separate `household-finances-connection-test-…json` file. It does not replace any prior file. These sample files can be removed manually from OneDrive after testing. Financial-data conflict detection, backup retention, imports, budgets, and migration remain later milestones.

## Configuration

The public client ID is configured in `src/config.js`. In the Microsoft application registration:

- Support personal Microsoft accounts.
- Register `https://gibsonjs27.github.io/household-finances/` as a **Single-page application (SPA)** redirect URI.
- Use delegated Microsoft Graph permission **Files.ReadWrite.AppFolder**.
- Do not create a client secret.

The full-page redirect returns to the root page, which calls `handleRedirectPromise()`. Token refresh is restricted to cached access tokens and refresh tokens; an interactive redirect is used when needed. Tokens use MSAL session storage; samples are not persisted in browser storage. Signed file URLs and tokens are never logged or displayed.

## Development and deployment

Run `npm ci`, `npm test`, and `npm run build`. Run `npm run dev` for a local appearance preview; sign-in is disabled locally because the registered redirect is the production address.

The generated `assets/app.js` is committed intentionally so GitHub Pages can serve the existing **main branch / root** setup without changing repository settings or requiring GitHub Actions. Run the build after changing source and commit the updated bundle. Only the repository's application files belong in a deployment; never add financial records, secrets, or the parent workspace.

## Validation status

Automated tests verify sample validation, Graph error handling, pagination, unique-file writes, and downloads without exposing access tokens to download hosts. They do not prove Microsoft registration, consent, real OneDrive access, or cross-device behavior. Those require completing the connection test above on the published site.

## Implementation references

- https://learn.microsoft.com/en-us/entra/msal/javascript/browser/initialization
- https://learn.microsoft.com/en-us/graph/onedrive-sharepoint-appfolder
- https://learn.microsoft.com/en-us/graph/api/driveitem-get-content
- https://learn.microsoft.com/en-us/graph/api/driveitem-put-content
