# Incident uploader — deployment and handover

Three files:

| File | What it is |
|---|---|
| `app.R` | The whole application. No other code files. |
| `manifest.json` | Tells the host which R version and packages to install. Already generated — you don't need a working R install to deploy. |
| `DEPLOY.md` | This file. |

---

## Read this first: the platform changed

I originally said to use shinyapps.io and store the GitHub token in its
dashboard. **That was wrong.** shinyapps.io has no secret-variable management —
Posit's own feature comparison lists it as unsupported. The only way to get a
token onto shinyapps.io is to bundle a `.Renviron` into the deployment, and that
bundle is downloadable from the dashboard.

Two other reasons to move:

- **shinyapps.io is being retired.** Posit is consolidating it into Posit
  Connect Cloud. Free accounts migrate automatically on **28 January 2027**.
- **Connect Cloud deploys from GitHub**, so you never need `rsconnect` working
  locally. Given the CRAN/SSL problems on your laptop, that matters.

**Deploy to Posit Connect Cloud.** Free plan, real secret management, GitHub
publishing, and it's the platform that still exists in two years.

---

## Credentials: use a GitHub App, not a token

This is the part that decides whether the app still works in three years.

A personal access token **expires**. GitHub caps fine-grained tokens at
**366 days** by default for organisation resources, and an org owner has to
relax that policy for anything longer. GitHub also **auto-removes any token
unused for a year**. So a token means someone has to remember to renew it, and
when they don't, uploads stop.

A **GitHub App** has no such limit. Its private key does not expire. The app
exchanges that key for a one-hour token every time it runs, automatically.
Nothing to renew, ever. It is also more tightly scoped than a token: one repo,
one permission.

The app supports both. It uses the App when `GITHUB_APP_ID` and `GITHUB_APP_KEY`
are set, and falls back to `GITHUB_PAT` when they aren't — so you can get running
on a token today and move to the App later without touching the code.

### Setting up the GitHub App (~10 minutes, once)

1. `PalaemonMaritime` → **Settings → Developer settings → GitHub Apps → New GitHub App**
2. Name it something like `Palaemon Incident Uploader`. Homepage URL can be the
   repo URL.
3. **Uncheck Webhook → Active.** The app doesn't use webhooks.
4. **Repository permissions → Contents: Read and write.** Nothing else.
5. **Where can this app be installed:** only this account.
6. Create it. Note the **App ID** shown at the top.
7. Scroll to **Private keys → Generate a private key**. A `.pem` file downloads.
8. **Install App** in the left sidebar → install on `PalaemonMaritime` →
   **Only select repositories** → `json.api`.

That's it. The `.pem` never expires.

### If you use a token instead

Fine-grained token, resource owner `PalaemonMaritime`, repository access limited
to `json.api`, permission **Contents: Read and write**, longest expiry allowed.

The app reads the expiry date GitHub reports and shows it on screen, warning
from **45 days out** and stating plainly once it has lapsed. It won't die
quietly — but someone still has to act on the warning, which is why the App is
the better answer.

---

## Deployment

### 1. Put the app in its own repository

Create a **new** repo, e.g. `PalaemonMaritime/incident-uploader`, with `app.R`
and `manifest.json` in the root.

Don't put it inside `json.api`. The app commits to that repo on every upload; if
the code lives there too, each data commit can retrigger a redeploy.

### 2. Publish

1. <https://connect.posit.cloud> → sign in with GitHub
2. **Publish** → your `incident-uploader` repo, branch `main`, primary file `app.R`
3. Deploy. The first build takes a few minutes while packages install.

### 3. Add the credentials

App → **Settings → Variables**. Add as **secrets**:

| Name | Value |
|---|---|
| `GITHUB_APP_ID` | the App ID from step 6 above |
| `GITHUB_APP_KEY` | the entire contents of the `.pem` file, including the `-----BEGIN` and `-----END` lines |

(Or just `GITHUB_PAT` if you're using a token.)

Restart the app. Credentials live only on Posit's servers — not in the repo, not
in the bundle, not on anyone's laptop.

### 4. Test before handing over

Open the app. It checks its credentials on load and shows their status at the
bottom of the panel. If that line is red, fix it before going further.

Then upload a file you have **already** uploaded. It should report every
incident as a duplicate and make no commit — check the repo to confirm nothing
changed. Then upload a genuinely new export.

Give the successor the URL. That's the handover.

---

## Two things to hand over besides the URL

**Who owns the Connect Cloud account.** If you publish under your own Posit
account, whoever takes over can't reach the settings. Publish from an account
the company controls, or transfer it before you go.

**Where the `.pem` is kept.** If the App's private key is lost, a new one can be
generated from the App settings — but only by someone with admin access to the
`PalaemonMaritime` organisation. Make sure at least one person who is staying
has that access.

---

## How the app decides what's a duplicate

It converts the Excel to JSON **first**, then compares JSON to JSON, so
formatting differences between spreadsheet and dataset can't produce a false
"new incident".

Two incidents match when **Date + Ship Name + IMO** match. Where there's no IMO
— every "Name Withheld" record, 19 of them currently — it falls back to
**Date + Ship Name + decimal position + the opening of the narrative**.

Two details that matter if this is ever modified:

- The fallback uses the **decimal** coordinates, not the `13° 36.80' N` strings.
  Text arriving from Excel and text arriving from JSON can hold identical bytes
  but carry different encoding marks, and R treats those as unequal. That caused
  a real false negative during testing — a record already in the dataset was
  reported as new. Numbers can't carry encoding marks.
- Every key component passes through `normalise()`, which forces UTF-8 and
  reduces to printable ASCII for the same reason.

Duplicates within a single upload are caught too. The function is `incident_key()`.

## What the app writes

- `Global_Incidents.json` — new incidents appended to the end. Existing records
  untouched, byte for byte (verified across all 213 current records).
- `upload_log.json` — one entry per successful upload. This is what makes the
  running total survive; Shiny forgets everything when the tab closes. Delete it
  and the history resets, but the dataset is unaffected.

Commit messages: `Update Global_Incidents.json - YYYY-MM-DD`.

## Field handling

The Excel `Boarded?` column is dropped — it isn't in the dataset schema.
Coordinates are kept in the original `1° 11.80' N` form and also converted to
decimal (`Latitude_dd` / `Longitude_dd`). Empty cells are omitted rather than
written as `null`, matching the existing file. The literal text `"NA"` is kept,
because it's a real value in `Coastal State Action Taken` — only genuinely empty
cells are dropped.

Dates are only reformatted when already in `YYYY-MM-DD` form. Anything else is
passed through exactly as written: left to guess, R reads `04/07/2026` as year 4
and would silently write `4-07-20`. An odd-looking date in the dataset is
recoverable; a silently mangled one isn't.

---

## If something goes wrong

**Red credential line on load** — the message says which setup is incomplete.
Fix the variable in Connect Cloud settings and restart.

**"The GitHub App is not installed on PalaemonMaritime/json.api"** — the App
exists but wasn't installed on the repo. Install App → select `json.api`.

**"The file changed on GitHub while this upload was in progress"** — someone
edited `Global_Incidents.json` mid-upload. Try again.

**Build fails on a package version.** The manifest was generated against
R 4.3.3. If Connect Cloud can't resolve something, regenerate it on any machine
with working R:

```r
install.packages("rsconnect")
rsconnect::writeManifest(appDir = "path/to/app/folder")
```

Then commit the new `manifest.json`.

**Logs** are in the Connect Cloud dashboard under the app's Logs tab.
