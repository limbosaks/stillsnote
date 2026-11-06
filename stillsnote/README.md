# Stillsnote

Members-only Instagram caption studio for stillsnote.com.

- **Website** (this folder): plain HTML, CSS and JavaScript, hosted on GitHub Pages.
- **Backend** (`google-script/Code.gs`): a Google Apps Script web app. It handles membership, every AI call, saved styles, drafts, and the daily trends refresh. It is the only place the Claude API key lives.
- **Data**: a Google Sheet called **Stillsnote members**, created by the script.

```
index.html          the whole site (sign-in, Studio, Planner, Trends, Creative, My style)
assets/styles.css   the look
assets/app.js       everything the pages do
config.js           paste your Web app URL here
CNAME               stillsnote.com (for GitHub Pages)
404.html            sends stray links back to the site
favicon.svg
google-script/Code.gs   the backend, pasted into script.google.com (not run from GitHub)
```

Preview without the backend: open `index.html?demo=1` (for example with the VS Code Live Server extension). Demo mode uses sample data and saves nothing.

---

## Setup (about 25 minutes, once)

### 1. Get a Claude API key (5 min)
1. Go to **console.anthropic.com**, sign in, and add a payment method under **Billing**.
2. **API keys → Create key**, name it `stillsnote`, copy it somewhere safe for step 3.
3. Optional but smart: under **Limits**, set a monthly spend limit.

### 2. Create the Google script (5 min)
1. Sign in to the Google account that should own Stillsnote and open **script.google.com → New project**. Rename it **Stillsnote**.
2. Delete the sample code in `Code.gs`, paste all of `google-script/Code.gs`, and click **Save**.
3. **Project Settings** (gear icon) → **Time zone** → `(GMT+09:00) Seoul`.
4. Back in the editor, choose **setup** in the function drop-down and click **Run**. Allow the permissions (if you see "Google hasn't verified this app", click **Advanced → Go to Stillsnote**; it is your own script).
5. The log shows **Ready. Your Stillsnote Sheet: https://docs.google.com/...** Open and bookmark it.

### 3. Add the API key (2 min)
1. **Project Settings → Script Properties → Add script property**.
2. Property: `ANTHROPIC_API_KEY`, Value: your key from step 1. **Save**.
3. In the editor, run **checkSetup**. The log should say `Claude says: ready` and `Time zone: Asia/Seoul`.

### 4. Deploy it as a web app (3 min)
1. **Deploy → New deployment**, gear icon → **Web app**.
2. Description `Stillsnote`, **Execute as: Me**, **Who has access: Anyone**.
3. **Deploy**, then copy the **Web app URL** (ends in `/exec`).

When you change the script later, use **Deploy → Manage deployments → ✎ Edit → Version: New version → Deploy**. That keeps the same URL. A *New deployment* gives you a new URL.

### 5. Connect the website (2 min)
Open `config.js` and paste the Web app URL between the quotes:
```js
endpoint: 'https://script.google.com/macros/s/AKfy.../exec'
```

### 6. Put it on GitHub Pages with your domain (5 min)
1. Create a GitHub repository (for example `stillsnote`) and push this folder to it.
2. Repository **Settings → Pages**: Source **Deploy from a branch**, Branch `main`, folder `/ (root)`. Save.
3. **Custom domain**: `stillsnote.com` (the `CNAME` file already says this). Tick **Enforce HTTPS** once it is offered.
4. At your domain registrar, add DNS records:
   - Four **A** records for `@`: `185.199.108.153`, `185.199.109.153`, `185.199.110.153`, `185.199.111.153`
   - One **CNAME** for `www` pointing to `<your-github-username>.github.io`
   DNS can take up to a few hours.

### 7. Test (3 min)
1. Open stillsnote.com in a private window → **Request access** with a personal email.
2. You get "We got your Stillsnote request"; the admin email gets **Review: approve or decline**.
3. Click it, **Approve and send the code**. The code arrives.
4. Sign in, upload a photo, press **Write it**.
5. In the Sheet, delete your test row if you like.

Trends: the first refresh runs at 6:00 the next morning. To fill it now, run **refreshTrends** in the editor (takes a minute or two).

---

## Day to day (all in the Sheet)

**Members tab**
| Column | What it does |
|---|---|
| Status | `Pending`, `Approved`, `Declined`, `Removed`. Setting someone to **Approved** emails their code. **Removed** locks them out on their next action. |
| Access code | Fixed. Never changes unless you edit it. |
| Plan | `Free` or `Paid`. Controls their monthly limit. |
| AI uses this month | Resets automatically each month. |
| Marketing opt-in | `Yes (date)` only if they ticked the box themselves. Only email news to these people. |

You can add a member by hand: type Name, Email and Status `Approved` in a new row and their code is emailed.

**Settings tab**
- **Sign-up mode**: `Approve` (you approve each person) or `Automatic` (codes are emailed instantly).
- **Free monthly limit** / **Paid monthly limit**: AI requests per month (photo captions, ideas, direction). `0` means unlimited.
- **Trend regions**, **Niches**: what the Trends page covers.

**Lost codes**: if an approved member requests access again with the same email, the script resends their code.

**Costs**: each caption is one Claude request with an image, roughly a few cents or less on Sonnet. The daily trends refresh uses web search (charged per search) for each region. Watch usage at console.anthropic.com.

## Going paid later
Members already have a **Plan** column. When you are ready, connect a payment service (Stripe, or Toss Payments for Korean cards) so a successful payment sets that member's Plan to `Paid`. Nothing else in the site needs to change.

## If something goes wrong
| What you see | Fix |
|---|---|
| "Stillsnote is not connected yet" | `config.js` still has the placeholder. Paste the Web app URL. |
| "Could not reach Stillsnote" | The deployment's access is not **Anyone**, or the URL is wrong. |
| Changed the script, nothing changed | Deploy a **new version** of the existing deployment. |
| Captions fail | In the script editor, open **Executions** to see the error. Run **checkSetup**. |
| No emails | Check spam. Google allows about 100 emails a day on a free Gmail account, 1,500 on Workspace. |
| iPhone photo will not open | HEIC files do not open in some browsers. On the iPhone: Settings → Camera → Formats → Most Compatible, or export as JPG. |
