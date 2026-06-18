# Company Name Matcher

A Streamlit app for fuzzy-matching company names between an **Indue** (reference) list and another list. Supports uploading:

- **1 Excel file with multiple tabs** (e.g. tabs `IndueD` and `KEV`), or
- **2 separate Excel files**

The app automatically detects which tab/file is "Indue" based on its name (containing the word `indue`, case-insensitive), auto-detects the column containing company names, and produces a result file with 3 tabs: `Result`, `Source - <Indue>`, `Source - <other>`.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open `http://localhost:8501` in your browser.

## Deploy to Streamlit Community Cloud (free)

### Step 1 — Push the code to GitHub

1. Create a new repository on GitHub (e.g. named `company-name-matcher`), either Public or Private.
2. From the folder containing these files (`app.py`, `requirements.txt`, `README.md`), run:

```bash
git init
git add .
git commit -m "Initial commit: company name matcher app"
git branch -M main
git remote add origin https://github.com/<YOUR_GITHUB_USERNAME>/company-name-matcher.git
git push -u origin main
```

> Replace `<YOUR_GITHUB_USERNAME>` with your GitHub username. If Git isn't configured yet, also run:
> `git config --global user.email "you@example.com"` and `git config --global user.name "Your Name"`.

If you'd rather not use the command line:
1. Go to [github.com/new](https://github.com/new), create a new repo, and **don't** tick "Add a README".
2. On the new repo's page, click **"uploading an existing file"**.
3. Drag and drop `app.py`, `requirements.txt`, and `README.md`, then click **Commit changes**.

### Step 2 — Deploy on Streamlit Cloud

1. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with your GitHub account.
2. Click **"New app"**.
3. Choose:
   - **Repository**: the repo you just created (e.g. `<your-username>/company-name-matcher`)
   - **Branch**: `main`
   - **Main file path**: `app.py`
4. Click **"Deploy!"**.

After a minute or two, Streamlit will give you a link like:

```
https://<app-name>-<random>.streamlit.app
```

Every time you push new code to GitHub, the app on Streamlit Cloud will redeploy automatically.

## File structure

```
.
├── app.py              # Main Streamlit app source code
├── requirements.txt    # Required Python packages
└── README.md            # This guide
```

## Notes

- The matching threshold (default 86/100) can be adjusted directly in the app using the slider.
- If the app doesn't correctly auto-detect the "Indue" tab/file (e.g. its name doesn't contain "Indue"), you can select it manually in the interface.
- The downloaded file never overwrites your original data — it's always a new Excel file.

## A note on data security

This app, when deployed on Streamlit Community Cloud's free tier, runs on shared public infrastructure with no built-in login — anyone with the app's link can open it and upload files. Uploaded data is processed in memory and isn't intentionally stored long-term, and traffic is encrypted in transit, but there's no enterprise-grade compliance guarantee (e.g. SOC2, GDPR data processing agreement) at the free tier. If your company names or related data are commercially sensitive or covered by confidentiality obligations, consider adding a login/password step, hosting the app on your own company infrastructure, or using a paid/enterprise Streamlit deployment instead of sharing the public link externally.
