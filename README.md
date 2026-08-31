# Article Generator — Setup Guide

This app turns a list of topics in an Excel file into ready-to-use articles
(with images) automatically. No coding needed to *use* it — just some
one-time setup.

---

## What you'll need

1. **Python** installed on your computer (free, one-time install)
2. **An Anthropic API key** (for Claude to write the articles) — required
3. **An OpenAI API key** (for generating images) — optional. If you skip
   this, the app will still write your articles, but instead of images it
   will save a text file listing what images it *would* have made.

---

## Step 1: Install Python (skip if already installed)

1. Go to https://www.python.org/downloads/
2. Download and run the installer for your system (Windows or Mac)
3. **Important (Windows only):** on the first installer screen, tick the
   box that says **"Add Python to PATH"** before clicking Install

To check it worked, open:
- **Windows:** search for "Command Prompt", open it, type `python --version`
- **Mac:** search for "Terminal", open it, type `python3 --version`

You should see something like `Python 3.12.x`.

---

## Step 2: Get your API keys

### Anthropic (Claude) key — required
1. Go to https://console.anthropic.com/
2. Sign up or log in
3. Go to "API Keys" and create a new key
4. Add a small amount of credit (a few dollars covers many articles)
5. Copy the key (starts with `sk-ant-...`) — you'll paste it into the app

### OpenAI key — optional, only if you want auto-generated images
1. Go to https://platform.openai.com/
2. Sign up or log in
3. Go to "API Keys" and create a new key
4. Add a small amount of credit
5. Copy the key (starts with `sk-...`)

Keep these keys private — treat them like a password. Don't share them or
post them anywhere public.

---

## Step 3: Install the app's dependencies (one-time)

1. Unzip the folder you downloaded (`article-generator`)
2. Open Command Prompt (Windows) or Terminal (Mac)
3. Navigate into the folder. For example:
   - Windows: `cd Downloads\article-generator`
   - Mac: `cd Downloads/article-generator`
4. Run this command:
   - Windows: `pip install -r requirements.txt`
   - Mac: `pip3 install -r requirements.txt`

This installs a couple of small helper libraries the app needs. You only
need to do this once.

---

## Step 4: Run the app

- **Windows:** double-click `run_windows.bat`
- **Mac:** double-click `run_mac.command`
  (If Mac blocks it the first time: right-click the file → Open → confirm)

A window titled "Article Generator" will open.

---

## Step 5: Use the app

1. **Settings section:** paste your Anthropic API key (and OpenAI key if
   you have one), choose how many images per article, then click
   **Save Settings**. You only need to do this once — it's remembered
   next time.
2. **Reference context (optional):** if you have an existing Claude chat
   or ChatGPT chat with style guidance, background facts, or a preferred
   image look, open that chat, select and copy the text you want to
   reuse, and paste it into the matching box here:
   - Top box → guides the **article's** writing style/tone/facts
   - Bottom box → guides the **image** style (colors, mood, art style)

   This is a manual copy-paste rather than a live link, because neither
   Claude.ai nor ChatGPT allow outside programs to log in and read a
   private chat automatically — pasting the text yourself is the reliable
   way to reuse it. Leave these blank if you don't need them.
3. **Choose Excel File:** select your spreadsheet of topics. It needs a
   column called **Topic** (required). You can optionally add columns
   called **Keywords** and **Notes** for more guidance per article. A
   sample file, `sample_topics.xlsx`, is included so you can see the
   expected format.
4. **Choose Output Folder:** defaults to your Desktop. This is where all
   the finished articles will be saved.
5. Click **Start**. Watch the log at the bottom — it will tell you what
   it's doing for each topic. You can click **Stop** at any time; it will
   finish the current article and then stop.

---

## What you'll get

For every topic, a new folder is created (named after the article title),
containing:
- `topic-name.blade.php` — the article, ready to drop into a Laravel site
- One image file per suggested image (if you added an OpenAI key)
- Or `image-suggestions.txt` (if you didn't add an OpenAI key)

---

## Troubleshooting

- **"Claude API error (401)"** — your Anthropic key is wrong or has no
  credit. Double check it in Settings.
- **"Image API error (401)"** — same, but for the OpenAI key.
- **App won't open on Windows** — make sure you ticked "Add Python to
  PATH" during install, then restart your computer and try again.
- **Mac says "cannot be opened because it is from an unidentified
  developer"** — right-click the file, choose Open, then confirm.

---

## Pushing this to GitHub

Your API keys are stored in `config.json`, which is listed in `.gitignore`
so it will never be uploaded — only the app code itself gets pushed.

1. Create a new repository on GitHub (don't add a README there — you
   already have one)
2. Open Command Prompt / Terminal, navigate into this folder, then run:
   ```
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/YOUR-USERNAME/YOUR-REPO-NAME.git
   git push -u origin main
   ```
3. Refresh your GitHub page — the files should now be there (and
   `config.json` should *not* be, which is correct)

