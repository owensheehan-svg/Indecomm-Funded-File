# Airtable to Excel Daily Export 📊

Automated daily export of your Airtable view to Excel with clean formatting.

## 🚀 Quick Setup

### 1. Create a GitHub Repository
1. Go to [GitHub](https://github.com/new)
2. Create a new repository (can be private)
3. Upload these files to your repo

### 2. Add Your Airtable Token as a Secret
1. Go to your repo → **Settings** → **Secrets and variables** → **Actions**
2. Click **New repository secret**
3. Name: `AIRTABLE_TOKEN`
4. Value: `pat0bs5RZuvipZYms.c10ca8f591ad5aa12c3f7eef00a3fd2b35b64333dcb22a037dbafba22f9cec85`
5. Click **Add secret**

### 3. Enable GitHub Actions
1. Go to the **Actions** tab in your repo
2. If prompted, click **"I understand my workflows, go ahead and enable them"**

### 4. Test the Export
1. Go to **Actions** tab
2. Click **Daily Airtable Export** workflow
3. Click **Run workflow** → **Run workflow** button
4. Wait ~30 seconds, then refresh
5. Click on the completed run
6. Download your Excel file from **Artifacts** at the bottom

## 📅 Schedule

The export runs automatically every day at **9 AM UTC**. 

To change the time, edit `.github/workflows/daily_export.yml` and modify the cron schedule:
```yaml
- cron: '0 9 * * *'  # Format: minute hour day month day-of-week
```

Examples:
- `0 14 * * *` = 2 PM UTC daily
- `0 9 * * 1` = 9 AM UTC every Monday
- `0 */6 * * *` = Every 6 hours

Use [crontab.guru](https://crontab.guru) to help with cron syntax.

## 📥 Getting Your Exports

### Option 1: Download from GitHub (Current Setup)
After each run, go to **Actions** → click the run → download from **Artifacts**

Files are kept for 30 days.

### Option 2: Auto-commit to Repo (Optional)
Uncomment the last section in `daily_export.yml` to automatically commit exports back to your repo. Then you can just browse your repo's files to see all exports.

### Option 3: Email or Upload to Cloud
You can extend the workflow to:
- Email the file (using a GitHub Action for sending emails)
- Upload to Google Drive, Dropbox, S3, etc.

## 🛠️ Local Testing

Want to test locally first?

```bash
# Install dependencies
pip install -r requirements.txt

# Set your token
export AIRTABLE_TOKEN="pat0bs5RZuvipZYms.c10ca8f591ad5aa12c3f7eef00a3fd2b35b64333dcb22a037dbafba22f9cec85"

# Run the script
python airtable_export.py
```

## 📝 Customization

### Change Table or View
Edit `airtable_export.py`:
```python
TABLE_NAME = 'Your Table Name'  # Line 15
VIEW_ID = 'shrYourViewID'       # Line 16
```

### Adjust Formatting
The script applies:
- Blue header row with white text
- Auto-sized columns (capped at 50 chars)
- Frozen header row

Modify the `format_excel()` function to customize styling.

## 🔒 Security Note

Your Airtable token is stored securely as a GitHub secret and never exposed in logs or files. Keep your repo private if it contains sensitive data.

## 💡 Tips

- Check the **Actions** tab regularly to ensure exports are running
- GitHub Actions are free for public repos and have generous limits for private repos
- You can manually trigger exports anytime via **Actions** → **Run workflow**

---

Questions? Need help? Let me know! 🎉
