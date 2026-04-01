# Pexip Outlook Add-in

Custom Outlook Add-in for adding Pexip VMR details to calendar invites.

## Setup Instructions

### 1. Enable GitHub Pages

1. Go to your repository **Settings**
2. Scroll down to **Pages**
3. Under **Source**, select **main** branch
4. Click **Save**
5. Wait a few minutes for GitHub to deploy
6. Your files will be available at: `https://[your-github-username].github.io/pexip-outlook-addin/`

### 2. Update the Manifest

Once GitHub Pages is live, you'll need to update the `manifest.xml` file with your GitHub Pages URLs:

- Replace `[YOUR-GITHUB-USERNAME]` with your actual GitHub username
- Replace `[REPO-NAME]` with your repository name

### 3. Install in Outlook

1. Go to Outlook Web (outlook.office.com)
2. Settings → Manage add-ins
3. Add from file → Select manifest.xml
4. Install

## Files

- `taskpane.html` - Main add-in interface
- `commands.html` - Required by Outlook
- `manifest.xml` - Add-in configuration (update with your URLs)

## How It Works

- First use: Configure your personal Pexip VMR
- Subsequent uses: One-click to insert meeting details
- Settings saved per user
- Works on Desktop, Web, iOS, Android

## Security

- Only accessible to authenticated Pexip Outlook users
- VMR settings stored in Office roaming settings (per user)
- No external data collection