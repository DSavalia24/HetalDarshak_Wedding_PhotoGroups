Wedding Photo Group Website With GET-Based Live Tracker Fix

Files included:
- index.html
- tracker.html
- style.css
- floral-left.png
- floral-right.png
- Code.gs

Add these yourself:
- guests.csv
- image1.jpg

IMPORTANT:
This version no longer uses POST. The coordinator page updates Google Sheets using GET links like:
YOUR_WEB_APP_URL?action=complete&group=1

Setup:
1. Rename your background image to image1.jpg.
2. Place image1.jpg and guests.csv in the same folder as index.html.
3. In Google Sheets, create a tab named exactly: PhotoTracker
4. Add headers: Group | Status
5. Add rows 1 through 44. Set Group 1 to Current and all others to Pending.
6. Replace your Apps Script code with Code.gs from this package.
7. Deploy Apps Script as Web App:
   - Execute as: Me
   - Who has access: Anyone
8. After changing Code.gs, click Deploy > Manage deployments > Edit > New version > Deploy.
9. Copy the Web App URL ending in /exec.
10. Replace PASTE_YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE in BOTH index.html and tracker.html.
11. Upload/push all website files to GitHub Pages.

Test before using the tracker page:
Open this in your browser:
YOUR_WEB_APP_URL?action=complete&group=1

If Google Sheets updates Group 1 to Done and Group 2 to Current, the backend works.
Then reset using:
YOUR_WEB_APP_URL?action=reset
