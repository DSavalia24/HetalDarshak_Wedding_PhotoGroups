Wedding Photo Group Website - Guest Iframe Status Fix

This version makes the guest page work like the coordinator page: it loads the live status directly from Apps Script in an iframe.

Tracker URL is already inserted:
https://script.google.com/macros/s/AKfycbyJenRcbXRDpt6vXL97EHQAdnqeH_9FHTRlCpfqA4nnE0_6Yc1Ur7RFqZI18HbXL1_Gdg/exec

Files included:
- index.html
- tracker.html
- style.css
- floral-left.png
- floral-right.png
- Code.gs
- README.txt

Add these yourself:
- guests.csv
- image1.jpg

IMPORTANT:
1. Replace Apps Script with Code.gs from this package.
2. Deploy a NEW Apps Script version:
   Deploy > Manage deployments > Edit pencil > Version: New version > Deploy
3. Push all website files to GitHub Pages.
4. Hard refresh the guest page with Ctrl + F5.

Test this direct iframe URL:
https://script.google.com/macros/s/AKfycbyJenRcbXRDpt6vXL97EHQAdnqeH_9FHTRlCpfqA4nnE0_6Yc1Ur7RFqZI18HbXL1_Gdg/exec?action=guestStatus&groups=10,44

It should show a styled status box by itself in the browser.
