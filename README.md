# KioskCMS
The front pages of KioskCMS work for the self-service devices with touch screen
and the backend pages are maintained on Desktop browser.

Installation:
1. Enable "ASP" feature for IIS
2. Set "Enable Parent Paths" as "True" in IIS
3. Add a new IIS site (not virtual directory)

The front page cloud use the Kiosk Mode of Chrome for fullscreen on os startup.

Backend Directory: "/admin"
Backend Account: "admin"
Password: "admin"
