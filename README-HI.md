# FAPL Maintenance Portal — Cloudflare Pages

## Is folder mein kya hai
- `index.html` — portal frontend
- `netlify-api-bridge.js` — purane `google.script.run` frontend calls ko API endpoint par bhejne ka bridge (filename purana hai; isse rename na karein)
- `functions/api/gas-api.js` — Cloudflare Pages API proxy

**`Code.gs` yahan jaan-boojhkar nahi rakha gaya hai.** Use Google Apps Script editor mein hi rakhein; public GitHub repo mein backend source ya secrets upload na karein.

## GitHub par upload
1. Is ZIP ko extract karein.
2. Is folder ke andar ki files/folder ko repository ke root mein upload karein — folder ko ek aur folder ke andar nest na karein.
3. GitHub root par `index.html`, `netlify-api-bridge.js`, aur `functions/api/gas-api.js` ka path hona chahiye.
4. Cloudflare Pages project ko GitHub repository se deploy karein. Framework preset `None`, build command blank, build output directory `.`.
5. Cloudflare Pages → Settings → Variables and Secrets → Production mein `GAS_WEB_APP_URL` naam ka variable banayein. Value Google Apps Script Web App URL ho jo `/exec` par khatam hota ho.
6. Save karke naya deployment trigger karein.

## Zaroori check
Google Apps Script deployment ko Web app ke roop mein deploy hona chahiye, access aapke portal users ke liye sahi set hona chahiye, aur `doPost(e)` ko `{ "fn": "loginUser", "args": [...] }` jaise JSON payload ka jawab dena chahiye.

Yeh package file structure aur Cloudflare Pages proxy ke liye hai; live Apps Script deployment aur login credentials ke saath end-to-end test karna abhi baaki hai. Agar login fail ho, error/screenshot bhejein — Apps Script ke handler allowlist aur return format ko verify karna hoga.
