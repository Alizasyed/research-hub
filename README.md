# Research Hub

A team space for field research data: upload recordings, note photos and typed notes, then get Urdu and English transcripts and summaries.

- `index.html`, `config.js`: the website (GitHub Pages). It runs in the browser and talks to the Google script.
- `apps-script/Code.gs`: the back end, a Google Apps Script web app. It stores everything in a shared Google Drive folder and a log sheet, and transcribes in the background (ElevenLabs Scribe or Gemini).

## Set up the back end

1. Paste `apps-script/Code.gs` into an Apps Script project and save.
2. In Project Settings > Script properties, add `TEAM_CODE`, `GEMINI_API_KEY`, `GEMINI_MODEL` (gemini-3.8-flash) and, optionally, `ELEVENLABS_API_KEY`.
3. Run `setup` once.
4. Deploy > New deployment > Web app, with Execute as Me and Who has access set to Anyone.
5. Put the web app URL in `config.js` as `window.HUB_API`.

People sign in on the website with their first name and the team code.
