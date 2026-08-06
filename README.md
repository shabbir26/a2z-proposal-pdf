# A2Z Proposal PDF service

A tiny web service that runs Shabbir's own proposal generator
(a2z_proposals_fpdf.py) and returns the PDF. His generator code is not
modified - it is imported and called exactly as-is.

Endpoints:
- POST /proposal   (JSON in, PDF out)
- GET  /health     (returns {"ok": true})

## Deploy on Render (free)
1. Put every file in this folder into a GitHub repository.
2. On Render: New -> Web Service -> connect the repository.
3. Settings:
   - Runtime: Python 3
   - Build command:  pip install -r requirements.txt
   - Start command:  gunicorn app:app
4. Deploy. Your live URL will look like:
   https://a2z-proposal-pdf.onrender.com
5. Send that URL back and the platform's proposal button will be wired to it.

Note: the free plan sleeps after 15 minutes idle, so the first request
after a quiet spell takes ~30-50 seconds to wake, then it is instant.
