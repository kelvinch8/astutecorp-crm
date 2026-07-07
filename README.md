# AstuteCorp CRM - Corporate Service Provider Deadline Manager

A simple, self-hosted CRM for corporate service providers to track client deadlines and maintain activity notes.

## Features

- Client management with UEN, contact details, and FYE
- Track key deadlines: AGM, Annual Return, ECI, Tax Filing
- Automatic suggestion of deadlines based on Financial Year End
- Activity / Notes log (great for recording ACRA updates, passport changes, etc.)
- Email reminders (to you + client) when deadlines approach
- Manual "Send Reminders Now" button + background daily check

## Quick Start

```bash
cd ~/AstuteCorpCRM
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python3 app.py
```

Visit: http://127.0.0.1:5001

Default login is not required (simple internal tool).

## Email Configuration

1. Go to **Settings**
2. Enter your email + SMTP details
3. For Gmail: Use an App Password (recommended)

## Recommended Deadlines (Singapore)

| Type          | Typical Timing                  |
|---------------|---------------------------------|
| ECI           | 3 months after FYE              |
| AGM           | Within 6 months of FYE          |
| Annual Return | ~1 month after AGM              |
| Tax Filing    | Usually 15 Nov / 15 Dec         |

You can always edit dates manually per client.

## Tips

- Use the **Calculate Defaults from FYE** button on a client's page.
- Add notes every time you do work for the client (very useful for audit trail).
- Click **Send Reminders Now** in the navbar regularly.

Built for Singapore corporate service providers.
