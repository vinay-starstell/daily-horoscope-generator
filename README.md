# Daily Horoscope Generator

Automatically generates daily horoscope Excel files using the DivineAPI and emails them to Vinay.

## Setup

1. Install Composer dependencies: `composer install`
2. Copy `.env.example` to `.env` and fill in your API credentials
3. Run: `php src/DailyPredictionsAggregator.php`

## GitHub Actions

The workflow runs on the 22nd of each month at 9 AM UTC and emails the generated file.

### Required Secrets (add in GitHub repo settings):
- `DIVINEAPI_AUTH_TOKEN`
- `DIVINEAPI_KEY`
- `GMAIL_ADDRESS`
- `GMAIL_APP_PASSWORD`
