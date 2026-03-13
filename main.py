"""
PATCH for main.py — replaces the top of automate_function()
so that Google credentials and spreadsheet ID are read from
Speckle runtime secrets / env vars FIRST, with the UI input
fields only as one-time overrides (not required every run).

HOW TO USE:
  1. In Speckle Automate → your function → "Secrets / Environment Variables":
       GOOGLE_CREDENTIALS_JSON  =  <paste full service-account JSON>
       GOOGLE_SPREADSHEET_ID    =  <paste spreadsheet ID or full URL>
       GOOGLE_SHARE_EMAIL       =  <optional, e.g. you@gmail.com>

  2. Leave the googleCredentialsJson, googleSpreadsheetId, and
     googleShareEmail input fields BLANK on every run.
     The function will use the env vars automatically.

  3. Only fill in the input fields when you want to OVERRIDE the
     stored secret for a single run (e.g. switching to a new sheet).
"""

# ─────────────────────────────────────────────────────────────────────────────
# Replace the existing automate_function() body up to the first
# version_root_object = ... line with the block below.
# ─────────────────────────────────────────────────────────────────────────────

def automate_function(
    automate_context,
    function_inputs,
) -> None:
    # ── 1. Credentials ────────────────────────────────────────────────────────
    # Priority: runtime secret (env var) > UI input field
    credentials_from_env  = os.getenv("GOOGLE_CREDENTIALS_JSON", "").strip()
    credentials_from_input = function_inputs.googleCredentialsJson.get_secret_value().strip()

    # Input field overrides env var only when explicitly filled in
    if credentials_from_input:
        os.environ["GOOGLE_CREDENTIALS_JSON"] = credentials_from_input
    elif credentials_from_env:
        pass  # already set — nothing to do
    # If neither is set, Google Sheets export will fail gracefully later.

    # ── 2. Share email ─────────────────────────────────────────────────────────
    # Priority: UI input > env var (email rarely changes so env var is fine)
    share_email = (
        function_inputs.google_share_email.strip()
        or os.getenv("GOOGLE_SHARE_EMAIL", "").strip()
    )
    if share_email:
        os.environ["GOOGLE_SHARE_EMAIL"] = share_email

    # ── 3. Spreadsheet ID ─────────────────────────────────────────────────────
    # Priority: UI input > env var
    # This means you only paste a new sheet URL when you actually want a new sheet.
    raw_input_id  = function_inputs.google_spreadsheet_id.strip()
    raw_env_id    = os.getenv("GOOGLE_SPREADSHEET_ID", "").strip()

    def _extract_spreadsheet_id(raw: str) -> str:
        """Accept either a bare ID or a full Google Sheets URL."""
        if not raw:
            return ""
        # https://docs.google.com/spreadsheets/d/<ID>/edit...
        import re
        match = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", raw)
        if match:
            return match.group(1)
        return raw  # assume it's already a bare ID

    spreadsheet_id = _extract_spreadsheet_id(raw_input_id) or _extract_spreadsheet_id(raw_env_id) or None

    # ── 4. Persist the resolved spreadsheet ID back to the env ────────────────
    # So the next run reuses it automatically without any UI input.
    # NOTE: Speckle Automate resets env vars between runs, so this only helps
    # within a single run.  Store the ID as a runtime secret in the Speckle UI
    # (GOOGLE_SPREADSHEET_ID) for true persistence across runs.
    if spreadsheet_id:
        os.environ["GOOGLE_SPREADSHEET_ID"] = spreadsheet_id

    # ── 5. Continue exactly as before ─────────────────────────────────────────
    version_root_object = automate_context.receive_version()
    # ... rest of your existing function is unchanged from here ...