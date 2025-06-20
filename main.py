import re
import pandas as pd
from datetime import datetime

CHAT_FILE = "chat.txt"
OUTPUT_EXCEL = "extracted_links.xlsx"

DATE_PATTERN = r'^(\d{2}/\d{2}/\d{2}), (\d{1,2}:\d{2}\s?[apAP][mM]) - (.*?): (.*)$'
URL_PATTERN = r'(https?://[^\s]+)'

entries = []

with open(CHAT_FILE, "r", encoding="utf-8") as file:
    for line in file:
        match = re.match(DATE_PATTERN, line)
        if match:
            date_str, time_str, sender, message = match.groups()
            links = re.findall(URL_PATTERN, message)
            for url in links:
                try:
                    dt = datetime.strptime(f"{date_str} {time_str.lower()}", "%d/%m/%y %I:%M %p")
                    entries.append({
                        "Date": dt.date(),
                        "Time": dt.time(),
                        "Sender": sender.strip(),
                        "URL": url.strip()  # Full URL text
                    })
                except Exception as e:
                    print(f"Failed to parse line: {line.strip()} | Error: {e}")

# Write raw URL (Excel will auto-detect and style as hyperlink)
df = pd.DataFrame(entries)

with pd.ExcelWriter(OUTPUT_EXCEL, engine='openpyxl') as writer:
    df.to_excel(writer, index=False, sheet_name="Links")

print(f"[✓] Exported {len(df)} full clickable links to '{OUTPUT_EXCEL}'")
