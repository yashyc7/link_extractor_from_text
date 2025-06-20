import re
import pandas as pd
from datetime import datetime
import requests
from bs4 import BeautifulSoup
import time
import sys
from urllib.parse import urlparse

CHAT_FILE = "chat.txt"
OUTPUT_EXCEL = "extracted_links.xlsx"

# Updated regex pattern to handle various WhatsApp export formats
DATE_PATTERN = (
    r"^(\d{1,2}/\d{1,2}/\d{2,4}),\s*(\d{1,2}:\d{2}\s*[apAP][mM])\s*-\s*(.*?):\s*(.*)$"
)
URL_PATTERN = r"(https?://[^\s]+)"

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
}


class ProgressAnimation:
    def __init__(self):
        self.spinner_chars = "⠋⠙⠹⠸⠼⠴⠦⠧⠇⠏"
        self.current = 0

    def spin(self):
        char = self.spinner_chars[self.current % len(self.spinner_chars)]
        self.current += 1
        return char

    def progress_bar(self, current, total, width=40):
        percentage = current / total
        filled = int(width * percentage)
        bar = "█" * filled + "░" * (width - filled)
        return f"[{bar}] {current}/{total} ({percentage:.1%})"


def animate_text(text, delay=0.05):
    """Print text with typewriter effect"""
    for char in text:
        print(char, end="", flush=True)
        time.sleep(delay)
    print()


def clean_time_string(time_str):
    """Clean time string by removing Unicode spaces and normalizing format"""
    # Remove various Unicode space characters
    time_str = time_str.replace("\u202f", " ")  # Narrow no-break space
    time_str = time_str.replace("\u00a0", " ")  # Non-breaking space
    time_str = time_str.replace("\u2009", " ")  # Thin space
    time_str = time_str.replace("\u200a", " ")  # Hair space

    # Normalize multiple spaces to single space
    time_str = re.sub(r"\s+", " ", time_str.strip())

    # Ensure there's a space before AM/PM
    time_str = re.sub(r"(\d)([apAP][mM])", r"\1 \2", time_str)

    return time_str


def parse_datetime(date_str, time_str):
    """Parse datetime with multiple format attempts"""
    # Clean the time string
    time_str = clean_time_string(time_str)

    # Try different datetime formats
    datetime_formats = [
        "%d/%m/%y %I:%M %p",
        "%d/%m/%Y %I:%M %p",
        "%m/%d/%y %I:%M %p",
        "%m/%d/%Y %I:%M %p",
        "%d/%m/%y %H:%M",
        "%d/%m/%Y %H:%M",
    ]

    for fmt in datetime_formats:
        try:
            full_datetime_str = f"{date_str} {time_str}".lower()
            return datetime.strptime(full_datetime_str, fmt)
        except ValueError:
            continue

    # If all formats fail, try without AM/PM
    try:
        # Extract just the time part and assume 24-hour format
        time_match = re.search(r"(\d{1,2}:\d{2})", time_str)
        if time_match:
            clean_time = time_match.group(1)
            return datetime.strptime(f"{date_str} {clean_time}", "%d/%m/%y %H:%M")
    except ValueError:
        pass

    raise ValueError(f"Unable to parse datetime: '{date_str} {time_str}'")


def fetch_title_description(url, progress_anim):
    """Fetch title and description with progress indication"""
    try:
        # Clean URL - remove trailing punctuation that might interfere
        url = url.rstrip(".,;!?")

        # Validate URL
        parsed = urlparse(url)
        if not parsed.scheme or not parsed.netloc:
            return "Invalid URL", "Invalid URL format"

        spinner = progress_anim.spin()
        print(
            f"\r{spinner} Fetching: {url[:60]}{'...' if len(url) > 60 else ''}",
            end="",
            flush=True,
        )

        response = requests.get(url, headers=headers, timeout=15)
        response.raise_for_status()

        soup = BeautifulSoup(response.text, "html.parser")

        # Extract title
        title = "N/A"
        if soup.title and soup.title.string:
            title = soup.title.string.strip()
        elif soup.find("meta", attrs={"property": "og:title"}):
            og_title = soup.find("meta", attrs={"property": "og:title"})
            title = og_title.get("content", "N/A").strip()

        # Extract description
        description = "N/A"
        desc_tag = (
            soup.find("meta", attrs={"name": "description"})
            or soup.find("meta", attrs={"property": "og:description"})
            or soup.find("meta", attrs={"name": "Description"})
        )

        if desc_tag and desc_tag.get("content"):
            description = desc_tag["content"].strip()

        return title, description

    except requests.exceptions.Timeout:
        return "Timeout Error", "Request timed out"
    except requests.exceptions.ConnectionError:
        return "Connection Error", "Failed to connect to URL"
    except requests.exceptions.HTTPError as e:
        return f"HTTP Error {e.response.status_code}", f"HTTP error occurred: {e}"
    except Exception as e:
        return "Fetch Error", f"Error: {str(e)[:100]}..."


def main():
    # Animated startup
    animate_text("🚀 Starting WhatsApp Chat Link Extractor...", 0.03)
    time.sleep(0.5)

    progress_anim = ProgressAnimation()
    entries = []

    # Check if chat file exists
    try:
        with open(CHAT_FILE, "r", encoding="utf-8") as file:
            lines = file.readlines()
    except FileNotFoundError:
        print(f"❌ Error: '{CHAT_FILE}' not found!")
        return
    except Exception as e:
        print(f"❌ Error reading file: {e}")
        return

    print(f"📁 Found {len(lines)} lines to process")
    time.sleep(0.5)

    # Parse lines and extract URLs
    print("📝 Parsing chat messages...")
    all_urls = []
    failed_lines = []

    for line_num, line in enumerate(lines, 1):
        spinner = progress_anim.spin()
        print(f"\r{spinner} Parsing line {line_num}/{len(lines)}", end="", flush=True)

        line = line.strip()
        if not line:
            continue

        match = re.match(DATE_PATTERN, line)
        if match:
            date_str, time_str, sender, message = match.groups()
            links = re.findall(URL_PATTERN, message)

            if links:  # Only process if there are links
                try:
                    dt = parse_datetime(date_str, time_str)

                    for url in links:
                        all_urls.append(
                            {
                                "url": url,
                                "sender": sender.strip(),
                                "datetime": dt,
                                "line_num": line_num,
                                "original_line": line[:100] + "..."
                                if len(line) > 100
                                else line,
                            }
                        )

                except ValueError as e:
                    failed_lines.append(
                        {
                            "line_num": line_num,
                            "error": str(e),
                            "line": line[:100] + "..." if len(line) > 100 else line,
                        }
                    )
                    continue

    print(f"\n✅ Found {len(all_urls)} URLs to process")

    if failed_lines:
        print(f"⚠️  {len(failed_lines)} lines couldn't be parsed (saved to debug file)")
        # Save failed lines for debugging
        with open("failed_lines.txt", "w", encoding="utf-8") as f:
            f.write("Failed to parse these lines:\n\n")
            for fail in failed_lines:
                f.write(f"Line {fail['line_num']}: {fail['error']}\n")
                f.write(f"Content: {fail['line']}\n\n")

    if not all_urls:
        print("❌ No URLs found in the chat file!")
        return

    time.sleep(0.5)

    # Fetch titles and descriptions
    print("\n🌐 Fetching website information...")

    for i, url_data in enumerate(all_urls):
        # Progress bar
        progress_bar = progress_anim.progress_bar(i, len(all_urls))
        print(f"\r{progress_bar}", end="", flush=True)

        title, description = fetch_title_description(url_data["url"], progress_anim)

        entries.append(
            {
                "Sender": url_data["sender"],
                "URL": url_data["url"],
                "Title": title,
                "Description": description,
                "Date": url_data["datetime"].date(),
                "Time": url_data["datetime"].time(),
                "Line_Number": url_data["line_num"],
                "Original_Line": url_data["original_line"],
            }
        )

        # Small delay to be respectful to servers
        time.sleep(0.1)

    # Final progress update
    progress_bar = progress_anim.progress_bar(len(all_urls), len(all_urls))
    print(f"\r{progress_bar}")

    # Export to Excel
    print("\n💾 Exporting to Excel...")
    try:
        df = pd.DataFrame(entries)

        # Sort by datetime
        df_sorted = df.sort_values(["Date", "Time"])

        with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:
            df_sorted.to_excel(writer, index=False, sheet_name="Links")

            # Auto-adjust column widths
            worksheet = writer.sheets["Links"]
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width

        print(f"✅ Successfully exported {len(df)} links to '{OUTPUT_EXCEL}'")

        # Summary statistics
        print("\n📊 Summary:")
        print(f"   • Total links processed: {len(df)}")
        print(f"   • Unique senders: {df['Sender'].nunique()}")
        print(f"   • Date range: {df['Date'].min()} to {df['Date'].max()}")

        # Show top senders
        top_senders = df["Sender"].value_counts().head(3)
        print(f"   • Top link sharers:")
        for sender, count in top_senders.items():
            print(f"     - {sender}: {count} links")

    except Exception as e:
        print(f"❌ Error exporting to Excel: {e}")
        return

    animate_text("\n🎉 All done! Check your Excel file.", 0.03)


if __name__ == "__main__":
    main()
