import re
import pandas as pd
from datetime import datetime
import aiohttp
import asyncio
import async_timeout
from bs4 import BeautifulSoup
from urllib.parse import urlparse
from tqdm import tqdm
import time
import sys

CHAT_FILE = "chat.txt"
OUTPUT_EXCEL = "extracted_links.xlsx"

# Regex patterns
DATE_PATTERN = (
    r"^(\d{1,2}/\d{1,2}/\d{2,4}),\s*(\d{1,2}:\d{2}\s*[apAP][mM])\s*-\s*(.*?):\s*(.*)$"
)
URL_PATTERN = r"(https?://[^\s]+)"

# Headers for HTTP requests
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
}

# Cache for storing fetched titles and descriptions
FETCH_CACHE = {}


def animate_text(text, delay=0.05):
    """Print text with typewriter effect"""
    for char in text:
        print(char, end="", flush=True)
        time.sleep(delay)
    print()


def clean_time_string(time_str):
    """Clean time string by removing Unicode spaces and normalizing format"""
    time_str = (
        time_str.replace("\u202f", " ")
        .replace("\u00a0", " ")
        .replace("\u2009", " ")
        .replace("\u200a", " ")
    )
    time_str = re.sub(r"\s+", " ", time_str.strip())
    time_str = re.sub(r"(\d)([apAP][mM])", r"\1 \2", time_str)
    return time_str


def parse_datetime(date_str, time_str):
    """Parse datetime with multiple format attempts"""
    time_str = clean_time_string(time_str)
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
            return datetime.strptime(f"{date_str} {time_str}".lower(), fmt)
        except ValueError:
            continue
    try:
        time_match = re.search(r"(\d{1,2}:\d{2})", time_str)
        if time_match:
            return datetime.strptime(
                f"{date_str} {time_match.group(1)}", "%d/%m/%y %H:%M"
            )
    except ValueError:
        pass
    raise ValueError(f"Unable to parse datetime: '{date_str} {time_str}'")


async def fetch_title_description(session, url, semaphore):
    """Fetch title and description asynchronously"""
    # Check cache first
    if url in FETCH_CACHE:
        return FETCH_CACHE[url]

    try:
        # Clean URL
        url = url.rstrip(".,;!?")
        parsed = urlparse(url)
        if not parsed.scheme or not parsed.netloc:
            return "Invalid URL", "Invalid URL format"

        async with semaphore:
            async with async_timeout.timeout(15):
                async with session.get(url, headers=HEADERS) as response:
                    if response.status != 200:
                        return (
                            f"HTTP Error {response.status}",
                            f"HTTP error occurred: {response.status}",
                        )
                    text = await response.text()

        soup = BeautifulSoup(text, "html.parser")

        # Extract title
        title = "N/A"
        if soup.title and soup.title.string:
            title = soup.title.string.strip()
        elif soup.find("meta", attrs={"property": "og:title"}):
            title = (
                soup.find("meta", attrs={"property": "og:title"})
                .get("content", "N/A")
                .strip()
            )

        # Extract description
        description = "N/A"
        desc_tag = (
            soup.find("meta", attrs={"name": "description"})
            or soup.find("meta", attrs={"property": "og:description"})
            or soup.find("meta", attrs={"name": "Description"})
        )
        if desc_tag and desc_tag.get("content"):
            description = desc_tag["content"].strip()

        # Cache the result
        FETCH_CACHE[url] = (title, description)
        return title, description

    except asyncio.TimeoutError:
        return "Timeout Error", "Request timed out"
    except aiohttp.ClientConnectionError:
        return "Connection Error", "Failed to connect to URL"
    except Exception as e:
        return "Fetch Error", f"Error: {str(e)[:100]}..."


async def fetch_all_urls(urls_data, max_concurrent=10):
    """Fetch titles and descriptions for all URLs concurrently"""
    semaphore = asyncio.Semaphore(max_concurrent)
    async with aiohttp.ClientSession() as session:
        tasks = []
        for url_data in urls_data:
            tasks.append(fetch_title_description(session, url_data["url"], semaphore))
        results = []
        for task, url_data in tqdm(
            zip(asyncio.as_completed(tasks), urls_data),
            total=len(urls_data),
            desc="Fetching URLs",
        ):
            title, description = await task
            results.append(
                {
                    "Sender": url_data["sender"],
                    "Date": url_data["datetime"].date(),
                    "URL": url_data["url"],
                    "Title": title,
                    "Description": description,
                    "Time": url_data["datetime"].time(),
                    "Line_Number": url_data["line_num"],
                    "Original_Line": url_data["original_line"],
                }
            )
        return results


def main():
    animate_text("🚀 Starting WhatsApp Chat Link Extractor...", 0.03)
    time.sleep(0.5)

    # Read chat file
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

    for line_num, line in enumerate(tqdm(lines, desc="Parsing lines")):
        line = line.strip()
        if not line:
            continue
        match = re.match(DATE_PATTERN, line)
        if match:
            date_str, time_str, sender, message = match.groups()
            links = re.findall(URL_PATTERN, message)
            if links:
                try:
                    dt = parse_datetime(date_str, time_str)
                    for url in links:
                        all_urls.append(
                            {
                                "url": url,
                                "sender": sender.strip(),
                                "datetime": dt,
                                "line_num": line_num + 1,
                                "original_line": line[:100] + "..."
                                if len(line) > 100
                                else line,
                            }
                        )
                except ValueError as e:
                    failed_lines.append(
                        {
                            "line_num": line_num + 1,
                            "error": str(e),
                            "line": line[:100] + "..." if len(line) > 100 else line,
                        }
                    )

    print(f"\n✅ Found {len(all_urls)} URLs to process")
    if failed_lines:
        print(f"⚠️  {len(failed_lines)} lines couldn't be parsed (saved to debug file)")
        with open("failed_lines.txt", "w", encoding="utf-8") as f:
            f.write("Failed to parse these lines:\n\n")
            for fail in failed_lines:
                f.write(f"Line {fail['line_num']}: {fail['error']}\n")
                f.write(f"Content: {fail['line']}\n\n")

    if not all_urls:
        print("❌ No URLs found in the chat file!")
        return

    time.sleep(0.5)

    # Fetch titles and descriptions asynchronously
    print("\n🌐 Fetching website information...")
    entries = asyncio.run(fetch_all_urls(all_urls, max_concurrent=10))

    # Export to Excel
    print("\n💾 Exporting to Excel...")
    try:
        df = pd.DataFrame(entries)
        df_sorted = df.sort_values(["Date", "Time"])
        with pd.ExcelWriter(OUTPUT_EXCEL, engine="openpyxl") as writer:
            df_sorted.to_excel(writer, index=False, sheet_name="Links")
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
        print("\n📊 Summary:")
        print(f"   • Total links processed: {len(df)}")
        print(f"   • Unique senders: {df['Sender'].nunique()}")
        print(f"   • Date range: {df['Date'].min()} to {df['Date'].max()}")
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
