import asyncio
import json
import os
from datetime import datetime
from playwright.async_api import async_playwright

# Setup Logger to write to both file and console
class Logger:
    def __init__(self, filename):
        self.terminal = open(filename, "w", encoding="utf-8")

    def log(self, message):
        print(message)
        self.terminal.write(message + "\n")
        self.terminal.flush()

async def run(url):
    # Create debug folder if not exists
    if not os.path.exists("debug"):
        os.makedirs("debug")

    # Generate unique filename for this run
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = os.path.join("debug", f"log_{timestamp}.txt")
    har_file = os.path.join("debug", f"network_{timestamp}.har")
    logger = Logger(log_file)

    async with async_playwright() as p:
        logger.log(f"--- Session Started at {timestamp} ---")
        logger.log(f"--- Logging text to: {log_file} ---")
        logger.log(f"--- Recording HAR (Full Network Debug) to: {har_file} ---")
        
        # Launch browser
        browser = await p.chromium.launch(headless=False)
        # Record HAR to fully capture everything (Network tab format)
        context = await browser.new_context(record_har_path=har_file)

        # 1. Capture Console Logs
        page = await context.new_page()
        page.on("console", lambda msg: logger.log(f"  [CONSOLE] {msg.type}: {msg.text}"))

        # 2. Capture Full Network Request
        async def handle_request(request):
            # Focus on XHR/Fetch/Document (ignore images, css, js files for cleaner text log)
            # But the HAR file will still capture EVERYTHING.
            if request.resource_type in ["fetch", "xhr", "document"]:
                logger.log(f"\n[>>> REQ] {request.method} {request.url}")
                logger.log(f"      Resource Type: {request.resource_type}")
                
                # Log Request Headers
                headers = await request.all_headers()
                if headers:
                    logger.log("      Request Headers:")
                    for k, v in headers.items():
                        logger.log(f"        {k}: {v}")

                # Log Request Payload / Post Data
                post_data = request.post_data
                if post_data:
                    try:
                        data = json.loads(post_data)
                        logger.log(f"      Request Payload (JSON):\n{json.dumps(data, indent=4)}")
                    except:
                        logger.log(f"      Request Payload (Raw):\n{post_data}")

        # 3. Capture Full Network Response
        async def handle_response(response):
            if response.request.resource_type in ["fetch", "xhr", "document"]:
                logger.log(f"\n[<<< RES] {response.status} {response.url}")
                
                # Log Response Headers
                headers = await response.all_headers()
                if headers:
                    logger.log("      Response Headers:")
                    for k, v in headers.items():
                        logger.log(f"        {k}: {v}")

                # Log Response Body
                try:
                    content_type = headers.get("content-type", "")
                    if "application/json" in content_type or "text" in content_type:
                        body = await response.text()
                        try:
                            # Attempt to prettify JSON
                            data = json.loads(body)
                            logger.log(f"      Response Body (JSON):\n{json.dumps(data, indent=4)}")
                        except:
                            # Fallback to plain text
                            logger.log(f"      Response Body (Text):\n{body}")
                except Exception as e:
                    logger.log(f"      [Cannot read response body: {e}]")

        page.on("request", handle_request)
        page.on("response", handle_response)

        logger.log(f"\n--- Navigating to: {url} ---")
        
        try:
            await page.goto(url, wait_until="networkidle", timeout=60000)
            logger.log("--- Waiting for extra network activity ---")
            await page.wait_for_timeout(5000)
            
            # Final Screenshot in debug folder
            screenshot_path = os.path.join("debug", f"screenshot_{timestamp}.png")
            await page.screenshot(path=screenshot_path)
            logger.log(f"\n--- Done. Screenshot saved to {os.path.abspath(screenshot_path)} ---")
            
        except Exception as e:
            logger.log(f"\n[ERROR] Navigation failed: {e}")
            error_screenshot = os.path.join("debug", f"error_{timestamp}.png")
            await page.screenshot(path=error_screenshot)
            logger.log(f"--- Error screenshot saved as {error_screenshot} ---")
        
        logger.log("\nBrowser remains open for manual inspection.")
        logger.log("NOTE: HAR file is being recorded. It will be finalized when you press Enter.")
        logger.log("Press Enter in this terminal to close browser and save HAR file...")
        await asyncio.get_event_loop().run_in_executor(None, input)
        
        # Important: Close context to flush and save the HAR file
        await context.close()
        await browser.close()
        logger.terminal.close()

if __name__ == "__main__":
    default_url = "https://coretaxdjp.pajak.go.id/registration-portal/id-ID/reg-home"
    target = input(f"Masukkan URL yang ingin di-debug [Default: {default_url}]: ").strip()
    
    if not target:
        target = default_url
        
    if not target.startswith("http"):
        target = "https://" + target
        
    asyncio.run(run(target))
