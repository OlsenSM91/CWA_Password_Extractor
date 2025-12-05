import requests
import pandas as pd
import argparse
import getpass
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
import time
import json
from datetime import datetime
import sys
import os
from pathlib import Path
import subprocess
import platform
import winreg

# ANSI Color codes
class Colors:
    RED = '\033[91m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    BLUE = '\033[94m'
    MAGENTA = '\033[95m'
    CYAN = '\033[96m'
    WHITE = '\033[97m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    RESET = '\033[0m'

def print_banner():
    """Print the CNS4U ASCII banner"""
    banner = f"""{Colors.CYAN}{Colors.BOLD}
 $$$$$$\\  $$\\   $$\\  $$$$$$\\  $$\\   $$\\ $$\\   $$\\
$$  __$$\\ $$$\\  $$ |$$  __$$\\ $$ |  $$ |$$ |  $$ |
$$ /  \\__|$$$$\\ $$ |$$ /  \\__|$$ |  $$ |$$ |  $$ |
$$ |      $$ $$\\$$ |\\$$$$$$\\  $$$$$$$$ |$$ |  $$ |
$$ |      $$ \\$$$$ | \\____$$\\ \\_____$$ |$$ |  $$ |
$$ |  $$\\ $$ |\\$$$ |$$\\   $$ |      $$ |$$ |  $$ |
\\$$$$$$  |$$ | \\$$ |\\$$$$$$  |      $$ |\\$$$$$$  |
 \\______/ \\__|  \\__| \\______/       \\__| \\______/
{Colors.RESET}
{Colors.YELLOW}    ConnectWise Automate - Client Offboarding Tool{Colors.RESET}
{Colors.WHITE}              Password Extraction Utility{Colors.RESET}
"""
    print(banner)

def print_success(message):
    """Print success message in green"""
    print(f"{Colors.GREEN}✓ {message}{Colors.RESET}")

def print_error(message):
    """Print error message in red"""
    print(f"{Colors.RED}✗ {message}{Colors.RESET}")

def print_warning(message):
    """Print warning message in yellow"""
    print(f"{Colors.YELLOW}⚠ {message}{Colors.RESET}")

def print_info(message):
    """Print info message in cyan"""
    print(f"{Colors.CYAN}ℹ {message}{Colors.RESET}")

def print_section(title):
    """Print section header"""
    print(f"\n{Colors.BOLD}{Colors.MAGENTA}{'='*80}{Colors.RESET}")
    print(f"{Colors.BOLD}{Colors.MAGENTA}{title}{Colors.RESET}")
    print(f"{Colors.BOLD}{Colors.MAGENTA}{'='*80}{Colors.RESET}\n")

def check_chrome_installed():
    """Check if Google Chrome is installed"""
    try:
        if platform.system() == 'Windows':
            # Check common Chrome installation paths
            paths = [
                r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
                os.path.expandvars(r"%PROGRAMFILES%\Google\Chrome\Application\chrome.exe"),
                os.path.expandvars(r"%PROGRAMFILES(X86)%\Google\Chrome\Application\chrome.exe")
            ]

            for path in paths:
                if os.path.exists(path):
                    return True, path

            # Also check registry
            try:
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe", 0, winreg.KEY_READ)
                chrome_path = winreg.QueryValue(key, None)
                winreg.CloseKey(key)
                if os.path.exists(chrome_path):
                    return True, chrome_path
            except:
                pass

        return False, None

    except Exception as e:
        return False, None

def is_running_as_exe():
    """Check if running as PyInstaller EXE"""
    return getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS')

def check_and_fix_dependencies():
    """Check for required dependencies and auto-fix if possible"""
    print_section("Dependency Check")

    # Check for Google Chrome
    print_info("Checking for Google Chrome...")
    chrome_installed, chrome_path = check_chrome_installed()

    if chrome_installed:
        print_success(f"Google Chrome found at: {chrome_path}")
    else:
        print_error("Google Chrome is NOT installed!")
        print(f"\n{Colors.YELLOW}Please install Google Chrome:{Colors.RESET}")
        print(f"{Colors.WHITE}Download from: https://www.google.com/chrome/{Colors.RESET}")
        print(f"\n{Colors.RED}Cannot continue without Chrome.{Colors.RESET}")
        return False

    # Check Selenium version (need 4.6+ for built-in driver management)
    print_info("Checking Selenium version...")
    try:
        import selenium
        from packaging import version

        selenium_version = version.parse(selenium.__version__)
        min_version = version.parse("4.6.0")

        if selenium_version >= min_version:
            print_success(f"Selenium {selenium.__version__} - Has built-in ChromeDriver management!")
        else:
            print_warning(f"Selenium {selenium.__version__} - Upgrade recommended for automatic ChromeDriver")
            print_info("Run: pip install --upgrade selenium")
    except Exception as e:
        print_warning(f"Could not check Selenium version: {e}")

    print(f"\n{Colors.BOLD}{Colors.GREEN}{'='*80}{Colors.RESET}")
    print(f"{Colors.BOLD}{Colors.GREEN}✓ All dependencies are ready!{Colors.RESET}")
    print(f"{Colors.BOLD}{Colors.GREEN}{'='*80}{Colors.RESET}\n")
    print_info("ChromeDriver will be automatically managed by Selenium")

    return True

class CWAOffboarding:
    def __init__(self, base_url=None, output_dir=None):
        # Normalize base URL - remove any /Automate or /automate suffix
        if base_url:
            base_url = base_url.rstrip('/')
            if base_url.lower().endswith('/automate'):
                base_url = base_url[:-9]  # Remove '/automate'

        self.base_url = base_url
        self.bearer_token = None
        self.clientid = None
        self.headers = None

        # Set output directory (default to current working directory)
        if output_dir:
            self.output_dir = Path(output_dir)
        else:
            # Use current working directory (where the exe is run from)
            self.output_dir = Path.cwd()

        # Create output directory if it doesn't exist
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def get_credentials_automated(self):
        """
        Open browser for manual login and extract bearer token using Selenium
        """
        print_section("Automated Bearer Token Extraction")
        print_info("This will open a browser window for you to log in manually.")
        print_info("After logging in, the tool will automatically extract your bearer token.")
        input(f"\n{Colors.BOLD}Press Enter to continue...{Colors.RESET}")

        # Setup Chrome options
        chrome_options = Options()
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)

        # Enable performance logging to capture network requests
        chrome_options.set_capability('goog:loggingPrefs', {'performance': 'ALL'})

        print_info("Opening browser...")
        print_info("Selenium will automatically download ChromeDriver if needed...")

        # Selenium 4.6+ automatically manages ChromeDriver
        # No need for webdriver-manager!
        try:
            driver = webdriver.Chrome(options=chrome_options)
        except Exception as e:
            print_error(f"Failed to start Chrome: {str(e)}")
            print_warning("If ChromeDriver is missing, Selenium should auto-download it.")
            print_warning("Ensure you have internet connection for first-time setup.")
            raise

        try:
            # Navigate to login page
            login_url = f"{self.base_url}/Automate"
            print_info(f"Navigating to {login_url}...")

            print(f"\n{Colors.BOLD}{Colors.CYAN}{'='*80}{Colors.RESET}")
            print(f"{Colors.BOLD}{Colors.CYAN}PLEASE LOG IN TO CONNECTWISE AUTOMATE IN THE BROWSER WINDOW{Colors.RESET}")
            print(f"{Colors.BOLD}{Colors.CYAN}{'='*80}{Colors.RESET}")
            print(f"\n{Colors.YELLOW}Instructions:{Colors.RESET}")
            print(f"{Colors.WHITE}1. Complete the login process (including MFA if prompted){Colors.RESET}")
            print(f"{Colors.WHITE}2. The tool will automatically detect when you're logged in{Colors.RESET}")
            print(f"{Colors.BOLD}{Colors.CYAN}{'='*80}{Colors.RESET}\n")

            driver.get(login_url)
            initial_url = driver.current_url

            # Wait for user to log in by monitoring URL changes
            print_warning("Waiting for you to log in...")
            wait = WebDriverWait(driver, 300)  # 5 minute timeout for login

            try:
                # More specific check: wait for URL to change AND contain browse/companies
                # This ensures we don't trigger on intermediate redirect pages
                def check_logged_in(driver):
                    current_url = driver.current_url.lower()
                    # Must have changed from login URL AND contain a known post-login path
                    url_changed = current_url != initial_url.lower()
                    has_browse = '/automate/browse/companies' in current_url
                    has_dashboard = '/dashboard' in current_url

                    # Also check if we can find typical post-login elements
                    is_logged_in = url_changed and (has_browse or has_dashboard)

                    return is_logged_in

                wait.until(check_logged_in)

                print_success("Login detected!")
                time.sleep(2)  # Brief pause to ensure session is established

            except Exception as e:
                print_warning(f"Timeout waiting for login. Current URL: {driver.current_url}")
                print_warning("Proceeding anyway...")

            # Clear any existing logs
            driver.get_log('performance')

            # Navigate to the computers page first
            print_info("Navigating to company computers page...")
            computers_url = f"{self.base_url}/automate/browse/companies/computers"
            driver.get(computers_url)
            time.sleep(3)

            # Now navigate to a company passwords page to trigger the API call
            print_info("Navigating to company passwords page to capture bearer token...")
            print_info("Trying company ID 377...")

            test_company_url = f"{self.base_url}/automate/browse/companies/company-passwords?companyId=377"
            driver.get(test_company_url)

            # Wait for page to load and API calls to be made
            print_warning("Waiting for API calls to be captured...")
            time.sleep(5)

            # Extract bearer token from performance logs
            print_info("Extracting bearer token from network requests...")
            logs = driver.get_log('performance')

            bearer_token = None
            clientid = None

            for log in logs:
                try:
                    log_data = json.loads(log['message'])
                    message = log_data['message']

                    # Look for network request events
                    if message['method'] == 'Network.requestWillBeSent':
                        request = message['params']['request']

                        # Check if it's a deploymentlogins API call
                        if 'deploymentlogins' in request['url'].lower():
                            headers = request.get('headers', {})

                            # Extract authorization header (case-insensitive)
                            for key, value in headers.items():
                                if key.lower() == 'authorization':
                                    bearer_token = value
                                elif key.lower() == 'clientid':
                                    clientid = value

                            if bearer_token and clientid:
                                break

                except Exception as e:
                    continue

            # If not found with company 377, try a few more companies
            if not bearer_token or not clientid:
                print_warning("Bearer token not found with company 377. Trying other companies...")

                for company_id in [1, 100, 438, 500]:
                    print_info(f"Trying company ID {company_id}...")

                    # Clear logs
                    driver.get_log('performance')

                    test_url = f"{self.base_url}/automate/browse/companies/company-passwords?companyId={company_id}"
                    driver.get(test_url)
                    time.sleep(3)

                    logs = driver.get_log('performance')

                    for log in logs:
                        try:
                            log_data = json.loads(log['message'])
                            message = log_data['message']

                            if message['method'] == 'Network.requestWillBeSent':
                                request = message['params']['request']

                                if 'deploymentlogins' in request['url'].lower():
                                    headers = request.get('headers', {})

                                    for key, value in headers.items():
                                        if key.lower() == 'authorization':
                                            bearer_token = value
                                        elif key.lower() == 'clientid':
                                            clientid = value

                                    if bearer_token and clientid:
                                        break

                        except Exception as e:
                            continue

                    if bearer_token and clientid:
                        break

            if bearer_token and clientid:
                print(f"\n{Colors.BOLD}{Colors.GREEN}{'='*80}{Colors.RESET}")
                print(f"{Colors.BOLD}{Colors.GREEN}✓ Successfully extracted credentials!{Colors.RESET}")
                print(f"{Colors.BOLD}{Colors.GREEN}{'='*80}{Colors.RESET}")
                print(f"{Colors.CYAN}Bearer Token: {Colors.WHITE}{bearer_token[:30]}...{Colors.RESET}")
                print(f"{Colors.CYAN}Client ID: {Colors.WHITE}{clientid}{Colors.RESET}")
                print(f"{Colors.BOLD}{Colors.GREEN}{'='*80}{Colors.RESET}")

                self.bearer_token = bearer_token
                self.clientid = clientid
                self.headers = {
                    "authorization": bearer_token,
                    "clientid": clientid,
                    "accept": "application/json"
                }

                return True
            else:
                print(f"\n{Colors.BOLD}{Colors.RED}{'='*80}{Colors.RESET}")
                print_error("Failed to extract bearer token from network requests.")
                print(f"{Colors.BOLD}{Colors.RED}{'='*80}{Colors.RESET}")
                print(f"\n{Colors.YELLOW}Possible reasons:{Colors.RESET}")
                print(f"{Colors.WHITE}1. The company might not have any passwords to load{Colors.RESET}")
                print(f"{Colors.WHITE}2. The API call might not have been made yet{Colors.RESET}")
                print(f"{Colors.WHITE}3. Network logging might not have captured the request{Colors.RESET}")
                print(f"\n{Colors.YELLOW}Please try manual extraction method.{Colors.RESET}")
                return False

        except Exception as e:
            print_error(f"Error during automated extraction: {str(e)}")
            print_warning("You may need to manually extract the bearer token.")
            import traceback
            traceback.print_exc()
            return False

        finally:
            print_info("Closing browser...")
            driver.quit()

    def get_credentials_manual(self):
        """
        Manual method for getting credentials
        """
        print_section("Manual Bearer Token Entry")
        self.clientid = input(f"{Colors.CYAN}Enter your Client ID: {Colors.RESET}")
        self.bearer_token = input(f"{Colors.CYAN}Enter your Bearer Token (e.g., 'bearer abc123...'): {Colors.RESET}")

        self.headers = {
            "authorization": self.bearer_token,
            "clientid": self.clientid,
            "accept": "application/json"
        }

        return True

    def get_all_clients(self):
        """
        Retrieve all clients from CWA
        """
        print_info("Fetching client list...")
        clientsurl = f"{self.base_url}/cwa/api/v1/clients?pageSize=-1&includeFields=Name&orderBy=Name%20asc"

        try:
            response = requests.get(clientsurl, headers=self.headers)

            if response.status_code == 200:
                clients_df = pd.DataFrame(response.json())
                print_success(f"Found {len(clients_df)} clients")
                return clients_df
            else:
                print_error(f"Failed to fetch clients. Status code: {response.status_code}")
                print(f"{Colors.RED}Response: {response.text}{Colors.RESET}")
                return None

        except Exception as e:
            print_error(f"Error fetching clients: {str(e)}")
            return None

    def search_clients(self, clients_df, search_term):
        """
        Search clients by name
        """
        if search_term:
            filtered = clients_df[clients_df['Name'].str.contains(search_term, case=False, na=False)]
            return filtered
        return clients_df

    def display_clients(self, clients_df, page=0, page_size=60):
        """
        Display clients in a 3-column paginated format
        """
        total = len(clients_df)
        start = page * page_size
        end = min(start + page_size, total)

        print(f"\n{Colors.BOLD}{Colors.MAGENTA}{'='*100}{Colors.RESET}")
        print(f"{Colors.BOLD}{Colors.CYAN}Clients {start+1}-{end} of {total}{Colors.RESET}")
        print(f"{Colors.BOLD}{Colors.MAGENTA}{'='*100}{Colors.RESET}\n")

        # Display in 3 columns
        clients_to_display = clients_df.iloc[start:end]
        rows_per_column = (len(clients_to_display) + 2) // 3  # Round up division

        for row_idx in range(rows_per_column):
            columns = []

            for col_idx in range(3):
                actual_idx = row_idx + (col_idx * rows_per_column)

                if actual_idx < len(clients_to_display):
                    original_idx = clients_to_display.index[actual_idx]
                    row = clients_to_display.loc[original_idx]
                    display_idx = start + actual_idx + 1
                    client_name = row['Name'][:28]  # Truncate for column width

                    # Format: "123  Company Name"
                    columns.append(f"{Colors.WHITE}{display_idx:<4}{Colors.RESET} {Colors.GREEN}{client_name:<28}{Colors.RESET}")
                else:
                    columns.append(" " * 33)  # Empty column

            print("  ".join(columns))

        print(f"\n{Colors.MAGENTA}{'='*100}{Colors.RESET}")

        return start, end, total

    def select_clients_interactive(self, clients_df):
        """
        Interactive client selection with search
        """
        print_section("Client Selection")
        print(f"{Colors.CYAN}Options:{Colors.RESET}")
        print(f"  {Colors.YELLOW}•{Colors.RESET} Enter {Colors.BOLD}'search'{Colors.RESET} to filter clients")
        print(f"  {Colors.YELLOW}•{Colors.RESET} Enter client number(s) separated by commas (e.g., {Colors.BOLD}1,3,5{Colors.RESET})")
        print(f"  {Colors.YELLOW}•{Colors.RESET} Enter a range (e.g., {Colors.BOLD}1-5{Colors.RESET})")
        print(f"  {Colors.YELLOW}•{Colors.RESET} Enter {Colors.BOLD}'all'{Colors.RESET} to select all clients")
        print(f"  {Colors.YELLOW}•{Colors.RESET} Enter {Colors.BOLD}'list'{Colors.RESET} to show more clients")
        print(f"  {Colors.YELLOW}•{Colors.RESET} Enter {Colors.BOLD}'quit'{Colors.RESET} to exit")

        filtered_df = clients_df.copy()
        page = 0
        page_size = 60  # Match the display page size (3 columns x 20 rows)
        selected_indices = []

        while True:
            self.display_clients(filtered_df, page, page_size)

            choice = input(f"\n{Colors.BOLD}{Colors.CYAN}Your choice: {Colors.RESET}").strip().lower()

            if choice == 'quit':
                return None

            elif choice == 'search':
                search_term = input(f"{Colors.CYAN}Enter search term: {Colors.RESET}").strip()
                filtered_df = self.search_clients(clients_df, search_term)
                page = 0
                print_success(f"Found {len(filtered_df)} matching clients")

            elif choice == 'list':
                page += 1
                if page * page_size >= len(filtered_df):
                    print_warning("No more clients to display")
                    page = 0

            elif choice == 'all':
                return filtered_df

            else:
                # Parse selection
                try:
                    selected_indices = []

                    # Handle ranges (e.g., 1-5)
                    if '-' in choice:
                        parts = choice.split('-')
                        start_idx = int(parts[0]) - 1
                        end_idx = int(parts[1])
                        selected_indices = list(range(start_idx, end_idx))
                    else:
                        # Handle comma-separated (e.g., 1,3,5)
                        parts = choice.split(',')
                        selected_indices = [int(p.strip()) - 1 for p in parts if p.strip().isdigit()]

                    if selected_indices:
                        selected_clients = filtered_df.iloc[selected_indices]

                        print(f"\n{Colors.CYAN}You selected {len(selected_clients)} client(s):{Colors.RESET}")
                        for _, row in selected_clients.iterrows():
                            print(f"  {Colors.GREEN}•{Colors.RESET} {Colors.WHITE}{row['Name']}{Colors.RESET}")

                        confirm = input(f"\n{Colors.BOLD}{Colors.YELLOW}Proceed with these clients? (y/n): {Colors.RESET}").strip().lower()
                        if confirm == 'y':
                            return selected_clients
                    else:
                        print_error("Invalid selection. Please try again.")

                except Exception as e:
                    print_error(f"Invalid input: {str(e)}. Please try again.")

    def export_passwords(self, selected_clients, output_file=None):
        """
        Export passwords for selected clients (creates individual CSV per client)
        """
        timestamp = datetime.now().strftime("%Y.%m.%d")
        print_info(f"Exporting passwords for {len(selected_clients)} client(s)...")

        deploymentloginidsurl = "/cwa/api/v1/clients/{cwclientid}/deploymentlogins?pagesize=-1&condition=&orderBy=title%20asc&includeFields=password,Username,Title,Notes,Url"

        success_count = 0
        fail_count = 0
        exported_files = []

        # Process each client individually
        for idx, row in selected_clients.iterrows():
            cwclientid = row['Id']
            client_name = row['Name']

            try:
                url = self.base_url + deploymentloginidsurl.format(cwclientid=cwclientid)
                resp = requests.get(url, headers=self.headers)

                if resp.status_code in [200, 201]:
                    passwords = resp.json()

                    if passwords:
                        # Generate filename for this specific client
                        safe_name = "".join(c for c in client_name if c.isalnum() or c in (' ', '-', '_')).strip()
                        safe_name = safe_name.replace(' ', '_')

                        # Use provided output_file as base if single client, otherwise generate
                        if len(selected_clients) == 1 and output_file:
                            client_output_file = self.output_dir / output_file
                        else:
                            client_output_file = self.output_dir / f"{safe_name}_{timestamp}.csv"

                        # Convert to string for pandas
                        client_output_file = str(client_output_file)

                        # Create DataFrame for this client
                        df = pd.DataFrame(passwords)
                        df['ClientId'] = df['Client'].apply(lambda x: x['ClientId'])
                        df['ClientName'] = client_name
                        df = df[["ClientName", "ClientId", "Title", "Username", "Password", "Notes", "Url"]]

                        # Try to save with error handling
                        save_successful = False
                        original_filename = client_output_file
                        attempt = 0

                        while not save_successful and attempt < 5:
                            try:
                                df.to_csv(client_output_file, index=False)
                                save_successful = True
                                exported_files.append(client_output_file)
                                success_count += 1
                                print(f"  {Colors.GREEN}✓{Colors.RESET} {Colors.WHITE}{client_name}{Colors.RESET} {Colors.CYAN}({len(passwords)} passwords){Colors.RESET} → {Colors.YELLOW}{client_output_file}{Colors.RESET}")

                            except PermissionError:
                                attempt += 1
                                if attempt < 5:
                                    base, ext = original_filename.rsplit('.', 1) if '.' in original_filename else (original_filename, 'csv')
                                    client_output_file = f"{base}_{attempt}.{ext}"
                                else:
                                    print(f"  {Colors.RED}✗{Colors.RESET} {Colors.WHITE}{client_name}{Colors.RESET} {Colors.RED}(Permission denied - file may be open){Colors.RESET}")
                                    fail_count += 1
                                    break

                            except Exception as e:
                                print(f"  {Colors.RED}✗{Colors.RESET} {Colors.WHITE}{client_name}{Colors.RESET} {Colors.RED}(Error saving: {str(e)}){Colors.RESET}")
                                fail_count += 1
                                break

                    else:
                        print(f"  {Colors.YELLOW}⚠{Colors.RESET} {Colors.WHITE}{client_name}{Colors.RESET} {Colors.YELLOW}(No passwords found){Colors.RESET}")
                        success_count += 1  # Still count as success since API call worked

                else:
                    fail_count += 1
                    print(f"  {Colors.RED}✗{Colors.RESET} {Colors.WHITE}{client_name}{Colors.RESET} {Colors.RED}(Failed - Status {resp.status_code}){Colors.RESET}")

            except Exception as e:
                fail_count += 1
                print(f"  {Colors.RED}✗{Colors.RESET} {Colors.WHITE}{client_name}{Colors.RESET} {Colors.RED}(Error: {str(e)}){Colors.RESET}")

        # Summary
        if success_count > 0:
            print(f"\n{Colors.BOLD}{Colors.GREEN}{'='*100}{Colors.RESET}")
            print(f"{Colors.BOLD}{Colors.GREEN}✓ Export complete!{Colors.RESET}")
            print(f"{Colors.BOLD}{Colors.GREEN}{'='*100}{Colors.RESET}")
            print(f"{Colors.CYAN}  • Successfully exported: {Colors.WHITE}{success_count} client(s){Colors.RESET}")
            print(f"{Colors.CYAN}  • Failed: {Colors.WHITE}{fail_count} client(s){Colors.RESET}")
            print(f"{Colors.CYAN}  • Files created: {Colors.WHITE}{len(exported_files)}{Colors.RESET}")

            if exported_files:
                print(f"\n{Colors.CYAN}Exported files:{Colors.RESET}")
                for file in exported_files:
                    print(f"  {Colors.GREEN}•{Colors.RESET} {Colors.WHITE}{file}{Colors.RESET}")

            print(f"{Colors.BOLD}{Colors.GREEN}{'='*100}{Colors.RESET}")

            return True
        else:
            print_error("No passwords retrieved. Export cancelled.")
            return False


def main():
    parser = argparse.ArgumentParser(
        description="CNS4U Offboarding Tool - Extract passwords from ConnectWise Automate",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  Interactive mode (recommended):
    python cns4u_offboarding.py

  Automated token extraction:
    python cns4u_offboarding.py --auto-login --base_url "https://cns4u.hostedrmm.com"

  Manual credentials:
    python cns4u_offboarding.py --manual --clientid "xxx" --bearer_token "bearer xxx"
        """
    )

    parser.add_argument("--auto-login", action="store_true",
                       help="Open browser for manual login and extract bearer token")
    parser.add_argument("--manual", action="store_true",
                       help="Manually enter bearer token and client ID")
    parser.add_argument("--clientid", type=str,
                       help="Client ID (for manual mode)")
    parser.add_argument("--bearer_token", type=str,
                       help="Bearer token (for manual mode)")
    parser.add_argument("--base_url", type=str,
                       help="Base URL (e.g., https://cns4u.hostedrmm.com)")
    parser.add_argument("--output_file", type=str,
                       help="Output CSV filename")
    parser.add_argument("--output_dir", type=str,
                       help="Directory to save CSV files (default: current directory)")
    parser.add_argument("--client-ids", type=str,
                       help="Comma-separated client IDs to export (skip interactive selection)")

    args = parser.parse_args()

    # Print the banner
    print_banner()

    # Check dependencies first
    if not check_and_fix_dependencies():
        print(f"\n{Colors.RED}Please fix the dependencies above and try again.{Colors.RESET}")
        input(f"\n{Colors.YELLOW}Press Enter to exit...{Colors.RESET}")
        return

    # Get base URL if not provided
    base_url = args.base_url
    if not base_url:
        print(f"\n{Colors.CYAN}Enter your ConnectWise Automate URL{Colors.RESET}")
        print(f"{Colors.YELLOW}Example: https://cns4u.hostedrmm.com{Colors.RESET}")
        base_url = input(f"{Colors.BOLD}{Colors.CYAN}Base URL: {Colors.RESET}").strip()

        # Clean up the URL
        if not base_url.startswith('http'):
            base_url = 'https://' + base_url
        base_url = base_url.rstrip('/')

    # Get output directory
    output_dir = args.output_dir
    if not output_dir:
        print(f"\n{Colors.CYAN}Where would you like to save the CSV files?{Colors.RESET}")
        print(f"{Colors.YELLOW}Press Enter for current directory, or type a path:{Colors.RESET}")
        output_dir_input = input(f"{Colors.BOLD}{Colors.CYAN}Output directory: {Colors.RESET}").strip()

        if output_dir_input:
            output_dir = output_dir_input
        else:
            output_dir = None  # Will use current directory

    if output_dir:
        print_info(f"CSV files will be saved to: {output_dir}")

    offboarding = CWAOffboarding(base_url=base_url, output_dir=output_dir)

    # Step 1: Get credentials
    if args.manual:
        if args.clientid and args.bearer_token:
            offboarding.clientid = args.clientid
            offboarding.bearer_token = args.bearer_token
            offboarding.headers = {
                "authorization": args.bearer_token,
                "clientid": args.clientid,
                "accept": "application/json"
            }
        else:
            offboarding.get_credentials_manual()
    elif args.auto_login:
        success = offboarding.get_credentials_automated()
        if not success:
            print_warning("Falling back to manual entry...")
            offboarding.get_credentials_manual()
    else:
        # Interactive mode - ask user preference
        print(f"\n{Colors.CYAN}How would you like to authenticate?{Colors.RESET}")
        print(f"  {Colors.YELLOW}1.{Colors.RESET} {Colors.WHITE}Automated token extraction (opens browser, you log in manually){Colors.RESET}")
        print(f"  {Colors.YELLOW}2.{Colors.RESET} {Colors.WHITE}Manual bearer token entry (you provide token and client ID){Colors.RESET}")

        choice = input(f"\n{Colors.BOLD}{Colors.CYAN}Your choice (1 or 2): {Colors.RESET}").strip()

        if choice == "1":
            success = offboarding.get_credentials_automated()
            if not success:
                print_warning("Falling back to manual entry...")
                offboarding.get_credentials_manual()
        else:
            offboarding.get_credentials_manual()

    # Step 2: Get all clients
    clients_df = offboarding.get_all_clients()

    if clients_df is None or len(clients_df) == 0:
        print_error("Failed to retrieve clients. Exiting.")
        return

    # Step 3: Select clients
    if args.client_ids:
        # Use provided client IDs
        client_id_list = [cid.strip() for cid in args.client_ids.split(',')]
        selected_clients = clients_df[clients_df['Id'].isin(client_id_list)]
        print_success(f"Selected {len(selected_clients)} client(s) from provided IDs")
    else:
        # Interactive selection
        selected_clients = offboarding.select_clients_interactive(clients_df)

    if selected_clients is None or len(selected_clients) == 0:
        print_warning("No clients selected. Exiting.")
        return

    # Step 4: Export passwords
    offboarding.export_passwords(selected_clients, args.output_file)


if __name__ == "__main__":
    main()
