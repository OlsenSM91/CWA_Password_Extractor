# CWA Password Extractor

<img width="752" height="607" alt="Screenshot 2025-12-05 160456" src="https://github.com/user-attachments/assets/b402c237-bfbc-4597-9a9d-b76b4b2e3ecc" />

Enhanced utility for extracting passwords from ConnectWise Automate for client offboarding. This tool automates the bearer token extraction and allows selective client password exports.

## Features

- **Automated Bearer Token Extraction**: Opens a browser for you to log in manually (handles MFA), then automatically captures bearer token
- **Interactive Client Selection**: Browse, search, and select specific clients for password export
- **Selective Export**: Export passwords for specific clients instead of all clients
- **Multiple Authentication Methods**: Choose between automated token extraction or manual token entry
- **Search Functionality**: Filter clients by name to quickly find the ones you need
- **Flexible Base URL**: Specify your Automate URL at runtime (not hardcoded)

## Quick Start

### Installation

1. Install Python dependencies:
```bash
pip install -r requirements.txt
```

2. Ensure you have Chrome installed

### Usage

#### Interactive Mode (Recommended)

Simply run the script and follow the prompts:

```bash
python CW_Automate_PW_Extractor.py
```

This will:
1. Ask for your ConnectWise Automate base URL
2. Ask how you want to authenticate (automated or manual)
3. If automated: Open a browser for you to log in manually (handles MFA)
4. Automatically extract your bearer token from network requests
5. Fetch all clients from ConnectWise Automate
6. Display clients with search and selection options
7. Export passwords for selected clients

#### Automated Token Extraction

Open browser for manual login and extract bearer token:

```bash
python CW_Automate_PW_Extractor.py --auto-login --base_url "https://cns4u.hostedrmm.com/Automate"
```

The browser will open to the login page. You log in manually (including MFA), then the tool captures your bearer token automatically.

#### Manual Bearer Token Entry

If you prefer to manually extract and enter the bearer token:

```bash
python CW_Automate_PW_Extractor.py --manual
```

Or provide credentials directly:

```bash
python CW_Automate_PW_Extractor.py --manual --clientid "46158abd-1452-5869-9abd-44b0edabe541" --bearer_token "bearer abc123..."
```

### Client Selection Options

Once the client list is loaded, you can:

- **Search**: Type `search` and enter a search term to filter clients
- **Select specific clients**: Enter numbers separated by commas (e.g., `1,3,5`)
- **Select a range**: Enter a range (e.g., `1-5`)
- **Select all clients**: Type `all`
- **View more clients**: Type `list` to see the next page
- **Exit**: Type `quit` to cancel

### Command-Line Options

```
--auto-login          Open browser for manual login and extract bearer token
--manual              Manually enter bearer token and client ID
--clientid            Client ID (for manual mode)
--bearer_token        Bearer token (for manual mode)
--base_url            Base URL (e.g., https://cns4u.hostedrmm.com)
--output_file         Output CSV filename
--client-ids          Comma-separated client IDs to export (skip interactive)
```

### Examples

Export passwords for specific client IDs:
```bash
python CW_Automate_PW_Extractor.py --manual --base_url "https://cns4u.hostedrmm.com/Automate" --clientid "xxx" --bearer_token "bearer xxx" --client-ids "123,456,789"
```

Use custom output filename:
```bash
python CW_Automate_PW_Extractor.py --auto-login --base_url "https://cns4u.hostedrmm.com/Automate" --output_file "client_offboarding_2024.csv"
```

Specify different base URL:
```bash
python CW_Automate_PW_Extractor.py --base_url "https://mycompany.hostedrmm.com/Automate"
```

## Manual Bearer Token Extraction (Old Method)

If you prefer to extract the bearer token manually:

1. Log into ConnectWise Automate at https://cns4u.hostedrmm.com/Automate
2. Navigate to any company's passwords page (e.g., https://cns4u.hostedrmm.com/automate/browse/companies/company-passwords?companyId=438)
3. Open Developer Tools (F12)
4. Go to Network tab
5. Look for a `deploymentlogins?` request
6. In the Headers tab, find the Request Headers section
7. Copy the `authorization` value (e.g., `bearer actualTokenDataHere==`)
8. Also copy the `clientid` value

## Requirements

- Python 3.7+
- pandas (for CSV export)
- requests (for API calls)
- selenium (for automated login)
- ChromeDriver (for Selenium)

## Legacy Tool

The original `CW_Automate_PW_Extractor.py` script was created by https://github.com/matthew-hoad/CWA_Password_Extractor

---
# Old Readme Below

# CWA Password Extractor
Utility for Extracting all passwords from ConnectWise Automate (E.g. while migrating to a new system). Outputs a csv file with all fields displayed in CWA for stored credentials.

ConnectWise don't seem to be fond of giving you an unencrypted dump of your passwords stored in CW Automate so I dug into Developer Tools in Chrome with the CWA web client open and discovered the following:

Passwords are retrieved and stored in memory as plaintext when you right-click on a credential in the passwords area. You can combine the parameters of that request with the one that retrieves the list of credentials to make a request for retreiving all credentials for a client using that client's ID. Then all you need to do is grab all client IDs and iterate over them. There's probably an even simpler way to get all credentials for all clients but I didn't really have time to look into that.

Usage: `python CW_Automate_PW_Extractor.py --clientid '00000000-0000-0000-0000-000000000000' --bearer_token 'bearer asdf1234asdf1234asdf1234==' --base_url 'https://myorgname.hostedrmm.com' --output_file 'cwa_passwords.csv'`

You can get your clientid and bearer token from the developer tools while you have the web client open (it might also work if you've got an existing API key and auth/token flow elsewhere but I haven't tried that, this is just quick and dirty). I tend to get these parameters from the "deploymentlogins?" request like so:
![image](https://user-images.githubusercontent.com/16311787/145383318-88f6fbf6-2d3f-4302-b45d-7ab9791de4e7.png)

# Requirements
- pandas (for exporting to csv)
- requests (for making API calls)

# Steven's Update
To obtain the bearer token and clientID, log into Automate's web portal and open up Inspect/Developer Options. Under the `Network` tab, load up the passwords under the company and open the password box. You will be able to locate the deploymentlogins and under the headers will have the `Authorization:` and `ClientID:` which is the necessary authentication you need to utilize this script

![wow](https://github.com/OlsenSM91/CWA_Password_Extractor/assets/130707762/e181804b-f358-41c7-981d-d25f5f6ba684)
