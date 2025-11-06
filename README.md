# snort-blocked-ip-analysis-automation
These scripts automates the retrieval, analysis, and management of Snort(PfSense package) blocked IPs to streamline security monitoring and whitelist management

check_ip.py downloads the Snort blocked hosts archive, extracts and compares IPs, enriches them using the ipinfo.io API, and saves results to an Excel report with organization-based highlights.

extract_ips_from_sheet.py reads the Excel report to extract red-highlighted IPs and adds them automatically to the pfSense Snort Pass List via the web interface using Selenium.

# Configuration
Before running the scripts, you must configure several variables at the top of each file.

In check_ip.py, edit lines 25–44.

In extract_ips_from_sheet.py, edit lines 14–26.
