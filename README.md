# IP_URL_OSINT_Checker

This is a simple script that can take a mixed list of IP addresses and URLs and run them through some OSINT tools.
First, it runs the whole list through VirusTotal. Then, any results with a high number of malicious indicators gets highlighted red and saved.
It then takes the saved results and runs them through AbuseIPDB for some whois information, CrowdStrike to find any associated actors or groups, and Shodan to find any associated Domains and open ports.

The script then takes all the information and creates a color coded report, with a separate sheet for each source, with a current Date/Time stamp to keep things organized.

#Future Plans:

1: Find a better solution than saving API credentials directly in the script
2: Find other OSINT tools to run through
