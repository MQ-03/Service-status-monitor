from re import findall
from subprocess import Popen, PIPE
from colorama import Fore, Back, Style
import csv, winrm
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime

service_name = 'esealservice'

print(f"""------------------------------------------------
        Tracking Server Service Status
------------------------------------------------""")

def check_service_status(ipaddr, username, password, service_name):
    """Check the status of a service on a remote PC."""
    try:
        # Construct WinRM URL
        remote_host = ipaddr
        
        # Create WinRM session
        session = winrm.Session(remote_host, auth=(username, password), transport='ntlm')
        
        # PowerShell script to get service status
        ps_script = f"""
        $service = Get-Service -Name '{service_name}'
        $service.Status
        """
        
        # Execute the PowerShell command
        result = session.run_ps(ps_script)
        
        # Decode and return the result
        return result.std_out.decode().strip()
    except Exception as e:
        return f"Error: {str(e)}"

def send_email(subject, body, to_email):
    """Send an email with the specified subject and body."""
    from_email = "mizbee.ai.info@gmail.com"  # Replace with your email
    from_password = "fjxp zbvb ztth zawp"  # Replace with your email password

    # Create the email message
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = ", ".join(to_email)
    msg['Subject'] = subject

    # Attach the email body
    msg.attach(MIMEText(body, 'plain'))

    # Send the email
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)  # Replace with your SMTP server and port
        server.starttls()
        server.login(from_email, from_password)
        text = msg.as_string()
        server.sendmail(from_email, to_email, text)
        server.quit()
        print(f"Email sent to {to_email}")
    except Exception as e:
        print(f"Failed to send email: {str(e)}")

# Read the CSV and check the service status
with open('hostname.csv', newline='') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        ipaddr = row['DSIP']
        loc = row['Location']
        plcod = row['Pl code']
        username = row['username']
        password = row['password']
        to_email = row['email'].split(';')  # Add this column to your CSV

        timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        status = check_service_status(ipaddr, username, password, service_name)
        if status == "Running":
            result = "Running"
        else:
            result = "Down"
        data = ""
        output = Popen(f"ping {ipaddr} -n 3", stdout=PIPE, encoding="utf-8")
        for line in output.stdout:
            data = data + line
            ping_test = findall("TTL", data)
        if ping_test:
            pc_status = "Active"
            print(f"""------------------------------------------------
Timestamp  : {Fore.MAGENTA}{timestamp}{Style.RESET_ALL}
IP Address : {Fore.YELLOW}{ipaddr} {Style.RESET_ALL}
Location   : {Fore.CYAN}{loc}{Style.RESET_ALL}       
Plant Code : {Fore.CYAN}{plcod}{Style.RESET_ALL}   
PC Status  : {Fore.GREEN}Active{Style.RESET_ALL}
Service    : {Fore.GREEN}{result}{Style.RESET_ALL}      
------------------------------------------------""")
        else:
            pc_status = "Down"
            print(f"""------------------------------------------------
Timestamp  : {Fore.MAGENTA}{timestamp}{Style.RESET_ALL}
IP Address : {Fore.YELLOW}{ipaddr} {Style.RESET_ALL}
Location   : {Fore.RED}{loc}{Style.RESET_ALL}       
Plant Code : {Fore.RED}{plcod}{Style.RESET_ALL}   
PC Status  : {Fore.RED}Down{Style.RESET_ALL}
Service    : {Fore.RED}{result}{Style.RESET_ALL}      
------------------------------------------------""")

        # Prepare email subject and body
        subject = f"Service Status for {service_name} on {ipaddr}"
        body = f"""
        Timestamp: {timestamp}
        IP Address: {ipaddr}
        Location: {loc}
        Plant Code: {plcod}
        PC Status: {pc_status}
        Service Status: {result}
        """
        # Send email
        send_email(subject, body, to_email)