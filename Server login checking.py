import winrm
import logging

##################cyberark login needed for creds

server ="172.18.161.196" #- DT
#"172.18.81.15" - prod 
username = "A1446991-A"  
password = ""  
#service_name = "ClosePendingCompleteOrderService"
service_name = "ProfSvc"
flag=0

logging.info("Establishing a cmd remoting session")
session = winrm.Session(server, auth=(username, password), transport='ntlm', server_cert_validation='ignore')
logging.info("Running the cmd command")

# Check if the service exists on the server
cmd = f"Get-Service {service_name}"
result = session.run_ps(cmd)
    
# Read the result from the command
output = result.std_out.decode().strip()

if result.status_code == 0:
    print( f"The service {service_name} is running successfully.")
else:
    print(f"The service {service_name} is not running.")
    start_cmd = f"Start-Service {service_name}"
    start_result = session.run_ps(start_cmd)
    start_cmd = f"Start-Service {service_name}"
    start_result = session.run_ps(start_cmd)
    if start_result.status_code == 0:
        print( f"The service {service_name} was stopped. It has been restarted successfully.") 

