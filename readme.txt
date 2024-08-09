Email Extractor Service
Overview
A Python-based Windows service that automatically extracts email bodies with the subject "Sent from AirCentre Movement Manager" from Microsoft Outlook. The extracted content is saved as XML files in the MQIN directory, with each file uniquely named for easy access.

Features
Automated extraction of specific emails.
Continuous operation as a Windows service.
Organized file storage with unique naming conventions.
Installation
Clone the Repo:

bash
Copy code
git clone https://github.com/your-username/email_extractor_project.git
cd email_extractor_project
Install Dependencies:

bash
Copy code
pip install -r requirements.txt
Install and Start the Service:

bash
Copy code
python email_extractor_service.py install
python email_extractor_service.py start
Usage
The service runs in the background, extracting and saving relevant emails as XML files in the MQIN folder.

Management
Use the batch scripts in the setup/ directory to install, start, stop, or uninstall the service.
