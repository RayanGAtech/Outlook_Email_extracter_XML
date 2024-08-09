<h1>📨 Email Extractor Service</h1>

<h2>📄 Overview</h2>
<p>A Python-based Windows service that automatically extracts email bodies with the subject <strong>"Sent from AirCentre Movement Manager"</strong> from Microsoft Outlook and saves them as XML files in the <code>MQIN</code> directory.</p>

<h2>🚀 Features</h2>
<ul>
  <li><strong>Automated Extraction:</strong> Automatically processes specific emails.</li>
  <li><strong>Runs as a Service:</strong> Operates continuously in the background.</li>
  <li><strong>Organized Storage:</strong> Saves files with unique names for easy access.</li>
</ul>

<h2>⚙️ Installation</h2>
<ol>
  <li><strong>Clone the Repository:</strong>
    <pre><code>git clone https://github.com/your-username/email_extractor_project.git
cd email_extractor_project
</code></pre>
  </li>
  <li><strong>Install Dependencies:</strong>
    <pre><code>pip install -r requirements.txt</code></pre>
  </li>
  <li><strong>Install and Start the Service:</strong>
    <pre><code>python email_extractor_service.py install
python email_extractor_service.py start
</code></pre>
  </li>
</ol>

<h2>🛠️ Usage</h2>
<p>The service extracts emails and saves them as XML files in the <code>MQIN</code> folder.</p>

<h2>📂 Management</h2>
<p>Use the scripts in the <code>setup/</code> directory to manage the service:</p>
<ul>
  <li><strong>Install:</strong> <code>install_service.bat</code></li>
  <li><strong>Start:</strong> <code>start_service.bat</code></li>
  <li><strong>Stop:</strong> <code>stop_service.bat</code></li>
  <li><strong>Uninstall:</strong> <code>uninstall_service.bat</code></li>
</ul>
