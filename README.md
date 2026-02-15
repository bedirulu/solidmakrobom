## Notes
This macro was developed with AI assistance.
I tested, edited, and adapted it for SolidWorks BOM workflows.

AI tools used during development: Claude Sonnet 4.5, Gemini Pro, Microsoft Copilot.

### Security Concerns ###
Note on Execution Policy:

This macro utilizes a dynamically generated PowerShell script to leverage Windows API for thumbnail extraction. To ensure the script runs smoothly on standard workstations without complex configuration, the VBA code uses the -ExecutionPolicy Bypass flag:

VBA

cmd = "powershell.exe -ExecutionPolicy Bypass ..."
Risk Analysis:
While efficient for personal use or isolated environments, using Bypass overrides the system's default safety execution policies. In a strict Enterprise Environment, this behavior might be flagged by EDR (Endpoint Detection and Response) systems or blocked by Group Policy Objects (GPO).

Recommended Mitigation for Production:
For deployment in a secure corporate network, the recommended approach is:

Remove the Bypass flag.

Digitally Sign the PowerShell script (Code Signing) using a trusted internal Certificate Authority (CA).

Configure the execution policy to AllSigned or RemoteSigned to allow only trusted scripts.
