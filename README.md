# Cover Letter Generator Powershell

**Cover Letter Generator**
This script is designed to automate the process of creating a cover letter by allowing the user to input company name, job title, and their interest in the company, which then gets filled in a pre-defined template.

**Usage**

* Create a template for the cover letter in Microsoft word, using placeholders for the company name, job title, and interest in the company.

* Copy the script to a local folder on your computer

* Open the script in PowerShell ISE or any other text editor

* Update the "Enter Template Path" and "Enter Destination Path" with the actual path of the cover letter template and the destination path where the final cover letter   should be saved.

* Run the script in PowerShell

* Follow the prompts to enter the company name, job title, and your interest in the company.

* The script will generate a cover letter in a pdf format, with the entered details, in the destination path specified.

* The script will also rename the generated pdf file with the company name and "Cover Letter.pdf"

**Dependencies**
This script requires the following to be installed on the machine:

* Microsoft Word

* PowerShell

Note: The user must create their own cover letter template and use placeholders for the company name, job title, and interest in the company. The placeholders used in the template must match the variables used in the script (case-sensitive).

It is important to note that the script uses -replace command which is case-sensitive so it is important to use the same case for the placeholders in the template and the variables in the script.

Also, the script uses Start-Sleep -Seconds .5 which is to wait for 500 milliseconds before renaming the file, it's just an arbitrary time to ensure that the file generation process is complete before renaming it, you may want to adjust it to the time that is suitable to your machine.

You can always test the script on a small scale before trying it on a large scale, and make any adjustments that may be necessary to suit your needs.
