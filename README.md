# OutlookFolder2Text_w_Attachments
When PST Export is Disabled and you don't want embedded links/images that trigger security warnings 

# ⚠️ LEGAL DISCLAIMER AND LIMITATION OF LIABILITY

**READ THIS SECTION CAREFULLY BEFORE USING THIS SOFTWARE.**

By downloading, accessing, installing, or executing the code contained in this repository ("the Software"), you ("the User") agree to be bound by the following terms and conditions. If you do not agree to these terms, **you must not use this Software.**

### 1. "AS IS" BASIS
THE SOFTWARE IS PROVIDED "AS IS," WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, TITLE, AND NON-INFRINGEMENT. IN NO EVENT SHALL THE AUTHORS, COPYRIGHT HOLDERS, OR CONTRIBUTORS BE LIABLE FOR ANY CLAIM, DAMAGES, OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT, OR OTHERWISE, ARISING FROM, OUT OF, OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

### 2. NO PROFESSIONAL ADVICE
The Software is intended for educational and informational purposes only. It is not intended to constitute professional legal, financial, or technical advice. The author(s) make no representations or warranties regarding the accuracy, completeness, or suitability of the Software for any specific purpose.

### 3. USER RESPONSIBILITY AND RISK ASSUMPTION
The User acknowledges that:
- **Data Loss Risk:** The execution of this script involves reading, writing, and moving files. There is an inherent risk of data corruption, accidental deletion, overwriting of existing files, or unintended modification of email data.
- **System Impact:** The script interacts directly with the Microsoft Outlook application and the local file system. Improper use may cause Outlook instability, crashes, or system performance degradation.
- **Environment Variability:** The Software may behave differently depending on the User's operating system version, Outlook build, security settings, network configuration, and existing file structures.

**THE USER ASSUMES ALL RISKS ASSOCIATED WITH THE USE OF THIS SOFTWARE.** The User agrees to:
1.  **BACKUP ALL DATA** before running the script. The Author(s) are not responsible for any loss of emails, attachments, or other data resulting from the use of this Software.
2.  Test the Software on a non-production environment or a small subset of data before applying it to large datasets.
3.  Ensure they have the necessary permissions and authority to access and modify the data being processed.

### 4. LIMITATION OF LIABILITY
TO THE MAXIMUM EXTENT PERMITTED BY APPLICABLE LAW, IN NO EVENT SHALL THE AUTHOR(S) BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

### 5. INDEMNIFICATION
The User agrees to indemnify, defend, and hold harmless the Author(s), contributors, and affiliates from and against any and all claims, liabilities, damages, losses, or expenses (including reasonable legal fees) arising out of or in any way connected with the User's use of the Software, violation of these terms, or infringement of any third-party rights.

### 6. NO SUPPORT GUARANTEE
While the Author(s) may provide community support or updates voluntarily, there is no obligation to provide maintenance, bug fixes, technical support, or updates. The Software may contain bugs or errors that could result in unexpected behavior.

### 7. GOVERNING LAW
These terms shall be governed by and construed in accordance with the laws of the jurisdiction in which the Author(s) reside, without regard to its conflict of law provisions.

---
**BY USING THIS SOFTWARE, YOU ACKNOWLEDGE THAT YOU HAVE READ, UNDERSTOOD, AND AGREED TO THESE TERMS IN THEIR ENTIRETY.**
# Outlook Email to Plain Text Exporter (VBA)

A robust Microsoft Outlook VBA macro designed to batch-export emails from a specific folder into plain text (`.txt`) files while preserving all attachments in a separate directory.

This tool is ideal for archiving, legal discovery, or migrating email data to a non-proprietary format without losing attachment integrity.

## 🚀 Features

- **Batch Processing**: Exports all emails in a selected folder automatically.
- **Plain Text Conversion**: Converts email bodies to `.txt` format (strips HTML/Rich Text).
- **Attachment Preservation**: Saves all attachments to a dedicated `Attachments` subfolder.
- **Metadata Capture**: Includes Sender, Recipient, CC, Subject, and Date in the text file header.
- **Safe Filenames**: Automatically sanitizes subject lines to remove illegal Windows characters (`\ / : * ? " < > |`).
- **Duplicate Handling**: Appends a counter to filenames if timestamps/subjects collide.
- **Robust Error Handling**: Skips corrupted emails gracefully and logs errors to the VBA Immediate Window.
- **Flexible Folder Navigation**: Supports folders located at the root level or nested deep within the Inbox hierarchy.

## ⚠️ Prerequisites & Warnings

> **⚠️ IMPORTANT**: This macro modifies files on your local drive. **Always test on a small subset of emails first.**

1.  **Outlook Version**: Works with Outlook 2016, 2019, 2021, and Microsoft 365 (Classic Desktop App).
2.  **Macros Enabled**: You must enable macros in Outlook's Trust Center.
3.  **Backup**: Create a backup of your Outlook data (PST/OST) before running batch operations.
4.  **Permissions**: Ensure you have write permissions to the destination folder (default: `C:\EmailExport\`).

## 📥 Installation

### Step 1: Enable the Developer Tab
1.  Open Outlook.
2.  Right-click anywhere on the Ribbon and select **Customize the Ribbon**.
3.  Check the box for **Developer** on the right side and click **OK**.

### Step 2: Open the VBA Editor
1.  Press `Alt + F11` to open the Visual Basic for Applications editor.
2.  In the menu, go to **Tools > References**.
3.  Ensure **Microsoft Outlook xx.x Object Library** is checked. Click **OK**.

### Step 3: Insert the Code
1.  In the Project Explorer (left pane), right-click on `VBAProject (YourName)`.
2.  Select **Insert > Module**.
3.  Copy the code from [`ExportEmails.vba`](./ExportEmails.vba) (or paste the code block below) into the new module window.

### Step 4: Configure the Destination
Edit the `saveFolder` variable in the code to match your desired output path:
```vba
saveFolder = "C:\EmailExport\"
Ensure the folder exists or the script will attempt to create it.

🛠️ Configuration
Target Folder Selection
The script is configured to look for a folder named FY24 containing a subfolder 11 May.

To change the target folder:

Locate the ExportEmailsAsPlainTextWithAttachments subroutine.
Modify the logic in the "Target Folder Selection" section.
If your folder is at the Root Level (same level as Inbox):
Set fy24Folder = objNamespace.Folders(1).Folders("YourFolderName")
If your folder is inside Inbox:
Set fy24Folder = objNamespace.GetDefaultFolder(olFolderInbox).Folders("YourFolderName")
Update the folder name string "11 May" to your specific subfolder name.
▶️ Usage
Run the Macro:
Press F5 inside the VBA editor, OR
Go to Developer > Macros, select ExportEmailsAsPlainTextWithAttachments, and click Run.
Monitor Progress:
The script runs silently. To see progress, press Ctrl + G in the VBA editor to open the Immediate Window.
It will print: Exported 10 emails..., Exported 20 emails..., etc.
Completion:
A popup message will confirm the total count and file locations.
Text Files: Located in C:\EmailExport\
Attachments: Located in C:\EmailExport\Attachments\
📂 Output Structure
C:\EmailExport\
├── 20240115_093045_Project_Update.txt
├── 20240116_141522_Invoice_Q1.txt
├── 20240117_100000_Meeting_Notes.txt
└── Attachments\
    ├── invoice.pdf
    ├── project_plan.docx
    └── screenshot.png
Text File Format Example:

FROM: John Doe
TO: Jane Smith
CC: Team Lead
SUBJECT: Project Update
DATE: 2024-01-15 09:30:45
----------------------------------------

[Email Body Content Here...]
🐛 Troubleshooting
Issue	Solution
"Type Mismatch" Error	Ensure the folder name in the code matches exactly (case-sensitive). Check if the folder is at the Root or inside Inbox and update the navigation logic accordingly.
"Permission Denied" Error	Check if C:\EmailExport\ exists and you have write permissions. Try changing the path to C:\Users\YourName\Documents\EmailExport\.
Script Stops Mid-Way	Check the Immediate Window (Ctrl+G) for specific error messages. This usually indicates a corrupted email or an invalid filename character that wasn't caught.
Attachments Overwritten	If multiple emails have attachments with the exact same name (e.g., image.jpg), later files will overwrite earlier ones. Future update planned to prefix attachment names.
Macros Disabled	Go to File > Options > Trust Center > Macro Settings and select Notifications for all macros.
📝 License
This project is provided "as is" without warranty of any kind. Use at your own risk.

🤝 Contributing
Feel free to submit issues or enhancement requests. Common improvements could include:

Adding a GUI form for folder selection.
Supporting CSV export for metadata.
Renaming attachments with unique prefixes to prevent overwrites.
Created for Outlook Classic (Windows 10/11)


### How to Publish to GitHub

1.  **Create a Repository**: Go to GitHub and create a new public/private repository (e.g., `outlook-email-exporter`).
2.  **Add Files**:
    *   Create a file named `README.md` and paste the content above.
    *   Create a file named `ExportEmails.vba` and paste the final VBA code we generated into it.
3.  **Commit**: Commit the changes to the main branch.
4.  **Optional**: Add a `.gitignore` file to exclude any temporary files (though VBA projects usually don't generate many).

This README provides professional context, clear instructions, and safety warnings, making it ready for anyone to use your tool.
