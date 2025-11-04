### **Contract Tracking Macro**

**English Version** | **[Deutsche Version](README.md)**

This VBA macro automatically monitors emails in the "Vertragswesen/Nachverfolgung" (Contracts/Follow-up) folder and checks for incoming replies in the inbox. When a reply is found, both the original message and the reply are moved to the "Vertragswesen/Abgeschlossen" (Contracts/Completed) folder. For messages without a reply, a reminder is sent after a specified number of days (default: 7 days).

**New Features:**
- ✅ Configurable waiting period (default: 7 days)
- ✅ Improved subject recognition for replies (supports RE:, AW:, FWD:, WG:)
- ✅ Error handling and folder validation
- ✅ More robust code structure with Option Explicit

---

### **Features**

1. **Email Monitoring:**
   - The macro checks emails in the "Vertragswesen/Nachverfolgung" folder that are older than N days (default: 7 days).

2. **Intelligent Reply Detection:**
   - Recognizes replies with various prefixes (RE:, AW:, FWD:, WG:)
   - Supports multiple prefixes (e.g., "RE: RE: Subject")
   - Case-insensitive subject comparison

3. **Reply Handling:**
   - When a reply with the same subject is found in the inbox, both the original email and the reply are moved to the "Vertragswesen/Abgeschlossen" folder.

4. **Reminder Sending:**
   - For messages without a reply, a reminder is sent to the original recipient if the message is older than the specified number of days.

5. **Automatic Summary:**
   - After completing the check, an email summary is created listing the people who replied and those who received reminders.

6. **Error Handling:**
   - Validates the existence of required folders
   - Displays user-friendly error messages
   - Prevents crashes in unexpected situations

---

### **Installation**

1. **Open VBA Editor:**
   - Open Microsoft Outlook.
   - Press `ALT + F11` to launch the VBA editor.

2. **Add Module:**
   - Select `Insert > Module` from the menu to create a new module.
   - Copy the complete VBA code and paste it into the new module.

3. **Save and Close:**
   - Save the code and close the VBA editor.

---

### **Usage**

1. **Run Macro Manually:**
   - In Outlook, go to `Developer Tools > Macros`, select `VertragErinnerungNachverfolgung`, and click `Run`.
   - The macro will check the emails and provide a summary.

2. **Automatic Execution (Optional):**
   - To run the macro automatically, you can use Windows Task Scheduler or Power Automate:
     - Create a .vbs file that executes the macro at regular intervals.
     - Schedule this file with your preferred scheduler.

---

### **Customization**

1. **Adjust Monitoring Period:**
   - The number of days after which reminders are sent can be adjusted in the line:  
     ```vba
     DaysToWait = 7 ' Customizable: Default is 7 days
     ```
     Replace `7` with your desired number of days.

2. **Change Folder Names:**
   - The target folders for follow-up and completed messages can be centrally defined:  
     ```vba
     Const VERTRAGSWESEN_ORDNER As String = "Vertragswesen"
     Const NACHVERFOLGUNG_ORDNER_NAME As String = "Nachverfolgung"
     Const ABGESCHLOSSEN_ORDNER_NAME As String = "Abgeschlossen"
     ```
     Adjust the names to match your folder structure.

3. **Summary Recipient:**
   - Change the recipient of the summary email in the line:
     ```vba
     Const ZUSAMMENFASSUNG_EMPFAENGER As String = "your.email@example.com"
     ```

4. **Customize Reminder Text:**
   - You can customize the subject and content of the reminder email as needed:
     ```vba
     ReminderMail.Subject = "Reminder: Please return the contract"
     ReminderMail.Body = "Please return the contract if you haven't already done so."
     ```

---

### **Requirements**

- **Microsoft Outlook:** With VBA enabled.
- **Email Folder Structure:** The folders "Vertragswesen/Nachverfolgung" and "Vertragswesen/Abgeschlossen" must exist.

---

### **Note**

- This macro accesses your Outlook data and sends emails. Make sure you only use it in trusted environments.
- Regularly check your email folders to ensure no important messages are lost.

---

**Created with ❤️ by João**
