### **Vertragsüberwachung Makro**

Dieses VBA-Makro überwacht automatisch E-Mails im Ordner "Vertragswesen/Nachverfolgung" und überprüft, ob es eingehende Antworten im Posteingang gibt. Wenn eine Antwort gefunden wird, werden die ursprüngliche Nachricht und die Antwort in den Ordner "Vertragswesen/Abgeschlossen" verschoben. Für Nachrichten ohne Antwort wird nach einer festgelegten Anzahl von Tagen eine Erinnerung versendet. 

---

### **Funktionen**

1. **Überwachung von E-Mails:**
   - Das Makro prüft E-Mails im Ordner "Vertragswesen/Nachverfolgung", die älter als N Tage sind.

2. **Behandlung von Antworten:**
   - Wenn eine Antwort mit demselben Betreff im Posteingang gefunden wird, werden die ursprüngliche E-Mail und die Antwort in den Ordner "Vertragswesen/Abgeschlossen" verschoben.

3. **Erinnerungsversand:**
   - Für Nachrichten ohne Antwort wird eine Erinnerung an den ursprünglichen Empfänger gesendet, sofern sie älter als die festgelegte Anzahl von Tagen ist.

4. **Automatische Zusammenfassung:**
   - Nach Abschluss der Prüfung wird eine E-Mail mit einer Übersicht erstellt, die die Personen auflistet, die geantwortet haben, sowie die Empfänger der Erinnerung.

---

### **Installation**

1. **VBA-Editor öffnen:**
   - Öffne Microsoft Outlook.
   - Drücke `ALT + F11`, um den VBA-Editor zu starten.

2. **Modul hinzufügen:**
   - Wähle im Menü `Einfügen > Modul`, um ein neues Modul zu erstellen.
   - Kopiere den vollständigen VBA-Code und füge ihn in das neue Modul ein.

3. **Speichern und Schließen:**
   - Speichere den Code und schließe den VBA-Editor.

---

### **Verwendung**

1. **Makro manuell starten:**
   - Gehe in Outlook zu `Entwicklertools > Makros`, wähle `VertragErinnerungMitOrdnern` aus und klicke auf `Ausführen`.
   - Das Makro überprüft die E-Mails und gibt eine Zusammenfassung aus.

2. **Automatische Ausführung (optional):**
   - Um das Makro automatisch auszuführen, kannst du den Windows Task Scheduler oder Power Automate verwenden:
     - Erstelle eine .vbs-Datei, die das Makro in regelmäßigen Abständen ausführt.
     - Plane diese Datei mit einem Zeitplaner deiner Wahl.

---

### **Anpassungen**

1. **Überwachungszeitraum anpassen:**
   - Die Anzahl der Tage, nach denen Nachrichten überprüft werden, kann in der Zeile angepasst werden:  
     ```vba
     If olMail.ReceivedTime <= DateAdd("d", -N, dtToday) Then
     ```
     Ersetze `N` durch die gewünschte Anzahl an Tagen.

2. **Ordnernamen ändern:**
   - Die Zielordner für Nachverfolgung und abgeschlossene Nachrichten können zentral definiert werden:  
     ```vba
     NachverfolgungOrdnerName = "Vertragswesen/Nachverfolgung"
     AbgeschlossenOrdnerName = "Vertragswesen/Abgeschlossen"
     ```
     Passe die Namen an deine Ordnerstruktur an.

3. **Empfänger der Zusammenfassung:**
   - Ändere den Empfänger der Zusammenfassungs-E-Mail in der Zeile:
     ```vba
     ZusammenfassungMail.To = "deine.email@domain.com"
     ```

---

### **Anforderungen**

- **Microsoft Outlook:** Mit aktiviertem VBA.
- **E-Mail-Ordnerstruktur:** Die Ordner "Vertragswesen/Nachverfolgung" und "Vertragswesen/Abgeschlossen" müssen existieren.

---

### **Hinweis**

- Dieses Makro greift auf deine Outlook-Daten zu und versendet E-Mails. Stelle sicher, dass du es nur in vertrauenswürdigen Umgebungen verwendest.
- Achte darauf, dass du regelmäßig deine E-Mail-Ordner überprüfst, um sicherzustellen, dass keine wichtigen Nachrichten verloren gehen.

---

**Erstellt mit ❤️ von João**
