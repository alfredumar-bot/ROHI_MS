ROHI Excel Auto-Sync Server

1. Install Python 3 and Flask on the ROHI server:
   pip install flask

2. Run:
   python excel_sync_server.py

3. The Android app sends Excel files to:
   http://SERVER_IP:8080/upload

4. In the app:
   Settings -> Server Connection -> Excel Auto-Sync Connections
   Put the same /upload URL into all four upload endpoint fields.

5. The server saves files into:
   excel_sync/attendance
   excel_sync/timesheet
   excel_sync/leave
   excel_sync/staff

6. If those folders are inside a Google Drive/OneDrive sync folder, the
   cloud service will automatically synchronize them.

IMPORTANT:
A Google Drive folder sharing URL by itself is not an upload API. The
folder links can be opened from the app, but actual automatic upload requires
an authenticated Drive API or an HTTP upload endpoint such as the one above.
For production, use HTTPS and authentication rather than exposing port 8080
directly to the internet.


v1.4: The Android app tests each configured report link individually. Green means the URL is reachable; red means the connection test failed. File upload still requires a configured HTTP upload endpoint or authenticated cloud API.
