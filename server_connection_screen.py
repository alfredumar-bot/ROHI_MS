from kivymd.uix.screen import MDScreen


class ServerConnectionScreen(MDScreen):
    """
    PostgreSQL server connection screen: connection fields, Test/Save/Connect/
    Disconnect/Synchronize Now actions, and a status card. All logic lives in
    main.py's ROHIAttendanceApp (calling into pg_sync.py), mirroring the rest
    of the app.
    """
    pass
