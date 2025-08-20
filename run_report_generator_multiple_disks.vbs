    Set WshShell = CreateObject("WScript.Shell")
    WshShell.Run "powershell.exe -NoProfile -WindowStyle Hidden -File ""report_generator_multiple_disks.ps1""", 0, False
    Set WshShell = Nothing