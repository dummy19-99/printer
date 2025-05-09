import win32print

def list_printers():
    printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
    default_printer = win32print.GetDefaultPrinter()

    print("接続されているプリンター一覧:")
    for printer in printers:
        print(f"- {printer[2]}")

    print(f"\nデフォルトプリンター: {default_printer}")

if __name__ == "__main__":
    list_printers()
