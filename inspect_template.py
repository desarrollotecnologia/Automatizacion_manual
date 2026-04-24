from __future__ import annotations

from pathlib import Path


def main() -> int:
    import win32com.client  # type: ignore

    template = Path(r"C:\Users\USUARIO\Downloads\FRF_Rend ACTUALIZADO.xlsm")
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        try:
            excel.AutomationSecurity = 3
        except Exception:
            pass

        wb = excel.Workbooks.Open(
            str(template),
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
            CorruptLoad=1,
        )
        try:
            ws = wb.Worksheets("Rendimientos")
            headers = ws.Range("A13:Y13").Value[0]
            print("HEADERS:")
            for i, h in enumerate(headers, start=1):
                if h is None:
                    continue
                s = str(h).strip()
                if s:
                    print(i, s)

            data = ws.Range("A14:Y80").Value
            print("\nEJEMPLO (filas con destino):")
            for r_i, row in enumerate(data, start=14):
                if row is None:
                    continue
                destinos = [row[8], row[10], row[12], row[14], row[16], row[18], row[20]]
                if any(
                    d is not None and ("CAVA" in str(d).upper() or "/" in str(d))
                    for d in destinos
                ):
                    animal = row[1]
                    print(
                        r_i,
                        animal,
                        [str(d).strip() if d is not None else "" for d in destinos],
                    )
        finally:
            wb.Close(SaveChanges=False)
    finally:
        excel.Quit()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

