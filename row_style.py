from openpyxl.styles import Border, Side, NamedStyle, Font, PatternFill, NumberFormatDescriptor, Alignment

# setting base styles of excel rows/cells
base = NamedStyle(name="base")
base.font = Font(name="Segoe UI Light", size=8)
border_style = Side(style="thin", color="000000")
base.border = Border(left=border_style, right=border_style,
                     top=border_style, bottom=border_style)
base.alignment = Alignment(horizontal="right")


# setting days_off style of excel rows/cells
days_off = NamedStyle(name="days_off")
days_off.font = Font(name="Segoe UI Light", size=8)
border_style = Side(style="thin", color="000000")
days_off.border = Border(left=border_style, right=border_style,
                         top=border_style, bottom=border_style)
days_off.fill = PatternFill("solid", fgColor="00C0C0C0")
days_off.alignment = Alignment(horizontal="right")
