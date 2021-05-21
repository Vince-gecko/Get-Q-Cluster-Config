from openpyxl.styles import NamedStyle, Font, Border, Side, Alignment, PatternFill

# Creating Style for table title
style_title = NamedStyle(name="style_title")
style_title.font = Font(bold=True, size=11, color='FFFFFF')
style_title.border = Border(left=Side(border_style='thin', color='000000'),
                            right=Side(border_style='thin', color='000000'),
                            top=Side(border_style='thin', color='000000'),
                            bottom=Side(border_style='thin', color='000000'),)
style_title.alignment = Alignment(horizontal='center',
                                  vertical='center',
                                  text_rotation=0,
                                  wrap_text=False,
                                  shrink_to_fit=False,
                                  indent=0)
style_title.fill = PatternFill(fill_type='solid',
                               start_color='0085B2',
                               end_color='0085B2')

# Creating Style for normal tab

style_normal = NamedStyle(name="style_normal")
style_normal.font = Font(size=11, color='000000')
style_normal.border = Border(left=Side(border_style='thin', color='000000'),
                             right=Side(border_style='thin', color='000000'),
                             top=Side(border_style='thin', color='000000'),
                             bottom=Side(border_style='thin', color='000000'),)
style_normal.alignment = Alignment(horizontal='center',
                                   vertical='center',
                                   text_rotation=0,
                                   wrap_text=True,
                                   shrink_to_fit=False,
                                   indent=0)
