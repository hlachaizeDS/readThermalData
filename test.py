from pywinauto import *
from time import *


app = Application().connect(title_re="Optris PI Connect")
print(app.windows())
window = app.window(title_re="Optris PI Connect")
window=window.Temperatures
window=window.child_window(auto_id='Item_0')

#window.set_focus()
#window.draw_outline()
#window.click()
print(window.legacy_properties())
#window.print_control_identifiers()

