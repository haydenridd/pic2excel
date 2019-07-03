import matplotlib.pyplot as plt

from skimage import data, color
from skimage.transform import rescale, resize, downscale_local_mean
from skimage.util import img_as_ubyte
from openpyxl import Workbook
from openpyxl.styles import PatternFill, colors, Color
from openpyxl.utils.cell import get_column_letter

# parameters
x_max = 250
y_max = 250
cell_pixel_dim = 12

# At 100% zoom 97*12 pixels vertical, 220*12 horizontal
default_x_zoom = 97 * 12
default_y_zoom = 220 * 12


# input image
input_image = data.astronaut()

# Get X/Y dimensions of image
l_x, l_y = input_image.shape[0], input_image.shape[1]

# If X or Y above max, scale image down maintaining ratio

# x scale factor
if l_x > x_max:
    x_sf = x_max / l_x
else:
    x_sf = 1

# y scale factor
if l_y > y_max:
    y_sf = y_max / l_y
else:
    y_sf = 1

# determine worst case scale factor
if x_sf < y_sf:
    sf = x_sf
else:
    sf = y_sf

# Scale image
if sf < 1:
    input_image = img_as_ubyte(rescale(input_image, sf, anti_aliasing=True))

# get new x/y size
l_x, l_y = input_image.shape[0], input_image.shape[1]

# Convert image matrix to excel matrix
wb = Workbook()
ws = wb.active
ws.title = "converted_image"

set_col_height = False

# Output excel workbook containing cell pixelated image
for row in range(0, l_x):

    ws.row_dimensions[row+1].height = 4.5

    for col in range(0, l_y):

        if not set_col_height:
            ws.column_dimensions[get_column_letter(col+1)].width = 0.83

        # Determine RGB from image array
        cell_hex = "{:02X}".format(input_image[row, col, 0]) + "{:02X}".format(input_image[row, col, 1])\
                    + "{:02X}".format(input_image[row, col, 2])

        # Set color using styles, Color takes ARGB hex input as AARRGGBB
        cell_color = Color(cell_hex)

        # Set cell fill
        sel_cell = ws.cell(column=col+1, row=row+1) # , value=1
        sel_cell.fill = PatternFill("solid", fgColor=cell_color)

    set_col_height = True

# Set zoom scale according to dimensions of photo

pixels_x = l_x * cell_pixel_dim
pixels_y = l_y * cell_pixel_dim

zoom_scale_x = int((default_x_zoom / pixels_x) * 100)
zoom_scale_y = int((default_y_zoom / pixels_y) * 100)

# if zoom_scale_x < zoom_scale_y:
#     ws.sheet_view.zoomScale = zoom_scale_x
# else:
#     ws.sheet_view.zoomScale = zoom_scale_y

ws.sheet_view.zoomToFit

# Output workbook
wb.save("TEST.xlsx")