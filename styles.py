"""
Custom styles for inhouse work

TESTS WRITTEN AND WORKING 12/2019 in tests.py PrintSettingsTests.
Check the files Testing_Borders and Testing_Printsettings in
unittests directory
"""

import re

from openpyxl.styles import Border, Side
from openpyxl.worksheet.properties import PageSetupProperties


def print_settings(sheet, widths=None, header=None, one_page=True, landscape=True):
    """
    With specifying the sheet, this can assign the sheets
    print settings to be one page, landscape/portrait,
    as well as assign column sizes with the width argument,
    where each item in the list should be a number,
    and defines x amount of columns.
    For excel, the width for each column in pixels is x*7
    """
    # Allowing header arg to not be set
    if header is None:
        header = []
    # Allowing widths arg to not be set
    if widths is None:
        widths = []
    # Creating a list of column indices for excel. From 'A' to 'Z'
    # Will expand alpha to accomodate if you need to set the width
    # or header of more than 26 columns
    alpha = []
    m_char = 65
    if len(header) > 26:
        m_char = 65+(len(header)+26)

    for letter in range(65, 91):
        alpha.append(chr(letter))
        for letter2 in range(65, m_char):
            a = chr(letter)
            b = chr(letter2)
            alpha.append(a + b)

    for i in range(len(widths)):
        sheet.column_dimensions[alpha[i]].width = widths[i]

    # Allowing header arg to not be set
    if header is None:
        header = []

    for i in range(len(header)):
        sheet[f'{alpha[i]}1'] = header[i]

    main_props = sheet.sheet_properties
    sheet.print_options.horizontalCentered = True
    sheet.print_options.verticalCentered = True

    if one_page:
        main_props.pageSetUpPr = PageSetupProperties(autoPageBreaks=False, fitToPage=True)
        sheet.print_options.horizontalCentered = True
        sheet.print_options.verticalCentered = True
    else:
        main_props.pageSetUpPr = PageSetupProperties(autoPageBreaks=False)
        sheet.page_setup.fitToHeight = False
        sheet.page_setup.fitToWidth = True

    if landscape:
        sheet.page_setup.orientation = sheet.ORIENTATION_LANDSCAPE
    else:
        sheet.page_setup.orientation = sheet.ORIENTATION_PORTRAIT

    sheet.print_title_rows = '1:1'


def full_border(ws, cell_range, border: list = None, rgb_hex: list = None):
    """
        Openpyxl can create a border on a single cell,
    but I’ve created this to create a border encompassing a range of cells.

    :param ws: Openpyxl worksheet object
    :param cell_range: Range of cells, i.e A12, D12:D15, B2:I16 etc
    :param border: Denotes border style - Passed as list of length 1-4,
        [All sides]
        [Right/Left, Top/Bottom]
        [Right, Top/bottom, Left]
        [Right, top, left, bottom]
        {‘double’, ‘mediumDashDotDot’,
        ‘slantDashDot’, ‘dashDotDot’,
        ‘dotted’, ‘hair’, ‘mediumDashed’,‘dashed’, ‘dashDot’, ‘thin’,
        ‘mediumDashDot’, ‘medium’, ‘thick’}
    :param rgb_hex: String to pass a colour value, in RGB hex format
        Passed as a list of length 1-4,
        [All sides]
        [Right/Left, Top/Bottom]
        [Right, Top/bottom, Left]
        [Right, top, left, bottom]
    :return: Returns an unsaved worksheet object with the new border
    """
    # New and improved, can draw full borders around any size box of cells
    # Cell range as a string 'A1:G7'

    # Assigning border side styles
    if border is None or len(border) == 0:
        border = ['thin']

    border = list_parser(border)

    # Assigning the border side colours
    if rgb_hex is None or len(rgb_hex) == 0:
        rgb_hex = ['000000']

    rgb_hex = list_parser(rgb_hex)

    # Single column box borders #
    # Top cell border
    vert_top_border = Border(right=Side(border_style=border['right'], color=rgb_hex['right']),
                             left=Side(border_style=border['left'], color=rgb_hex['left']),
                             top=Side(border_style=border['top'], color=rgb_hex['top']))
    # Middle cells border
    vert_middle_border = Border(left=Side(border_style=border['left'], color=rgb_hex['left']),
                                right=Side(border_style=border['right'], color=rgb_hex['right']))
    # Bottom cell border
    vert_bottom_border = Border(left=Side(border_style=border['left'], color=rgb_hex['left']),
                                right=Side(border_style=border['right'], color=rgb_hex['right']),
                                bottom=Side(border_style=border['bottom'], color=rgb_hex['bottom']))

    # Single row box borders #
    # Leftmost cell border
    horiz_top_border = Border(left=Side(border_style=border['left'], color=rgb_hex['left']),
                              bottom=Side(border_style=border['bottom'], color=rgb_hex['bottom']),
                              top=Side(border_style=border['top'], color=rgb_hex['top']))
    # Middle cells border
    horiz_middle_border = Border(top=Side(border_style=border['top'], color=rgb_hex['top']),
                                 bottom=Side(border_style=border['bottom'], color=rgb_hex['bottom']))
    # Bottom cell border
    horiz_bottom_border = Border(top=Side(border_style=border['top'], color=rgb_hex['top']),
                                 bottom=Side(border_style=border['bottom'], color=rgb_hex['bottom']),
                                 right=Side(border_style=border['right'], color=rgb_hex['right']))

    # Single cell border #
    full_outer_border = Border(left=Side(border_style=border['left'], color=rgb_hex['left']),
                               top=Side(border_style=border['top'], color=rgb_hex['top']),
                               right=Side(border_style=border['right'], color=rgb_hex['right']),
                               bottom=Side(border_style=border['bottom'], color=rgb_hex['bottom']))

    # Square & rectangle box borders #
    # Top left cell corner
    top_left_border = Border(left=Side(border_style=border['left'], color=rgb_hex['left']),
                             top=Side(border_style=border['top'], color=rgb_hex['top']))
    # Top right cell corner
    top_right_border = Border(right=Side(border_style=border['right'], color=rgb_hex['right']),
                              top=Side(border_style=border['top'], color=rgb_hex['top']))
    # Top cells borders, in between the two corners
    top_middle_border = Border(top=Side(border_style=border['top'], color=rgb_hex['top']))
    # Leftmost cell borders, in between upperleft and bottom left corners
    left_border = Border(left=Side(border_style=border['left'], color=rgb_hex['left']))
    # Rightmost cell borders, in between upperright and bottom right corners
    right_border = Border(right=Side(border_style=border['right'], color=rgb_hex['right']))
    # Bottom left cell corner
    bottom_left_border = Border(left=Side(border_style=border['left'], color=rgb_hex['left']),
                                bottom=Side(border_style=border['bottom'], color=rgb_hex['bottom']))
    # Bottom right cell corner
    bottom_right_border = Border(right=Side(border_style=border['right'], color=rgb_hex['right']),
                                 bottom=Side(border_style=border['bottom'], color=rgb_hex['bottom']))
    # Bottom cells borders, in between bottom two corners
    bottom_middle_border = Border(bottom=Side(border_style=border['bottom'], color=rgb_hex['bottom']))

    # If the cell range is just one cell, create a dummy range i.e 'A2' -> 'A2:A2'
    if len(cell_range) < 5:
        cell_range = cell_range + ':' + cell_range

    # Uppercases the cell range to handle lowercase parameters
    cell_range = cell_range.upper()
    # Split up the two ends of the cell range argument
    cr_endpoints = str.split(cell_range, ':')
    cr_start = re.split(r'(\d+)', cr_endpoints[0])
    cr_end = re.split(r'(\d+)', cr_endpoints[1])

    width = ord(cr_end[0]) - ord(cr_start[0]) + 1
    height = int(cr_end[1]) - int(cr_start[1]) + 1

    total_cells = width * height

    rows = ws[cell_range]
    count = 1
    if width != 1:
        if height == 1:
            for row in rows:
                for cell in row:
                    if count == 1:
                        cell.border = horiz_top_border

                    elif count == width:
                        cell.border = horiz_bottom_border

                    else:
                        cell.border = horiz_middle_border

                    count += 1

        else:
            for row in rows:
                for cell in row:
                    if count <= width:
                        if count == 1:
                            cell.border = top_left_border

                        elif count == width:
                            cell.border = top_right_border

                        else:
                            cell.border = top_middle_border

                    if (count > width) & (count <= (total_cells - width)):
                        if count % width == 1:
                            cell.border = left_border

                        elif count % width == 0:
                            cell.border = right_border

                    if count > (total_cells - width):
                        if count == (total_cells - width) + 1:
                            cell.border = bottom_left_border

                        elif count == total_cells:
                            cell.border = bottom_right_border

                        else:
                            cell.border = bottom_middle_border

                    count += 1

    else:
        for row in rows:
            for cell in row:
                if len(rows) == 1:
                    cell.border = full_outer_border
                elif count == 1:
                    cell.border = vert_top_border

                elif count == len(rows):
                    cell.border = vert_bottom_border

                else:
                    cell.border = vert_middle_border

                count += 1


def list_parser(list_one: list):
    """
    :param list_one: Passed as list of length 1 or more
    :return: Returns a dictionary of length at 4, determined by the length of
        list_one
        Length 1: [list_one(0), list_one(0), list_one(0), list_one(0)]
        Length 2: [list_one(0), list_one(1), list_one(0), list_one(1)]
        Length 3: [list_one(0), list_one(1), list_one(2), list_one(0)]
        Length 4+: [list_one(0), list_one(1), list_one(2), list_one(3)]
    """
    std_dict = {'right': '', 'top': '', 'left': '', 'bottom': ''}

    if len(list_one) == 1:
        list_one *= 4

    elif len(list_one) == 2:
        list_one *= 2

    elif len(list_one) == 3:
        list_one.append(list_one[1])

    else:
        list_one = list_one[0:4]

    std_dict['right'] = list_one[0]
    std_dict['top'] = list_one[1]
    std_dict['left'] = list_one[2]
    std_dict['bottom'] = list_one[3]
    return std_dict
