#!/usr/bin/env python3

import sys

import xlsxwriter
from PIL import Image


if __name__ == '__main__':

	if len(sys.argv) < 3:
		print('USAGE {} <input-file> <output-file>.xlsx'.format(sys.argv[0]))
		sys.exit(1)

	input_file = sys.argv[1]
	output_file = sys.argv[2]

	workbook = xlsxwriter.Workbook(output_file)
	worksheet = workbook.add_worksheet()

	formats = {}

	i = Image.open(input_file)
	s = i.size
	width = s[0]
	height = s[1]
	
	MAX_WIDTH = MAX_HEIGHT = 255
	MAX_WIDTH = width - 1
	MAX_HEIGHT = height

	image_data = list(i.getdata())

	current_column = 0
	current_row = 0

	worksheet.set_column('A:ZZ', 1.5)



	for pxl in image_data:

		if current_row >= MAX_HEIGHT:
			break


		if current_column > MAX_WIDTH:
			current_column = 0
			current_row += 1

		color_hex = '#{:02X}{:02X}{:02X}'.format(*pxl)

		if color_hex not in formats.keys():
			frmt = workbook.add_format({'bg_color': color_hex, 'pattern': 1})
			formats[color_hex] = frmt
		else:
			frmt = formats[color_hex]
		
		worksheet.write_blank(current_row, current_column, None, frmt)

		current_column += 1

	workbook.close()

