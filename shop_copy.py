# This class is contains the logic of the shop copy itself. 
# There are setter and getter methods for the order number, which is stored here.
# There are three primary methods for creating the shop copy:
# The first queries the Syteline SQL database to get the part numbers, line numbers,
# and quantities for each order. The second organizes them into the correct order
# and combines part numbers that appear on multiple lines. The third looks up the 
# PDF drawings for each part, performs OCR to determine where the Job, Item, and Qty
# fields are on the page, draws the correct values on a new, temporary PDF package, prints the
# package to a physical printer, and deletes the temporary package.

import os
import csv
import tempfile
from pdf2image import convert_from_path
from PIL import Image, ImageDraw, ImageFont
import img2pdf
import pytesseract

class ShopCopy:
    def __init__(self):
        # Values for the order number (as string) and for the table of parts, line numbers, 
        # and quantities (as list of strings)
        self.order_number = None
        self.order_file = None
        self.order_data_table = None

    def set_order_number(self, order_number):
        # Sets the order number
        self.order_number = order_number
        self.order_file = f"{order_number}.csv"

    def get_order_number(self):
        # Gets the order number
        return self.order_number

    def query_customer_order_table(self):
        # Open connection to SQL database, query table, put data in a list, close connection, return list
        pass

    def organize_shop_copy_data(self):
    # This method acquires the data in list form, and reorganizes it so that duplicate parts are combined.
    # Right now, it reads this from a CSV file with dummy data. 
    # Eventually this list will be returned by the query_customer_order_table method.
        
        # Open CSV file with dummy data (for testing)
        with open(self.order_file, 'r') as shop_copy_file:
            reader = csv.reader(shop_copy_file)
            # Extract and skip over the header
            header = next(reader)
            # Read CSV into list where each item in the list is a string containing the contents of each "cell"
            shop_copy_input_table = [row for row in reader]

        # Create a dictionary of part numbers where each part number has an associated line item and quantity field.
        part_number_dict = {}

        # For each part number, check to see if it's in the list already. If it is, append its line and quantity
        # data to the respective dictionary entries. If it isn't, create a new dictionary entry for the part number.
        for row in shop_copy_input_table:
            line_item, quantity, part_number = row
            if part_number in part_number_dict:
                part_number_dict[part_number]['line_items'].append(line_item)
                part_number_dict[part_number]['quantities'].append(quantity)
            else:
                part_number_dict[part_number] = {
                    'line_items': [line_item],
                    'quantities': [quantity]
                }

        # Convert to a list for ease of displaying in the UI and using in the print_shop_copy method
        shop_copy_output_table = []
        for part_number, data in part_number_dict.items():
            line_items = ",".join(data['line_items'])
            quantities = ",".join(data['quantities'])
            shop_copy_output_table.append([part_number, line_items, quantities])
            self.order_data_table = shop_copy_output_table
     
        return shop_copy_output_table

    def print_shop_copy(self):
        # Iterate through each drawing to be printed
        for i, row in enumerate(self.order_data_table):
            drawing_filename = f"{row[0]}.pdf"
            job_text = self.order_number
            item_text = row[1]
            qty_text = row[2]
            
            img = convert_from_path(drawing_filename)
            img = img[0]

            # This line is what runs OCR on the part drawing.
            ocr_data = pytesseract.image_to_data(img, output_type='dict')

            # These are the strings we are looking for on the drawings.
            job_str = "Job"
            item_str = "Item(s)"
            qty_str = "Qty."

            # The OCR returns a dictionary with headings for the word found, and then location, height, and width in pixels.
            # Establish Job block position
            for i, text in enumerate(ocr_data['text']):
                if job_str in text:
                    job_left = ocr_data['left'][i]
                    job_top = ocr_data['top'][i]
                    job_width = ocr_data['width'][i]
                    job_height = ocr_data['height'][i]

            # Establish Item block position
            for i, text in enumerate(ocr_data['text']):
                if item_str in text:
                    item_left = ocr_data['left'][i]
                    item_top = ocr_data['top'][i]
                    item_width = ocr_data['width'][i]
                    item_height = ocr_data['height'][i]

            # Establish Quantity block position
            for i, text in enumerate(ocr_data['text']):
                if qty_str in text:
                    qty_left = ocr_data['left'][i]
                    qty_top = ocr_data['top'][i]
                    qty_width = ocr_data['width'][i]
                    qty_height = ocr_data['height'][i]

            # Create object on which to write data
            draw = ImageDraw.Draw(img)
            font = ImageFont.truetype("arial.ttf", size=30)

            # Determine start points of text to be inserted. Determining these x and y coordinates just took 
            # trial-and-error. Ideally, we will test for drawing type (A, B, or C), and have specific values
            # optimized for each type. For now, these work well enough for all three.
            job_x = job_left + (3 * job_width) + 20
            job_y = job_top - (job_height / 2)

            item_x = job_left + (3 * job_width) + 20
            item_y = item_top - (job_height / 2)

            qty_x = job_left + (3 * job_width) + 20
            qty_y = qty_top - (job_height / 2)

            # Draw text on the shop copy drawing
            draw.text((job_x, job_y), job_text, font=font)
            draw.text((item_x, item_y), item_text, font=font)
            draw.text((qty_x, qty_y), qty_text, font=font)

            modified_image_path = []
            output_pdf_path = f"modified_{os.path.basename(drawing_filename)}"

            # The below is a less-than-ideal way of saving but I was running into an issue saving directly as a PDF.
            # Save the image to a temporary file
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as f:
                img.save(f, format="PNG")
                modified_image_path.append(f.name)

            # Save the modified images as a PDF
            with open(output_pdf_path, "wb") as f:
                f.write(img2pdf.convert(modified_image_path))

            # Remove temporary image files
            for path in modified_image_path:
                os.remove(path)


            
