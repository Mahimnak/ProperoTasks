# import openpyxl
# import xlrd
# from xls2xlsx import XLS2XLSX
# import re

# class Order:
    
#     def __init__(self) -> None:
#         pass
    
#     def create_output_file(self):
#         workbook = openpyxl.Workbook()
#         sheet = workbook.active

#         workbook.save('output\Output.xlsx')

#     def convert_to_xlsx_file(self):
#         #Converting .xls file to .xlsx file for easier formatting.
#         x2x = XLS2XLSX("input\Consolidated Order Statement 10022024.xls")
#         x2x.to_xlsx("Order_Statement.xlsx")
    
#     def extract_integer(self, string):
#         match = re.search(r'(\d+)', string)
#         if match:
#             return int(match.group(1))
#         else:
#             return None


#     def get_data(self):
#         """Data retrieval from the input file"""
#         vegetable_names = []
#         unit_quant = []
#         veg_quant = dict()

#         #Creating the Output file in the output folder.
#         Order.create_output_file(self)
        
#         #Accessing the input file from the input folder and retrieving the data.
#         wb = openpyxl.load_workbook('input\input.xlsx')
#         ws = wb.active

#         #Storing all the data into lists.
#         col_a = [cell.value for cell in ws['A'][1:] if cell.value is not None]
#         col_q = [cell.value for cell in ws['Q'][1:] if cell.value is not None]
#         col_r = [cell.value for cell in ws['R'][1:] if cell.value is not None]
#         col_y = [cell.value for cell in ws['Y'][1:] if cell.value is not None]
        
#         number_of_clients = [x for i, x in enumerate(col_a) if col_a.index(x) == i] # remove duplicates from list
        
#         for vegetable in col_r:
#             vegetable_names.append(vegetable.split("|")[0].strip())
#             veg_quant[vegetable.split("|")[0].strip()] = vegetable.split("-")[-1].strip()
#             unit_quant.append(vegetable.split("-")[-1].strip())

#         #All the unique vegetables that have been ordered
#         vegetables = {x for x in vegetable_names}
#         vegetables = list(vegetables)
#         sorted_vegetables = sorted(vegetables)

#         Order.create_template(self, sorted_vegetables, veg_quant, number_of_clients)
#         Order.add_data(self, number_of_clients)

#     def create_template(self, vegetables, uom, clients):
#         """Create a template excel file with headers."""
#         unit = ""
#         wb = openpyxl.load_workbook('output\Output.xlsx')
#         ws = wb.active

#         ws.cell(row = 1, column = 1, value = "Sr. No.")
#         ws.cell(row = 1, column = 2, value = "Product")
#         ws.cell(row = 1, column = 3, value = "UOM")
#         ws.cell(row = 1, column = 4, value = "Qty.")

#         for i in range(len(clients)):
#             ws.cell(row = 1, column = 5+i, value = clients[i])

#         for i in range(len(vegetables)):
#             if "Gram" in uom[vegetables[i]] or "gram" in uom[vegetables[i]] or "Kg" in uom[vegetables[i]] or "kg" in uom[vegetables[i]]:
#                 unit = "Kg"
#             elif "Pc" in uom[vegetables[i]] or "pc" in uom[vegetables[i]]:
#                 unit = "Pc"
#             ws.append([i+1, vegetables[i], unit])

#         wb.save('output\Output.xlsx')
        
#     def add_data(self, clients):
#         """Add data to existing work book."""
#         #16 - Q
#         #17 - R
#         col_counter = 1
#         row_counter = 1

#         wb1 = openpyxl.load_workbook('Order_Statement.xlsx')
#         ws1 = wb1.get_sheet_by_name("Input File")
        
#         wb2 = openpyxl.load_workbook('output\Output.xlsx')
#         ws2 = wb2.active

#         for row in ws1.iter_rows(values_only=True):
#             if row[0] in clients:
#                 for col_idx, cell_value in enumerate(ws2[1], start=1):  # Iterate over cells in the first row
#                     if cell_value.value == row[0]:  # Check if the cell value matches the client identifier
#                         col_counter = col_idx  # Update the column counter with the column index
#                         break
#                 for row_value in ws2.iter_rows(values_only = True):
#                     if any(cell == row[17].split("|")[0].strip() for cell in row_value):
#                         row_of_veg = row_counter
#                         break
#                     else:
#                         row_counter += 1
#                 qty = Order.extract_integer(self, row[17].split("-")[-1].strip())
#                 if qty != None:
#                     if "Gram" in row[17].split("-")[-1].strip() or "gram" in row[17].split("-")[-1].strip():
#                         qty = float(qty/1000)
#                         qty *= row[16]
#                     else:
#                         qty *= row[16]
#                     ws2.cell(row = row_of_veg, column = col_counter, value = qty)
#                 wb2.save('output\Output.xlsx')
#             row_counter = 1


                
#     def testing(self):
#         wb1 = openpyxl.load_workbook('Order_Statement.xlsx')
#         ws1 = wb1.get_sheet_by_name("Input File")

#         for row in ws1.iter_rows(values_only=True):
#             print(row)
#             break
# # Create a new order object and call its method to retrieve data from Excel file
# order = Order()
# order.get_data()


import openpyxl
import xlrd
from xls2xlsx import XLS2XLSX
import re

class Order:
    def __init__(self):
        pass

    def create_output_file(self, filename="Output.xlsx"):
        """Create a new output file with a default name 'Output.xlsx'."""
        try:
            workbook = openpyxl.Workbook()
            sheet = workbook.active

            workbook.save(f'output/{filename}')
        except Exception as e:
            print(f"Error creating output file: {e}")

    def convert_to_xlsx_file(self, filename="Consolidated Order Statement 10022024.xls"):
        """Convert a .xls file to .xlsx file for easier formatting."""
        try:
            x2x = XLS2XLSX(f"input/{filename}")
            x2x.to_xlsx(f"Order_Statement.xlsx")
        except Exception as e:
            print(f"Error converting to xlsx file: {e}")

    def extract_integer(self, string):
        """Extract integer from a string."""
        try:
            match = re.search(r'(\d+)', string)
            if match:
                return int(match.group(1))
            else:
                return None
        except Exception as e:
            print(f"Error extracting integer: {e}")

    def get_data(self):
        """Retrieve data from the input file."""
        try:
            vegetable_names = []
            unit_quant = []
            veg_quant = {}

            # Create the output file
            self.create_output_file()

            # Load the input file
            wb = openpyxl.load_workbook(f'Order_Statement.xlsx')
            ws = wb.get_sheet_by_name("Input File")

            # Store data in lists
            col_a = [cell.value for cell in ws['A'][1:] if cell.value is not None]
            col_q = [cell.value for cell in ws['Q'][1:] if cell.value is not None]
            col_r = [cell.value for cell in ws['R'][1:] if cell.value is not None]
            col_y = [cell.value for cell in ws['Y'][1:] if cell.value is not None]

            number_of_clients = list(set(col_a))  # remove duplicates from list

            for vegetable in col_r:
                veg_name = vegetable.split("|")[0].strip()
                veg_quant[veg_name] = vegetable.split("-")[-1].strip()
                unit_quant.append(vegetable.split("-")[-1].strip())
                vegetable_names.append(veg_name)

            # Get unique vegetables
            vegetables = list(set(vegetable_names))
            sorted_vegetables = sorted(vegetables)

            # Create template
            self.create_template(sorted_vegetables, veg_quant, number_of_clients)

            # Add data
            self.add_data(number_of_clients)

        except Exception as e:
            print(f"Error retrieving data: {e}")

    def create_template(self, vegetables, uom, clients):
        """Create a template excel file with headers."""
        try:
            unit = ""
            wb = openpyxl.load_workbook(f'output/Output.xlsx')
            ws = wb.active

            ws.cell(row=1, column=1, value="Sr. No.")
            ws.cell(row=1, column=2, value="Product")
            ws.cell(row=1, column=3, value="UOM")
            ws.cell(row=1, column=4, value="Qty.")

            for i in range(len(clients)):
                ws.cell(row=1, column=5+i, value=clients[i])

            for i in range(len(vegetables)):
                if "Gram" in uom[vegetables[i]] or "gram" in uom[vegetables[i]] or "Kg" in uom[vegetables[i]] or "kg" in uom[vegetables[i]]:
                    unit = "Kg"
                elif "Pc" in uom[vegetables[i]] or "pc" in uom[vegetables[i]]:
                    unit = "Pc"
                ws.append([i+1, vegetables[i], unit])

            wb.save(f'output/Output.xlsx')
        except Exception as e:
            print(f"Error creating template: {e}")

    def add_data(self, clients):
        """Add data to existing work book."""
        try:
            col_counter = 1
            row_counter = 1

            wb1 = openpyxl.load_workbook('Order_Statement.xlsx')
            ws1 = wb1.get_sheet_by_name("Input File")

            wb2 = openpyxl.load_workbook(f'output/Output.xlsx')
            ws2 = wb2.active

            for row in ws1.iter_rows(values_only=True):
                if row[0] in clients:
                    for col_idx, cell_value in enumerate(ws2[1], start=1):  # Iterate over cells in the first row
                        if cell_value.value == row[0]:  # Check if the cell value matches the client identifier
                            col_counter = col_idx  # Update the column counter with the column index
                            break
                    for row_value in ws2.iter_rows(values_only=True):
                        if any(cell == row[17].split("|")[0].strip() for cell in row_value):
                            row_of_veg = row_counter
                            break
                        else:
                            row_counter += 1
                    qty = self.extract_integer(row[17].split("-")[-1].strip())
                    if qty is not None:
                        if "Gram" in row[17].split("-")[-1].strip() or "gram" in row[17].split("-")[-1].strip():
                            qty = float(qty/1000)
                            qty *= row[16]
                        else:
                            qty *= row[16]
                        ws2.cell(row=row_of_veg, column=col_counter, value=qty)
                    wb2.save(f'output/Output.xlsx')
                row_counter = 1
            row_counter = 1
            total = 0

            for row in ws2.iter_rows(values_only=True):
                if row_counter==1:
                    row_counter+=1
                    pass
                else:
                    for i in range(4,len(row)):
                        if row[i]  != None:
                            total += float(row[i])
                    ws2.cell(row=row_counter, column=4, value=total)
                    row_counter+=1
                    total = 0
            wb2.save(f'output/Output.xlsx')

        except Exception as e:
            print(f"Error adding data: {e}")

    def testing(self):
        try:
            wb1 = openpyxl.load_workbook('Order_Statement.xlsx')
            ws1 = wb1.get_sheet_by_name("Input File")

            for row in ws1.iter_rows(values_only=True):
                print(row)
                break
        except Exception as e:
            print(f"Error testing: {e}")

# Create a new order object and call its method to retrieve data from Excel file
order = Order()
order.get_data()