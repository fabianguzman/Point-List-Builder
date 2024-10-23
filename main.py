from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment
import wiring

#DEFAULT_NUM_POINTS_D20S = 64
#DEFAULT_NUM_POINTS_D20A = 32
#DEFAULT_NUM_POINTS_D20K = 32


class Sheet_Handler:
  pass

class D20_Module:

  # add calc alias
  
  def __init__(self, type, number, address):
    self.type = type
    self.number = number
    self.address = address

    match self.type:
      case 'S':
        self.num_points = 64
        self.wiring_list = wiring.status_wiring
      case 'A':
        self.num_points = 32
        self.wiring_list = wiring.analog_wiring
      case 'K':
        self.num_points = 32
        self.wiring_list = wiring.control_wiring

    '''if self.type == 'S':
      self.num_points = 64
      self.wiring_list = wiring.status_wiring
    elif self.type == 'A':
      self.num_points = 32
      self.wiring_list = wiring.analog_wiring
    else:
      self.num_points = 32
      self.wiring_list = wiring.control_wiring'''
      



last_status_pl_pt = last_analog_pl_pt = last_control_pl_pt = 0

module_points = {'S': 64, 'A': 32, 'K': 32}
#module_sheets = {'S': 64, 'A': 32, 'K': 32}


def calculate_points():
  num_of_D20S_modules = 0


def add_D20_module(module_num, module_type, module_address):
  
  if module_type == 'S':
    ws = status_sheet
    pl_pt = last_status_pl_pt
  elif module_type == 'A':
    ws = analog_sheet
    pl_pt = last_analog_pl_pt
  else:
    ws = control_sheet
    pl_pt = last_control_pl_pt
    
  for i in range(1, module_points[module_type]+1):
    pl_pt = pl_pt + 1
    ws.append([0, pl_pt, '', '', '', 'D20 LINK', '', 'IED', 
      f'I/O MODULE #{module_num} (D20{module_type})', f'{module_address}', 
      'D20 LINK 1', '', '', '', i, 'SPARE', '', '', '', ''] 
       + wiring.get_wiring(module_type, i))


'''
def add_D20S_module(module_num, module_address):
  pass#for row in status_sheet.iter_rows(min_row=10, max_row=)
'''

def format_cells(sheet, cell_range, skips):
  center_align = Alignment(horizontal='center')
  left_align = Alignment(horizontal='general')
  thin_border = Side(border_style="thin", color="000000")
  for row in sheet[cell_range]:
    for cell in row:
      cell.border = Border(top=thin_border, bottom=thin_border,
                           right=thin_border, left=thin_border)
      if cell.coordinate[0] in skips:
        cell.alignment = left_align
      else:
        cell.alignment = center_align


TOTAL_STATUS_POINTS = 64
TOTAL_ANALOG_POINTS = 32
TOTAL_CONTROL_POINTS = 32

output_filename = 'RESULT'

wb = load_workbook(filename = 'Template G500.xlsx')


#d20module_points = {'A': 32, 'C': 32, 'S':64}

wb_sheets = {
  'S': wb['Status & Alarms'],
  'A': wb['Analogs'],
  'K': wb['Controls']
}

d20_addresses = ['0x03', '0x05', '0x06', '0x09', '0x0A','0x0C', '0x0F']

status_sheet = wb['Status & Alarms']
analog_sheet = wb['Analogs']
control_sheet = wb['Controls']

# READ INPUT
with open('input.txt', 'r') as f:
  for line in f:
    if line.startswith('D20S'):
      add_D20_module(1, 'S','0x03')
    elif line.startswith('D20A'):
      add_D20_module(2, 'A','0x05')
    else:
      add_D20_module(3, 'K','0x06')


format_cells(status_sheet, 'A10:AD73', 'PQ')
format_cells(analog_sheet, 'A10:AH41', 'NO')
format_cells(control_sheet, 'A10:Z41', 'MN')


wb.save(f'./{output_filename}.xlsx')
print(f'Points list successfully created: {output_filename}.xlsx')