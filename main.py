from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment
import wiring


class Sheet_Handler:

  def __init__(self, sheet, total_points, text_columns, last_column):
    self.sheet = sheet
    self.total_points = total_points
    self.text_columns = text_columns
    self.last_column = last_column
    self.current_pl_pt = 1

  # def add_D20_module():
  #   pass

  def format_cells(self):
    center_align = Alignment(horizontal='center')
    left_align = Alignment(horizontal='general')
    thin_border = Side(border_style="thin", color="000000")
    cell_range = f'A{10}:{self.last_column}{self.total_points+9}'
    for row in self.sheet[cell_range]:
      for cell in row:
        cell.border = Border(top=thin_border, bottom=thin_border,
                            right=thin_border, left=thin_border)
        if cell.coordinate[0] in self.text_columns:
          cell.alignment = left_align
        else:
          cell.alignment = center_align


class D20_Module:

  # add calc alias
  
  def __init__(self, type, number, address, dummy_boards):
    self.type = type
    self.number = number
    self.address = address
    self.dummy_boards = dummy_boards

    match self.type[-1]:
      case 'S':
        self.num_points = 64
        #self.wiring_list = wiring.status_wiring
      case 'A':
        self.num_points = 32
        #self.wiring_list = wiring.analog_wiring
      case 'K':
        self.num_points = 32
        #self.wiring_list = wiring.control_wiring
      
 

# def add_D20_module(sheet, module):
#   for i in range(module.num_points):
#     pl_pt = pl_pt + 1
#     sheet.append([0, pl_pt, '', '', '', 'D20 LINK', '', 'IED', 
#       f'I/O MODULE #{module.number} (D20{module.type})', f'{module.address}', 
#       'D20 LINK 1', '', '', '', i, 'SPARE', '', '', '', ''] 
#        + wiring.get_wiring(module, i))



# def format_cells(sheet, cell_range, skips):
#   center_align = Alignment(horizontal='center')
#   left_align = Alignment(horizontal='general')
#   thin_border = Side(border_style="thin", color="000000")
#   for row in sheet[cell_range]:
#     for cell in row:
#       cell.border = Border(top=thin_border, bottom=thin_border,
#                            right=thin_border, left=thin_border)
#       if cell.coordinate[0] in skips:
#         cell.alignment = left_align
#       else:
#         cell.alignment = center_align


#######################################################################################

output_filename = 'RESULT'

wb = load_workbook(filename = 'Template G500.xlsx')


d20_module_points = {'A': 32, 'C': 32, 'S':64}

wb_sheets = {
  'S': wb['Status & Alarms'],
  'A': wb['Analogs'],
  'K': wb['Controls']
}

d20_addresses = ['0x03', '0x05', '0x06', '0x09', '0x0A', '0x0C', '0x0F', '0x11',
                 '0x12', '0x14', '0x17', '0x18', '0x1B', '0x1D', '0x1E', '0x21',
                 '0x22', '0x24', '0x27', '0x28', '0x2B', '0x2D', '0x2E', '0x30' ]

# READ INPUT
D20S_modules = []
D20A_modules = []
D20K_modules = []
total_status_points = 0
total_analog_points = 0
total_control_points = 0

with open('input.txt', 'r') as f:
  for i, line in enumerate(f):
    line_inputs = line.split(',')
    in_D20_type = line_inputs[0].strip(' \n')
    in_dummy_boards = [board.strip(' \n') for board in line_inputs[1:]]
    if in_D20_type[-1] == 'S':
      D20S_modules.append(D20_Module(in_D20_type, i+1, d20_addresses[i], in_dummy_boards))
      total_status_points = total_status_points + 64
      # add_D20_module(1, 'S','0x03')
    elif in_D20_type[-1] == 'A':
      D20A_modules.append(D20_Module(in_D20_type, i+1, d20_addresses[i], in_dummy_boards))
      total_analog_points = total_analog_points + 32
      # add_D20_module(2, 'A','0x05')
    else:
      D20K_modules.append(D20_Module(in_D20_type, i+1, d20_addresses[i], in_dummy_boards))
      total_control_points = total_control_points + 32
      # add_D20_module(3, 'K','0x06')


status_sh = Sheet_Handler(wb['Status & Alarms'], total_status_points, 'PQ', 'AD')
analog_sh = Sheet_Handler(wb['Analogs'], total_analog_points, 'NO', 'AH')
control_sh = Sheet_Handler(wb['Controls'], total_control_points, 'MN', 'Z')

print('processing status')
for module in D20S_modules:
  for i in range(module.num_points):
    status_sh.sheet.append([0, status_sh.current_pl_pt, '', '', '', 'D20 LINK', '', 'IED', 
      f'I/O MODULE #{module.number} ({module.type})', f'{module.address}', 
      'D20 LINK 1', '', '', '', i+1, 'SPARE', '', '', '', '']
      + wiring.get_wiring(module, i+1))
    status_sh.current_pl_pt = status_sh.current_pl_pt + 1


print('processing analogs')
for module in D20A_modules:
  for i in range(module.num_points):
    analog_sh.sheet.append([0, analog_sh.current_pl_pt, '', '', '', 'D20 LINK', '', 'IED', 
      f'I/O MODULE #{module.number} ({module.type})', f'{module.address}', 
      'D20 LINK 1', '', i+1, 'SPARE', '', '', '', 2032/32767, 0, '', '' , '', '', '']
      + wiring.get_wiring(module, i+1))
    analog_sh.current_pl_pt = analog_sh.current_pl_pt + 1


print('processing control')
for module in D20K_modules:
  for i in range(module.num_points):
    #print('got here, module.num_points = ', module.num_points)
    control_sh.sheet.append([0, control_sh.current_pl_pt, '', '', '', 'D20 LINK', '', 'IED', 
      f'I/O MODULE #{module.number} ({module.type})', f'{module.address}', 
      'D20 LINK 1', i+1, 'SPARE', '']
      + wiring.get_wiring(module, i+1))
    control_sh.current_pl_pt = control_sh.current_pl_pt + 1
    

# format_cells(status_sh.sheet, f'A{10}:AD{status_sh.total_points+9}', 'PQ')
# format_cells(analog_sh.sheet, f'A{10}:AH{analog_sh.total_points+9}', 'NO')
# format_cells(control_sh.sheet, f'A{10}:Z{control_sh.total_points+9}', 'MN')

status_sh.format_cells()
analog_sh.format_cells()
control_sh.format_cells()

wb.save(f'./{output_filename}.xlsx')
print(f'Points list successfully created: {output_filename}.xlsx')