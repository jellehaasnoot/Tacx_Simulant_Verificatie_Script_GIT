import xlsxwriter
from ValueConverter import ValueConverter


def logfile_analyser(logfile):
    global velocity_list, power_list
    sentences = []
    value_list = []
    velocity_list = []
    power_list = []
    speed = True
    power = True

    """
    Hier wordt het logfile geopend. 
    """
    log = open(logfile)
    with open(logfile) as f:
        for lines, l in enumerate(f):
            pass

    """
    Hier zullen de belangrijke zinnen uit het logbestand worden onttrokken, door te kijken naar een karakter combinatie die
    in iedere zin voor komt.
    """
    for n in range(lines):
        sentence = log.readline()
        if "Rx:" in sentence:
            sentences.append(sentence)

    """
    Dit deel splitst de gesorteerde regels in delen (delen zijn de stukken die zich tussen 2 spaties bevinden). Daarna wordt
    het hexadecimale getal uit deze zin gefilterd.
    """
    time_value = []
    for i in range(len(sentences)):
        sentence = sentences[i].split()
        index = sentence.index("Rx:")
        value_raw = sentence[index+1]
        value = value_raw.replace("[", "").replace("]", "")     # This will removes the useless characters
        value_list.append(value)
        time_value.append(sentence[index-2])

    print(time_value)

    """
    Deze waardes worden gesorteerd per indexatie in het begin van het hexadecimale getal, in de index '10' staat de snelheid
    , in index '19' de power en de rest is voor dit bestand niet belangrijk en zal daarom niet geanalyseerd worden.
    """
    for i in range(len(value_list)):
        value_list_characters = list(value_list[i])
        if value_list_characters[0] == '1' and value_list_characters[1] == '0' and speed:      # index waar snelheid in staat
            speed = False
            power = True
            speed_values_raw = [value_list_characters[10], value_list_characters[11], value_list_characters[8], value_list_characters[9]]
            speed_values_raw_string = "".join(speed_values_raw)
            value_converter = ValueConverter()
            velocity_bin = value_converter.hex_to_bin(speed_values_raw_string)
            velocity = value_converter.bin_to_dec(velocity_bin)*3.6/1000
            velocity_list.append(velocity)
        elif value_list_characters[0] == '1' and value_list_characters[1] == '9' and power:      # index waar power in staat
            speed = True
            power = False
            power_values_raw = [value_list_characters[13], value_list_characters[10], value_list_characters[11]]
            power_values_raw_string = "".join(power_values_raw)
            value_converter = ValueConverter()
            power_bin = value_converter.hex_to_bin(power_values_raw_string)
            wattage = value_converter.bin_to_dec(power_bin)
            power_list.append(wattage)
        else:
            pass


logfile_analyser('log.txt')
velocity_list_high = velocity_list
power_list_high = power_list
number_of_measurements_high = min(len(power_list), len(velocity_list))
logfile_analyser('log1.txt')
velocity_list_low = velocity_list
power_list_low = power_list
number_of_measurements_low = min(len(power_list), len(velocity_list))

"""
Initialize writing an excel file.
"""
name = input("Which trainer has been tested?  ")
excel = xlsxwriter.Workbook(name + ".xlsx")
graph = excel.add_chart({'type': 'scatter'})
worksheet = excel.add_worksheet()

"""
Setting variables for excel file.
"""
bold = excel.add_format({'bold': True})
underline = excel.add_format({'bold': True, 'underline': True})
graph.set_y_axis({'name': 'Power [W]'})
graph.set_x_axis({'name': 'Velocity [km/h]'})
graph.set_title({'name': 'Operating range ' + name})
graph.set_size({'width': 720, 'height': 576})
worksheet.set_column('A:E', 14)

"""
Writing to excel file.
"""
worksheet.write('A1', 'Tested with highest gradient (without slip)', underline)
worksheet.write('A2', 'Velocity [km/h]', bold)
worksheet.write('B2', 'Power [W]', bold)
worksheet.write('D1', 'Tested with lowest gradient (without slip)', underline)
worksheet.write('D2', 'Velocity [km/h]', bold)
worksheet.write('E2', 'Power [W]', bold)

worksheet.write_column(2, 0, velocity_list_high)
worksheet.write_column(2, 1, power_list_high)
worksheet.write_column(2, 3, velocity_list_low)
worksheet.write_column(2, 4, power_list_low)

"""
Writing to graph.
"""
graph.add_series({
    'categories': [worksheet.name] + [2, 3] + [number_of_measurements_low + 2, 3],
    'values':     [worksheet.name] + [2, 4] + [number_of_measurements_low + 2, 4],
    'line':       {'color': 'black'},
    'name':       'lowest gradient',
})

graph.add_series({
    'categories': [worksheet.name] + [2, 0] + [number_of_measurements_high + 2, 0],
    'values':     [worksheet.name] + [2, 1] + [number_of_measurements_high + 2, 1],
    'line':       {'color': 'black'},
    'name':       'Highest gradient',
})

worksheet.insert_chart('H4', graph)
excel.close()
