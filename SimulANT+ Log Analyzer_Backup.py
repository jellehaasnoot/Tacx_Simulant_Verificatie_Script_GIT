import wx as wx
from ValueConverter import ValueConverter
import xlsxwriter
import os
import numpy as np


class Main(wx.Frame):
    def __init__(self, parent, title):
        """
        Initializing the program:

        1:
        :param parent:
        :param title:
        """
        wx.Frame.__init__(self, parent, title=title,
                          style=wx.DEFAULT_FRAME_STYLE & ~(wx.RESIZE_BORDER | wx.MAXIMIZE_BOX), size=(720, 480))
        self.CreateStatusBar()

        self.top_panel = wx.Panel(self)
        self.SetBackgroundColour("white")

        # Create the file menu
        file_menu = wx.Menu()

        menu_about = file_menu.Append(wx.ID_ABOUT, "&About", "Information about this program")
        file_menu.AppendSeparator()
        menu_file_open = file_menu.Append(wx.ID_FILE, "&Open files...", "Open a text file with this program")
        menu_exit = file_menu.Append(wx.ID_EXIT, "E&xit", "Terminate the program")

        # Create the menu bar
        menu_bar = wx.MenuBar()
        menu_bar.Append(file_menu, "&File")
        self.SetMenuBar(menu_bar)

        # Creating buttons
        self.exit_button = wx.Button(self.top_panel, -1, label='Exit', pos=(590, 360), size=(100, 30))
        self.reset_button = wx.Button(self.top_panel, -1, label='Reset Program', pos=(480, 360), size=(100, 30))
        self.open_xlsx_button = wx.Button(self.top_panel, -1, label='Open Excel File', pos=(370, 360), size=(100, 30))

        # Creating panels
        self.font_header = wx.Font(12, family=wx.DECORATIVE, style=wx.NORMAL, weight=wx.BOLD)
        self.font_normal = wx.Font(10, family=wx.DECORATIVE, style=wx.NORMAL, weight=wx.NORMAL)

        self.path_panel_1 = wx.Panel(self.top_panel, -1, style=wx.SUNKEN_BORDER, size=(680, 50), pos=(10, 10))
        self.path_header_display = wx.StaticText(self.path_panel_1, label="Path to first selected LOG-file: ", pos=(4, 0))
        self.path_header_display.SetFont(self.font_header)

        self.path_panel_2 = wx.Panel(self.top_panel, -1, style=wx.SUNKEN_BORDER, size=(680, 50), pos=(10, 70))
        self.path_header_display = wx.StaticText(self.path_panel_2, label="Path to second selected LOG-file: ", pos=(4, 0))
        self.path_header_display.SetFont(self.font_header)

        self.some_data_panel_1 = wx.Panel(self.top_panel, -1, style=wx.SUNKEN_BORDER, size=(680, 70), pos=(10, 130))
        self.data_panel_1_header_display = wx.StaticText(self.some_data_panel_1,
                                                         label="Some statistics about the first file: ", pos=(4, 0))
        self.data_panel_1_header_display.SetFont(self.font_header)

        self.some_data_panel_2 = wx.Panel(self.top_panel, -1, style=wx.SUNKEN_BORDER, size=(680, 70), pos=(10, 210))
        self.data_panel_2_header_display = wx.StaticText(self.some_data_panel_2,
                                                         label="Some statistics about the second file: ", pos=(4, 0))
        self.data_panel_2_header_display.SetFont(self.font_header)

        # Set events
        self.Bind(wx.EVT_MENU, self.on_open, menu_file_open)
        self.Bind(wx.EVT_MENU, self.on_about, menu_about)
        self.Bind(wx.EVT_MENU, self.on_exit, menu_exit)
        self.exit_button.Bind(wx.EVT_BUTTON, self.on_exit_button)
        self.reset_button.Bind(wx.EVT_BUTTON, self.on_reset)
        self.open_xlsx_button.Bind(wx.EVT_BUTTON, self.on_xlsx_button)

        self.data_1 = []
        self.data_2 = []
        self.folder_pathname = ""
        self.user_file_name = ""

    def panel_layout(self):
        """
        Assign panels to the main panel. This includes the path to both files and some basic data about the files.

        New fonts are created to create some diversity on the screen, making the application more appealing to look at.
        """

        # Create panels



        self.path_display = wx.StaticText(self.path_panel_1, label=str(self.pathname_1), pos=(4, 25))
        self.path_display.SetFont(self.font_normal)

        self.path_display = wx.StaticText(self.path_panel_2, label=str(self.pathname_2), pos=(4, 25))
        self.path_display.SetFont(self.font_normal)

        self.data_panel_1_display = wx.StaticText(self.some_data_panel_1, label="Average power at high slope:    " + str(
            self.power_high_avg) + " W\n"
                                   "Average velocity at high slope:    " + str(self.velocity_high_avg) + " km/h\n",
                                                  pos=(4, 24))
        self.data_panel_1_display.SetFont(self.font_normal)

        self.data_panel_2_display = wx.StaticText(self.some_data_panel_2,
                                                  label="Average power at low (negative) slope:    " + str(
                                                      self.power_low_avg) + " W\n"
                                                                            "Average velocity at low (negative) slope:    " + str(
                                                      self.velocity_low_avg) + " km/h\n", pos=(4, 24))
        self.data_panel_2_display.SetFont(self.font_normal)

        xlsx_path_panel = wx.Panel(self.top_panel, -1, style=wx.SUNKEN_BORDER, size=(680, 50), pos=(10, 290))
        self.xlsx_path_panel_header_display = wx.StaticText(xlsx_path_panel,
                                                            label="Path to " + self.user_file_name + ".xslx: ",
                                                            pos=(4, 0))
        self.xlsx_path_panel_header_display.SetFont(self.font_header)
        self.xlsx_path_panel_display = wx.StaticText(xlsx_path_panel, label=str(self.folder_pathname), pos=(4, 25))

    def on_open(self, e):
        """"
        This function is used to open LOG-files, selected by the user.

        Workflow is as follows, per file:
        1: Call the FileDialog function to open a dialog screen
        2: Enable the user to go back without crashing the program
        3: Retrieve the pathname and the entire file
        4: Prepare for possible exceptions
        5: Close file

        """

        # Opening File 1
        self.directory_name_1 = ""

        with wx.FileDialog(self, "Choose the logged SimulANT+ file with the HIGHEST slope...",
                           wildcard="Text files (*.txt)|*.txt|" "Comma Separated Value-files (*.csv)|*.csv",
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as prompted_dialog:

            if prompted_dialog.ShowModal() == wx.ID_CANCEL:
                return

            self.pathname_1 = prompted_dialog.GetPath()

        # Opening File 2
        self.directory_name_2 = ""

        with wx.FileDialog(self, "Choose the second logged SimulANT+ file with the LOWEST (negative) slope...",
                           wildcard="Text files (*.txt)|*.txt|" "Comma Separated Value-files (*.csv)|*.csv",
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as prompted_dialog:

            if prompted_dialog.ShowModal() == wx.ID_CANCEL:
                return

            self.pathname_2 = prompted_dialog.GetPath()

        self.folder_pathname = os.path.dirname(self.pathname_2)

        self.user_file_name_dialog = wx.TextEntryDialog(self,
                                                        "What do you want the .xslx file to be named? Enter here: ",
                                                        "Enter file name...")
        self.user_file_name_dialog.CenterOnParent()

        if self.user_file_name_dialog.ShowModal() == wx.ID_CANCEL:
            return
        self.user_file_name = self.user_file_name_dialog.GetValue()

        self.logfile_analyser(self.pathname_1)
        velocity_list_high = velocity_list
        power_list_high = power_list
        number_of_measurements_high = min(len(power_list), len(velocity_list))
        self.logfile_analyser(self.pathname_2)
        velocity_list_low = velocity_list
        power_list_low = power_list
        number_of_measurements_low = min(len(power_list), len(velocity_list))

        self.velocity_high_avg = np.mean(velocity_list_high)
        self.velocity_high_avg = round(float(self.velocity_high_avg), 3)
        self.power_high_avg = np.mean(power_list_high)
        self.power_high_avg = round(float(self.power_high_avg), 3)
        self.velocity_low_avg = np.mean(velocity_list_low)
        self.velocity_low_avg = round(float(self.velocity_low_avg), 3)
        self.power_low_avg = np.mean(power_list_low)
        self.power_low_avg = round(float(self.power_low_avg), 3)

        """
        Initialize writing an excel file.
        """
        # name = input("Which trainer has been tested?  ")
        excel = xlsxwriter.Workbook(str(self.user_file_name) + ".xlsx")
        graph = excel.add_chart({'type': 'scatter'})
        worksheet = excel.add_worksheet()

        """
        Setting variables for excel file.
        """
        bold = excel.add_format({'bold': True})
        underline = excel.add_format({'bold': True, 'underline': True})
        graph.set_y_axis({'name': 'Power [W]'})
        graph.set_x_axis({'name': 'Velocity [km/h]'})
        graph.set_title({'name': 'Operating range ' + str(self.user_file_name)})
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
            'values': [worksheet.name] + [2, 4] + [number_of_measurements_low + 2, 4],
            'line': {'color': 'black'},
            'name': 'lowest gradient',
        })

        graph.add_series({
            'categories': [worksheet.name] + [2, 0] + [number_of_measurements_high + 2, 0],
            'values': [worksheet.name] + [2, 1] + [number_of_measurements_high + 2, 1],
            'line': {'color': 'black'},
            'name': 'Highest gradient',
        })

        worksheet.insert_chart('H4', graph)
        excel.close()

        self.panel_layout()

    def on_about(self, e):
        """"Message box with OK button"""
        prompted_dialog = wx.MessageDialog(self, "A file which converts SimulANT+ data into an Excel file, which \n"
                                                 "can be analyzed and compared to other log-files. \n"
                                                 "\n"
                                                 "Built in Python 3.6.6, compiled with PyInstaller"
                                                 "\n"
                                                 "\n"
                                                 "Created by Tim de Jong and Jelle Haasnoot at Tacx B.V.",
                                           "About SimulANT+ Log Analyzer", wx.OK)
        prompted_dialog.ShowModal()
        prompted_dialog.Destroy()

    def on_exit(self, e):
        self.Close(True)

    def logfile_analyser(self, logfile):
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
        for i in range(len(sentences)):
            sentence = sentences[i].split()
            index = sentence.index("Rx:")
            value_raw = sentence[index + 1]
            value = value_raw.replace("[", "").replace("]", "")  # This will removes the useless characters
            value_list.append(value)

        """
        Deze waardes worden gesorteerd per indexatie in het begin van het hexadecimale getal, in de index '10' staat de snelheid
        , in index '19' de power en de rest is voor dit bestand niet belangrijk en zal daarom niet geanalyseerd worden.
        """
        for i in range(len(value_list)):
            value_list_characters = list(value_list[i])
            if value_list_characters[0] == '1' and value_list_characters[
                1] == '0' and speed:  # index waar snelheid in staat
                speed = False
                power = True
                speed_values_raw = [value_list_characters[10], value_list_characters[11], value_list_characters[8],
                                    value_list_characters[9]]
                speed_values_raw_string = "".join(speed_values_raw)
                value_converter = ValueConverter()
                velocity_bin = value_converter.hex_to_bin(speed_values_raw_string)
                velocity = value_converter.bin_to_dec(velocity_bin) * 3.6 / 1000
                velocity_list.append(velocity)
            elif value_list_characters[0] == '1' and value_list_characters[
                1] == '9' and power:  # index waar power in staat
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

    def on_exit_button(self, event):
        self.Close()

    def on_reset(self, event):
        if __name__ == '__main__':
            self.Close()
            frame = Main(None, 'SimulANT+ Log Analyzer').Show()

    def on_xlsx_button(self,event):
        if os.path.isfile(self.folder_pathname + "\\" + self.user_file_name + ".xlsx"):
            os.startfile(self.folder_pathname + "\\" + self.user_file_name + ".xlsx")
        elif self.folder_pathname == "":
            no_file_dialog = wx.MessageDialog(self.top_panel, message="The file does not exist. Please try again.", caption="Warning!")
            no_file_dialog.CenterOnParent()
            if no_file_dialog.ShowModal() == wx.OK:
                no_file_dialog.Destroy()
                return

if __name__ == '__main__':
    Application = wx.App(False)
    frame = Main(None, 'SimulANT+ Log Analyzer').Show()
    Application.MainLoop()
