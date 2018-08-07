import wx as wx
from ValueConverter import ValueConverter
import xlsxwriter
import sys
from os import path
from os import startfile
from numpy import mean
from numpy import array
from numpy import arctan
from numpy import sin
from numpy import cos
from scipy.optimize import curve_fit
#
class Main(wx.Frame):
    def __init__(self, parent, title):
        """
        Initializing the program:

        1: First the file menu will be configured. This is the top bar which holds the options file, open, about, etc...
        2: Then the menu bar is created
        3: Buttons with their names and positions are created, as well as checkboxes
        4: An image for in the program is loaded and positioned
        5: Panels which will hold text are created and positioned. Static text is written, other text will come later
        6: Events which connect functionality with button and checkbox-clicks are created
        7: The startup message is created
        8: The status bar is created
        9: Empty parameters are defined.
        """

        wx.Frame.__init__(self, parent, title=title,
                          style=wx.DEFAULT_FRAME_STYLE & ~(wx.RESIZE_BORDER | wx.MAXIMIZE_BOX), size=(720, 788))

        self.top_panel = wx.Panel(self)
        self.SetBackgroundColour("white")

        # 1: Create the file menu
        file_menu = wx.Menu()

        menu_about = file_menu.Append(wx.ID_ABOUT, "&About", "Information about this program")
        file_menu.AppendSeparator()
        menu_file_open = file_menu.Append(wx.ID_FILE, "&Open files...", "Open a text file with this program")
        menu_exit = file_menu.Append(wx.ID_EXIT, "E&xit", "Terminate the program")

        # 2: Create the menu bar
        menu_bar = wx.MenuBar()
        menu_bar.Append(file_menu, "&File")
        self.SetMenuBar(menu_bar)

        # 3: Creating buttons
        self.exit_button = wx.Button(self.top_panel, -1, label='Exit', pos=(590, 664), size=(100, 30))
        self.reset_button = wx.Button(self.top_panel, -1, label='Reset Program', pos=(480, 664), size=(100, 30))
        self.open_xlsx_button = wx.Button(self.top_panel, -1, label='Open Excel File', pos=(370, 664), size=(100, 30))
        self.open_files_butten = wx.Button(self.top_panel, -1, label='Open LOG\'s', pos=(260, 664), size=(100, 30))


        # 4: Loading images
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = path.abspath('.')
        image_path = path.join(base_path, 'tacx-logo.png')

        image_file_png = wx.Image(image_path, wx.BITMAP_TYPE_PNG)
        image_file_png.Rescale(image_file_png.GetWidth() * 0.28, image_file_png.GetHeight() * 0.28)
        image_file_png = wx.Bitmap(image_file_png)
        self.image = wx.StaticBitmap(self.top_panel, -1, image_file_png, pos=(9, 604),
                                     size=(image_file_png.GetWidth(), image_file_png.GetHeight()))

        # 5: Creating panels
        self.font_header = wx.Font(12, family=wx.FONTFAMILY_DECORATIVE, style=wx.FONTSTYLE_NORMAL, weight=wx.FONTWEIGHT_BOLD)
        self.font_normal = wx.Font(10, family=wx.FONTFAMILY_DECORATIVE, style=wx.FONTSTYLE_NORMAL, weight=wx.FONTWEIGHT_NORMAL)
        self.font_green = wx.Font(12, family=wx.FONTFAMILY_DEFAULT, style=wx.FONTSTYLE_NORMAL, weight=wx.FONTWEIGHT_BOLD)

        self.panel_titles = ["Path to directory first selected LOG-file: ", "Path to directory second selected LOG-file: ", "Path to directory third selected LOG-file: ", "Path to directory fourth selected LOG-file: "]
        self.statistics_titles = ["Some statistics about the first file: ", "Some statistics about the second file: ", "Some statistics about the third file: ", "Some statistics about the fourth file: "]
        for i in range(len(self.statistics_titles)):
            # self.path_panel = wx.Panel(self.top_panel, -1, style=wx.TAB_TRAVERSAL | wx.SUNKEN_BORDER, size=(685, 50), pos=(10, 10 + i * 55))
            # self.path_panel_header = wx.StaticText(self.path_panel, label=self.panel_titles[i], pos=(4, 2))
            # self.path_panel_header.SetFont(self.font_header)

            self.data_panel = wx.Panel(self.top_panel, -1, style=wx.SUNKEN_BORDER, size=(685, 80), pos=(10, 10 + i * 90))
            self.data_panel_header = wx.StaticText(self.data_panel, label=self.statistics_titles[i], pos=(4, 2))
            self.data_panel_header.SetFont(self.font_header)


        self.xlsx_path_panel = wx.Panel(self.top_panel, -1, style=wx.SUNKEN_BORDER, size=(685, 50), pos=(10, 537))
        self.user_input_panel = wx.Panel(self.top_panel, -1, style=wx.BORDER_RAISED, size=(685, 156), pos=(10, 370))

        self.sim_mass_panel = wx.Panel(self.top_panel, -1, style=wx.SUNKEN_BORDER, size=(433, 50), pos=(262, 601))
        self.sim_mass_panel_header_display = wx.StaticText(self.sim_mass_panel, label="Simulated Mass (calculated / user-given): ", pos=(4, 2))
        self.sim_mass_panel_header_display.SetFont(self.font_header)



        # 7: Set start-up message
        welcome_dialog = wx.MessageDialog(self.top_panel,
                                          message="Welcome to SimulANT+ Log Analyzer. \nIf you have read the README.pdf, you're good to go. \nIf you haven't yet, please do.",
                                          caption="Welcome!")
        welcome_dialog.CenterOnParent()
        if welcome_dialog.ShowModal() == wx.OK:
            welcome_dialog.Destroy()
            return


        # 8: Create status bar
        self.statusbar = self.CreateStatusBar()


        # 9: Create TextCtrl boxes
        self.gear_front_ask = wx.StaticText(self.user_input_panel, label="No. of teeth front sprocket: ", pos=(20, 10))
        self.gear_front_ask.SetFont(self.font_normal)
        self.edit_gear_front_text = wx.TextCtrl(self.user_input_panel, size=(80, -1), pos=(180, 8.5))
        self.gear_front_sizer = wx.BoxSizer()
        self.gear_front_sizer.Add(self.user_input_panel, 1, wx.ALL | wx.EXPAND)
        self.sizer = wx.GridBagSizer(5, 5)
        self.sizer.Add(self.gear_front_ask, (0, 0))
        self.border = wx.BoxSizer()
        self.border.Add(self.sizer, 1, wx.ALL | wx.EXPAND, 5)

        self.gear_rear_ask = wx.StaticText(self.user_input_panel, label="No. of teeth rear sprocket: ", pos=(20, 45))
        self.gear_rear_ask.SetFont(self.font_normal)
        self.edit_gear_rear_text = wx.TextCtrl(self.user_input_panel, size=(80, -1), pos=(180, 43.5))
        self.gear_rear_sizer = wx.BoxSizer()
        self.gear_rear_sizer.Add(self.user_input_panel, 1, wx.ALL | wx.EXPAND)
        self.sizer = wx.GridBagSizer(5, 5)
        self.sizer.Add(self.gear_rear_ask, (0, 0))
        self.border = wx.BoxSizer()
        self.border.Add(self.sizer, 1, wx.ALL | wx.EXPAND, 5)

        self.trainer_deviation_ask = wx.StaticText(self.user_input_panel, label="Trainer accuracy [%]: ", pos=(20, 80))
        self.trainer_deviation_units = wx.StaticText(self.user_input_panel, label="%", pos=(270, 80))
        self.trainer_deviation_ask.SetFont(self.font_normal)
        self.trainer_deviation_units.SetFont(self.font_normal)
        self.trainer_deviation_text = wx.TextCtrl(self.user_input_panel, size=(80, -1), pos=(180, 78.5))
        self.trainer_deviation_sizer = wx.BoxSizer()
        self.trainer_deviation_sizer.Add(self.user_input_panel, 1, wx.ALL | wx.EXPAND)
        self.sizer = wx.GridBagSizer(5, 5)
        self.sizer.Add(self.trainer_deviation_ask, (0, 0))
        self.border = wx.BoxSizer()
        self.border.Add(self.sizer, 1, wx.ALL | wx.EXPAND, 5)

        self.sensor_deviation_ask = wx.StaticText(self.user_input_panel, label="Sensor accuracy [%]: ", pos=(20, 115))
        self.sensor_deviation_units = wx.StaticText(self.user_input_panel, label="%", pos=(270, 115))
        self.sensor_deviation_ask.SetFont(self.font_normal)
        self.sensor_deviation_units.SetFont(self.font_normal)
        self.sensor_deviation_text = wx.TextCtrl(self.user_input_panel, size=(80, -1), pos=(180, 113.5))
        self.sensor_deviation_sizer = wx.BoxSizer()
        self.sensor_deviation_sizer.Add(self.user_input_panel, 1, wx.ALL | wx.EXPAND)
        self.sizer = wx.GridBagSizer(5, 5)
        self.sizer.Add(self.sensor_deviation_ask, (0, 0))
        self.border = wx.BoxSizer()
        self.border.Add(self.sizer, 1, wx.ALL | wx.EXPAND, 5)

        self.simulated_mass_ask = wx.StaticText(self.user_input_panel, label="User input simulated mass: ", pos=(322, 10))
        self.simulated_mass_units = wx.StaticText(self.user_input_panel, label="kg", pos=(605, 10))
        self.simulated_mass_units.SetFont(self.font_normal)
        self.simulated_mass_ask.SetFont(self.font_normal)
        self.edit_simulated_mass_text = wx.TextCtrl(self.user_input_panel, size=(80, -1), pos=(510, 8.5))
        self.simulated_mass_sizer = wx.BoxSizer()
        self.simulated_mass_sizer.Add(self.user_input_panel, 1, wx.ALL | wx.EXPAND)
        self.sizer = wx.GridBagSizer(5, 5)
        self.sizer.Add(self.simulated_mass_ask, (0, 0))
        self.border = wx.BoxSizer()
        self.border.Add(self.sizer, 1, wx.ALL | wx.EXPAND, 5)

        self.simulated_mass_ask_alt_1 = wx.StaticText(self.user_input_panel, label="User input moment of inertia: ", pos=(322, 45))
        self.sim_mass_units_alt_1 = wx.StaticText(self.user_input_panel, label="kg m²",pos=(605, 45))
        self.sim_mass_units_alt_1.SetFont(self.font_normal)
        self.simulated_mass_ask_alt_1.SetFont(self.font_normal)
        self.simulated_mass_alt_1_text = wx.TextCtrl(self.user_input_panel, size=(80, -1), pos=(510, 43.5))
        self.simulated_mass_alt_1_sizer = wx.BoxSizer()
        self.simulated_mass_alt_1_sizer.Add(self.user_input_panel, 1, wx.ALL | wx.EXPAND)
        self.sizer = wx.GridBagSizer(5, 5)
        self.sizer.Add(self.simulated_mass_ask_alt_1, (0, 0))
        self.border = wx.BoxSizer()
        self.border.Add(self.sizer, 1, wx.ALL | wx.EXPAND, 5)

        self.simulated_mass_ask_alt_2 = wx.StaticText(self.user_input_panel, label="User input rot. / lin. ratio: ", pos=(322, 80))
        self.sim_mass_units_alt_2 = wx.StaticText(self.user_input_panel, label="rad / m", pos=(605, 80))
        self.sim_mass_units_alt_2.SetFont(self.font_normal)
        self.simulated_mass_ask_alt_2.SetFont(self.font_normal)
        self.simulated_mass_alt_2_text = wx.TextCtrl(self.user_input_panel, size=(80, -1), pos=(510, 78.5))
        self.simulated_mass_alt_2_sizer = wx.BoxSizer()
        self.simulated_mass_alt_2_sizer.Add(self.user_input_panel, 1, wx.ALL | wx.EXPAND)
        self.sizer = wx.GridBagSizer(5, 5)
        self.sizer.Add(self.simulated_mass_ask_alt_1, (0, 0))
        self.border = wx.BoxSizer()
        self.border.Add(self.sizer, 1, wx.ALL | wx.EXPAND, 5)


        self.save_inputs_button = wx.Button(self.user_input_panel, -1, label="Save Inputs", pos=(500, 112), size=(100, 30))
        self.saved_text = wx.StaticText(self.user_input_panel, -1, label="", pos=(320, 119), style=wx.ALIGN_CENTER_HORIZONTAL)
        self.saved_text.SetFont(self.font_green)
        self.saved_text.SetLabel("VALUES NOT SAVED")
        self.saved_text.SetForegroundColour((255, 10, 10))
        self.saved = False

        # 9: Create empty parameters
        self.data_1 = []
        self.data_2 = []
        self.data_3 = []
        self.directory_name_1 = ""
        self.directory_name_2 = ""
        self.directory_name_3 = ""
        self.directory_name_4 = ""

        self.folder_pathname = ""
        self.user_file_name = ""
        self.velocity_list_high = []
        self.velocity_list_low = []
        self.velocity_list_const = []
        self.power_list_low = []
        self.power_list_high = []
        self.power_list_const = []
        self.check_counter = 0
        self.wheel_radius = 0.3395 # TODO CHECKEN OF DIT WERKT BINNEN SIMULANT, ANDERS DIT EN README AANPASSEN

        self.checkbox = wx.CheckBox(self.user_input_panel, -1, '', pos=(300, 10.7))
        self.checkbox.SetValue(False)
        self.edit_simulated_mass_text.SetEditable(False)
        self.simulated_mass_alt_1_text.SetEditable(False)
        self.simulated_mass_alt_2_text.SetEditable(False)

        self.edit_simulated_mass_text.SetBackgroundColour((220, 220, 220))
        self.simulated_mass_alt_1_text.SetBackgroundColour((220, 220, 220))
        self.simulated_mass_alt_2_text.SetBackgroundColour((220, 220, 220))

        # 6: Set events
        self.Bind(wx.EVT_MENU, self.on_open, menu_file_open)
        self.Bind(wx.EVT_MENU, self.on_about, menu_about)
        self.Bind(wx.EVT_MENU, self.on_exit, menu_exit)
        self.exit_button.Bind(wx.EVT_BUTTON, self.on_exit_button)
        self.exit_button.Bind(wx.EVT_ENTER_WINDOW, self.on_exit_widget_enter)
        self.reset_button.Bind(wx.EVT_BUTTON, self.on_reset)
        self.reset_button.Bind(wx.EVT_ENTER_WINDOW, self.on_reset_widget_enter)
        self.open_xlsx_button.Bind(wx.EVT_BUTTON, self.on_xlsx_button)
        self.open_xlsx_button.Bind(wx.EVT_ENTER_WINDOW, self.on_excel_widget_enter)
        self.open_files_butten.Bind(wx.EVT_BUTTON, self.on_open)
        self.open_files_butten.Bind(wx.EVT_ENTER_WINDOW, self.on_open_widget_enter)
        self.checkbox.Bind(wx.EVT_ENTER_WINDOW, self.on_check_hover)
        self.Bind(wx.EVT_CHECKBOX, self.on_check)
        self.save_inputs_button.Bind(wx.EVT_BUTTON, self.on_save_inputs)
        self.save_inputs_button.Bind(wx.EVT_ENTER_WINDOW, self.on_save_hover)

    def panel_layout(self):
        """
        Assign panels to the main panel. This includes the path to both files and some basic data about the files.
        New fonts are created to create some diversity on the screen, making the application more appealing to look at.
        """
        # for i in range(len(self.statistics_titles)):
        #     self.path_display_panel = wx.Panel(self.top_panel, -1, style=wx.NO_BORDER, size=(660, 22), pos=(14, 35 + i * 55))
        #     self.path_display = wx.StaticText(self.path_display_panel, label=str(path.dirname(self.pathname[i])), pos=(4, 2))

        data_display_strings = [["Average power at high power:                ", "Average velocity at high power:              ", "Amount of received ANT+ messages:     "],
                                ["Average power at low power:                ", "Average velocity at low power:              ", "Amount of received ANT+ messages:   "],
                                ["Average power at constant velocity:     ", "Average velocity at constant velocity:   ", "Amount of received ANT+ messages:   "],
                                ["Average power from power / cadence sensor:     ", "Average velocity from power / cadence sensor:   ", "Amount of received ANT+ messages:                "]]
        for i in range(len(self.statistics_titles)):
            self.data_display_panel = wx.Panel(self.top_panel, -1, style=wx.NO_BORDER, size=(660, 50), pos=(18, 35 + i * 90))
            self.data_display = wx.StaticText(self.data_display_panel, label=data_display_strings[i][0] + str(self.all_averages[i][0]) + "W\n" + data_display_strings[i][1] + str(self.all_averages[i][1]) + "km/h\n" + data_display_strings[i][2] + str(self.all_averages[i][2]))
            self.data_display.SetFont(self.font_normal)

        # Create panels
        self.xlsx_path_panel_header_display = wx.StaticText(self.xlsx_path_panel,
                                                            label="Path to " + self.user_file_name + ".xslx: ",
                                                            pos=(4, 0))
        self.xlsx_path_panel_header_display.SetFont(self.font_header)
        self.xlsx_path_panel_display = wx.StaticText(self.xlsx_path_panel, label=str(self.folder_pathname), pos=(4, 25))

        self.sim_mass_panel_display = wx.StaticText(self.sim_mass_panel, label=str(round(float(self.simulated_mass_guess), 2)) + " [kg m²]", pos=(4, 25))
        self.sim_mass_panel_display.SetFont(self.font_normal)

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

        # Opening all files with the use of a dialog. Add a question to dummy_strings to make this code open more files.
        # File 1 will contain the ANT+ data of the measurements with a high resistance (wattage). This will be used to
        # calculate the maximal brake power. File 2 will contain the ANT+ data of the measurements with 0 W as goal
        # power. This will be used to calculate the minimal brake power. Opening File 3 with the
        # use of a dialog. File 3 will contain the ANT+ data of the measurements with a power goal of 0W while cycling
        # at some multiple constant velocities. This will be used to see the residual brake power if no brake is used.
        # Opening File 4 with the use of a dialog. File 4 will contain the ANT+ data of the measurements with a power
        # goal of 0W while cycling at some multiple constant velocities. This will be measured with an external power
        # meter, this way the accuracy can be calculated.

        if self.sensor_deviation_text.GetValue() == "" or self.trainer_deviation_text.GetValue() == "" or self.edit_gear_front_text.GetValue() == "" or self.edit_gear_rear_text.GetValue() == "" or self.saved == False:
            no_entry_dialog = wx.MessageDialog(self.top_panel, style=wx.ICON_ERROR, message="(some of) The fields are still left empty, or haven't been saved. \n\nPlease fill out the gear ratio, deviations and optionally inertia data.")
            no_entry_dialog.CenterOnParent()
            if no_entry_dialog.ShowModal() == wx.OK:
                no_entry_dialog.Destroy()
                return
            return

        if self.edit_simulated_mass_text.GetValue() == "" and self.simulated_mass_alt_1_text.GetValue() == "" and self.saved == False:
            no_entry_dialog = wx.MessageDialog(self.top_panel, style=wx.ICON_ERROR, message="(some of) The fields are still left empty, or haven't been saved. \n\nPlease fill out the gear ratio, deviations and optionally inertia data.")
            no_entry_dialog.CenterOnParent()
            if no_entry_dialog.ShowModal() == wx.OK:
                no_entry_dialog.Destroy()
                return
            return

        self.pathname = []
        dummy_strings = ["Choose the logged SimulANT+ file with high power...",
                         "Choose the second logged SimulANT+ file with low power...",
                         "Choose the third logged SimulANT+ file with the constant velocity...",
                         "Choose the fourth logged SimulANT+ file with the constant velocity - read from the external power sensor..."]
        for i in range(len(dummy_strings)):
            with wx.FileDialog(self, dummy_strings[i],
                               wildcard="Text files (*.txt)|*.txt|" "Comma Separated Value-files (*.csv)|*.csv",
                               style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as prompted_dialog:
                if prompted_dialog.ShowModal() == wx.ID_CANCEL:
                    return
                self.pathname.append(prompted_dialog.GetPath())


        # Naming the excel file which will be made by the program.
        self.user_file_name_dialog = wx.TextEntryDialog(self,
                                                        "What do you want the .xslx file to be named? Enter here: ",
                                                        "Enter file name...")
        self.folder_pathname = path.dirname(self.pathname[2])
        self.user_file_name_dialog.CenterOnParent()
        if self.user_file_name_dialog.ShowModal() == wx.ID_CANCEL:
            return
        self.user_file_name = self.user_file_name_dialog.GetValue()

        # Analyse the log-files. This will be used to retrieve the data from the four selected log files above. This
        # will be done by ANTlogfileconverter.py. Raw data will be stored in data_.... and used in further calculations.
        # Analysing log file 1:
        self.velocity_list = []
        self.power_list = []
        self.velocity_time_list = []
        self.power_time_list = []
        data = []

        for j in range(3):
            data_dummy = []
            self.logfile_analyser_trainer(self.pathname[j])
            if len(power_list) < len(velocity_list):
                velocity_list.pop()
                velocity_time_list.pop()
            elif len(velocity_list) < len(power_list):
                power_list.pop()
                power_time_list.pop()
            else:
                pass

            for i in range(len(velocity_list)):
                data_dummy.append([velocity_list[i], power_list[i], velocity_time_list[i], power_time_list[i]])

            data.append(data_dummy)
            self.velocity_list.append(velocity_list)
            self.power_list.append(power_list)
            self.velocity_time_list.append(velocity_time_list)
            self.power_time_list.append(power_time_list)

        data_high = data[0]
        data_low = data[1]
        data_const = data[2]

        self.velocity_list_high = self.velocity_list[0]
        self.power_list_high = self.power_list[0]
        self.velocity_list_low = self.velocity_list[1]
        self.power_list_low = self.power_list[1]
        self.velocity_list_const = self.velocity_list[2]
        self.power_list_const = self.power_list[2]
        self.power_time_list_const = self.power_time_list[2]
        self.velocity_time_list_const = self.velocity_time_list[2]

        # analysing log file 4:
        self.logfile_analyser_sensor(self.pathname[3])
        data_sensor = []

        if len(power_list) < len(cadence_list):
            cadence_list.pop()
            sensor_time_list.pop()
        elif len(cadence_list) < len(power_list):
            power_list.pop()
            power_time_list.pop()
        else:
            pass

        for i in range(len(cadence_list)):
            data_sensor.append([cadence_list[i], power_list[i], sensor_time_list[i]])

        self.power_list_sensor = power_list
        self.cadence_list_sensor = cadence_list
        self.power_time_list_sensor = sensor_time_list

        # Correcting time-stamps between log file 3 and log file 4
        if int(self.time_stamp_trainer[0]) > int(self.time_stamp_sensor[0]):
            correcting_time = abs(int(self.time_stamp_sensor[0]) - int(self.time_stamp_trainer[0]))
            for i in range(len(self.power_time_list_const)):
                self.power_time_list_const[i] = float(self.power_time_list_const[i]) + correcting_time/1000
            for i in range(len(self.velocity_time_list_const)):
                self.velocity_time_list_const[i] = float(self.velocity_time_list_const[i]) + correcting_time/1000
        elif int(self.time_stamp_trainer[0]) < int(self.time_stamp_sensor[0]):
            correcting_time = abs(int(self.time_stamp_sensor[0]) - int(self.time_stamp_trainer[0]))
            for i in range(len(self.power_time_list_sensor)):
                self.power_time_list_sensor[i] = float(self.power_time_list_sensor[i]) + correcting_time/1000

        power_time_const_check = self.power_time_list_const
        power_sensor_check = self.power_list_sensor
        power_sensor_check_upper_bound = []
        power_sensor_check_lower_bound = []
        for i in range(len(power_sensor_check)):
            power_sensor_check_upper_bound.append(power_sensor_check[i] * (1 + sensor_deviation_perc/100))
            power_sensor_check_lower_bound.append(power_sensor_check[i] * (1 - sensor_deviation_perc/100))

        power_time_sensor_check = self.power_time_list_sensor
        power_const_check = self.power_list_const
        power_const_check_upper_bound = []
        power_const_check_lower_bound = []
        for i in range(len(power_const_check)):
            power_const_check_upper_bound.append(power_const_check[i] * (1 + trainer_deviation_perc/100))
            power_const_check_lower_bound.append(power_const_check[i] * (1 - trainer_deviation_perc/100))

        error = []
        if len(power_time_const_check) > len(power_time_sensor_check):
            for i in range(len(power_sensor_check)):
                closest_value = min(enumerate(power_time_const_check), key=lambda x: abs(x[1] - power_time_sensor_check[i]))
                index_closest_value = closest_value[0]
                if power_sensor_check[i] == 0 or power_const_check[i] == 0:
                    pass
                elif power_time_const_check[index_closest_value] > power_time_sensor_check[i]:
                    dummy = (power_const_check[index_closest_value] - power_const_check[index_closest_value - 1]) / \
                            (power_time_const_check[index_closest_value]- power_time_const_check[index_closest_value
                                                                                                    - 1])
                    dummy2 = power_time_const_check[index_closest_value] - power_time_sensor_check[i]
                    power_const_at_index = power_const_check[index_closest_value] - dummy * dummy2
                    # print(power_time_const_check[index_closest_value], power_const_check[index_closest_value])
                    # print(power_time_sensor_check[i], power_sensor_check[i])
                    # print(power_time_const_check[index_closest_value-1], power_const_check[index_closest_value-1])
                    # print(power_const_at_index)
                    # print('--------')
                    error.append(abs(power_sensor_check[i] - power_const_at_index) / power_sensor_check[i] * 100)
                elif power_time_const_check[index_closest_value] < power_time_sensor_check[i] and index_closest_value < len(power_const_check)-1:
                    print(index_closest_value, len(power_const_check))
                    dummy = (power_const_check[index_closest_value] - power_const_check[index_closest_value + 1])/ \
                            (power_time_const_check[index_closest_value] - power_time_const_check[index_closest_value
                                                                                                    + 1])
                    dummy2 = power_time_const_check[index_closest_value] - power_time_sensor_check[i]
                    power_const_at_index = power_const_check[index_closest_value] - dummy * dummy2

                    error.append(abs(power_sensor_check[i] - power_const_at_index) / power_sensor_check[i] * 100)
                    # print(abs(power_sensor_check[i] - power_const_at_index) / power_sensor_check[i] * 100)
                    # print(power_time_const_check[index_closest_value], power_const_check[index_closest_value])
                    # print(power_time_sensor_check[i], power_sensor_check[i])
                    # print(power_time_const_check[index_closest_value+1], power_const_check[index_closest_value+1])
                    # print(power_const_at_index)
                    # print('--------')

        elif len(power_time_const_check) < len(power_time_sensor_check):
             for i in range(len(power_time_const_check)-1):
                closest_value = min(enumerate(power_time_sensor_check), key=lambda x: abs(x[1] - power_time_const_check[i]))
                index_closest_value = closest_value[0]
                if power_sensor_check[i] == 0 or power_const_check[i] == 0:
                    pass
                elif power_time_const_check[i] < power_time_sensor_check[index_closest_value]:
                    dummy = (power_sensor_check[index_closest_value] - power_sensor_check[index_closest_value - 1])/ \
                            (power_time_sensor_check[index_closest_value] - power_time_sensor_check[index_closest_value
                                                                                                    - 1])
                    dummy2 = power_time_sensor_check[index_closest_value] - power_time_const_check[i]
                    power_sensor_at_index = power_sensor_check[index_closest_value] - dummy * dummy2
                    # print(power_time_const_check[index_closest_value], power_const_check[index_closest_value])
                    # print(power_time_sensor_check[i], power_sensor_check[i])
                    # print(power_time_const_check[index_closest_value-1], power_const_check[index_closest_value-1])
                    # print(power_const_at_index)
                    # print('--------')
                    error.append(abs(power_const_check[i] - power_sensor_at_index) / power_sensor_at_index * 100)
                elif power_time_const_check[i] > power_time_sensor_check[index_closest_value] and index_closest_value < len(power_sensor_check)-1:
                    dummy = (power_sensor_check[index_closest_value] - power_sensor_check[index_closest_value + 1])/ \
                            (power_time_sensor_check[index_closest_value] - power_time_sensor_check[index_closest_value
                                                                                                    + 1])
                    dummy2 = power_time_sensor_check[index_closest_value] - power_time_const_check[i]
                    power_sensor_at_index = power_sensor_check[index_closest_value] - dummy * dummy2
                    # print(power_time_const_check[index_closest_value], power_const_check[index_closest_value])
                    # print(power_time_sensor_check[i], power_sensor_check[i])
                    # print(power_time_const_check[index_closest_value-1], power_const_check[index_closest_value-1])
                    # print(power_const_at_index)
                    # print('--------')
                    error.append(abs(power_const_check[i] - power_sensor_at_index) / power_sensor_at_index * 100)

        mean_error_power = mean(error)
        # print(mean_error_power)

        speed_sensor_check = []
        speed_sensor_check_upper_bound = []
        speed_sensor_check_lower_bound = []
        for i in range(len(self.cadence_list_sensor)):
            speed_sensor_check.append(self.cadence_list_sensor[i] * self.sprocket_ratio * self.wheel_radius)
            speed_sensor_check_upper_bound.append(speed_sensor_check[i] * (1 + sensor_deviation_perc/100))
            speed_sensor_check_lower_bound.append(speed_sensor_check[i] * (1 - sensor_deviation_perc/100))

        speed_trainer_check = []
        speed_trainer_check_time = []
        speed_trainer_check_upper_bound = []
        speed_trainer_check_lower_bound = []
        for i in range(len(self.velocity_time_list_const)):
            speed_trainer_check.append(self.velocity_list_const[i])
            speed_trainer_check_time.append(self.velocity_time_list_const[i])
            speed_trainer_check_upper_bound.append(speed_trainer_check[i] * (1 + trainer_deviation_perc/100))
            speed_trainer_check_lower_bound.append(speed_trainer_check[i] * (1 - trainer_deviation_perc/100))


        # Calculating the averages of every file, this is not necessary for the calculations below, but this will give
        # a quick overview of the used files to the user.
        self.all_averages = [[], [], [], []]

        self.power_high_avg = mean(self.power_list_high)
        self.all_averages[0].append(round(float(self.power_high_avg), 1))
        self.velocity_high_avg = mean(self.velocity_list_high)
        self.all_averages[0].append(round(float(self.velocity_high_avg), 1))
        self.all_averages[0].append(len(self.velocity_list_high))

        self.power_low_avg = mean(self.power_list_low)
        self.all_averages[1].append(round(float(self.power_low_avg), 1))
        self.velocity_low_avg = mean(self.velocity_list_low)
        self.all_averages[1].append(round(float(self.velocity_low_avg), 1))
        self.all_averages[1].append(len(self.velocity_list_low))

        self.power_const_avg = mean(self.power_list_const)
        self.all_averages[2].append(round(float(self.power_const_avg), 1))
        self.velocity_const_avg = mean(self.velocity_list_const)
        self.all_averages[2].append(round(float(self.velocity_const_avg), 1))
        self.all_averages[2].append(len(self.velocity_list_const))
        self.all_averages = array(self.all_averages)

        self.power_sensor_avg = mean(self.power_list_sensor)
        self.all_averages[3].append(round(float(self.power_sensor_avg), 1))
        self.cadence_sensor_avg = mean(self.cadence_list_sensor)
        self.all_averages[3].append(round(float(self.cadence_sensor_avg * self.wheel_radius * self.sprocket_ratio), 1))
        self.all_averages[3].append(len(self.cadence_list_sensor))
        self.all_averages = array(self.all_averages)

        # TODO No mean error up to this point

        # Some constants needed for the next part of code
        global index_low_below_zero_1, index_low_below_zero_2, index_low_below_zero, fitted_power_high, fitted_power_low, poplin, graph
        velocity_raw_high = []
        power_raw_high = []
        velocity_time_raw_high = []
        power_time_raw_high = []
        velocity_raw_low = []
        power_raw_low = []
        velocity_time_raw_low = []
        power_time_raw_low =[]
        power_clean_high = []
        velocity_clean_high = []
        velocity_time_clean_high = []
        power_time_clean_high = []
        power_clean_low = []
        velocity_clean_low = []
        velocity_time_clean_low = []
        power_time_clean_low = []
        velocity_time_clean_const = []
        power_time_clean_const = []
        power_clean_const = []
        velocity_clean_const = []
        power_clean_sensor = []
        sensor_time_raw = []
        first_non_zero_power = []
        power_no_int_res_high_imd = []
        power_no_int_res_high = []
        velocity_time_raw_const = []
        power_const_1 = []
        velocity_const_1 = []
        power_time_raw_const = []
        power_flywheel = []
        power_flywheel_high = []
        power_flywheel_high_imd = []
        power_flywheel_low = []
        power_clean_low_brake = []
        power_clean_high_brake = []
        power_trainer = []
        power_const = [0]
        index_low_below_zero_1 = 0
        error_lin_low = 0
        error_quadratic_low = 0
        index_low_below_zero_2 = 0
        error_lin_high = 0
        error_quadratic_high = 0
        first_limit = 12*3.6
        range_half = 0.5

        # Convert the raw data from the file to named lists for the FIRST file. This file gives the information for the
        # maximal braking power.
        for j in range(len(data_high)):
            if round(data_high[j][0]) == 0:
                power_raw_high.append(0)
                velocity_raw_high.append(0)
                velocity_time_raw_high.append(data_high[j][2])
                power_time_raw_high.append(data_high[j][3])
            else:
                power_raw_high.append(data_high[j][1])
                velocity_raw_high.append(data_high[j][0])
                velocity_time_raw_high.append(data_high[j][2])
                power_time_raw_high.append(data_high[j][3])

        for i in range(power_raw_high.index(max(power_raw_high))):
            power_clean_high.append(power_raw_high[i])
            velocity_clean_high.append(velocity_raw_high[i])
            velocity_time_clean_high.append(velocity_time_raw_high[i])
            power_time_clean_high.append(power_time_raw_high[i])

        # Calculate errors and make a fit for the data of the FIRST file. This fit will be the main information in the
        # P-v plot.
        popt1_high, pcov = curve_fit(self.func_lin, array(velocity_clean_high), array(power_clean_high))
        fitted_power_high_1 = self.func_lin(array(velocity_clean_high), *popt1_high)
        for i in range(len(fitted_power_high_1)):
            if fitted_power_high_1[i] < 0:
                fitted_power_high_1[i] = 0
        for i in range(len(fitted_power_high_1)):
            if fitted_power_high_1[i] > 0:
                error_lin_high += abs(fitted_power_high_1[i] - power_clean_high[i])

        popt2_high, pcov = curve_fit(self.func_quadratic, array(velocity_clean_high), array(power_clean_high))
        fitted_power_high_2 = self.func_quadratic(array(velocity_clean_high), *popt2_high)
        for i in range(len(fitted_power_high_2)):
            if fitted_power_high_2[i] < 0:
                fitted_power_high_2[i] = 0
        for i in range(len(fitted_power_high_2)):
            if fitted_power_high_2[i] > 0:
                error_quadratic_high += abs(fitted_power_high_2[i] - power_clean_high[i])

        errors = {
            '1': error_lin_high,
            '2': error_quadratic_high
        }
        lowest_error = min(errors, key=errors.get)
        if lowest_error == '1':
            fitted_power_high = fitted_power_high_1
        elif lowest_error == '2':
            fitted_power_high = fitted_power_high_2

        # Convert the raw data from the file to named lists for the SECOND file. The second file gives the information
        # for the lowest slope gradient and therefore lowest resistance.
        for i in range(len(data_low)):
            if data_low[i][1] != 0:
                first_non_zero_power.append([a[1] for a in data_low].index(data_low[i][1]))
                first_non_zero_power.append([a[1] for a in data_low].index(data_low[i][1]))
        first_non_zero_power = first_non_zero_power[0]

        try:
            poplin, pcov = curve_fit(self.func_lin, array([4, 10]), array([0, data_low[first_non_zero_power][1]]), absolute_sigma=True)
        except Warning:
            pass
        except Exception:
            pass

        for j in range(len(data_low)):
            if round(data_low[j][0]) == 0:
                power_raw_low.append(0)
                velocity_raw_low.append(0)
                velocity_time_raw_low.append(data_low[j][2])
                power_time_raw_low.append(data_low[j][3])

            elif round(data_low[j][1]) == 0 and 4 <= data_low[j][0] <= 14:
                power_raw_low.append(poplin[0] * data_low[j][0] + poplin[1])
                velocity_raw_low.append(data_low[j][0])
                velocity_time_raw_low.append(data_low[j][2])
                power_time_raw_low.append(data_low[j][3])

            else:
                power_raw_low.append(data_low[j][1])
                velocity_raw_low.append(data_low[j][0])
                velocity_time_raw_low.append(data_low[j][2])
                power_time_raw_low.append(data_low[j][3])

        for i in range(power_raw_low.index(max(power_raw_low))):
            power_clean_low.append(power_raw_low[i])
            velocity_clean_low.append(velocity_raw_low[i])
            velocity_time_clean_low.append(velocity_time_raw_low[i])
            power_time_clean_low.append(power_time_raw_low[i])

        # Calculate errors and make a fit for the data of the SECOND file. This fit will be the main information
        # in the P-v plot.
        popt1_low, pcov = curve_fit(self.func_lin, array(velocity_clean_low), array(power_clean_low))
        fitted_power_low_1 = self.func_lin(array(velocity_clean_low), *popt1_low)
        for i in range(len(fitted_power_low_1)):
            if fitted_power_low_1[i] < 0:
                fitted_power_low_1[i] = 0
                index_low_below_zero_1 = i
        for i in range(len(fitted_power_low_1)):
            if fitted_power_low_1[i] > 0:
                error_lin_low += abs(fitted_power_low_1[i] - power_clean_low[i])

        popt2_low, pcov = curve_fit(self.func_quadratic, array(velocity_clean_low), array(power_clean_low))
        fitted_power_low_2 = self.func_quadratic(array(velocity_clean_low), *popt2_low)
        for i in range(len(fitted_power_low_2)):
            if fitted_power_low_2[i] < 0:
                fitted_power_low_2[i] = 0
                index_low_below_zero_2 = i
        for i in range(len(fitted_power_low_2)):
            if fitted_power_low_2[i] > 0:
                error_quadratic_low += abs(fitted_power_low_2[i] - power_clean_low[i])

        errors = {
            '1': error_lin_low,
            '2': error_quadratic_low
        }
        lowest_error = min(errors, key=errors.get)
        global index_low_below_zero

        if lowest_error == '1':
            fitted_power_low = fitted_power_low_1
            index_low_below_zero = index_low_below_zero_1
        elif lowest_error == '2':
            fitted_power_low = fitted_power_low_2
            index_low_below_zero = index_low_below_zero_2

        # Convert the raw data from the file to named lists for the THIRD file. This file will be used to calculate the
        #  basic resistance and simulated mass if this is not given by the user.
        velocity_const = [velocity_clean_low[index_low_below_zero]]
        for j in range(len(data_const)):
            if first_limit - range_half < (data_const[j][0]) < first_limit + range_half:
                power_const_1.append(data_const[j][1])
                velocity_const_1.append(data_const[j][0])
                velocity_time_raw_const.append(data_const[j][2])
                power_time_raw_const.append(data_const[j][3])
        power_trainer.append(mean(power_const_1))

        velocity_const.append(mean(velocity_const_1))
        power_const.append(mean(power_const_1))

        # Convert the raw data from the file to named lists for the FOURTH file. This file will be used to calculate the
        # precision of the trainer.
        # for j in range(len(data_sensor)):
        #     if first_limit - range_half < (data_sensor[j][0] * self.sprocket_ratio * self.wheel_radius) < first_limit + range_half:
        #         power_clean_sensor.append(data_sensor[j][1])
        #         sensor_time_raw.append(data_sensor[j][2])
        # power_clean_sensor_mean = mean(power_clean_sensor)

        # Start calculations on the THIRD file to calculate the SIMULATED MASS. This includes fitting the
        # data.
        for i in range(len(power_const)):
            power_clean_const.append(power_const[i])
            velocity_clean_const.append(velocity_const[i])
            velocity_time_clean_const.append(velocity_time_raw_const[i])
            power_time_clean_const.append(power_time_raw_const[i])

        # Start calculations on the THIRD file to calculate the deviation limits.


        # Calculating the best possible fit, we only consider quadratic and linear fits at this moment. The error with
        # the original data is calculated and the best fit will be drawn. A dictionary is used to track the variable
        # with the highest value without a big if statement structure.
        popt1, pcov = curve_fit(self.func_lin, array(velocity_clean_const), array(power_clean_const), absolute_sigma=True)
        fitted_power_const_1 = self.func_lin(array(velocity_clean_const), *popt1)
        for i in range(len(fitted_power_const_1)):
            if fitted_power_const_1[i] < 0:
                fitted_power_const_1[i] = 0

        fitted_power_const = self.func_lin(array(velocity_clean_low), *popt1)
        for i in range(len(fitted_power_const)):
            if fitted_power_const[i] < 0:
                fitted_power_const[i] = 0
        for i in range(len(velocity_clean_high)):
            power_to_substract = popt1[0] * velocity_clean_high[i]
            power_no_int_res_high_imd.append(fitted_power_high[i] - power_to_substract)

        popt4, pcov = curve_fit(self.func_lin, array(velocity_time_clean_low), array(velocity_clean_low) / 3.6)


        # Calculation of the power which is needed to accelerate the flywheel. For the first and second file. The
        # simulated mass will also be calculated if there is no user input. If the simulated mass is out of given
        # bounds, the used simulated mass will be zero. Acceleration will be assumed constant, otherwise it is not
        # possible to calculate the needed values, because of the fluctuations in he trainers' output data. There will
        # also be a plot for the maximum brake force when substracting all the internal resistances.
        j = 0
        if self.checkbox.GetValue() != True:
            self.simulated_mass_guess = []
            for i in range(len(fitted_power_const)):
                if velocity_clean_low[i] > 15:
                    power_flywheel.append(fitted_power_low[i] - fitted_power_const[i])
                    if (velocity_clean_low[i] * popt4[0]) == 0:
                        continue
                    else:
                        self.simulated_mass_guess.append(power_flywheel[j] / (velocity_clean_low[i]/3.6 * popt4[0]))
                    j += 1
            self.simulated_mass_guess = mean(self.simulated_mass_guess)
            if 0 > self.simulated_mass_guess:
                self.simulated_mass_guess = 0

        popt5, pcov = curve_fit(self.func_lin, array(velocity_time_clean_high), array(velocity_clean_high) / 3.6)
        popt6, pcov = curve_fit(self.func_lin, array(velocity_time_clean_low), array(velocity_clean_low) / 3.6)

        for i in range(len(velocity_clean_high)):
            power_flywheel_high.append(float(self.simulated_mass_guess) * velocity_clean_high[i] / 3.6 * popt5[0])
            power_clean_high_brake.append(fitted_power_high[i] - power_flywheel_high[i])
        for i in range(len(velocity_clean_low)):
            power_flywheel_low.append(float(self.simulated_mass_guess) * velocity_clean_low[i] / 3.6 * popt6[0])
            power_clean_low_brake.append(fitted_power_low[i] - power_flywheel_low[i])
        for i in range(len(power_no_int_res_high_imd)):
            power_flywheel_high_imd.append(float(self.simulated_mass_guess) * velocity_clean_high[i] / 3.6 * popt5[0])
            power_no_int_res_high.append(power_no_int_res_high_imd[i] - power_flywheel_high_imd[i])
        for i in range(len(power_no_int_res_high)):
            if power_no_int_res_high[i] < 0:
                power_no_int_res_high[i] = 0

        # Calculate the resistance which should be present when cycling a conventional road.
        velocity_x = []
        percentage_lines = [1, 2, 5, 10, 20, 30]
        theoretical_power_values = [[], [], [], [], [], []]
        velocities_for_percentages = [[], [], [], [], [], []]
        for i in range(10 * round(max(velocity_clean_low))):
            velocity_x.append(i / 10)
        for j in range(len(percentage_lines)):
            for i in range(len(velocity_x)):
                if self.theoretical_power_at_velocity(velocity_x[i], percentage_lines[j]) <= max(power_clean_high):
                    theoretical_power_values[j].append(self.theoretical_power_at_velocity(velocity_x[i], percentage_lines[j]))
                    velocities_for_percentages[j].append(velocity_x[i])
                else:
                    pass

        # Initialize writing an excel file. This file will be used to store all the necessary information which is
        # analysed in the code.
        excel = xlsxwriter.Workbook(self.folder_pathname + "\\" + self.user_file_name + ".xlsx")
        try:
            graph = excel.add_chart({'type': 'scatter', 'subtype': 'straight'})
        except Exception:
            print('NO')
        graph_2 = excel.add_chart({'type': 'scatter', 'subtype': 'straight'})
        graph_3 = excel.add_chart({'type': 'scatter', 'subtype': 'straight'})
        worksheet_charts = excel.add_worksheet('Charts')
        worksheet_data = excel.add_worksheet('Data')

        # Setting variables for excel file.
        bold = excel.add_format({'bold': True})
        underline = excel.add_format({'bold': True, 'underline': True})
        header = excel.add_format({'bold': True, 'font_size': 24})
        superscript = excel.add_format({'bold': True, 'font_size': '24', 'font_script': True})
        graph.set_y_axis({'name': 'Power [W]'})
        graph.set_x_axis({'name': 'Velocity [km/h]'})
        graph.set_title({'name': 'Operating range ' + self.user_file_name})
        graph.set_size({'width': 1080, 'height': 720})
        graph_2.set_y_axis({'name': 'Power [W]'})
        graph_2.set_y2_axis({'name': 'Velocity [km/h]'})
        graph_2.set_x_axis({'name': 'Time [s]'})
        graph_2.set_title({'name': 'Power ' + self.user_file_name + ' vs. Power External Sensor'})
        graph_2.set_size({'width': 1080, 'height': 720})
        graph_3.set_y_axis({'name': 'Velocity [km/h]'})
        graph_3.set_x_axis({'name': 'Time [s]'})
        graph_3.set_title({'name': 'Velocity ' + self.user_file_name + ' vs. Velocity External Sensor'})
        graph_3.set_size({'width': 1080, 'height': 720})
        worksheet_data.set_column('A:Q', 14)
        worksheet_charts.set_column('X:X', 16)

        # Writing to excel file.
        worksheet_data.write('A1', 'Tested with highest gradient (without slip)', underline)
        worksheet_data.write('A2', 'Velocity [km/h]', bold)
        worksheet_data.write('B2', 'Power [W]', bold)
        worksheet_data.write('C2', 'Fitted Power [W]', bold)
        worksheet_data.write_column(2, 0, velocity_clean_high)
        worksheet_data.write_column(2, 1, power_clean_high)
        worksheet_data.write_column(2, 2, fitted_power_high)

        worksheet_data.write('E1', 'Tested with lowest gradient (without slip)', underline)
        worksheet_data.write('E2', 'Velocity [km/h]', bold)
        worksheet_data.write('F2', 'Power [W]', bold)
        worksheet_data.write('G2', 'Fitted Power [W]', bold)
        worksheet_data.write_column(2, 4, velocity_clean_low)
        worksheet_data.write_column(2, 5, power_clean_low)
        worksheet_data.write_column(2, 6, fitted_power_low)

        worksheet_data.write('I1', 'Tested with 0W Power, combined with other files', underline)
        worksheet_data.write('I2', 'Upper Limit [W]', bold)
        worksheet_data.write('J2', 'Lower Limit [W]', bold)
        worksheet_data.write_column(2, 8, power_clean_high_brake)
        worksheet_data.write_column(2, 9, power_clean_low_brake)

        worksheet_data.write('K2', 'Brake Limit [W]', bold)
        worksheet_data.write_column(2, 10, power_no_int_res_high)

        worksheet_data.write_column(2, 12, power_time_const_check)
        worksheet_data.write_column(2, 13, power_const_check)
        worksheet_data.write_column(2, 14, power_const_check_lower_bound)
        worksheet_data.write_column(2, 15, power_const_check_upper_bound)

        worksheet_data.write_column(2, 16, power_time_sensor_check)
        worksheet_data.write_column(2, 17, power_sensor_check)
        worksheet_data.write_column(2, 18, power_sensor_check_lower_bound)
        worksheet_data.write_column(2, 19, power_sensor_check_upper_bound)
        worksheet_data.write_column(2, 20, speed_sensor_check)
        worksheet_data.write_column(2, 21, speed_sensor_check_lower_bound)
        worksheet_data.write_column(2, 22, speed_sensor_check_upper_bound)

        worksheet_data.write_column(2, 23, speed_trainer_check_time)
        worksheet_data.write_column(2, 24, speed_trainer_check)
        worksheet_data.write_column(2, 25, speed_trainer_check_lower_bound)
        worksheet_data.write_column(2, 26, speed_trainer_check_upper_bound)

        # Fill graph 1
        count = [0]
        i = 0
        velocity_check = []
        pop = []
        fill_power = []
        fill_power_low = []
        power_flywheel_h = []
        power_flywheel_l = []

        for i in range(1000):
            velocity_check.append(i * 0.2)
            if velocity_check[i] > max(velocity_clean_low):
                break
        worksheet_data.write_column(2, 449, velocity_check)
        fill_power_raw_low = self.func_lin(array(velocity_check), *popt1_low)
        fill_power_raw_high = self.func_lin(array(velocity_check), *popt1_high)
        for i in range(len(velocity_check)):
            power_flywheel_h.append(float(self.simulated_mass_guess) * velocity_check[i] / 3.6 * popt5[0])
            power_flywheel_l.append(float(self.simulated_mass_guess) * velocity_check[i] / 3.6 * popt6[0])
            fill_power.append(fill_power_raw_high[i] - power_flywheel_h[i])
            fill_power_low.append(fill_power_raw_low[i] - power_flywheel_l[i])


        for j in range(240):
                count.append(count[j] + 1)
                pops = 0
                fill_power_list = []
                for i in range(len(fill_power)):
                    fill_power_list.append(fill_power[i] - j * 5)
                    if fill_power_list[i - pops] < fill_power_low[i]:
                        fill_power_list.pop()
                        pops += 1
                    elif fill_power_list[i - pops] > max(power_clean_high_brake):
                        fill_power_list.pop()
                        break
                worksheet_data.write_column(2 + pops, 450 + j, fill_power_list)
                pop.append(pops)
                if len(fill_power_list) > 0:
                    if max(fill_power_list) <= max(fill_power_low):
                        break

        # Fill graph 2
        count_1 = [4]
        for i in range(int(trainer_deviation_perc * 2)):
            for j in range(len(power_const_check)):
                worksheet_data.write(2 + j, 50 + i, power_const_check[j] * (1 + i / 200))
                worksheet_data.write(2 + j, 50 + int(trainer_deviation_perc) * 2 + i, power_const_check[j] * (1 - i / 200))

        for i in range(int(sensor_deviation_perc * 2)):
            for j in range(len(power_sensor_check)):
                worksheet_data.write(2 + j, 150 + i, power_sensor_check[j] * (1 + i / 200))
                worksheet_data.write(2 + j, 150 + int(sensor_deviation_perc) * 2 + i, power_sensor_check[j] * (1 - i / 200))

        # Fill graph 3
        count_2 = [4]
        for i in range(int(trainer_deviation_perc * 2)):
            for j in range(len(power_const_check)):
                worksheet_data.write(2 + j, 250 + i, speed_trainer_check[j] * (1 + i / 200))
                worksheet_data.write(2 + j, 250 + int(trainer_deviation_perc) * 2 + i, speed_trainer_check[j] * (1 - i / 200))

        for i in range(int(sensor_deviation_perc * 2)):
            for j in range(len(power_sensor_check)):
                worksheet_data.write(2 + j, 350 + i, speed_sensor_check[j] * (1 + i / 200))
                worksheet_data.write(2 + j, 350 + int(sensor_deviation_perc) * 2 + i, speed_sensor_check[j] * (1 - i / 200))


        # Percentage Lines
        for i in range(len(velocities_for_percentages)):
            worksheet_data.write_column(2, i + 27, velocities_for_percentages[i])
            worksheet_data.write_column(2, i + 33, theoretical_power_values[i])

        # Writing to graph.
        for i in range(len(velocities_for_percentages)):
            graph.add_series({
                'categories': [worksheet_data.name] + [2, i + 27] + [len(velocities_for_percentages[i]) + 2, i + 27],
                'values': [worksheet_data.name] + [2, i + 33] + [len(theoretical_power_values[i]) + 2, i + 33],
                'line': {'color': '#67bfe7', 'width': 2.5, 'transparency': 70},
                'name': "Reference Lines: 1, 2, 5, 10, 20, 30%"
            })
        for i in range(len(count)):
            graph.add_series({
            'categories': [worksheet_data.name] + [2, 449] + [len(velocity_check) + 2, 449],
            'values': [worksheet_data.name] + [2, 450 + i] + [len(velocity_check) + 2, 450 + i],
            'line': {'color': 'red', 'width': 10, 'transparency': 99},
        })
        graph.add_series({
            'categories': [worksheet_data.name] + [2, 0] + [len(velocity_clean_high) + 2, 0],
            'values': [worksheet_data.name] + [2, 1] + [len(power_clean_high) + 2, 1],
            'line': {'color': 'blue', 'dash_type': 'dash', 'width': 1.5},
            'name': 'Highest Power',
        })


        graph.add_series({
            'categories': [worksheet_data.name] + [2, 0] + [len(velocity_clean_high) + 2, 0],
            'values': [worksheet_data.name] + [2, 8] + [len(power_clean_high) + 2, 8],
            'line': {'color': 'red', 'width': 3},
            'name': 'Highest Power - Without Flywheel Effects',
        })
        graph.add_series({
            'categories': [worksheet_data.name] + [2, 0] + [len(velocity_clean_high) + 2, 0],
            'values': [worksheet_data.name] + [2, 10] + [len(power_no_int_res_high) + 2, 10],
            'line': {'color': 'green', 'dash_type': 'dash', 'width': 2},
            'name': 'Highest Power - Without Internal Friction - Without Flywheel Effects',
        })
        graph.add_series({
            'categories': [worksheet_data.name] + [2, 4] + [len(velocity_clean_low) + 2, 4],
            'values': [worksheet_data.name] + [2, 5] + [len(power_clean_low) + 2, 5],
            'line': {'color': 'blue', 'dash_type': 'dash', 'width': 1.5},
            'name': 'Lowest Power',
        })
        graph.add_series({
            'categories': [worksheet_data.name] + [2, 4] + [len(velocity_clean_low) + 2, 4],
            'values': [worksheet_data.name] + [2, 9] + [len(power_clean_low_brake) + 2, 9],
            'line': {'color': 'red', 'width': 3},
            'name': 'Lowest Power - Without Flywheel Effects',
        })
        graph.add_series({
            'categories': [worksheet_data.name] + [2, 4] + [len(velocity_clean_low) + 2, 4],
            'values': [worksheet_data.name] + [2, 6] + [len(power_clean_low) + 2, 6],
            'line': {'color': 'blue', 'width': 2},
            'name': 'Fitted Lowest Power',
        })
        graph.add_series({
            'categories': [worksheet_data.name] + [2, 0] + [len(velocity_clean_high) + 2, 0],
            'values': [worksheet_data.name] + [2, 2] + [len(power_clean_high) + 2, 2],
            'line': {'color': 'blue', 'width': 2},
            'name': 'Fitted Highest Power',
        })

        graph_2.add_series({
            'values': [worksheet_data.name] + [2, 14] + [len(power_const_check) + 2, 14],
            'categories': [worksheet_data.name] + [2, 12] + [len(power_time_const_check) + 2, 12],
            'line': {'color': 'blue', 'width': 1.5},
            'name': 'Trainer Power Lower Bound (-' + str(trainer_deviation_perc) + '%)',
        })
        graph_2.add_series({
            'values': [worksheet_data.name] + [2, 15] + [len(power_const_check) + 2, 15],
            'categories': [worksheet_data.name] + [2, 12] + [len(power_time_const_check) + 2, 12],
            'line': {'color': 'blue', 'width': 1.5},
            'name': 'Trainer Power Upper Bound (' + str(trainer_deviation_perc) + '%)',
        })
        graph_2.add_series({
            'values': [worksheet_data.name] + [2, 18] + [len(power_sensor_check) + 2, 18],
            'categories': [worksheet_data.name] + [2, 16] + [len(power_time_sensor_check) + 2, 16],
            'line': {'color': 'red', 'width': 1.5},
            'name': 'Sensor Power Lower Bound (-' + str(sensor_deviation_perc) + '%)',
        })
        graph_2.add_series({
            'values': [worksheet_data.name] + [2, 19] + [len(power_sensor_check) + 2, 19],
            'categories': [worksheet_data.name] + [2, 16] + [len(power_time_sensor_check) + 2, 16],
            'line': {'color': 'red', 'width': 1.5},
            'name': 'Sensor Power Upper Bound (' + str(sensor_deviation_perc) + '%)',
        })
        for i in range(int(4 * sensor_deviation_perc)):
            graph_2.add_series({
                'values': [worksheet_data.name] + [2, 150 + i] + [len(power_const_check) + 2, 150 + i],
                'categories': [worksheet_data.name] + [2, 16] + [len(power_const_check) + 2, 16],
                'line': {'color': 'red', 'width': 3, 'transparency': 98},
            })
            count_1.append(count_1[i] + 1)
        for i in range(int(4 * trainer_deviation_perc)):
            graph_2.add_series({
                'values': [worksheet_data.name] + [2, 50 + i] + [len(power_const_check) + 2, 50 + i],
                'categories': [worksheet_data.name] + [2, 12] + [len(power_const_check) + 2, 12],
                'line': {'color': 'blue', 'width': 3, 'transparency': 98},
            })
            count_1.append(count_1[4*int(sensor_deviation_perc) + i] + 1)


        graph_3.add_series({
            'categories': [worksheet_data.name] + [2, 16] + [len(speed_sensor_check_lower_bound) + 2, 16],
            'values': [worksheet_data.name] + [2, 21] + [len(speed_sensor_check_lower_bound) + 2, 21],
            'line': {'color': 'red', 'width': 1.5},
            'name': 'Sensor Velocity Lower Bound (-' + str(sensor_deviation_perc) + '%)',
        })
        graph_3.add_series({
            'categories': [worksheet_data.name] + [2, 16] + [len(speed_sensor_check_upper_bound) + 2, 16],
            'values': [worksheet_data.name] + [2, 22] + [len(speed_sensor_check_upper_bound) + 2, 22],
            'line': {'color': 'red', 'width': 1.5},
            'name': 'Sensor Velocity Upper Bound (' + str(sensor_deviation_perc) + '%)',
        })
        graph_3.add_series({
            'categories': [worksheet_data.name] + [2, 23] + [len(speed_trainer_check_lower_bound) + 2, 23],
            'values': [worksheet_data.name] + [2, 25] + [len(speed_trainer_check_lower_bound) + 2, 25],
            'line': {'color': 'blue', 'width': 1.5},
            'name': 'Trainer Velocity Lower Bound (-' + str(trainer_deviation_perc) + '%)',
        })
        graph_3.add_series({
            'categories': [worksheet_data.name] + [2, 23] + [len(speed_trainer_check_lower_bound) + 2, 23],
            'values': [worksheet_data.name] + [2, 26] + [len(speed_trainer_check_lower_bound) + 2, 26],
            'line': {'color': 'blue', 'width': 1.5},
            'name': 'Trainer Velocity Upper Bound (' + str(trainer_deviation_perc) + '%)',
        })
        for i in range(int(4 * sensor_deviation_perc)):
            graph_3.add_series({
                'values': [worksheet_data.name] + [2, 350 + i] + [len(power_const_check) + 2, 350 + i],
                'categories': [worksheet_data.name] + [2, 16] + [len(power_const_check) + 2, 16],
                'line': {'color': 'red', 'width': 3, 'transparency': 97},
            })
            count_2.append(count_2[i] + 1)
        for i in range(int(4 * trainer_deviation_perc)):
            graph_3.add_series({
                'values': [worksheet_data.name] + [2, 250 + i] + [len(power_const_check) + 2, 250 + i],
                'categories': [worksheet_data.name] + [2, 23] + [len(power_const_check) + 2, 23],
                'line': {'color': 'blue', 'width': 3, 'transparency': 97},
            })
            count_2.append(count_2[4* int(sensor_deviation_perc) + i] + 1)


        list =[*range(1, 7 + max(count))]

        worksheet_charts.insert_chart('B2', graph)
        worksheet_charts.insert_chart('B40', graph_2)
        worksheet_charts.insert_chart('S40', graph_3)
        graph_2.set_legend({'position': 'bottom', 'delete_series': count_1})
        graph_3.set_legend({'position': 'bottom', 'delete_series': count_2})
        graph.set_legend({'position': 'bottom', 'delete_series': list})
        worksheet_charts.write('T2', 'Simulated Mass:', header)
        worksheet_charts.write('X2', str(round(float(self.simulated_mass_guess), 2)), header)
        worksheet_charts.write_rich_string('Y2', header, '[kg]')
        worksheet_charts.write('T3', 'Mean deviation power:', header)
        worksheet_charts.write('X3', str(round(mean_error_power, 2)), header)
        # TODO: uncomment (only uncomment mean if there are more measurements.
        # worksheet_charts.write('T3', 'Minimal precision:', header)
        # worksheet_charts.write('X3', str(round(float(precision_trainer_mean), 2)), header)
        # worksheet_charts.write_rich_string('Y3', header, '[%]')
        # worksheet_charts.write('T4', 'Mean precision:', header)
        # worksheet_charts.write('X4', str(round(float(precision_trainer_max), 2)), header)
        # worksheet_charts.write_rich_string('Y4', header, '[%]')


        try:
            excel.close()
        except Exception:
            excel_open_dialog = wx.MessageDialog(self.top_panel, style= wx.OK | wx.CANCEL | wx.ICON_ERROR,
                                                 message="Excel seems to be still running. It needs to be closed for this application to be able to save a new file.\n\nClick ""Ok"" to try again, after having closed Excel.",
                                                 caption="Error!")
            excel_open_dialog.CenterOnParent()
            if excel_open_dialog.ShowModal() == wx.OK:
                excel_open_dialog.Destroy()
                try:
                    excel.close()
                except:
                    excel_open_dialog = wx.MessageDialog(self.top_panel, style= wx.OK | wx.ICON_ERROR,
                                     message="Something still went wrong. \n\n Please restart, and make sure no other instances of this application are running.",
                                     caption="Error!")
                    excel_open_dialog.CenterOnParent()
                    if excel_open_dialog.ShowModal() == wx.OK:
                        excel_open_dialog.Destroy()

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

    def logfile_analyser_trainer(self, logfile):
        global velocity_list, power_list, time_list, velocity_time_list, power_time_list
        sentences = []
        value_list = []
        time_list = []
        time_values_raw = []
        self.time_stamp_trainer = []
        sentences_1 = []
        velocity_list = []
        velocity_time_list = []
        power_time_list = []
        power_list = []
        speed = True
        power = True

        # The log-file is opened here.
        log = open(logfile)
        with open(logfile) as f:
            for lines, l in enumerate(f):
                pass

        # The important lines will be retrieved from the log-file here, by looking at the lines which start with 'Rx'. These are received messages.
        for n in range(lines):
            sentence = log.readline()
            if "Rx:" in sentence:
                sentences.append(sentence)
                sentences_1.append(sentence)
            elif "Rx" in sentence:
                sentences_1.append(sentence)

        # for i in range(len(sentences_1)):
        #     sentence_1 = sentences_1[i].split()
        #     try:
        #         index_1 = sentence_1.index("Rx")
        #     except:
        #         index_1 = sentence_1.index("Rx:")
        #     self.time_stamp_trainer.append(sentence_1[index_1 - 2])

        # This part splits the retrieved lines in subparts, after which the hexadecimals will be read.
        for i in range(len(sentences)):
            sentence = sentences[i].split()
            index = sentence.index("Rx:")
            value_raw = sentence[index + 1]
            time_raw = sentence[index - 2]
            value = value_raw.replace("[", "").replace("]", "")  # This will removes the useless characters
            value_list.append(value)
            time_values_raw.append(time_raw)
        self.time_stamp_trainer = time_values_raw

        for i in range(len(value_list)):
            value_list_characters = list(value_list[i])

        # These subparts will be categorized according to their first character: When this is '10', this means the velocity is recorded in that line. When the first character is '19', this means power is recorded in that line. The other characters are not important for the functionality of this file, which means they will be left out.
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
                velocity_time_list.append((float(time_values_raw[i]) - float(time_values_raw[0])) / 1000)

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
                power_time_list.append((float(time_values_raw[i]) - float(time_values_raw[0])) / 1000)
            else:
                pass

    def logfile_analyser_sensor(self, logfile):
        global cadence_list, power_list, time_list, sensor_time_list
        self.time_stamp_sensor = []
        sentences = []
        sentences_1 = []
        value_list = []
        time_list = []
        time_values_raw = []
        cadence_list = []
        sensor_time_list = []
        power_list = []

        # The log-file is opened here.
        log = open(logfile)
        with open(logfile) as f:
            for lines, l in enumerate(f):
                pass

        # The important lines will be retrieved from the log-file here, by looking at the lines which start with 'Rx'. These are received messages.
        for n in range(lines):
            sentence = log.readline()
            if "Rx:" in sentence:
                sentences.append(sentence)
                sentences_1.append(sentence)
            elif "Rx" in sentence:
                sentences_1.append(sentence)

        # for i in range(len(sentences_1)):
        #     sentence_1 = sentences_1[i].split()
        #     try:
        #         index_1 = sentence_1.index("Rx")
        #     except:
        #         index_1 = sentence_1.index("Rx:")
        #     self.time_stamp_sensor.append(sentence_1[index_1 - 2])

        # This part splits the retrieved lines in subparts, after which the hexadecimals will be read.
        for i in range(len(sentences)):
            sentence = sentences[i].split()
            index = sentence.index("Rx:")
            value_raw = sentence[index + 1]
            time_raw = sentence[index - 2]
            value = value_raw.replace("[", "").replace("]", "")  # This will removes the useless characters
            value_list.append(value)
            time_values_raw.append(time_raw)

        self.time_stamp_sensor = time_values_raw

        # These subparts will be categorized according to their first character: When this is '10', this means the cadence is recorded in that line. When the first character is '19', this means power is recorded in that line. The other characters are not important for the functionality of this file, which means they will be left out.
        for i in range(len(value_list)):
            value_list_characters = list(value_list[i])

            if value_list_characters[0] == '1' and value_list_characters[1] == '0':
                cadence_values_raw = [value_list_characters[6], value_list_characters[7]]
                power_values_raw = [value_list_characters[15], value_list_characters[12], value_list_characters[13]]
                cadence_values_raw_string = "".join(cadence_values_raw)
                power_values_raw_string = "".join(power_values_raw)
                value_converter = ValueConverter()
                cadence_bin = value_converter.hex_to_bin(cadence_values_raw_string)
                power_bin = value_converter.hex_to_bin(power_values_raw_string)
                cadence_list.append(value_converter.bin_to_dec(cadence_bin))
                power_list.append(value_converter.bin_to_dec(power_bin))
                sensor_time_list.append((float(time_values_raw[i]) - float(time_values_raw[0])) / 1000)

            else:
                pass

    def on_exit_button(self, event):
        self.Close()

    def on_reset(self, event):
        if __name__ == '__main__':
            self.Close()
            frame = Main(None, 'SimulANT+ Log Analyzer').Show()

    def on_xlsx_button(self, event):
        if path.isfile(self.folder_pathname + "\\" + self.user_file_name + ".xlsx"):
            startfile(self.folder_pathname + "\\" + self.user_file_name + ".xlsx")
        elif self.folder_pathname == "":
            no_file_dialog = wx.MessageDialog(self.top_panel,
                                              message="The file does not exist. Please try selecting files with Open LOG's",
                                              caption="Warning!")
            no_file_dialog.CenterOnParent()
            if no_file_dialog.ShowModal() == wx.OK:
                no_file_dialog.Destroy()
                return

    def func_powerlaw(self, x, m, c):
        return x ** m * c

    def func_quadratic(self, x, a, b, c):
        return a * x ** 2 + b * x + c

    def func_quadratic_without_constant(self, x, a, b):
        return a * x ** 2 + b * x

    def func_lin(self, x, a, b):
        return a * x + b

    def func_lin_without_constant(self, x, a):
        return a * x

    def on_check(self, event):
        self.check_counter += 1
        if self.check_counter % 2 != 0:
            self.edit_simulated_mass_text.SetEditable(True)
            self.simulated_mass_alt_1_text.SetEditable(True)
            self.simulated_mass_alt_2_text.SetEditable(True)

            self.edit_simulated_mass_text.SetBackgroundColour((255, 255, 255))
            self.simulated_mass_alt_1_text.SetBackgroundColour((255, 255, 255))
            self.simulated_mass_alt_2_text.SetBackgroundColour((255, 255, 255))

        elif self.check_counter % 2 == 0:
            self.edit_simulated_mass_text.SetEditable(False)
            self.simulated_mass_alt_1_text.SetEditable(False)
            self.simulated_mass_alt_2_text.SetEditable(False)

            self.edit_simulated_mass_text.SetBackgroundColour((220, 220, 220))
            self.simulated_mass_alt_1_text.SetBackgroundColour((220, 220, 220))
            self.simulated_mass_alt_2_text.SetBackgroundColour((220, 220, 220))

            self.edit_simulated_mass_text.SetValue("")
            self.simulated_mass_alt_1_text.SetValue("")
            self.simulated_mass_alt_2_text.SetValue("")

    def on_exit_widget_enter(self, event):
        self.statusbar.SetStatusText('Exit the program')
        event.Skip()

    def on_open_widget_enter(self, event):
        self.statusbar.SetStatusText('Open multiple LOG-files')
        event.Skip()

    def on_reset_widget_enter(self, event):
        self.statusbar.SetStatusText('Reset the program')
        event.Skip()

    def on_excel_widget_enter(self, event):
        self.statusbar.SetStatusText('Open the created Excel-file')
        event.Skip()

    def on_check_hover(self, event):
        self.statusbar.SetStatusText('Enable the option to use user-input simulated mass')
        event.Skip()

    def on_save_hover(self, event):
        self.statusbar.SetStatusText('Save the entered inputs to use them in calculations')
        event.Skip()

    def on_save_inputs(self, event):
        try:
            self.front_gear_value = float(self.edit_gear_front_text.GetValue())
            self.rear_gear_value = float(self.edit_gear_rear_text.GetValue())
            global trainer_deviation_perc, sensor_deviation_perc
            trainer_deviation_perc = float(self.trainer_deviation_text.GetValue())
            sensor_deviation_perc = float(self.sensor_deviation_text.GetValue())

            self.sprocket_ratio = self.front_gear_value / self.rear_gear_value

            if self.edit_simulated_mass_text.GetValue() == "" and self.checkbox.GetValue() == True:
                self.inertia_value = float(self.simulated_mass_alt_1_text.GetValue())
                self.ratio_value = float(self.simulated_mass_alt_2_text.GetValue())
                self.simulated_mass_guess = self.inertia_value * self.ratio_value ** 2
            elif self.edit_simulated_mass_text.GetValue() == "" and self.checkbox.GetValue() == False:
                self.simulated_mass_guess = 0
            else:
                self.simulated_mass_guess = float(self.edit_simulated_mass_text.GetValue())
            self.saved_text.SetFont(self.font_green)
            self.saved_text.SetLabel("VALUES SAVED")
            self.saved_text.SetForegroundColour((10, 255, 10))
            self.saved = True

        except ValueError:
            no_number_dialog = wx.MessageDialog(self.top_panel, style=wx.ICON_ERROR, message="This doesn't appear to be a number. \n\nPlease try again.")
            no_number_dialog.CenterOnParent()
            if no_number_dialog.ShowModal() == wx.OK:
                no_number_dialog.Destroy()
                pass


    def theoretical_power_at_velocity(self, velocity, theta):
        # Variables
        frontal_area = 0.4 # m^2
        air_density = 1.226 # kg / m^3
        drag_coefficient = 0.85
        mass = 82.5 # kg
        grav = 9.81 # m / s^2
        angle = arctan(float(theta) / 100)
        roll_coefficient = 0.004

        return (0.5 * frontal_area * air_density * drag_coefficient * (velocity / 3.6) ** 2 + sin(angle) * mass * grav + cos(angle) * mass * grav * roll_coefficient) * (velocity / 3.6)


if __name__ == '__main__':
    Application = wx.App(False)
    frame = Main(None, 'SimulANT+ Log Analyzer                                                                       '
                       '                                                    [v1.1]').Show()
    Application.MainLoop()
