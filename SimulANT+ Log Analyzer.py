import wx as wx
from ValueConverter import ValueConverter
import xlsxwriter
import os
from numpy import mean
from numpy import array
from scipy.optimize import curve_fit


class Main(wx.Frame):
    def __init__(self, parent, title):
        """
        Initializing the program:

        1:
        :param parent:
        :param title:
        """
        wx.Frame.__init__(self, parent, title=title,
                          style=wx.DEFAULT_FRAME_STYLE & ~(wx.RESIZE_BORDER | wx.MAXIMIZE_BOX), size=(720, 920))
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
        self.exit_button = wx.Button(self.top_panel, -1, label='Exit', pos=(590, 795), size=(100, 30))
        self.reset_button = wx.Button(self.top_panel, -1, label='Reset Program', pos=(480, 795), size=(100, 30))
        self.open_xlsx_button = wx.Button(self.top_panel, -1, label='Open Excel File', pos=(370, 795), size=(100, 30))
        self.open_files_butten = wx.Button(self.top_panel, -1, label='Open LOG\'s', pos=(260, 795), size=(100, 30))

        # Loading images
        # image_file_png = wx.Image("tacx-logo.png", wx.BITMAP_TYPE_ANY).ConvertToBitmap()

        # Creating panels
        self.font_header = wx.Font(12, family=wx.DECORATIVE, style=wx.NORMAL, weight=wx.BOLD)
        self.font_normal = wx.Font(10, family=wx.DECORATIVE, style=wx.NORMAL, weight=wx.NORMAL)

        self.path_panel_1 = wx.Panel(self.top_panel, -1, style=wx.SUNKEN_BORDER, size=(680, 50), pos=(10, 10))
        self.path_header_display = wx.StaticText(self.path_panel_1, label="Path to first selected LOG-file: ",
                                                 pos=(4, 2))
        self.path_header_display.SetFont(self.font_header)

        self.path_panel_2 = wx.Panel(self.top_panel, -1, style=wx.SUNKEN_BORDER, size=(680, 50), pos=(10, 70))
        self.path_header_display = wx.StaticText(self.path_panel_2, label="Path to second selected LOG-file: ",
                                                 pos=(4, 2))
        self.path_header_display.SetFont(self.font_header)

        self.path_panel_3 = wx.Panel(self.top_panel, -1, style=wx.SUNKEN_BORDER, size=(680, 50), pos=(10, 130))
        self.path_header_display = wx.StaticText(self.path_panel_3, label="Path to third selected LOG-file: ",
                                                 pos=(4, 2))
        self.path_header_display.SetFont(self.font_header)

        self.path_panel_4 = wx.Panel(self.top_panel, -1, style=wx.SUNKEN_BORDER, size=(680, 50), pos=(10, 190))
        self.path_header_display = wx.StaticText(self.path_panel_4, label="Path to fourth selected LOG-file: ",
                                                 pos=(4, 2))
        self.path_header_display.SetFont(self.font_header)

        self.some_data_panel_1 = wx.Panel(self.top_panel, -1, style=wx.SUNKEN_BORDER, size=(680, 80), pos=(10, 260))
        self.data_panel_1_header_display = wx.StaticText(self.some_data_panel_1,
                                                         label="Some statistics about the first file: ", pos=(4, 2))
        self.data_panel_1_header_display.SetFont(self.font_header)

        self.some_data_panel_2 = wx.Panel(self.top_panel, -1, style=wx.SUNKEN_BORDER, size=(680, 80), pos=(10, 350))
        self.data_panel_2_header_display = wx.StaticText(self.some_data_panel_2,
                                                         label="Some statistics about the second file: ", pos=(4, 2))
        self.data_panel_2_header_display.SetFont(self.font_header)

        self.some_data_panel_3 = wx.Panel(self.top_panel, -1, style=wx.SUNKEN_BORDER, size=(680, 80), pos=(10, 440))
        self.data_panel_3_header_display = wx.StaticText(self.some_data_panel_3,
                                                         label="Some statistics about the third file: ", pos=(4, 2))
        self.data_panel_3_header_display.SetFont(self.font_header)
        
        self.some_data_panel_4 = wx.Panel(self.top_panel, -1, style=wx.SUNKEN_BORDER, size=(680, 80), pos=(10, 530))
        self.data_panel_4_header_display = wx.StaticText(self.some_data_panel_4,
                                                         label="Some statistics about the fourth file: ", pos=(4, 2))
        self.data_panel_4_header_display.SetFont(self.font_header)
        # self.image_panel = wx.Panel(self.top_panel, -1, style=wx.BORDER_SIMPLE, size=())

        self.checkbox_sim = wx.CheckBox(self.top_panel, -1, 'User Input Simulated Mass', pos=(30, 795))
        self.checkbox_sim.SetValue(False)

        self.checkbox_inertia = wx.CheckBox(self.top_panel, -1, 'User Input Moment of Inertia', pos=(30, 815))
        self.checkbox_inertia.SetValue(False)

        # Set events
        self.Bind(wx.EVT_MENU, self.on_open, menu_file_open)
        self.Bind(wx.EVT_MENU, self.on_about, menu_about)
        self.Bind(wx.EVT_MENU, self.on_exit, menu_exit)
        self.exit_button.Bind(wx.EVT_BUTTON, self.on_exit_button)
        self.reset_button.Bind(wx.EVT_BUTTON, self.on_reset)
        self.open_xlsx_button.Bind(wx.EVT_BUTTON, self.on_xlsx_button)
        self.open_files_butten.Bind(wx.EVT_BUTTON, self.on_open)
        self.Bind(wx.EVT_CHECKBOX, self.on_check_sim_mass)
        self.Bind(wx.EVT_CHECKBOX, self.on_check_inertia)


        # Set start-up message
        welcome_dialog = wx.MessageDialog(self.top_panel,
                                          message="Welcome to SimulANT+ Log Analyzer. \nIf you have read the README.txt, you're good to go. \nIf you haven't yet, please do.",
                                          caption="Welcome!")
        welcome_dialog.CenterOnParent()
        if welcome_dialog.ShowModal() == wx.OK:
            welcome_dialog.Destroy()
            return

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
        self.velocity_list_zero = []
        self.power_list_low = []
        self.power_list_high = []
        self.power_list_zero = []
        self.time_list_high = []
        self.time_list_low = []
        self.time_list_zero = []

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

        self.path_display = wx.StaticText(self.path_panel_3, label=str(self.pathname_3), pos=(4, 25))
        self.path_display.SetFont(self.font_normal)

        self.path_display = wx.StaticText(self.path_panel_4, label=str(self.pathname_4), pos=(4, 25))
        self.path_display.SetFont(self.font_normal)

        self.data_panel_1_display = wx.StaticText(self.some_data_panel_1,
                                                  label="Average power at high slope:     " + str(
                                                      self.power_high_avg) + " W\n" + "Average velocity at high slope:   " + str(
                                                      self.velocity_high_avg) + " km/h\n" + "Amount of received ANT+ messages:   " + str(
                                                      len(self.velocity_list_high)), pos=(4, 24))
        self.data_panel_1_display.SetFont(self.font_normal)

        self.data_panel_2_display = wx.StaticText(self.some_data_panel_2,
                                                  label="Average power at low (negative) slope:     " + str(
                                                      self.power_low_avg) + " W\n""Average velocity at low (negative) slope:   " + str(
                                                      self.velocity_low_avg) + " km/h\n""Amount of received ANT+ messages:   " + str(
                                                      len(self.velocity_list_low)), pos=(4, 24))
        self.data_panel_2_display.SetFont(self.font_normal)

        self.data_panel_3_display = wx.StaticText(self.some_data_panel_3,
                                                  label="Average power at 0 Watt programming - low accelaration:     " + str(
                                                      self.power_zero_avg) + " W\n""Average velocity at 0 Watt programming:   " + str(
                                                      self.velocity_zero_avg) + " km/h\n""Amount of received ANT+ messages:   " + str(
                                                      len(self.velocity_list_zero)), pos=(4, 24))
        self.data_panel_3_display.SetFont(self.font_normal)

        self.data_panel_4_display = wx.StaticText(self.some_data_panel_4,
                                                  label="Average power at 0 Watt programming - high acceleration:     " + str(
                                                      self.power_zero_acc_avg) + " W\n""Average velocity at 0 Watt programming:   " + str(
                                                      self.velocity_zero_acc_avg) + " km/h\n""Amount of received ANT+ messages:   " + str(
                                                      len(self.velocity_list_zero_acc)), pos=(4, 24))
        self.data_panel_4_display.SetFont(self.font_normal)

        xlsx_path_panel = wx.Panel(self.top_panel, -1, style=wx.SUNKEN_BORDER, size=(680, 50), pos=(10, 630))
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
        # Opening File 1 with the use of a dialog. File 1 will contain the ANT+ data of the measurements with a high
        # gradient (slope). This will be used to calculate the maximal brake power.
        with wx.FileDialog(self, "Choose the logged SimulANT+ file with the HIGHEST slope...",
                           wildcard="Text files (*.txt)|*.txt|" "Comma Separated Value-files (*.csv)|*.csv",
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as prompted_dialog:

            if prompted_dialog.ShowModal() == wx.ID_CANCEL:
                return
            self.pathname_1 = prompted_dialog.GetPath()

        # Opening File 2 with the use of a dialog. File 2 will contain the ANT+ data of the measurements with a low
        # (negative) gradient (slope). This will be used to calculate the minimal brake power.
        with wx.FileDialog(self, "Choose the second logged SimulANT+ file with the LOWEST (negative) slope...",
                           wildcard="Text files (*.txt)|*.txt|" "Comma Separated Value-files (*.csv)|*.csv",
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as prompted_dialog:

            if prompted_dialog.ShowModal() == wx.ID_CANCEL:
                return
            self.pathname_2 = prompted_dialog.GetPath()

        # Opening File 3 with the use of a dialog. File 3 will contain the ANT+ data of the measurements with a power
        # goal of 0W, this will be used to see the residual brake power if no brake is used.
        with wx.FileDialog(self, "Choose the third logged SimulANT+ file with the 0 W Power program - low acceleration...",
                           wildcard="Text files (*.txt)|*.txt|" "Comma Separated Value-files (*.csv)|*.csv",
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as prompted_dialog:

            if prompted_dialog.ShowModal() == wx.ID_CANCEL:
                return
            self.pathname_3 = prompted_dialog.GetPath()

        self.folder_pathname = os.path.dirname(self.pathname_3)
        # Opening File 4 with the use of a dialog. File 4 will contain the ANT+ data of the measurements with a power
        # goal of 0W. In contrary with the other files, the acceleration needs to be high. This way, it is possible to
        # calculate the simulated mass (inertia).
        with wx.FileDialog(self, "Choose the third logged SimulANT+ file with the 0 W Power program - high acceleration...",
                           wildcard="Text files (*.txt)|*.txt|" "Comma Separated Value-files (*.csv)|*.csv",
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as prompted_dialog:

            if prompted_dialog.ShowModal() == wx.ID_CANCEL:
                return
            self.pathname_4 = prompted_dialog.GetPath()

        self.folder_pathname = os.path.dirname(self.pathname_4)
        # Retrieving the filename the user wants to use. This will be the file name of the new excel file which will be
        # created after running this program.
        self.user_file_name_dialog = wx.TextEntryDialog(self,
                                                        "What do you want the .xslx file to be named? Enter here: ",
                                                        "Enter file name...")
        self.user_file_name_dialog.CenterOnParent()

        if self.user_file_name_dialog.ShowModal() == wx.ID_CANCEL:
            return
        self.user_file_name = self.user_file_name_dialog.GetValue()

        # Analyse the log-files. This will be used to retrieve the data from the four selected log files above. This
        # will be done by ANTlogfileconverter.py. Raw data will be stored in data_.... and used in further calculations.
        # Analysing log file 1:
        self.logfile_analyser(self.pathname_1)
        data_high = []
        max_velocity_high = max(velocity_list)
        if len(power_list) < len(velocity_list):
            velocity_list.pop()
        elif len(velocity_list) < len(power_list):
            power_list.pop()
        elif len(velocity_list) < len(time_list):
            time_list.pop()
        else:
            pass

        for i in range(len(velocity_list)):
            data_high.append([velocity_list[i], power_list[i], time_list[i]])

        self.velocity_list_high = velocity_list
        self.power_list_high = power_list
        self.time_list_high = time_list

        # Analysing log file 2:
        self.logfile_analyser(self.pathname_2)
        data_low = []
        max_velocity_low = max(velocity_list)
        if len(power_list) < len(velocity_list):
            velocity_list.pop()
        elif len(velocity_list) < len(power_list):
            power_list.pop()
        elif len(velocity_list) < len(time_list):
            time_list.pop()
        else:
            pass

        for i in range(len(velocity_list)):
            data_low.append([velocity_list[i], power_list[i], time_list[i]])
        self.velocity_list_low = velocity_list
        self.power_list_low = power_list
        self.time_list_low = time_list

        # Analysing log file 3:
        self.logfile_analyser(self.pathname_3)
        data_zero = []
        max_velocity_zero = max(velocity_list)
        if len(power_list) < len(velocity_list):
            velocity_list.pop()
        elif len(velocity_list) < len(power_list):
            power_list.pop()
        elif len(velocity_list) < len(time_list):
            time_list.pop()
        else:
            pass

        for i in range(len(velocity_list)):
            data_zero.append([velocity_list[i], power_list[i], time_list[i]])

        self.power_list_zero = power_list
        self.time_list_zero = time_list
        self.velocity_list_zero_acc = velocity_list

        # Analysing log file 4:
        self.logfile_analyser(self.pathname_4)
        data_zero_acc = []
        max_velocity_zero_acc = max(velocity_list)
        if len(power_list) < len(velocity_list):
            velocity_list.pop()
        elif len(velocity_list) < len(power_list):
            power_list.pop()
        elif len(velocity_list) < len(time_list):
            time_list.pop()
        else:
            pass

        for i in range(len(velocity_list)):
            data_zero_acc.append([velocity_list[i], power_list[i], time_list[i]])

        self.velocity_list_zero = velocity_list
        self.power_list_zero_acc = power_list
        self.time_list_zero_acc = time_list

        # Calculating the averages of every file, this is not necessary for the calculations below, but this will give
        # a quick overview of the used files to the user.
        self.velocity_high_avg = mean(self.velocity_list_high)
        self.velocity_high_avg = round(float(self.velocity_high_avg), 1)
        self.power_high_avg = mean(self.power_list_high)
        self.power_high_avg = round(float(self.power_high_avg), 1)
        self.velocity_low_avg = mean(self.velocity_list_low)
        self.velocity_low_avg = round(float(self.velocity_low_avg), 1)
        self.power_low_avg = mean(self.power_list_low)
        self.power_low_avg = round(float(self.power_low_avg), 1)
        self.velocity_zero_avg = mean(self.velocity_list_zero)
        self.velocity_zero_avg = round(float(self.velocity_zero_avg), 1)
        self.power_zero_avg = mean(self.power_list_zero)
        self.power_zero_avg = round(float(self.power_zero_avg), 1)
        self.velocity_zero_acc_avg = mean(self.velocity_list_zero_acc)
        self.velocity_zero_acc_avg = round(float(self.velocity_zero_acc_avg), 1)
        self.power_zero_acc_avg = mean(self.power_list_zero_acc)
        self.power_zero_acc_avg = round(float(self.power_zero_acc_avg), 1)

        velocity_raw_high = []
        power_raw_high = []
        time_raw_high = []
        velocity_raw_low = []
        power_raw_low = []
        time_raw_low = []
        power_clean_high = []
        velocity_clean_high = []
        time_clean_high = []
        power_clean_low = []
        velocity_clean_low = []
        time_clean_low = []
        velocity_zero = []
        power_zero = []
        time_clean_zero = []
        velocity_zero_acc = []
        power_zero_acc = []
        time_clean_zero_acc = []

        """
        Convert the raw data from the file to named lists for the FIRST file
        """
        for j in range(len(data_high)):
            if round(data_high[j][0]) == 0:
                power_raw_high.append(0)
                velocity_raw_high.append(0)
                time_raw_high.append(data_high[j][2] * 2)
            else:
                power_raw_high.append(data_high[j][1])
                velocity_raw_high.append(data_high[j][0])
                time_raw_high.append(data_high[j][2] * 2)

        for i in range(power_raw_high.index(max(power_raw_high))):
            power_clean_high.append(power_raw_high[i])
            velocity_clean_high.append(velocity_raw_high[i])
            time_clean_high.append(time_raw_high[i])

        """
        Make a fit for the data of the FIRST file 
        """
        error_lin_high = 0
        error_quadratic_high = 0
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
            '1' : error_lin_high,
            '2' : error_quadratic_high
        }
        lowest_error = min(errors, key=errors.get)
        if lowest_error == '1':
            fitted_power_high = fitted_power_high_1
        elif lowest_error == '2':
            fitted_power_high = fitted_power_high_2

        """
        Convert the raw data from the file to named lists for the SECOND file
        """
        for j in range(len(data_low)):
            if round(data_low[j][0]) == 0:
                power_raw_low.append(0)
                velocity_raw_low.append(0)
                time_raw_low.append(data_low[j][2] * 2)

            else:
                power_raw_low.append(data_low[j][1])
                velocity_raw_low.append(data_low[j][0])
                time_raw_low.append(data_low[j][2] * 2)

        for i in range(power_raw_low.index(max(power_raw_low))):
            power_clean_low.append(power_raw_low[i])
            velocity_clean_low.append(velocity_raw_low[i])
            time_clean_low.append(time_raw_low[i])

        """
        Make a fit for the data of the SECOND file 
        """
        error_lin_low = 0
        error_quadratic_low = 0
        popt1_low, pcov = curve_fit(self.func_lin, array(velocity_clean_low), array(power_clean_low))
        fitted_power_low_1 = self.func_lin(array(velocity_clean_low), *popt1_low)
        for i in range(len(fitted_power_low_1)):
            if fitted_power_low_1[i] < 0:
                fitted_power_low_1[i] = 0
        for i in range(len(fitted_power_low_1)):
            if fitted_power_low_1[i] > 0:
                error_lin_low += abs(fitted_power_low_1[i] - power_clean_low[i])

        popt2_low, pcov = curve_fit(self.func_quadratic, array(velocity_clean_low), array(power_clean_low))
        fitted_power_low_2 = self.func_quadratic(array(velocity_clean_low), *popt2_low)
        for i in range(len(fitted_power_low_2)):
            if fitted_power_low_2[i] < 0:
                fitted_power_low_2[i] = 0
        for i in range(len(fitted_power_low_2)):
            if fitted_power_low_2[i] > 0:
                error_quadratic_low += abs(fitted_power_low_2[i] - power_clean_low[i])

        errors = {
            '1' : error_lin_low,
            '2' : error_quadratic_low
        }
        lowest_error = min(errors, key=errors.get)
        if lowest_error == '1':
            fitted_power_low = fitted_power_low_1
        elif lowest_error == '2':
            fitted_power_low = fitted_power_low_2

        """
        Convert the raw data from the file to named lists for the THIRD file
        """
        time_raw_zero = []
        for j in range(len(data_zero)):
            if round(data_zero[j][0]) == 0:
                power_zero.append(0)
                velocity_zero.append(0)
                time_raw_zero.append(data_zero[j][2] * 2)
            else:
                power_zero.append(data_zero[j][1])
                velocity_zero.append(data_zero[j][0])
                time_raw_zero.append(data_zero[j][2] * 2)

        """
        Convert the raw data from the file to named lists for the FOURTH file
        """
        for j in range(len(data_zero_acc)):
            if round(data_zero_acc[j][0]) == 0:
                power_zero_acc.append(0)
                velocity_zero_acc.append(0)
                time_clean_zero_acc.append(data_zero_acc[j][2] * 2)
            else:
                power_zero_acc.append(data_zero_acc[j][1])
                velocity_zero_acc.append(data_zero_acc[j][0])
                time_clean_zero_acc.append(data_zero_acc[j][2] * 2)

        """
        Start calculations on the THIRD AND FOURTH file to calculate the SIMULATED MASS.
        This includes fitting the data. 
        """
        power_clean_zero = []
        velocity_clean_zero = []
        power_clean_zero_acc = []
        velocity_clean_zero_acc = []
        power_clean_zero_acc_dummy = []
        velocity_clean_zero_acc_dummy = []

        for i in range(power_zero.index(max(power_zero))):
            power_clean_zero.append(power_zero[i])
            velocity_clean_zero.append(velocity_zero[i])
            time_clean_zero.append(time_raw_zero[i])
        for i in range(power_zero_acc.index(max(power_zero_acc))):
            power_clean_zero_acc_dummy.append(power_zero_acc[i])
        for i in range(power_zero_acc.index(max(power_zero_acc))):
            velocity_clean_zero_acc_dummy.append(velocity_zero_acc[i])

        for i in range(len(power_clean_zero_acc_dummy)):
            if power_clean_zero_acc_dummy[i] > 50:
                power_clean_zero_acc.append(power_clean_zero_acc_dummy[i])
                velocity_clean_zero_acc.append(velocity_clean_zero_acc_dummy[i])

        # Calculating the best possible fit, we only consider quadratic and linear fits at this moment. The error with
        # the original data is calculated and the best fit will be drawn. A dictionary is used to track the variable
        # with the highest value without a big if statement structure.
        error_lin = 0
        error_quadratic = 0
        popt1, pcov = curve_fit(self.func_lin, array(velocity_clean_zero), array(power_clean_zero))
        fitted_power_zero_1 = self.func_lin(array(velocity_clean_zero), *popt1)
        for i in range(len(fitted_power_zero_1)):
            if fitted_power_zero_1[i] < 0:
                fitted_power_zero_1[i] = 0
        for i in range(len(fitted_power_zero_1)):
            if fitted_power_zero_1[i] > 0:
                error_lin += abs(fitted_power_zero_1[i] - power_clean_zero[i])

        popt2, pcov = curve_fit(self.func_quadratic, array(velocity_clean_zero), array(power_clean_zero))
        fitted_power_zero_2 = self.func_quadratic(array(velocity_clean_zero), *popt2)
        for i in range(len(fitted_power_zero_2)):
            if fitted_power_zero_2[i] < 0:
                fitted_power_zero_2[i] = 0
        for i in range(len(fitted_power_zero_2)):
            if fitted_power_zero_2[i] > 0:
                error_quadratic += abs(fitted_power_zero_2[i] - power_clean_zero[i])

        errors = {
            '1' : error_lin,
            '2' : error_quadratic
        }
        lowest_error = min(errors, key=errors.get)

        power_compensated = []
        if lowest_error == '1':
            fitted_power_zero = fitted_power_zero_1
            for i in range(len(velocity_clean_zero_acc)):
                power_to_substract = popt1[0] * velocity_clean_zero_acc[i] + popt1[1]
                power_compensated.append(power_clean_zero_acc[i] - power_to_substract)

        elif lowest_error == '2':
            fitted_power_zero = fitted_power_zero_2
            for i in range(len(velocity_clean_zero_acc)):
                power_to_substract = popt2[0] * velocity_clean_zero_acc[i] ** 2 + popt2[1] * velocity_clean_zero_acc[i] + popt2[2]
                power_compensated.append(power_clean_zero_acc[i] - power_to_substract)

        popt3, pcov = curve_fit(self.func_lin, array(velocity_clean_zero_acc) / 3.6, array(power_compensated))
        fitted_power_zero_acc = self.func_lin(array(velocity_clean_zero_acc) / 3.6, *popt3)

        for i in range(len(time_clean_zero_acc)):
            if len(velocity_clean_zero_acc) < len(time_clean_zero_acc):
                time_clean_zero_acc.pop()

        popt4, pcov = curve_fit(self.func_lin, array(time_clean_zero_acc), array(velocity_clean_zero_acc) / 3.6)
        fitted_velocity_zero_acc = self.func_lin(array(time_clean_zero_acc), *popt4)

        self.simulated_mass = []
        for i in range(len(fitted_velocity_zero_acc)):
            self.simulated_mass.append(popt3[0] / popt4[0])

        """
        Calculation of the power which is needed to accelerate the flywheel. For the first and second file.
        """
        self.simulated_mass = 22
        power_flywheel_high = []
        power_flywheel_low = []
        power_flywheel_zero = []
        power_clean_low_brake = []
        power_clean_high_brake = []
        power_clean_zero_brake = []

        # Acceleration needs to be constant to use this.
        popt5, pcov = curve_fit(self.func_lin, array(time_clean_high), array(velocity_clean_high) / 3.6)
        popt6, pcov = curve_fit(self.func_lin, array(time_clean_low), array(velocity_clean_low) / 3.6)
        popt7, pcov = curve_fit(self.func_lin, array(time_clean_zero), array(velocity_clean_zero) / 3.6)

        for i in range(len(velocity_clean_high)):
            power_flywheel_high.append(self.simulated_mass * velocity_clean_high[i]/3.6 * popt5[0])
            power_clean_high_brake.append(fitted_power_high[i] - power_flywheel_high[i])

        for i in range(len(velocity_clean_low)):
            power_flywheel_low.append(self.simulated_mass * velocity_clean_low[i]/3.6 * popt6[0])
            power_clean_low_brake.append(fitted_power_low[i] - power_flywheel_low[i])

        for i in range(len(velocity_clean_zero)):
            power_flywheel_zero.append(self.simulated_mass * velocity_clean_zero[i]/3.6 * popt7[0])
            power_clean_zero_brake.append(fitted_power_zero[i] - power_flywheel_zero[i])

        """
        Initialize writing an excel file.
        """
        excel = xlsxwriter.Workbook(self.user_file_name + ".xlsx")
        graph = excel.add_chart({'type': 'scatter', 'subtype': 'straight'})
        graph_2 = excel.add_chart({'type': 'scatter', 'subtype': 'straight'})
        worksheet_charts = excel.add_worksheet('Charts')
        worksheet_data = excel.add_worksheet('Data')

        """
        Setting variables for excel file.
        """
        bold = excel.add_format({'bold': True})
        underline = excel.add_format({'bold': True, 'underline': True})
        graph.set_y_axis({'name': 'Power [W]'})
        graph.set_x_axis({'name': 'Velocity [km/h]'})
        graph.set_title({'name': 'Operating range ' + self.user_file_name})
        graph.set_size({'width': 1080, 'height': 720})
        graph_2.set_y_axis({'name': 'Power [W]'})
        graph_2.set_y2_axis({'name': 'Velocity [km/h]'})
        graph_2.set_x_axis({'name': 'Time [s]'})
        graph_2.set_title({'name': '0 Watt acceleration ' + self.user_file_name})
        graph_2.set_size({'width': 1080, 'height': 720})
        worksheet_data.set_column('A:Q', 14)

        """
        Writing to excel file.
        """
        worksheet_data.write('A1', 'Tested with highest gradient (without slip)', underline)
        worksheet_data.write('A2', 'Velocity [km/h]', bold)
        worksheet_data.write('B2', 'Power [W]', bold)
        worksheet_data.write('C2', 'Fitted Power [W]', bold)
        worksheet_data.write_column(2, 0, velocity_clean_high)
        worksheet_data.write_column(2, 1, power_clean_high)
        worksheet_data.write_column(2, 2, fitted_power_high)

        worksheet_data.write('D1', 'Tested with lowest gradient (without slip)', underline)
        worksheet_data.write('D2', 'Velocity [km/h]', bold)
        worksheet_data.write('E2', 'Power [W]', bold)
        worksheet_data.write('F2', 'Fitted Power [W]', bold)
        worksheet_data.write_column(2, 3, velocity_clean_low)
        worksheet_data.write_column(2, 4, power_clean_low)
        worksheet_data.write_column(2, 5, fitted_power_low)

        worksheet_data.write('G1', 'Tested with 0 W program - low acceleration (without slip)', underline)
        worksheet_data.write('G2', 'Time [s]', bold)
        worksheet_data.write('H2', 'Power [W]', bold)
        worksheet_data.write('I2', 'Fitted Power [W]', bold)
        worksheet_data.write('J2', 'Velocity [km/h]', bold)
        worksheet_data.write_column(2, 6, time_clean_zero)
        worksheet_data.write_column(2, 7, power_clean_zero)
        worksheet_data.write_column(2, 8, fitted_power_zero)
        worksheet_data.write_column(2, 9, velocity_clean_zero)

        worksheet_data.write('L1', 'Tested with 0 W program - high acceleration (without slip)', underline)
        worksheet_data.write('L2', 'Time [s]', bold)
        worksheet_data.write('M2', 'Power [km/h]', bold)
        worksheet_data.write('N2', 'Velocity [km/h]', bold)
        worksheet_data.write('O2', 'Fitted Velocity [km/h]', bold)
        worksheet_data.write('P2', 'Theor. Power [W]', bold)
        worksheet_data.write('Q2', 'Fitted Compensated Power [W]', bold)
        worksheet_data.write_column(2, 11, time_clean_zero_acc)
        worksheet_data.write_column(2, 12, power_clean_zero_acc)
        worksheet_data.write_column(2, 13, velocity_clean_zero_acc)
        worksheet_data.write_column(2, 14, fitted_velocity_zero_acc)
        worksheet_data.write_column(2, 15, power_compensated)
        worksheet_data.write_column(2, 16, fitted_power_zero_acc)

        worksheet_data.write('S2', 'Brake Power Trainer Upper Limit [W]', bold)
        worksheet_data.write('T2', 'Brake Power Trainer Lower Limit [W]', bold)
        worksheet_data.write_column(2, 18, power_clean_high_brake)
        worksheet_data.write_column(2, 19, power_clean_low_brake)

        """
        Writing to graph.
        """
        graph.add_series({
            'categories': [worksheet_data.name] + [2, 3] + [len(velocity_clean_low) + 2, 3],
            'values': [worksheet_data.name] + [2, 4] + [len(power_clean_low) + 2, 4],
            'line': {'color': '#67bfe7', 'dash_type': 'dash'},
            'name': 'Lowest Gradient Power',
        })

        graph.add_series({
            'categories': [worksheet_data.name] + [2, 3] + [len(velocity_clean_low) + 2, 3],
            'values': [worksheet_data.name] + [2, 5] + [len(power_clean_low) + 2, 5],
            'line': {'color': '#67bfe7'},
            'name': 'Lowest Gradient Power',
        })

        graph.add_series({
            'categories': [worksheet_data.name] + [2, 3] + [len(velocity_clean_low) + 2, 3],
            'values': [worksheet_data.name] + [2, 19] + [len(power_clean_low) + 2, 19],
            'line': {'color': 'black'},
            'name': 'Lowest Gradient Power',
        })

        graph.add_series({
            'categories': [worksheet_data.name] + [2, 0] + [len(velocity_clean_high) + 2, 0],
            'values': [worksheet_data.name] + [2, 1] + [len(power_clean_high) + 2, 1],
            'line': {'color': '#67bfe7', 'dash_type': 'dash'},
            'name': 'Highest Gradient Power',
        })

        graph.add_series({
            'categories': [worksheet_data.name] + [2, 0] + [len(velocity_clean_high) + 2, 0],
            'values': [worksheet_data.name] + [2, 2] + [len(power_clean_high) + 2, 2],
            'line': {'color': '#67bfe7'},
            'name': 'Highest Gradient Power',
        })

        graph.add_series({
            'categories': [worksheet_data.name] + [2, 0] + [len(velocity_clean_high) + 2, 0],
            'values': [worksheet_data.name] + [2, 18] + [len(power_clean_high) + 2, 18],
            'line': {'color': 'black'},
            'name': 'Highest Gradient Power',
        })
        # graph.add_series({
        #     'categories': [worksheet_data.name] + [2, 13] + [len(velocity_clean_zero_acc) + 2, 13],
        #     'values': [worksheet_data.name] + [2, 12] + [len(power_clean_zero_acc) + 2, 12],
        #     'line': {'color': 'red'},
        #     'name': 'constant power',
        # })
        # graph.add_series({
        #     'categories': [worksheet_data.name] + [2, 13] + [len(velocity_clean_zero_acc) + 2, 13],
        #     'values': [worksheet_data.name] + [2, 15] + [len(power_compensated) + 2, 15],
        #     'line': {'color': 'black'},
        #     'name': 'Power vs. Velocity',
        # })
        # graph.add_series({
        #     'categories': [worksheet_data.name] + [2, 13] + [len(velocity_clean_zero_acc) + 2, 13],
        #     'values': [worksheet_data.name] + [2, 16] + [len(fitted_power_zero_acc) + 2, 16],
        #     'line': {'color': 'purple'},
        #     'name': 'Fitted Compensated Power',
        # })
        #
        # graph.add_series({
        #     'categories': [worksheet_data.name] + [2, 9] + [len(velocity_clean_zero) + 2, 9],
        #         'values': [worksheet_data.name] + [2, 8] + [len(fitted_power_zero) + 2, 8],
        #     'line': {'color': 'red'},
        #     'name': 'Power Zero',
        # })

        graph_2.add_series({
            'categories': [worksheet_data.name] + [2, 6] + [len(time_clean_zero) + 2, 6],
            'values': [worksheet_data.name] + [2, 7] + [len(power_clean_zero) + 2, 7],
            'line': {'color': '#ff0000'},
            'name': '0 W Program Low Acceleration Power',
        })

        graph_2.add_series({
            'categories': [worksheet_data.name] + [2, 6] + [len(time_clean_zero) + 2, 6],
            'values': [worksheet_data.name] + [2, 9] + [len(velocity_clean_zero) + 2, 9],
            'line': {'color': '#0000ff', 'dash_type': 'dash'},
            'name': '0 W Program Low Acceleration Velocity',
            'y2_axis': True,
        })

        graph_2.add_series({
            'categories': [worksheet_data.name] + [2, 11] + [len(time_clean_zero_acc) + 2, 11],
            'values': [worksheet_data.name] + [2, 13] + [len(velocity_clean_zero_acc) + 2, 13],
            'line': {'color': '#0000ff', 'dash_type': 'dash'},
            'name': '0 W Program High Acceleration Velocity',
            'y2_axis': True,
        })

        graph_2.add_series({
            'categories': [worksheet_data.name] + [2, 11] + [len(time_clean_zero_acc) + 2, 11],
            'values': [worksheet_data.name] + [2, 14] + [len(fitted_velocity_zero_acc) + 2, 14],
            'line': {'color': '#0000ff'},
            'name': '0 W Program Fitted Acceleration Velocity',
            'y2_axis': True,
        })

        graph_2.add_series({
            'categories': [worksheet_data.name] + [2, 11] + [len(time_clean_zero_acc) + 2, 11],
            'values': [worksheet_data.name] + [2, 15] + [len(power_compensated) + 2, 15],
            'line': {'color': '#ff0000'},
            'name': '0 W Program High Acceleration Power Compensated',
        })

        worksheet_charts.insert_chart('B2', graph)
        worksheet_charts.insert_chart('B40', graph_2)
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
        global velocity_list, power_list, time_list
        sentences = []
        value_list = []
        time_list = []
        time_values_raw = []
        velocity_list = []
        power_list = []
        speed = True
        power = True

        """
        The log-file is opened here. 
        """
        log = open(logfile)
        with open(logfile) as f:
            for lines, l in enumerate(f):
                pass

        """
        The important lines will be retrieved from the log-file here, by looking at the lines which start with 'Rx'. These are received messages. 
        """
        for n in range(lines):
            sentence = log.readline()
            if "Rx:" in sentence:
                sentences.append(sentence)

        """
        This part splits the retrieved lines in subparts, after which the hexadecimals will be read. 
        """
        for i in range(len(sentences)):
            sentence = sentences[i].split()
            index = sentence.index("Rx:")
            value_raw = sentence[index + 1]
            time_raw = sentence[index - 2]
            value = value_raw.replace("[", "").replace("]", "")  # This will removes the useless characters
            value_list.append(value)
            time_values_raw.append(time_raw)

        """
        These subparts will be categorized according to their first character: When this is '10', this means the velocity is recorded in that line. When the first character is '19', this means power is recorded in that line. The other characters are not important for the functionality of this file, which means they will be left out.
        """
        for i in range(len(value_list)):
            value_list_characters = list(value_list[i])

            time_list.append((float(time_values_raw[i]) - float(time_values_raw[0])) / 1000)

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

    def on_xlsx_button(self, event):
        if os.path.isfile(self.folder_pathname + "\\" + self.user_file_name + ".xlsx"):
            os.startfile(self.folder_pathname + "\\" + self.user_file_name + ".xlsx")
        elif self.folder_pathname == "":
            no_file_dialog = wx.MessageDialog(self.top_panel,
                                              message="The file does not exist. Please try selecting files in File -> Open files...",
                                              caption="Warning!")
            no_file_dialog.CenterOnParent()
            if no_file_dialog.ShowModal() == wx.OK:
                no_file_dialog.Destroy()
                return

    def func_powerlaw(self, x, m, c):
        return x ** m * c

    def func_quadratic(self, x, a, b, c):
        return a * x**2 + b*x + c

    def func_lin(self, x, a, b):
        return a * x + b

    def on_check_sim_mass(self, event):
        if self.checkbox_sim.GetValue():
            self.sim_mass_dialog = wx.TextEntryDialog(self,
                                                            "What is the value for the simulated mass [kg] (use '.' as decimal separator): ",
                                                            "Enter simulated mass value...")
            self.sim_mass_dialog.CenterOnParent()

            if self.sim_mass_dialog.ShowModal() == wx.ID_CANCEL:
                self.checkbox_sim.SetValue(False)
                return
            self.simulated_mass = float(self.sim_mass_dialog.GetValue())

    def on_check_inertia(self, event):
        if self.checkbox_inertia.GetValue():
            self.inertia_dialog = wx.TextEntryDialog(self,
                                                            "What is the value for the moment of inertia [kg * m^2] (use '.' as decimal separator): ",
                                                            "Enter inertia value...")
            self.inertia_dialog.CenterOnParent()

            if self.inertia_dialog.ShowModal() == wx.ID_CANCEL:
                self.checkbox_inertia.SetValue(False)
                return

            self.conversion_dialog = wx.TextEntryDialog(self,
                                                            "What is the value for the conversion factor (use '.' as decimal separator, see README.txt for explanation): ",
                                                            "Enter conversion value...")
            self.conversion_dialog.CenterOnParent()

            if self.conversion_dialog.ShowModal() == wx.ID_CANCEL:
                self.checkbox_inertia.SetValue(False)
                return

            self.inertia = float(self.inertia_dialog.GetValue())
            self.conversion = float(self.conversion_dialog.GetValue())

            self.simulated_mass = (self.conversion ** 2) * self.inertia

if __name__ == '__main__':
    Application = wx.App(False)
    frame = Main(None, 'SimulANT+ Log Analyzer').Show()
    Application.MainLoop()
