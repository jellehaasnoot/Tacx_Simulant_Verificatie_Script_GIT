import wx as wx
from ValueConverter import ValueConverter
import xlsxwriter
import os
import numpy as np
from scipy import signal
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
                          style=wx.DEFAULT_FRAME_STYLE & ~(wx.RESIZE_BORDER | wx.MAXIMIZE_BOX), size=(720, 790))
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
        self.exit_button = wx.Button(self.top_panel, -1, label='Exit', pos=(590, 665), size=(100, 30))
        self.reset_button = wx.Button(self.top_panel, -1, label='Reset Program', pos=(480, 665), size=(100, 30))
        self.open_xlsx_button = wx.Button(self.top_panel, -1, label='Open Excel File', pos=(370, 665), size=(100, 30))
        self.open_files_butten = wx.Button(self.top_panel, -1, label='Open LOG\'s', pos=(260, 665), size=(100, 30))

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

        self.some_data_panel_1 = wx.Panel(self.top_panel, -1, style=wx.SUNKEN_BORDER, size=(680, 80), pos=(10, 270))
        self.data_panel_1_header_display = wx.StaticText(self.some_data_panel_1,
                                                         label="Some statistics about the first file: ", pos=(4, 2))
        self.data_panel_1_header_display.SetFont(self.font_header)

        self.some_data_panel_2 = wx.Panel(self.top_panel, -1, style=wx.SUNKEN_BORDER, size=(680, 80), pos=(10, 360))
        self.data_panel_2_header_display = wx.StaticText(self.some_data_panel_2,
                                                         label="Some statistics about the second file: ", pos=(4, 2))
        self.data_panel_2_header_display.SetFont(self.font_header)

        self.some_data_panel_3 = wx.Panel(self.top_panel, -1, style=wx.SUNKEN_BORDER, size=(680, 80), pos=(10, 450))
        self.data_panel_3_header_display = wx.StaticText(self.some_data_panel_3,
                                                         label="Some statistics about the third file: ", pos=(4, 2))
        self.data_panel_3_header_display.SetFont(self.font_header)
        
        self.some_data_panel_4 = wx.Panel(self.top_panel, -1, style=wx.SUNKEN_BORDER, size=(680, 80), pos=(10, 540))
        self.data_panel_4_header_display = wx.StaticText(self.some_data_panel_4,
                                                         label="Some statistics about the fourth file: ", pos=(4, 2))
        self.data_panel_4_header_display.SetFont(self.font_header)
        # self.image_panel = wx.Panel(self.top_panel, -1, style=wx.BORDER_SIMPLE, size=())

        # Set events
        self.Bind(wx.EVT_MENU, self.on_open, menu_file_open)
        self.Bind(wx.EVT_MENU, self.on_about, menu_about)
        self.Bind(wx.EVT_MENU, self.on_exit, menu_exit)
        self.exit_button.Bind(wx.EVT_BUTTON, self.on_exit_button)
        self.reset_button.Bind(wx.EVT_BUTTON, self.on_reset)
        self.open_xlsx_button.Bind(wx.EVT_BUTTON, self.on_xlsx_button)
        self.open_files_butten.Bind(wx.EVT_BUTTON, self.on_open)

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
                                                      self.power_zero_avg) + " W\n""Average velocity at 0 Watt programming:   " + str(
                                                      self.velocity_zero_avg) + " km/h\n""Amount of received ANT+ messages:   " + str(
                                                      len(self.velocity_list_zero)), pos=(4, 24))
        self.data_panel_4_display.SetFont(self.font_normal)

        xlsx_path_panel = wx.Panel(self.top_panel, -1, style=wx.SUNKEN_BORDER, size=(680, 50), pos=(10, 620))
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
        self.directory_name_1 = ""

        with wx.FileDialog(self, "Choose the logged SimulANT+ file with the HIGHEST slope...",
                           wildcard="Text files (*.txt)|*.txt|" "Comma Separated Value-files (*.csv)|*.csv",
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as prompted_dialog:

            if prompted_dialog.ShowModal() == wx.ID_CANCEL:
                return

            self.pathname_1 = prompted_dialog.GetPath()


        # Opening File 2 with the use of a dialog. File 2 will contain the ANT+ data of the measurements with a low
        # (negative) gradient (slope). This will be used to calculate the minimal brake power.
        self.directory_name_2 = ""

        with wx.FileDialog(self, "Choose the second logged SimulANT+ file with the LOWEST (negative) slope...",
                           wildcard="Text files (*.txt)|*.txt|" "Comma Separated Value-files (*.csv)|*.csv",
                           style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST) as prompted_dialog:

            if prompted_dialog.ShowModal() == wx.ID_CANCEL:
                return

            self.pathname_2 = prompted_dialog.GetPath()


        # Opening File 3 with the use of a dialog. File 3 will contain the ANT+ data of the measurements with a power
        # goal of 0W, this will be used to see the residual brake power if no brake is used.
        self.directory_name_3 = ""

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
        self.directory_name_4 = ""

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
        self.velocity_high_avg = np.mean(self.velocity_list_high)
        self.velocity_high_avg = round(float(self.velocity_high_avg), 1)
        self.power_high_avg = np.mean(self.power_list_high)
        self.power_high_avg = round(float(self.power_high_avg), 1)
        self.velocity_low_avg = np.mean(self.velocity_list_low)
        self.velocity_low_avg = round(float(self.velocity_low_avg), 1)
        self.power_low_avg = np.mean(self.power_list_low)
        self.power_low_avg = round(float(self.power_low_avg), 1)
        self.velocity_zero_avg = np.mean(self.velocity_list_zero)
        self.velocity_zero_avg = round(float(self.velocity_zero_avg), 1)
        self.power_zero_avg = np.mean(self.power_list_zero)
        self.power_zero_avg = round(float(self.power_zero_avg), 1)
        self.velocity_zero_acc_avg = np.mean(self.velocity_list_zero_acc)
        self.velocity_zero_acc_avg = round(float(self.velocity_zero_acc_avg), 1)
        self.power_zero_acc_avg = np.mean(self.power_list_zero_acc)
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
        power_zero_coefficients = []
        power_zero_scalars = []
        energy_zero = []
        simulated_mass = []

        vertical_x = []
        vertical_y = []
        horizontal_x = []
        horizontal_y = []
        horizontal_2_x = []
        horizontal_2_y = []

        # Raw data will be converted in the first file.
        for i in range(round(max_velocity_high)):
            dummy_velocity = []
            dummy_power = []
            dummy_time = []
            for j in range(len(data_high) - 1):
                if i <= data_high[j][0] < i + 1:
                    dummy_velocity.append(data_high[j][0])
                    dummy_power.append(data_high[j][1])
                    dummy_time.append(data_high[j][2])
                else:
                    pass
            if len(dummy_velocity) > 0 and len(dummy_power) > 0 and len(dummy_time) > 0:
                index = dummy_power.index(max(dummy_power))
                velocity_raw_high.append(dummy_velocity[index])
                power_raw_high.append(dummy_power[index])
                time_raw_high.append(dummy_time[index])
            else:
                pass

        for i in range(len(power_raw_high)):
            if round(velocity_raw_high[i]) == 0:
                power_clean_high.append(0)
                velocity_clean_high.append(0)
                time_clean_high.append(time_raw_high[i])
            elif power_raw_high[i] - power_raw_high[i - 1] < -15:
                pass
            else:
                power_clean_high.append(power_raw_high[i])
                velocity_clean_high.append(velocity_raw_high[i])
                time_clean_high.append(time_raw_high[i])

        """
        Convert the raw data from the file to named lists for the SECOND file
        """

        for i in range(round(max_velocity_low)):
            dummy_velocity = []
            dummy_power = []
            dummy_time = []
            for j in range(len(data_low) - 1):
                if i <= data_low[j][0] < i + 1:
                    dummy_velocity.append(data_low[j][0])
                    dummy_power.append(data_low[j][1])
                    dummy_time.append(data_low[j][2])
                else:
                    pass
            if len(dummy_velocity) > 0 and len(dummy_power) > 0 and len(dummy_time) > 0:
                index = dummy_power.index(max(dummy_power))
                velocity_raw_low.append(dummy_velocity[index])
                power_raw_low.append(dummy_power[index])
                time_raw_low.append(dummy_time[index])
            else:
                pass

        for i in range(len(power_raw_low)):
            if round(velocity_raw_low[i]) == 0:
                power_clean_low.append(0)
                velocity_clean_low.append(0)
                time_clean_low.append(time_raw_low[i])
            elif power_raw_low[i] - power_raw_low[i - 1] < -15:
                pass
            else:
                power_clean_low.append(power_raw_low[i])
                velocity_clean_low.append(velocity_raw_low[i])
                time_clean_low.append(time_raw_low[i])

        """
        Convert the raw data from the file to named lists for the THIRD file
        """

        for j in range(len(data_zero)):
            if round(data_zero[j][0]) == 0:
                power_zero.append(0)
                velocity_zero.append(0)
                time_clean_zero.append(data_zero[j][2])
            else:
                power_zero.append(data_zero[j][1])
                velocity_zero.append(data_zero[j][0])
                time_clean_zero.append(data_zero[j][2])
        max_velocity_zero_index = velocity_zero.index(max(velocity_zero))



        """
        Convert the raw data from the file to named lists for the FOURTH file
        """
        for j in range(len(data_zero_acc)):
            if round(data_zero_acc[j][0]) == 0:
                power_zero_acc.append(0)
                velocity_zero_acc.append(0)
                time_clean_zero_acc.append(data_zero_acc[j][2])
            else:
                power_zero_acc.append(data_zero_acc[j][1])
                velocity_zero_acc.append(data_zero_acc[j][0])
                time_clean_zero_acc.append(data_zero_acc[j][2])
        max_velocity_zero_acc_index = velocity_zero_acc.index(max(velocity_zero_acc))

        """
        Start calculations on the THIRD AND FOURTH file to calculate the SIMULATED MASS.
        This includes fitting the data. 
        """
        power_clean_zero = []
        for i in range(max_velocity_zero_index):
            power_clean_zero.append(power_zero[i])
        velocity_clean_zero = []
        for i in range(max_velocity_zero_index):
            velocity_clean_zero.append(velocity_zero[i])

        power_clean_zero_acc = []
        for i in range(max_velocity_zero_acc_index):
            power_clean_zero_acc.append(power_zero_acc[i])
        velocity_clean_zero_acc = []
        for i in range(max_velocity_zero_acc_index):
            velocity_clean_zero_acc.append(velocity_zero_acc[i])

        popt, pcov = curve_fit(self.func_lin, np.array(velocity_clean_zero), np.array(power_clean_zero))
        fitted_power_zero = self.func_lin(np.array(velocity_clean_zero), *popt)

        power_needed_to_accelerate = []
        acceleration = []
        cycle_mass = float(22.8)

        for i in range(len(time_clean_zero)):
            if len(time_clean_zero) > len(velocity_clean_zero):
                time_clean_zero.pop()

        filtered_velocity_zero = velocity_clean_zero

        for i in range(len(filtered_velocity_zero) - 1):
            acceleration.append((filtered_velocity_zero[i + 1] - filtered_velocity_zero[i]) / (2 * (time_clean_zero[i + 1] - time_clean_zero[i]))) # TODO '2 * ' verwijderen
            power_needed_to_accelerate.append(((filtered_velocity_zero[i + 1] + filtered_velocity_zero[i]) / 3.6 / 2) * acceleration[i] / 3.6 * cycle_mass)

        for j in range(len(power_clean_zero) - 1):
            if power_clean_zero[j] > power_clean_zero[j + 1]:
                energy_zero.append((0.5 * (power_clean_zero[j] - power_clean_zero[j + 1]) / (
                        time_clean_zero[j + 1] - time_clean_zero[j]) * (
                                            time_clean_zero[j + 1] - time_clean_zero[j])) + (
                                           power_clean_zero[j + 1] * (time_clean_zero[j + 1] - time_clean_zero[j])))
            else:
                energy_zero.append((0.5 * (power_clean_zero[j + 1] - power_clean_zero[j]) / (
                        time_clean_zero[j + 1] - time_clean_zero[j]) * (
                                            time_clean_zero[j + 1] - time_clean_zero[j])) + (
                                           power_clean_zero[j] * (time_clean_zero[j + 1] - time_clean_zero[j])))

        for j in range(len(energy_zero)):
            try:
                simulated_mass.append(
                    ((abs(float(velocity_clean_zero[j]) + float(velocity_clean_zero[j + 1])) / 3.6 / 2) ** -2) * 2 *
                    energy_zero[j])
            except:
                continue

        velocity_clean_zero.pop()
        popt, pcov = curve_fit(self.func_powerlaw, np.array(velocity_clean_zero), np.array(power_needed_to_accelerate))
        filtered = self.func_powerlaw(np.array(time_clean_zero), *popt)

        filtered_calculated_power = signal.savgol_filter(np.array(power_needed_to_accelerate), 3, 1)

        max_for_average = simulated_mass.index(max(simulated_mass))

        estimate_range = []
        for i in range(max_for_average + 1):
            if simulated_mass[i] > 0:
                estimate_range.append(simulated_mass[i])

        self.simulated_mass_estimate = np.mean(estimate_range)

        """
        Initialize writing an excel file.
        """
        excel = xlsxwriter.Workbook(self.user_file_name + ".xlsx")
        graph = excel.add_chart({'type': 'scatter', 'subtype': 'straight'})
        graph_2 = excel.add_chart({'type': 'scatter', 'subtype': 'straight'})
        worksheet = excel.add_worksheet()

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
        worksheet.set_column('A:H', 14)

        for i in range(len(power_clean_low)):
            vertical_y.append(0)
            vertical_y.append(power_clean_low[i])
            vertical_y.append(0)
            vertical_x.append(velocity_clean_low[i])
            vertical_x.append(velocity_clean_low[i])
            vertical_x.append(velocity_clean_low[i])

        for i in range(len(power_clean_high)):
            horizontal_x.append(0)
            horizontal_x.append(velocity_clean_high[i])
            horizontal_x.append(0)
            horizontal_y.append(power_clean_high[i])
            horizontal_y.append(power_clean_high[i])
            horizontal_y.append(power_clean_high[i])

        for i in range(len(power_clean_zero)):
            horizontal_2_x.append(0)
            horizontal_2_x.append(time_clean_zero[i])
            horizontal_2_x.append(0)
            horizontal_2_y.append(power_clean_zero[i])
            horizontal_2_y.append(power_clean_zero[i])
            horizontal_2_y.append(power_clean_zero[i])

        for i in range(len(power_clean_zero_acc)):
            horizontal_2_x.append(0)
            horizontal_2_x.append(time_clean_zero_acc[i])
            horizontal_2_x.append(0)
            horizontal_2_y.append(power_clean_zero_acc[i])
            horizontal_2_y.append(power_clean_zero_acc[i])
            horizontal_2_y.append(power_clean_zero_acc[i])

        """
        Writing to excel file.
        """
        worksheet.write('A1', 'Tested with highest gradient (without slip)', underline)
        worksheet.write('A2', 'Velocity [km/h]', bold)
        worksheet.write('B2', 'Power [W]', bold)
        worksheet.write('D1', 'Tested with lowest gradient (without slip)', underline)
        worksheet.write('D2', 'Velocity [km/h]', bold)
        worksheet.write('E2', 'Power [W]', bold)
        worksheet.write('G1', 'Tested with 0 W program - low acceleration (without slip)', underline)
        worksheet.write('G2', 'Time [s]', bold)
        worksheet.write('H2', 'Power [W]', bold)
        worksheet.write('I2', 'Fitted Power [W]')
        worksheet.write('J2', 'Calculated Power [W]', bold)
        worksheet.write('K2', 'Filtered Calculated Power [W]', bold)
        worksheet.write('L2', 'Fitted Velocity [km/h]', bold)
        worksheet.write('O1', 'Tested with 0 W program - high acceleration (without slip)', underline)
        worksheet.write('O2', 'Time [s]', bold)
        worksheet.write('P2', 'Power [km/h]', bold)
        worksheet.write('Q2', 'Velocity [km/h]', bold)
        worksheet.write('R2', 'Theor. Power [W]', bold)

        worksheet.write_column(2, 0, velocity_clean_high)
        worksheet.write_column(2, 1, power_clean_high)
        worksheet.write_column(2, 3, velocity_clean_low)
        worksheet.write_column(2, 4, power_clean_low)
        worksheet.write_column(2, 6, time_clean_zero)
        worksheet.write_column(2, 7, power_clean_zero)
        worksheet.write_column(2, 8, fitted_power_zero)
        worksheet.write_column(2, 9, power_needed_to_accelerate)
        worksheet.write_column(2, 10, filtered_calculated_power)
        worksheet.write_column(2, 11, filtered_velocity_zero)

        worksheet.write_column(2, 14, time_clean_zero_acc)
        worksheet.write_column(2, 15, power_clean_zero_acc)
        worksheet.write_column(2, 16, velocity_clean_zero_acc)
        worksheet.write_column(2, 99, vertical_x)
        worksheet.write_column(2, 100, vertical_y)
        worksheet.write_column(2, 101, horizontal_x)
        worksheet.write_column(2, 102, horizontal_y)
        worksheet.write_column(2, 103, horizontal_2_x)
        worksheet.write_column(2, 104, horizontal_2_y)

        """
        Writing to graph.
        """
        graph.add_series({
            'categories': [worksheet.name] + [2, 3] + [len(velocity_clean_low) + 2, 3],
            'values': [worksheet.name] + [2, 4] + [len(velocity_clean_low) + 2, 4],
            'line': {'color': '#67bfe7'},
            'name': 'Lowest Gradient Power',
            'fill': {'color': 'red'},
        })

        graph.add_series({
            'categories': [worksheet.name] + [2, 0] + [len(velocity_clean_high) + 2, 0],
            'values': [worksheet.name] + [2, 1] + [len(velocity_clean_high) + 2, 1],
            'line': {'color': '#67bfe7'},
            'name': 'Highest Gradient Power',
        })

        graph.add_series({
            'categories': [worksheet.name] + [2, 8] + [len(velocity_clean_zero) + 2, 8],
            'values': [worksheet.name] + [2, 9] + [len(power_needed_to_accelerate) + 2, 9],
            'line': {'color': '#ff0000'},
            'name': '0 W Theoretical Power Low Acceleration Program',
        })

        graph_2.add_series({
            'categories': [worksheet.name] + [2, 6] + [len(time_clean_zero) + 2, 6],
            'values': [worksheet.name] + [2, 7] + [len(time_clean_zero) + 2, 7],
            'line': {'color': '#ff0000'},
            'name': '0 W Program Low Acceleration Power',
        })

        graph_2.add_series({
            'categories': [worksheet.name] + [2, 6] + [len(time_clean_zero) + 2, 6],
            'values': [worksheet.name] + [2, 9] + [len(time_clean_zero) + 2, 9],
            'line': {'color': '#00ff00'},
            'name': '0 W Program Low Acceleration Theoretical Power',
        })

        graph_2.add_series({
            'categories': [worksheet.name] + [2, 6] + [len(time_clean_zero) + 2, 6],
            'values': [worksheet.name] + [2, 11] + [len(time_clean_zero) + 2, 11],
            'line': {'color': '#66ff00'},
            'name': '0 W Program Filtered Low Acceleration Theoretical Power',
        })

        graph_2.add_series({
            'categories': [worksheet.name] + [2, 6] + [len(time_clean_zero) + 2, 6],
            'values': [worksheet.name] + [2, 8] + [len(time_clean_zero) + 2, 8],
            'line': {'color': '#0000ff'},
            'name': '0 W Program Low Acceleration Velocity',
            'y2_axis': True,
        })

        graph_2.add_series({
            'categories': [worksheet.name] + [2, 6] + [len(time_clean_zero) + 2, 6],
            'values': [worksheet.name] + [2, 10] + [len(time_clean_zero) + 2, 10],
            'line': {'color': '#0000ff'},
            'name': '0 W Program Filtered Low Acceleration Velocity',
            'y2_axis': True,
        })

        graph_2.add_series({
            'categories': [worksheet.name] + [2, 13] + [len(time_clean_zero_acc) + 2, 13],
            'values': [worksheet.name] + [2, 15] + [len(velocity_clean_zero_acc) + 2, 15],
            'line': {'color': '#0000ff'},
            'name': '0 W Program High Acceleration Velocity',
            'y2_axis': True,
        })

        # graph.set_legend({'delete_series': [0, 1, 2, 3]})
        # graph_2.set_legend({'delete_series': [0, 1]})
        worksheet.insert_chart('U2', graph)
        worksheet.insert_chart('U40', graph_2)
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
        time = True

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
                time = True
                speed_values_raw = [value_list_characters[10], value_list_characters[11], value_list_characters[8],
                                    value_list_characters[9]]
                speed_values_raw_string = "".join(speed_values_raw)
                value_converter = ValueConverter()
                velocity_bin = value_converter.hex_to_bin(speed_values_raw_string)
                velocity = value_converter.bin_to_dec(velocity_bin) * 3.6 / 1000
                velocity_list.append(velocity)

                # time_values_raw = [value_list_characters[4], value_list_characters[5]]
                # time_values_raw_string = "".join(time_values_raw)
                # value_converter = ValueConverter()
                # time_bin = value_converter.hex_to_bin(time_values_raw_string)
                # time = value_converter.bin_to_dec(time_bin)
                # time_list.append(time)

            elif value_list_characters[0] == '1' and value_list_characters[
                1] == '9' and power:  # index waar power in staat
                speed = True
                power = False
                time = False
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

    def func_lin(self, x, a, b):
        return a * x + b

if __name__ == '__main__':
    Application = wx.App(False)
    frame = Main(None, 'SimulANT+ Log Analyzer').Show()
    Application.MainLoop()
