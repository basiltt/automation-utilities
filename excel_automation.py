# Developer : Basil T T(basil.tt@hpe.com)
# Created At: 24th April 2024
# Updated At: 21st May 2024
# Version: 1.2
# Note:  The module is not complete tested and may have some bugs and improvements needed
#######################################

import os
import time
from typing import Optional

import pandas as pd
from loguru import logger

# from loguru import self.logger
from win32com.client import Dispatch, constants, CDispatch, gencache
from win32com.universal import com_error


class ExcelAutomation:
    """
    Class to automate Excel operations using the win32com.client library.
    """

    def __init__(self, **kwargs):
        """
        Initializes the Excel object with the given keyword arguments.

        Parameters:
            **kwargs: Additional keyword arguments.

        Returns:
            None
        """
        # self.excel = Dispatch("Excel.Application")
        self.excel = gencache.EnsureDispatch("Excel.Application")
        self.logger = kwargs.pop("logger", logger)
        self.logger.debug("Excel application started.")
        # self.workbook = None
        self._workbook: Optional[CDispatch, None] = None
        self._worksheet: Optional[CDispatch, None] = None

        if kwargs:
            self.set_window_state(kwargs.get("window_state", -4137))
            self.set_display_alerts(kwargs.get("display_alerts", False))
            self.set_visibility(kwargs.get("visibility", True))
            self.set_screen_updating(kwargs.get("screen_updating", False))
            # self.set_calculation(kwargs.get("calculation", -4105))
            self.set_cut_copy_paste(kwargs.get("cut_copy_paste", False))

    @property
    def workbook(self):
        """
        Return the workbook property.
        """
        return self._workbook

    @workbook.setter
    def workbook(self, value):
        """
        Setter for the workbook attribute.

        Parameters:
            value: The new value to set for the workbook attribute.

        Return:
            None
        """
        self._workbook = value

    @property
    def worksheet(self):
        """
        A property function that returns the worksheet attribute.
        """
        return self._worksheet

    @worksheet.setter
    def worksheet(self, value):
        """
        A setter method for updating the worksheet attribute.

        Parameters:
            value: The new value to set for the worksheet attribute.

        Returns:
            None
        """
        self._worksheet = value

    def set_window_state(self, state=-4137):
        """
        Set the window state of the Excel application.

        :param state: int, the state to set the window to (default is -4137)
        :return: None
        """
        try:
            self.excel.WindowState = state
            self.logger.debug(f"Window state set to {state}.")
        except Exception as e:
            self.logger.warning(f"Error setting window state: {str(e)}")

    def set_screen_updating(self, value=False):
        """
        Set screen updating.

        :param value: Boolean value to set screen updating (default is False).
        :return: None
        """
        self.excel.ScreenUpdating = value
        self.logger.debug(f"Screen updating set to {value}.")

    def set_cut_copy_paste(self, value=False):
        """
        Set the CutCopyMode property of the excel object to the specified value.

        Parameters:
            value (bool): The value to set the CutCopyMode property to.

        Returns:
            None
        """
        self.excel.CutCopyMode = value
        self.logger.debug(f"Cut copy paste set to {value}.")

    def set_calculation(self, value=-4105):
        """
        Set the calculation mode for the Excel object.

        :param value: int, the calculation mode to set (default is -4105).
        :return: None
        """
        self.excel.Calculation = value
        self.logger.debug(f"Calculation set to {value}.")

    def set_display_alerts(self, value=False):
        """
        Set the display alerts property of the Excel object.

        :param value: bool - The value to set the display alerts property to. Default is False.
        :return: None
        """
        self.excel.DisplayAlerts = value
        self.logger.debug(f"Display alerts set to {value}.")

    def set_visibility(self, value=False):
        """
        Set the visibility of the excel object.

        Parameters:
            value (bool): A boolean indicating the visibility to set, default is False.
        """
        self.excel.Visible = value
        self.logger.debug(f"Visibility set to {value}.")

    def open_workbook(self, filename, update_links=0, read_only=False):
        """
        Check if file is absolute path or not and convert to absolute path

        Parameters:
            filename (str): The path to the file
            update_links (int): Flag to update external links
            read_only (bool): Flag to open the file in read-only mode

        Returns:
            workbook: The opened workbook object
        """

        # Check if file is absolute path or not and convert to absolute path
        if not os.path.isabs(filename):
            filename = os.path.join(os.getcwd(), filename)

        # Check file exists
        if not os.path.exists(filename):
            self.logger.error(f"File {filename} does not exist.")
            raise FileNotFoundError(f"File {filename} does not exist.")

        self.workbook = self.excel.Workbooks.Open(
            filename, UpdateLinks=update_links, ReadOnly=read_only
        )
        self.logger.debug(f"Workbook {filename} opened.")
        return self.workbook

    def set_worksheet(self, name):
        """
        Set the worksheet with the given name in the workbook.

        :param name: The name of the worksheet to set.
        :return: The worksheet that was set.
        """
        self.worksheet = self.workbook.Worksheets(name)
        self.logger.debug(f"Worksheet {name} set.")
        return self.worksheet

    def run_macro(self, macro_name):
        """
        Runs a macro specified by the macro_name parameter using the workbook's Application.Run method.

        Parameters:
            macro_name (str): The name of the macro to run.

        Returns:
            None
        """
        self.workbook.Application.Run(macro_name)
        self.logger.debug(f"Macro {macro_name} run.")

    def activate_workbook(self, workbook: Optional[CDispatch] = None):
        """
        A description of the entire function, its parameters, and its return types.
        """
        if workbook:
            workbook.Activate()
            self.logger.debug("Workbook activated.")
        else:
            self.workbook.Activate()
            self.logger.debug("Workbook activated.")

    def save_as(self, filename, workbook: Optional[CDispatch] = None):
        """
        Save the workbook with the specified filename.

        :param filename: The name of the file to save the workbook as.
        :param workbook: An optional workbook object to save. Default is None.
        """
        if workbook:
            workbook.SaveAs(Filename=filename)
            self.logger.debug(f"Workbook saved as {filename}.")
        else:
            self.workbook.SaveAs(Filename=filename)
            self.logger.debug(f"Workbook saved as {filename}.")

    def close_workbook(
        self, save_changes=False, workbook: Optional[CDispatch] = None
    ):
        """
        A function to close a workbook, with an optional parameter to save changes.

        :param save_changes: bool, indicates whether to save changes made to the workbook
        :param workbook: Optional[CDispatch], the workbook to be closed
        """
        if workbook:
            workbook.Close(SaveChanges=save_changes)
            self.logger.debug("Workbook closed.")
        else:
            if self.workbook:
                try:
                    self.workbook.Close(SaveChanges=save_changes)
                except com_error as e:
                    if (
                        "The object invoked has disconnected from its clients."
                        in str(e)
                    ):
                        pass
                    else:
                        self.logger.warning(f"Error closing workbook: {e}")
                self.logger.debug("Workbook closed.")

    def save(self):
        """
        Save the workbook and log a debug message.
        """
        self.workbook.Save()
        self.logger.debug("Workbook saved.")

    def quit(self, retries=5, delay=2):
        """
        Quit the Excel application.

        Parameters:
            retries (int): Number of retries to attempt to quit the application.
            delay (int): Delay in seconds between retry attempts.

        Returns:
            None
        """
        for i in range(retries):
            try:
                self.excel.Quit()
                self.logger.debug("Excel application quit.")
                break
            except com_error as e:
                if e.strerror == (
                    "The message filter indicated that the "
                    "application is busy."
                ):
                    time.sleep(delay)
                else:
                    # Try killing the Excel forcefully for the current logged-in user
                    self.logger.warning(
                        "Unable to quit Excel application gracefully. "
                        "Killing the Excel application forcefully."
                    )
                    username = os.getlogin()
                    # Kill the Excel application only for the current user
                    os.system(
                        f'taskkill /F /FI "USERNAME eq {username}" '
                        f"/IM excel.exe"
                    )
                    self.logger.debug("Excel application killed forcefully.")

    def get_worksheet(self, name):
        """
        Retrieves a worksheet by name from the workbook.

        Parameters:
            name (str): The name of the worksheet to retrieve.

        Returns:
            Worksheet: The worksheet object retrieved by name.
        """
        worksheet = self.workbook.Worksheets(name)
        self.logger.debug(f"Worksheet {name} retrieved.")
        return worksheet

    def paste_to_range(self, worksheet, range_start, paste_type=-4163):
        """
        A function to paste the copied content to a specified range in the given worksheet.

        Parameters:
            worksheet: The worksheet to paste the content into.
            range_start: The starting range where the content will be pasted.
            paste_type: The type of paste operation to perform (default is -4163).

        Returns:
            None
        """
        worksheet.Range(range_start).PasteSpecial(Paste=paste_type)
        self.logger.debug(f"Pasted to range {range_start}.")

    def set_cell_value(
        self, cell, value, worksheet: Optional[CDispatch] = None
    ):
        """
        Set the value of a cell in a worksheet.

        Parameters:
            cell: Tuple representing the cell coordinates (row, col).
            value: The value to be set in the cell.
            worksheet: Optional[CDispatch], the worksheet object where the cell is located.
                If not provided, the default worksheet is used.

        Returns:
            None
        """
        row, col = cell
        if worksheet:
            worksheet.Cells(row, col).Value = value
            self.logger.debug(f"Set value of cell '{cell}' to '{value}'.")
        else:
            self.worksheet.Cells(row, col).Value = value
            self.logger.debug(f"Set value of cell '{cell}' to '{value}'.")

    def copy_range(self, data_range, worksheet: Optional[CDispatch] = None):
        """
        A function to copy a specified data range, either within the given worksheet or the instance's worksheet.

        Parameters:
            data_range: str - The range of data to be copied.
            worksheet: Optional[CDispatch] - The worksheet to copy the data to. If not provided, data will be copied to the instance's worksheet.
        """
        self.logger.debug(f"Copying data from the range '{data_range}'.")
        if worksheet:
            data_range.Copy()
            self.logger.debug(f"Copied data from the range '{data_range}'.")
        else:
            self.worksheet.Range(data_range).Copy()
            self.logger.debug(f"Copied data from the range '{data_range}' .")

        # Add a sleep to allow the copy to complete
        time.sleep(3)

    def paste_range_as_special(
        self,
        range_start,
        worksheet: Optional[CDispatch] = None,
        paste_type=-4163,
    ):
        """
        A description of the entire function, its parameters, and its return types.
        """
        if worksheet:
            worksheet.Range(range_start).PasteSpecial(Paste=paste_type)
            self.logger.debug(f"Pasted to range {range_start}.")
        else:
            self.worksheet.Range(range_start).PasteSpecial(Paste=paste_type)
            self.logger.debug(f"Pasted to range {range_start}.")

    # Create  a function to write a data  frame starting from a row and column
    def write_dataframe_to_excel_with_a_start_row_and_start_column(
        self, df, start_row, start_col, worksheet: Optional[CDispatch] = None
    ):
        """
        Write a dataframe to an Excel worksheet starting from a specified start row and start column.

        Parameters:
            df (pandas.DataFrame): The dataframe to write to the worksheet.
            start_row (int): The row to start writing the dataframe.
            start_col (int): The column to start writing the dataframe.
            worksheet (Optional[CDispatch]): The worksheet to write the dataframe to. Defaults to None.
        """
        if worksheet:
            for i in range(len(df)):
                for j in range(len(df.columns)):
                    value = df.iloc[i, j]
                    if pd.isna(value):
                        worksheet.Cells(start_row + i, start_col + j).Value = (
                            None
                        )
                    else:
                        worksheet.Cells(start_row + i, start_col + j).Value = (
                            value
                        )
            self.logger.debug("Dataframe written to worksheet.")
        else:
            for i in range(len(df)):
                for j in range(len(df.columns)):
                    value = df.iloc[i, j]
                    if pd.isna(value):
                        self.worksheet.Cells(
                            start_row + i, start_col + j
                        ).Value = None
                    else:
                        self.worksheet.Cells(
                            start_row + i, start_col + j
                        ).Value = value
            self.logger.debug("Dataframe written to worksheet.")

    def sleep(self, seconds):
        """
        Sleep for a specified number of seconds.

        Parameters:
            seconds (int): The number of seconds to sleep.

        Returns:
            None
        """
        time.sleep(seconds)
        self.logger.debug(f"Slept for {seconds} seconds.")

    def get_active_sheet(self):
        """
        A method to retrieve the active sheet and return it.

        :return: The active sheet object.
        """
        active_sheet = self.excel.ActiveSheet
        self.logger.debug("Retrieved active sheet.")
        return active_sheet

    def add_worksheet(self, name=None):
        """
        Add a worksheet to the workbook with an optional name.

        Parameters:
            name (str): The name of the worksheet (optional).

        Returns:
            Worksheet: The newly added worksheet.
        """
        if name:
            worksheet = self.workbook.Worksheets.Add()
            worksheet.Name = name
            self.logger.debug(f"Added worksheet with name {name}.")
            return worksheet
        else:
            worksheet = self.workbook.Worksheets.Add()
            self.logger.debug("Added worksheet.")
            return worksheet

    def delete_worksheet(self, name):
        """
        Delete a worksheet by name from the workbook.

        :param name: The name of the worksheet to delete.
        :return: None
        """
        self.workbook.Worksheets(name).Delete()
        self.logger.debug(f"Deleted worksheet {name}.")

    def rename_worksheet(self, old_name, new_name):
        """
        Renames a worksheet in the workbook.

        :param old_name: The name of the worksheet to be renamed.
        :param new_name: The new name for the worksheet.
        """
        self.workbook.Worksheets(old_name).Name = new_name
        self.logger.debug(f"Renamed worksheet {old_name} to {new_name}.")

    def select_range(self, worksheet, range_value):
        """
        A method to select a range in a worksheet.

        Parameters:
            worksheet: The worksheet where the range will be selected.
            range_value: The value representing the range to be selected.

        Returns:
            The selected range object.
        """
        range_selected = worksheet.Range(range_value)
        self.logger.debug(f"Selected range {range_value}.")
        return range_selected

    # @staticmethod
    def get_cell_value(self, cell, worksheet: Optional[CDispatch] = None):
        """
        A function to get the value of a specific cell in a worksheet.

        Parameters:
            cell: Tuple containing the row and column indices of the cell.
            worksheet: Optional[CDispatch], the specific worksheet to retrieve the cell value from.

        Returns:
            The value of the specified cell.
        """
        row, col = cell
        if worksheet:
            value = worksheet.Cells(row, col).Value
            self.logger.debug(
                f"Retrieved value of cell '{cell}' and value is '{value}'."
            )
            return value
        else:
            value = self.worksheet.Cells(row, col).Value
            self.logger.debug(
                f"Retrieved value of cell '{cell}' and value is '{value}'."
            )
            return value

    def get_cell_value_with_title_and_row_index(
        self,
        title,
        row_index,
        worksheet: Optional[CDispatch] = None,
        title_row_index=1,
    ):
        """
        A function to retrieve the cell value based on the title and row index from a given worksheet.

        Parameters:
            title (str): The title to search for in the worksheet.
            row_index (int): The row index where the value needs to be retrieved.
            worksheet (Optional[CDispatch]): The worksheet to search for the title. Defaults to None.
            title_row_index (int): The row index where the titles are located. Defaults to 1.

        Returns:
            The value of the cell if found, None otherwise.
        """
        if worksheet is None:
            worksheet = self.worksheet

        title_range = worksheet.Rows(title_row_index).Find(
            What=title, LookAt=constants.xlWhole
        )
        if title_range is not None:
            col = title_range.Column
            value = worksheet.Cells(row_index, col).Value
            self.logger.debug(
                f"Retrieved value of cell with title {title} and value is {value}."
            )
            return value
        else:
            self.logger.debug(
                f"Title {title} not found in row {title_row_index}."
            )
            return None

    def set_range_values(self, worksheet, range_start, values):
        """
        Set values to a specified range in the given worksheet.

        Parameters:
            worksheet (object): The worksheet to set the values in.
            range_start (str): The starting cell of the range.
            values (list): The values to set in the range.

        Returns:
            None
        """
        worksheet.Range(range_start).Value = values
        self.logger.debug(f"Set values of range {range_start}.")

    def get_range_values(self, worksheet, range_start):
        """
        Retrieves values from a given range in a worksheet.

        Parameters:
            worksheet (object): The worksheet object to retrieve values from.
            range_start (str): The starting range to retrieve values from.

        Returns:
            object: The values retrieved from the specified range.
        """
        values = worksheet.Range(range_start).Value
        self.logger.debug(f"Retrieved values of range {range_start}.")
        return values

    def clear_range(self, worksheet, range_start):
        """
        Clears a specified range in the worksheet.

        Parameters:
            worksheet (object): The worksheet object.
            range_start (str): The starting cell of the range to clear.
        """
        worksheet.Range(range_start).ClearContents()
        self.logger.debug(f"Cleared range {range_start}.")

    # @staticmethod
    def get_used_range(self, worksheet: Optional[CDispatch] = None):
        """
        A function to get the used range of a worksheet.

        Parameters:
            worksheet (Optional[CDispatch]): The worksheet to get the used range from.

        Returns:
            The used range of the worksheet.
        """
        if worksheet:
            used_range = worksheet.UsedRange
            self.logger.debug(
                f"Retrieved used range of worksheet {worksheet} and used range is {used_range}."
            )
            return used_range
        else:
            used_range = self.worksheet.UsedRange
            self.logger.debug(
                f"Retrieved used range of worksheet {worksheet} and used range is {used_range}."
            )
            return used_range

    # @staticmethod
    def get_used_row_count(self, worksheet: Optional[CDispatch] = None):
        """
        Get the used row count of the specified worksheet. If no worksheet is provided,
        it defaults to using the instance's worksheet. Returns the total row count.

        :param worksheet: The worksheet to retrieve the row count from. Defaults to None.
        :type worksheet: Optional[CDispatch]
        :return: The total row count of the specified worksheet.
        :rtype: int
        """
        if worksheet:
            row_count = worksheet.UsedRange.Rows.Count
            self.logger.debug(
                f"Retrieved row count of worksheet {worksheet} and row count is {row_count}."
            )
            return row_count
        else:
            row_count = self.worksheet.UsedRange.Rows.Count
            self.logger.debug(
                f"Retrieved row count of worksheet {worksheet} and row count is {row_count}."
            )
            return row_count

    def get_column_count(self, worksheet):
        """
        Get the count of columns in the given worksheet.

        Parameters:
            worksheet (object): The worksheet object to retrieve the column count from.

        Returns:
            int: The number of columns in the worksheet.
        """
        column_count = worksheet.UsedRange.Columns.Count
        self.logger.debug("Retrieved column count.")
        return column_count

    def get_cell_formula(self, worksheet, cell):
        """
        Get the formula of a specific cell in the given worksheet.

        Parameters:
            worksheet (worksheet): The worksheet object where the cell is located.
            cell (int): The cell number to retrieve the formula from.

        Returns:
            str: The formula of the specified cell.
        """
        formula = worksheet.Cells(cell).Formula
        self.logger.debug(f"Retrieved formula of cell {cell}.")
        return formula

    def set_cell_formula(self, worksheet, cell, formula):
        """
        Set the formula of a cell in the given worksheet.

        :param worksheet: The worksheet object where the cell is located.
        :param cell: The cell reference where the formula will be set.
        :param formula: The formula to set in the cell.
        """
        worksheet.Cells(cell).Formula = formula
        self.logger.debug(f"Set formula of cell {cell} to {formula}.")

    def protect_worksheet(self, name, password):
        """
        Protects a worksheet with a password.

        :param name: The name of the worksheet to protect.
        :param password: The password to use for protection.
        """
        self.workbook.Worksheets(name).Protect(password)
        self.logger.debug(f"Protected worksheet {name}.")

    def unprotect_worksheet(self, name, password):
        """
        Unprotect a worksheet using the provided password.

        :param name: The name of the worksheet to unprotect.
        :param password: The password required to unprotect the worksheet.
        """
        self.workbook.Worksheets(name).Unprotect(password)
        self.logger.debug(f"Unprotected worksheet {name}.")

    def protect_workbook(self, password):
        """
        Protects the workbook with a password.

        :param password: str - The password to protect the workbook.
        :return: None
        """
        self.workbook.Protect(password)
        self.logger.debug("Protected workbook.")

    def unprotect_workbook(self, password):
        """
        A description of the entire function, its parameters, and its return types.
        """
        self.workbook.Unprotect(password)
        self.logger.debug("Unprotected workbook.")

        # Worksheet operations
        def get_worksheet_count(self):
            """
            Method to retrieve the count of worksheets in the workbook.
            No parameters.
            Returns the count of worksheets in the workbook.
            """
            count = self.workbook.Worksheets.Count
            self.logger.debug(f"Worksheet count: {count}")
            return count

        def get_worksheet_names(self):
            """
            Get the names of all the worksheets in the workbook.

            :return: List of worksheet names
            """
            names = [sheet.Name for sheet in self.workbook.Worksheets]
            self.logger.debug(f"Worksheet names: {names}")
            return names

        def hide_worksheet(self, name):
            """
            Hides a specific worksheet in the workbook.

            :param name: The name of the worksheet to hide.
            :return: None
            """
            self.workbook.Worksheets(name).Visible = False
            self.logger.debug(f"Worksheet {name} hidden.")

        def show_worksheet(self, name):
            """
            Shows a specific worksheet in the workbook.

            Parameters:
                name (str): The name of the worksheet to be shown.

            Returns:
                None
            """
            self.workbook.Worksheets(name).Visible = True
            self.logger.debug(f"Worksheet {name} shown.")

        # Range and Cell operations
        def get_range_address(self, worksheet, range_start):
            """
            Get the address of a range in the given worksheet starting from the specified range_start.

            :param worksheet: The worksheet object where the range is located.
            :param range_start: The starting point of the range.
            :return: Address of the range.
            """
            address = worksheet.Range(range_start).Address
            self.logger.debug(f"Range address: {address}")
            return address

        def get_cell_address(self, worksheet, cell):
            """
            Get the address of a specific cell in the worksheet.

            Args:
                worksheet: The worksheet object where the cell is located.
                cell: The cell reference.

            Returns:
                str: The address of the cell.
            """
            address = worksheet.Cells(cell).Address
            self.logger.debug(f"Cell address: {address}")
            return address

        def get_cell_format(self, worksheet, cell):
            """
            Get the format of a specific cell in a worksheet.

            Args:
                worksheet: The worksheet containing the cell.
                cell: The cell to retrieve the format from.

            Returns:
                The format of the specified cell.
            """
            format = worksheet.Cells(cell).NumberFormat
            self.logger.debug(f"Cell format: {format}")
            return format

        def set_cell_format(self, worksheet, cell, format):
            """
            Set the format of a specific cell in the given worksheet.

            Parameters:
                worksheet (object): The worksheet object where the cell is located.
                cell (object): The cell to set the format for.
                format (str): The format to apply to the cell.

            Returns:
                None
            """
            worksheet.Cells(cell).NumberFormat = format
            self.logger.debug(f"Set cell format to {format}.")

        # Utility methods
        def calculate(self):
            """
            Calculate method that triggers the calculation in the excel object and logs the action.
            """
            self.excel.Calculate()
            self.logger.debug("Excel calculated.")

        def save_copy_as(self, filename):
            """
            Save a copy of the workbook with the specified filename.

            Parameters:
                filename (str): The name of the file to save the workbook copy as.
            """
            self.workbook.SaveCopyAs(Filename=filename)
            self.logger.debug(f"Workbook saved as copy: {filename}")

        def refresh_all(self):
            """
            Refreshes all data in the workbook and logs the action.
            """
            self.workbook.RefreshAll()
            self.logger.debug("Workbook refreshed.")

        def get_named_ranges(self):
            """
            Get all named ranges from the workbook.
            No parameters.
            Returns a list of named ranges.
            """
            names = [name.Name for name in self.workbook.Names]
            self.logger.debug(f"Named ranges: {names}")
            return names

        def get_named_range_value(self, name):
            """
            Get the value of a named range in the workbook.

            :param name: The name of the range.
            :return: The value of the named range.
            """
            value = self.workbook.Names(name).RefersToRange.Value
            self.logger.debug(f"Named range value: {value}")
            return value

        def set_named_range_value(self, name, value):
            """
            Set the value of a named range in the workbook.

            Parameters:
                name (str): The name of the range to set the value for.
                value (Any): The value to set for the named range.

            Returns:
                None
            """
            self.workbook.Names(name).RefersToRange.Value = value
            self.logger.debug(f"Set named range value to {value}.")

        def add_named_range(self, name, refers_to):
            """
            Adds a named range to the workbook.

            Parameters:
                name (str): The name of the named range.
                refers_to (str): The cell or range that the named range refers to.
            """
            self.workbook.Names.Add(Name=name, RefersTo=refers_to)
            self.logger.debug(f"Added named range: {name}")

        def delete_named_range(self, name):
            """
            Delete a named range from the workbook.

            Args:
                name (str): The name of the range to be deleted.
            """
            self.workbook.Names(name).Delete()
            self.logger.debug(f"Deleted named range: {name}")

        def protect_range(self, worksheet, range_start, password):
            """
            This function protects a specified range in a worksheet using a given password.

            Parameters:
                worksheet (object): The worksheet object where the range is located.
                range_start (str): The starting cell of the range to be protected.
                password (str): The password to protect the range.

            Returns:
                None
            """
            worksheet.Range(range_start).Protect(password)
            self.logger.debug(f"Protected range: {range_start}")

        def unprotect_range(self, worksheet, range_start, password):
            """
            Unprotect a range in the worksheet using the provided password.

            :param worksheet: The worksheet object where the range is located.
            :param range_start: The starting range to be unprotected.
            :param password: The password needed to unprotect the range.
            """
            worksheet.Range(range_start).Unprotect(password)
            self.logger.debug(f"Unprotected range: {range_start}")

        def find(self, worksheet, value):
            """
            Find a specific value in the given worksheet.

            :param worksheet: The worksheet to search in.
            :param value: The value to find in the worksheet.
            :return: The result of the search operation.
            """
            result = worksheet.Cells.Find(What=value, LookAt=constants.xlWhole)
            self.logger.debug(f"Found value: {value}")
            return result
