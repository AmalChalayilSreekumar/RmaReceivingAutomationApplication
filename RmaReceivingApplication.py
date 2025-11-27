import pyautogui as pa
import win32gui
import tkinter
import tkinter as tk
from tkinter import *
import re
from queue import Queue
import os

"""
RMA Receiving Program
======================
Author: Amal Chalayil Sreekumar  
Date: August 21st, 2025  

Description:  
This program automates the processing of RMAs in  AS400 environment.  
It identifies the AS400 window, collects information about RMAs, processes barcodes, and organizes data and images 
into structured folder paths. 
A GUI interface built with Tkinter allows users to input RMA numbers, navigate through 
barcodes, and track critical data such as return types, SLA status, and damaged product details.

Key Features:  
1. Identifies the AS400 application window and interacts with it.  
2. Extracts information such as return type, part numbers, SLA status, etc.  
3. Creates and organizes folder structures for received items and damaged items based on RMA numbers.  
4. Provides an intuitive GUI for input, progress tracking, and user control during the process.  
5. Writes processed data into organized `.txt` files for tracking purposes.  
6. Handles cases of damaged RMAs with optional folder creation for damaged items.

Frameworks/Libraries Used:  
- PyAutoGUI: Automates keyboard and mouse interactions to navigate the AS400 system.  
- Win32GUI: Retrieves and interacts with open windows in the operating system.  
- Tkinter: Builds the user interface for the program.  
- OS: Manages file paths and folder creations.  
- Regex: Used to extract barcodes for RMA.  

Usage Instructions:  
1. Run the program to launch the GUI.  
2. Enter the RMA number and start the process using the "Start" button.  
3. Navigate through individual serial numbers using the "Next Serial Number" button or the space bar.  
4. Follow on-screen messages for progress and errors.  

Developed in collaboration with:  
- Majority of the AccessAS400 class functionality written by Deivy Munoz.  
- Jezu Mario Palackal Stanley contributed to folder structure creation logic.

"""


class AccessAS400:  # Majority written by Deivy Munoz
    """
    Class to identify the AS400 application window.
    """

    def find_as400(self, list_of_windows):
        """
        Searches a list of open windows for the AS400 application.

        Args:
            list_of_windows (list): A list of currently open windows

        Returns:
            int: The window handle (`hwnd`) of the AS400 application if found, or -1 if not found.
        """
        for window in list_of_windows:
            if "as400" in window[2].lower():  # Check if AS400 is in the title (case insensitive)
                return window[0]  # Return the window handle ('hwnd')
        return -1  # Return -1 if no AS400 window found

    def list_window_names(self):
        """
        Lists the names (titles) of all currently open windows.

        Uses the WinAPI to enumerate all visible windows and retrieve their
        handles, hex representations, and names.

        Returns:
            list: A list of currently open windows.
        """
        list_of_windows = []

        def win_enum_handler(hwnd, ctx):
            """
            Callback function for use with `win32gui.EnumWindows`.

            Captures all visible windows and appends their information
            (handle, hex handle, title) to the `list_of_windows`.

            Args:
                hwnd (int): The handle to the window being enumerated.
                ctx: Additional context (not used).
            """
            if win32gui.IsWindowVisible(hwnd):  # Check if the window is visible
                list_of_windows.append([hwnd, hex(hwnd), win32gui.GetWindowText(hwnd)])

        win32gui.EnumWindows(win_enum_handler, None)  # Enumerate all windows
        return list_of_windows

    def as400_main_screen(self, data):
        """
        Checks whether the AS400 is in the main menu for Failure analysis.

        Args:
            data (str): The screen data that it takes in

        Returns:
            bool: `True` if the content contains "Failure Analysis Menu",
                  otherwise `False`.
        """
        if "Failure Analysis Menu" in data:
            return True 
        else:
            return False  

class ProcessRMA(AccessAS400):
    def __init__(self, RMA):
        """
        Initializes Finds the AS400 window and sets it to the foreground. Initializes the tkinter root, checks if it is in AS400 homescreen 
        then initializes the date and receiver variables by copying the 

        Args:
            RMA(str): The RMA number

        Returns: Nothing
        """
        self.RMA = RMA #Initialize the RMA


        list_of_windows = self.list_window_names()
        as400 = self.find_as400(list_of_windows)
        win32gui.SetForegroundWindow(as400)
        pa.hotkey('ctrl', 'c')

        self.root = tkinter.Tk()
        self.root.withdraw()
        mdata = self.root.clipboard_get().split('\n')
        # self.root.destroy()

        mdata = mdata[0]

        if self.as400_main_screen(mdata):
            mdata = mdata.split(" ")
            cleaned = [s.strip() for s in mdata if s.strip()]
            self.receiver = cleaned[1]
            self.date = cleaned[5]
            pa.typewrite("02")
            pa.hotkey(["return"], interval=0.01)
        else:
            pa.typewrite("e")
            pa.hotkey(["return"])
            return

    def getAssigned(self):
        """"
        Gets who the RMA is assigned to by analyzing the screen data.
        
        Args: None

        Returns:
            str: The name of the person who the RMA is assigned to, or "Not assigned" if it is not assigned yet.
        """
        screenClean = self.screenCopy()
        for i in range(len(screenClean)):
            if screenClean[i] == 'Note':
                #If it is not assigned screenClean[i+2] would return the date (MM/DD/YY)(December = 12, Jan = 01) 
                # which is why the condition for the if statement is whether screenClean[i+2] starts with 1 or a 0 
                if screenClean[i+2].startswith("1") or screenClean[i+2].startswith('0'): 
                    return "Not assigned"
                else:
                    return screenClean[i+2]

    def dateFormat(self):
        """"
        Formats the date from eg. 08/27/25 --> Aug/27/2025. 

        Args: None

        Returns:
            str: The date formatted.
        """
        self.months = {
        1: "Jan",
        2: "Feb",
        3: "Mar",
        4: "Apr",
        5: "May",
        6: "Jun",
        7: "Jul",
        8: "Aug",
        9: "Sep",
        10: "Oct",
        11: "Nov",
        12: "Dec"
        }

        date = self.date.split("/")
        date[2] = f"20{date[2]}"
        date[0] = self.months[int(date[0])]

        return date

    def trackRMA(self, serialNum, returnType, partNum):
        """
        Searches for an existing folder in the specified parent directory whose name starts with a given prefix.
        If no such folder exists, it generates a default folder name using the prefix filled with additional 'x' characters then
        creates or opens a txt file based on if it exists for the current day and writes in the following RMA information:
        RMA#, Return Type, SerialNum, Part Number, Person Assigned To, and Received by".

        Args:
            serialNum (str): The serial number of the current part being proccessed
            returnType (str): The return type of the current part being proccessed
            partNum(str): The part number of the current part being proccessed

        Returns:
            dayFile (str): The path of the txt file that was either created or opened and written into
        """
        date = self.dateFormat()
        yearFolder = date[2]  
        monthFolder = date[0]  
        dayFile = f"{date[0]} {date[1]}, {date[2]}.txt"  

        basePath = r"\\panther\RMA\RMA_Repairs\RMAs_Received"

        # Define the full path structure
        yearPath = os.path.join(basePath, yearFolder)  
        monthPath = os.path.join(yearPath, monthFolder)  
        dayFile = os.path.join(monthPath, dayFile)  

        # Ensure the folder structure exists
        if not os.path.exists(yearPath):
            os.makedirs(yearPath)
        if not os.path.exists(monthPath):
            os.makedirs(monthPath)

        # Write content multiple times into the file

        with open(dayFile, "a") as f:
            f.write(f"RMA#: {self.RMA}   Type: {returnType}   S/N: {serialNum}	P/N: {partNum}	Assigned To: {self.assignedTo}	Received by: {self.receiver}\n")

        # Notify the user
        return dayFile

    def find_existing_folder(self,parent, folder_prefix, total_length): #Function written with Jezu Mario Palackal Stanley
        """
        Searches for an existing folder in the specified parent directory whose name starts with a given prefix.
        If no such folder exists, it generates a default folder name using the prefix filled with additional 'x' characters 
        to match the required total length.

        Args:
            parent (str): Path to the parent directory where the folders are located.
            folder_prefix (str): The prefix of the folder to search for.
            total_length (int): The total length of the desired folder name (including the prefix).

        Returns:
            str: Name of the matching folder if found, or a generated folder name based on 
                the prefix if no match exists.
        """
        candidates = []
        if os.path.exists(parent):
            for fname in os.listdir(parent):
                fpath = os.path.join(parent, fname)
                if os.path.isdir(fpath) and fname.startswith(folder_prefix):
                    candidates.append(fname)

        if candidates:
            match = min(candidates, key=len)
            return match
        return folder_prefix + ('x' * (total_length - len(folder_prefix)))

    def create_rma_folder_structure(self): #Function written with Jezu Mario Palackal Stanley
        """
        Creates a nested folder structure for RMA images based on the RMA number.Each RMA is organized into three levels: a root folder, 
        a subfolder, and finally a unique folder for the specific RMA. 

        Args:
            None.

        Returns:
            str: Path to the third-level folder corresponding to the specific RMA. 
                Creates the folder if it does not exist.
        
        Raises:
            ValueError: If the RMA number is not in the expected format (e.g., does not start with 'RMA' and does not have at least
            six digits following 'RMA').
        """
        if not (self.RMA.startswith('RMA') and len(self.RMA) >= 9 and self.RMA[3:].isdigit()):
            raise ValueError("RMA code must start with 'RMA' and have at least 6 digits after RMA")

        base_path = r"\\panther\RMA\RMA_Repairs\RMA_Received_Pictures"

        first_folder_prefix = self.RMA[:5]
        second_folder_prefix = self.RMA[:6]
        third_folder = self.RMA

        # Find or create root folder
        first_folder = self.find_existing_folder(base_path, first_folder_prefix, 9)
        first_path = os.path.join(base_path, first_folder)
        if not os.path.exists(first_path):
            os.makedirs(first_path)

        # Find or create second level folder
        second_folder = self.find_existing_folder(first_path, second_folder_prefix, 9)
        second_path = os.path.join(first_path, second_folder)
        if not os.path.exists(second_path):
            os.makedirs(second_path)

        #create third level folder for unique RMA

        third_path = os.path.join(second_path, third_folder)
        if not os.path.exists(third_path):
            os.makedirs(third_path, exist_ok=True)
            #Return the folder path of the third folder created/opened
            return third_path
        else:
            return third_path

    def createDamagedRmaFolder(self):
        """
        Same logic as the create_rma_folder_structure(self) method.  

        Args:
            None.

        Returns:
            str: Path to the third-level folder corresponding to the specific RMA. 
                Creates the folder if it does not exist.
        
        Raises:
            ValueError: If the RMA number is not in the expected format 
            (e.g., does not start with 'RMA' and does not have at least six digits following 'RMA').
        """
        if not (self.RMA.startswith('RMA') and len(self.RMA) >= 9 and self.RMA[3:].isdigit()):
            raise ValueError("RMA code must start with 'RMA' and have at least 6 digits after RMA")

        base_path = r"\\panther\RMA\RMA_Repairs\RMA_Damage"

        first_folder_prefix = self.RMA[:5]
        second_folder_prefix = self.RMA[:6]
        third_folder = self.RMA

        # Find or create root folder
        first_folder = self.find_existing_folder(base_path, first_folder_prefix, 9)
        first_path = os.path.join(base_path, first_folder)
        if not os.path.exists(first_path):
            os.makedirs(first_path)

        # Find or create second level folder
        second_folder = self.find_existing_folder(first_path, second_folder_prefix, 9)
        second_path = os.path.join(first_path, second_folder)
        if not os.path.exists(second_path):
            os.makedirs(second_path)

        #create third level folder for unique RMA

        third_path = os.path.join(second_path, third_folder)
        if not os.path.exists(third_path):
            os.makedirs(third_path, exist_ok=True)
            return third_path
        else:
            return third_path

    def getBarcodes(self):
        """
        Navigates through and enters the RMA number into FA02 then gets and returns all barcodes in the RMA entered. If more than one page of 
        barcodes in RMA then it will press page down until at last page while still collecting all the barcodes.

        Args:
            None.

        Returns:
            self.barcodeList (Queue): A queue containing all the serial numbers in the RMA
            **If Queue is empty** return (str): RMA not open
        
        """
        
        #Navigates to section to search up RMA number in FA02
        pa.typewrite("I")
        pa.hotkey("down")
        pa.hotkey("down")
        pa.hotkey("end") #Deletes the user
        pa.hotkey("up")

        #Enters RMA number
        pa.typewrite(f"{self.RMA}")
        pa.hotkey("return")

        #Gets who the RMA is assigned to and stores it in an instance variable
        self.assignedTo = self.getAssigned()

        #Store screen data in a variable
        pa.hotkey('ctrl', 'a')
        pa.hotkey('ctrl', 'c')
        screen = self.root.clipboard_get()

        #Gets all barcodes if there is more than one page of serial numbers in an RMA
        if "More..." in screen:
            while "Bottom" not in screen:
                pa.hotkey("pagedown")
                pa.hotkey('ctrl', 'a')
                pa.hotkey('ctrl', 'c')
                new = self.root.clipboard_get()
                screen+=new

                if "Bottom" in screen:
                    break
        
        #Find all serial numbers in the screen
        self.barcodes = re.findall(r'\b\d{10}\b', screen)

        # Print the extracted barcodes
        self.barcodeList = Queue()
        for barcode in self.barcodes:
            self.barcodeList.put(barcode)

        #For the case that the RMA is not open therefore having no serial numbers when RMA is searched into FA02
        if self.barcodeList.empty() == True:
            return "RMA not open"

        #Put whitespace at the end of the queue (important of GUI section)
        self.barcodeList.put(" ")
        return self.barcodeList

    def enterDate(self):
        """
        Once inside the failure analysis processing screen for a product in an RMA, this method enters the date into the "Other" category. 

        Args:
            None.

        Returns:
            Nothing
        """
        #Format self.date
        date = self.dateFormat() 
        date = f"{date[0]} {date[1]}, {date[2]}"

        #Navigate to the "Other" section
        pa.hotkey("down")
        pa.hotkey("down")
        pa.hotkey("down")
        pa.hotkey("down")
        pa.hotkey("right")
        pa.hotkey("right")
        pa.hotkey("right")
        pa.hotkey("right")

        #Write the date into the "Other" section
        pa.typewrite(f"{date}")

    def deleteDate(self):
        """
        ***This method is for testing***
        Once inside the failure analysis processing screen for a product in an RMA, this method deletes the date in the "Other" category. 

        Args:
            None.

        Returns:
            Nothing
        """
        #Navigates to the "Other" section
        pa.hotkey("down")
        pa.hotkey("down")
        pa.hotkey("down")
        pa.hotkey("down")
        pa.hotkey("right")
        pa.hotkey("right")
        pa.hotkey("right")
        pa.hotkey("right")

        #Deletes the date
        pa.hotkey('end') 
        pa.hotkey('return')

    def screenCopy(self):
        """
        ***Helper Method***
        This method copys the content onto the screen, removes all of the whitespace and then returns a list of everthing that is on the screen when 
        function is called. 

        Args:
            None.

        Returns:
            screenClean (List): List of all content (words, numbers, characters) as a string without whitespace 
        """
        pa.hotkey('ctrl', 'a')
        pa.hotkey('ctrl', 'c')
        screen = self.root.clipboard_get()
        screen = screen.split(" ")
        screenClean = []
        for i in screen:
            if i != "":
                screenClean.append(i)
        return screenClean

    def isSLA(self):
        """
        Checks if the current product is an SLA.

        This method scans the screen data for the keyword "SLA" and evaluates the response following it. 
        If the response is "Y", it indicates that the product is under SLA.

        Args:
            None.

        Returns:
            str: "Yes" if the product is SLA, otherwise "No".
        """

        screenClean = self.screenCopy()
        for i in range(len(screenClean)):
            if screenClean[i] == "SLA":
                m = screenClean[i+2]
                if m == "Y":
                    return "Yes"
                else:
                    return "No"

    def returnType(self):
        """
        Determines the return type associated with the current RMA.

        This method scans the screen data for the keyword "RMA#" and retrieves 
        the type of return listed directly after it.

        Args:
            None.

        Returns:
            str: The return type of the RMA (e.g., "Repair", "Replace", etc.)
                if found, or "Unknown" if not found.
        """
        screenClean = self.screenCopy()

        for i in range(len(screenClean)):
            if screenClean[i] == "RMA#":
                returnType = screenClean[i+2]
                # print(f"Return Type: {returnType}")
                return returnType

    def partNum(self):
        """
        Extracts the part number from the screen data.

        This method scans the screen data for the keywords "Part" and "Number" 
        in sequence and retrieves the part number listed after them.

        Args:
            None.

        Returns:
            str: The part number if found, or "N/A" (Not Available) if not found.
        """

        screenClean = self.screenCopy()

        for i in range(len(screenClean)):
            if screenClean[i] == "Part" and screenClean[i+1] == "Number":
                partNumber = screenClean[i+3]
                # print(f"Part Number: {partNumber}")
                return partNumber

    def dateEntered(self):
        """
        Checks if the date has been entered in other section.

        This method scans the screen data for the keywords "Other"
        in sequence and retrieves whether or not a date has been entered.

        Args:
            None.

        Returns:
            bool: True if date has not been entered, False if date has not been entered.
        """
        screenClean = self.screenCopy()

        for i in range(len(screenClean)):
            if screenClean[i] == "Other:":
                dateEntered = screenClean[i+1]
                break

        if dateEntered == "\n":
            return True
        else:
            return False

class GUI(ProcessRMA):
    def __init__(self):
        """
        Initializes the GUI application and sets up the main window.

        Args:
            None.
        
        Attributes:
            root (tk.Tk): The main Tkinter root window.
            done (bool): Tracks whether the process is completed.
            rma_number_var (tk.StringVar): Stores the user-inputted RMA number.
            allow_next_step (bool): Flag to control whether the next step can proceed.
            backend (ProcessRMA): Backend instance for processing RMA-related tasks.
            current_serial (str): Keeps track of the current serial number being processed.
        """

        self.root = tk.Tk()
        self.root.title("RMA Receiving Program")
        self.root.geometry("700x800")
        self.root.resizable(False, False)
        self.root.config(bg="light blue")
        self.done = False

        # Variables
        self.rma_number_var = tk.StringVar()  # To store user-inputted RMA number
        self.allow_next_step = False  # Controls whether the next loop step can proceed
        self.backend = None          # Instance of the backend
        self.current_serial = None   # Track the current serial being processed

        # Build the GUI layout
        self.build_gui()

    def build_gui(self):
        """
        Sets up the user interface components for the GUI.

        This method creates labels, entry fields, checkboxes, and buttons for 
        interacting with the user. Additionally, it binds specific actions to keys 
        and buttons to enhance functionality.

        Args:
            None.

        Returns:
            None.
        """
        # Title
        tk.Label(self.root, text="RMA Receiving Program",
                font=("Rockwell Extra Bold", 22), bg="light blue").pack(pady=5)
        
        tk.Label(self.root, text="Created By: Amal Chalayil Sreekumar",
        font=("Rockwell Extra Bold", 13), bg = "light blue").pack(pady=15)

        tk.Label(self.root, text="Verify AS400 in Failure Analysis Main Menu Before Running",
        font=("Rockwell", 13), bg = "yellow").pack(pady=15)

        # RMA Entry Field
        tk.Label(self.root, text="Enter RMA Number:", font=("Rockwell Extra Bold", 18), bg="light blue").pack(pady=5)
        self.rma_entry = tk.Entry(self.root, textvariable=self.rma_number_var, font=("Arial", 12), width=30)
        self.rma_entry.pack()

        # Bind Enter key to start_processing() when focused on RMA entry field
        self.rma_entry.bind("<Return>", lambda event: self.start_processing())

        # Status/Instruction Label
        self.message_label = tk.Label(self.root, text="Please enter an RMA number and press 'Start'.",
                                    font=("Rockwell", 14), bg="light blue", fg="black")
        self.message_label.pack(pady=10)

        self.rmaDamaged = False  # Initialize rmaDamaged as False
        self.rmaDamaged_var = tk.BooleanVar(value=False)  # BooleanVar holds the GUI checkbox state

        def update_rmaDamaged():
            """
            Updates the RMA damaged status based on checkbox state.

            Args: None

            Returns: None
            """
            self.rmaDamaged = self.rmaDamaged_var.get()
            print(f"RMA Damaged is set to {self.rmaDamaged}")  # Debug print

        # Create the checkbox and bind it to rmaDamaged_var
        self.damaged_checkbox = tk.Checkbutton(
            self.root,
            text="RMA is Damaged",
            variable=self.rmaDamaged_var,
            onvalue=True,       # Value when checked
            offvalue=False,     # Value when unchecked
            font=("Rockwell", 14),
            bg="light blue",
            command=update_rmaDamaged  # Callback function to update value
        )
        self.damaged_checkbox.pack(pady=10)

        # Start Button to Begin the Process
        self.start_button = tk.Button(self.root, text="Start", font=("Rockwell", 14), bg="green", fg="white", command=self.start_processing)
        self.start_button.pack(pady=10)

        # Bind Enter key to start_processing() when focused on Start button
        self.start_button.bind("<Return>", lambda event: self.start_processing())

        # Button to Continue the Loop in Front-End
        self.next_button = tk.Button(self.root, text="Next Serial Number", font=("Rockwell", 14), bg="yellow", command=self.allow_next_iteration, state=tk.DISABLED)
        self.next_button.pack(pady=10)

        # Bind Enter key to allow_next_iteration() when focused on Next button
        self.next_button.bind("<Return>", lambda event: self.allow_next_iteration())

        # Quit Button
        quit_button = tk.Button(self.root, text="Quit", font=("Rockwell", 14), bg="red", fg="white", command=self.root.quit)
        quit_button.pack(pady=10)

        self.prevSerialNumInfo = tk.Label(self.root, text="Previous Serial Number Information", bg="light blue", font=("Rockwell", 14))
        self.prevSerialNumInfo.pack(pady=10)

        # Bind spacebar to move to the next serial number
        self.root.bind("<space>", lambda event: self.allow_next_iteration())


    def create_copyable_label(self, text):
        """
        Creates a text widget styled as a label that allows its content to be copied.

        Args:
            text (str): The text content to display in the widget.

        Returns:
            text_widget (tk.Text): A Text widget configured to look like a label.
        """
        text_widget = tk.Text(
            self.root,
            wrap="none",        # Disable wrapping
            height=1,           # Restrict to a single line
            width=50,           # Set max characters per line
            font=("Arial", 12), # Matching font size
            bg="lightgray",     # Match color to a label
            fg="black",
            bd=0                # No border
        )
        text_widget.insert("1.0", text)       # Insert text
        text_widget.configure(state="disabled")  # Prevent manual edits
        text_widget.pack(pady=10, padx=10, anchor="center")  # Constrain behavior with fixed padding

        # Allow copying (select the entire text when clicked)
        def enable_selection(event):
            text_widget.configure(state="normal")  # Temporarily allow interaction
            text_widget.tag_add("sel", "1.0", "end")  # Select all
            text_widget.configure(state="disabled")  # Disable it again to restore immutability

        # Bind mouse click to enable selection
        text_widget.bind("<Button-1>", enable_selection)

        return text_widget

    def start_processing(self):
        """
        Starts the backend RMA processing process, enabling barcode processing.

        Retrieves user-entered RMA data, validates it, and initializes backend processing.
        If errors occur (e.g., RMA not open), displays appropriate error messages
        in the GUI.
        """
        rma_number = self.rma_number_var.get().strip()

        if not rma_number:  # Validation: Check if RMA number is empty
            self.message_label.config(text="Error: RMA number cannot be empty.", fg="red")
            return

        try:
            self.backend = ProcessRMA(rma_number)  # Create the backend instance
            self.backend.barcodeList = self.backend.getBarcodes()

            if self.backend.barcodeList == "RMA not open":
                self.message_label.config(
                    text="Error: RMA not open or entered incorrectly. Ensure the RMA is open.",
                    fg="red",
                    bg = "white"
                )
                self.backend = None  # Reset backend instance
                self.start_button.config(state=tk.NORMAL)  # Re-enable the Start button to retry
                self.next_button.config(state=tk.DISABLED)  # Disable the Next button
                return
            
            self.message_label.config(text="RMA Serial Number Saved.", fg="blue")
            
            # Disable the start button and enable the next step button
            self.start_button.config(state=tk.DISABLED)
            self.next_button.config(state=tk.NORMAL)

            # Fetch barcodes and start the main processing loop
            self.backend.getBarcodes()
            self.root.after(100, self.main_loop)

        except Exception as e:
            # Handle any exception during initialization
            self.message_label.config(text=f"Error starting process: {e}", fg="red")

    def allow_next_iteration(self):
        """
        Allows the next iteration in the main processing loop by setting the control variable.

        Args:
            None.

        Returns:
            None.
        """
        self.allow_next_step = True


    def create_dynamic_textbox(self):
        """
        Create an editable Text widget for displaying multiline content.

        The Text widget supports:
            - Multiline content with support for wrapping.
            - Formatting, including bold text.

        The Text widget is primarily used to dynamically display information
        such as serial numbers and other RMA details.

        Args:
            None.

        Returns:
            tkinter.Text: A Text widget configured with custom styles and 
                        disabled manual editing.
        """

        dynamic_textbox = tk.Text(
            self.root,
            wrap="word",         # Enable word wrapping
            height=12,           # Height (number of lines visible)
            width=70,            # Width (number of characters visible per line)
            font=("Arial", 12),  # Default font style and size
            bg="lightgray",      # Background color
            fg="black",          # Text color
            bd=2                 # Border width
        )

        # Define a tag for bold text
        dynamic_textbox.tag_configure("bold", font=("Arial", 12, "bold"))

        dynamic_textbox.configure(state="disabled")  # Disable manual editing
        dynamic_textbox.pack(pady=10, padx=10)

        return dynamic_textbox

    def update_dynamic_textbox(self, textbox, content_dict):
        """
        Updates the content of the dynamic Text widget with formatted text.

        The content in the Text widget is cleared and replaced with new data 
        provided via a dictionary `content_dict`. Labels are displayed in bold,
        while their corresponding content is displayed in a regular font.

        Args:
            textbox (tkinter.Text): The Text widget to update.
            content_dict (dict): A dictionary where the keys are labels (e.g., 
                                "Serial Number") and their values are the 
                                associated content.

        Returns:
            None.
        """
         # Temporarily enable editing
        textbox.configure(state="normal") 
        
        # Clear the textbox
        textbox.delete("1.0", tk.END)

        for label, value in content_dict.items():
            # Insert the label as bold
            textbox.insert(tk.END, f"{label}: ", "bold")

            # Insert the value normally
            textbox.insert(tk.END, f"{value}\n")

        # Lock the Text widget to prevent manual edits
        textbox.configure(state="disabled")  


    def main_loop(self):
        """
        Main processing loop for handling barcode serial numbers.

        This method processes all serial numbers in the backend's barcode queue.
        For each serial:
        - Processes the associated backend information.
        - Sends updates to the dynamic Text widget.
        - Responds to the user pressing "Next Step" to move to the next serial.

        Stops processing when the barcode queue is empty or an "empty marker"
        (empty string) is encountered.

        Args:
            None.

        Returns:
            None.
        """

        #Breaks loop and returns if RMA is entered incorrectly or RMA is not open
        if self.backend == None: 
            return
        
        if self.backend and self.backend.barcodeList and not self.backend.barcodeList.empty():

            # Check if the user has pressed "Next Step"
            if self.allow_next_step:
                self.allow_next_step = False  # Reset the control variable

                # Get the next serial from the backend's barcode queue
                self.current_serial = self.backend.barcodeList.get()
                self.message_label.config(text=f"Processing Serial: {self.current_serial}", fg="blue")

                # Check if the empty string (marker) is reached
                if self.current_serial == " ":
                    self.message_label.config(text="Processing complete.", fg="red")
                    self.backend.barcodeList.queue.clear()  # Clear the queue to trigger `else`
                    return self.main_loop()  # Immediately transition to the `else` block

                # Simulate backend processing for the serial (extend as needed)
                list_of_windows = self.list_window_names()
                as400 = self.find_as400(list_of_windows)
                win32gui.SetForegroundWindow(as400)

                pa.hotkey(["return"])
                pa.typewrite('p')
                pa.hotkey(["return"])
                pa.typewrite('r')
                pa.hotkey(["return"]) 
                pa.typewrite("I")
                pa.typewrite(f"{self.current_serial}", interval=0.01)
                pa.hotkey(["return"])
                pa.typewrite("s")
                pa.hotkey("down")
                pa.hotkey("down")
                pa.hotkey("down")
                pa.typewrite("s")
                pa.hotkey(["return"])

                pa.hotkey('ctrl', 'a')
                pa.hotkey('ctrl', 'c')
                screen = self.root.clipboard_get()

                #In the case that to get to the RMA processing screen you need to type OK
                if "Type OK" in screen:
                    pa.typewrite("OK")

                if self.backend.dateEntered() == True:
                    self.backend.enterDate()

                isSLA = self.backend.isSLA()
                returnType = self.backend.returnType()
                partNum = self.backend.partNum()
                informationTxt = self.backend.trackRMA(self.current_serial, returnType, partNum)
                folderPath = self.backend.create_rma_folder_structure()
                if self.rmaDamaged == True:
                    damagedPath = self.backend.createDamagedRmaFolder()
                else:
                    damagedPath = "Not Damaged"
                # If the dynamic textbox doesn’t yet exist, create it
                if not hasattr(self, 'information_textbox'):
                    self.information_textbox = self.create_dynamic_textbox()

                # Update the textbox with formatted content
                info_content_dict = {
                    "Serial Number": self.current_serial,
                    "SLA": isSLA,
                    "Return Type": returnType, 
                    "Part Number": partNum,
                    "Content Written To": informationTxt,
                    "Folder Path " : folderPath,
                    "Damaged Path" : damagedPath
                }

                # Call the updated function
                self.update_dynamic_textbox(self.information_textbox, info_content_dict)
                
                # Schedule the next iteration after processing
                self.root.after(100, self.main_loop)
            else:
                # If the user hasn't pressed the "Next Step" button, check again later
                self.root.after(100, self.main_loop)
        else:
            # All barcodes processed, finish the process
            self.message_label.config(text="All barcodes processed. Process completed!", fg="green")
            self.start_button.config(state=tk.NORMAL)  # Re-enable the "Start" button
            self.next_button.config(state=tk.DISABLED)  # Disable the "Next Step" button
            list_of_windows = self.list_window_names()
            as400 = self.find_as400(list_of_windows)
            win32gui.SetForegroundWindow(as400)
            pa.hotkey(["return"])
            pa.typewrite('p')
            pa.hotkey(["return"])
            pa.typewrite('r')
            pa.hotkey(["return"]) 
            
            # Uncheck the "RMA is Damaged" checkbox after processing
            self.rmaDamaged_var.set(False)
            self.rmaDamaged = False  # Set the internal value to False as well

            pa.typewrite('e')
            pa.hotkey(["return"])

    def run(self):
        #Start the Main Loop
        self.root.mainloop()


if __name__ == "__main__":
    # Create the front-end application instance and run it
    frontend = GUI()
    frontend.run()
