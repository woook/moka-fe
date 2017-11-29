'''
v1.1 - AB 2017/11/28
Updated to open Access file in runtime mode

Usage:
	This script is run when a user starts Moka. It updates and opens the selected front end.
	It is invoked from a locally installed batch script.
'''

import os
import datetime
import shutil
import subprocess
import threading
from Tkinter import *
import ttk


class SelectFE(object):
	def __init__(self):
		#Create root window
		self.root = Tk()
		#Set window icon to Moka icon. In different locations on Win 10 and XP so have to check file exists.
		if os.path.exists(r"C:\Program Files (x86)\GSTT\Moka\moka.ico"):
			self.root.iconbitmap(r"C:\Program Files (x86)\GSTT\Moka\moka.ico")
		elif os.path.exists(r"C:\Program Files\GSTT\Moka\moka.ico"):
			self.root.iconbitmap(r"C:\Program Files\GSTT\Moka\moka.ico")
		self.root.title("Moka") #Window title
		self.app = Frame(self.root) #Create frame to hold content
		self.app.grid(padx=10, pady=10) #Add padding
		self.choice = IntVar() #Set choice variable to integer
		self.choice.set(2) #Default is 2 (Laboratory)
		#Create radio buttons to get users choice of front end
		self.L1 = Label(self.app, text = "Please select your level of Moka access:", font = (None, 14))
		self.L1.grid(row = 0, column = 0, columnspan = 2, sticky = W)
		self.R1 = Radiobutton(self.app, text="Clinic", variable=self.choice, value=1, font = (None, 12))
		self.R1.grid(row = 1, column = 0, columnspan = 2, sticky = W)
		self.R2 = Radiobutton(self.app, text="Lab", variable=self.choice, value=2, font = (None, 12))
		self.R2.grid(row = 2, column = 0, columnspan = 2, sticky = W)
		self.R3 = Radiobutton(self.app, text="Admin", variable=self.choice, value=3, font = (None, 12))
		self.R3.grid(row = 3, column = 0, columnspan = 2, sticky = W)
		#Add text 
		self.L2 = Label(self.app, 
						text = "If you are adding or editing information, please make sure you are logged into this computer before using Moka", 
						font = (None, 8, "italic"))
		self.L2.grid(row = 4, column = 0, columnspan = 2, sticky = W)
		#Add buttons
		self.B1 = Button(self.app, text="OK", command=self.sel, font = (None, 10), width = 10)
		self.B1.grid(row = 5, column = 0, columnspan = 1, sticky = N)
		self.B2 = Button(self.app, text="Cancel", command=self.cancel, font = (None, 10), width = 10)
		self.B2.grid(row = 5, column = 1, columnspan = 1, sticky = N)
		self.root.mainloop()

	def close(self):
		#Close interface
		self.root.destroy()

	def sel(self):
		# Capture users selection, close window and call script to update and open Moka
		selection = self.choice.get()
		self.close()
		o = OpenMoka(selection)
		o.openLatest()

	def cancel(self):
		self.close()

class waitMessage(object):
	'''
	Displays a wait message whilst new copy of front end is copied from server
	'''
	def __init__(self, selection):
		#Create text string for selected front end
		self.types = {1: "Clinic", 2: "Lab", 3: "Admin"}
		self.message = "Retrieving updated %s file. This may take several minutes to complete..." % (self.types[selection])
		#Create root window, set icon and title
		self.root = Tk() 
		if os.path.exists(r"C:\Program Files (x86)\GSTT\Moka\moka.ico"):
			self.root.iconbitmap(r"C:\Program Files (x86)\GSTT\Moka\moka.ico")
		elif os.path.exists(r"C:\Program Files\GSTT\Moka\moka.ico"):
			self.root.iconbitmap(r"C:\Program Files\GSTT\Moka\moka.ico")
		self.root.title("Moka")
		self.app = Frame(self.root) #Create frame to hold content
		self.app.grid(padx=10, pady=10) #Add padding
		#Add text set above
		self.L1 = Label(self.app, text = self.message, font = (None, 12))
		self.L1.grid()
		#Add indeterminate progress bar 
		#doesn't actually display progress but bounces back and forth so user knows 'something is happening'
		self.P1 = ttk.Progressbar(self.app, length = 100, mode = 'indeterminate')
		self.P1.grid()
		self.P1.start(15)
		self.root.mainloop()

class OpenMoka(object):
	def __init__(self, selection):
		#Locations for front end and timestamp files
		self.selection = selection
		self.locFE = {1: r"H:\Moka\clinicial.mdb",
						   2: r"H:\Moka\master.mdb",
						   3: r"H:\Moka\admin.mdb"}
		self.sourceFE = {1: r"\\gstt.local\apps\Moka\FrontEnd\clinicial.mdb",
		 					  2: r"\\gstt.local\apps\Moka\FrontEnd\master.mdb",
							  3: r"\\gstt.local\apps\Moka\FrontEnd\admin.mdb"}
		self.timestamps = {1: r"H:\Moka\clinicial.txt",
						   2: r"H:\Moka\master.txt",
						   3: r"H:\Moka\Admin.txt"}

	def waitMsgBox(self):
		# Displays a wait message box while front end is updated. Must be run on separate thread so rest of code not stalled. 
		self.w = waitMessage(self.selection)

	def retrieveFile(self):
		# Open message box informing user file is being updated. Run in a separate thread so rest of code not stalled.
		t = threading.Thread(target=self.waitMsgBox)
		t.setDaemon(True)
		t.start()
		# Copy access fe file from server, overwriting existing local copy
		if not os.path.isdir(os.path.dirname(self.locFE[self.selection])):
			os.makedirs(os.path.dirname(self.locFE[self.selection]))
		new_timestamp = os.path.getmtime(self.sourceFE[self.selection])
		shutil.copy(self.sourceFE[self.selection], self.locFE[self.selection])
		#Update timestamp file
		str_newtimestamp = datetime.datetime.fromtimestamp(new_timestamp).strftime('%d-%m-%Y %H:%M:%S')
		with open(self.timestamps[self.selection], 'w') as ts:
			ts.write(str_newtimestamp)

	def latestVersion(self):
		# Checks whether the local copy matches the latest version stored on server
		# If either the access file or timestamp file are not present locally, return False so they will be copied from server
		if not os.path.isfile(self.locFE[self.selection]) or not os.path.isfile(self.timestamps[self.selection]):
			return False
		# If timestamps don't match, return False so new copy will be made from server
		# Get modifed time from file on server
		sourceTimestamp = os.path.getmtime(self.sourceFE[self.selection])
		str_sourceTimestamp = datetime.datetime.fromtimestamp(sourceTimestamp).strftime('%d-%m-%Y %H:%M:%S')
		# Read timestamp from local text file and compare to source timestamp
		# If local file has been corrupted and this step can't complete, return False so new copy is downloaded.
		try:
			with open(self.timestamps[self.selection], 'r') as ts:
				localTimestamp = ts.readline().rstrip()
				if str_sourceTimestamp != localTimestamp:
					return False 
		except:
			return False
		return True # If all steps above complete, return True to indicate local copy is already latest file.

	def openLatest(self):
		#If local copy is not latest version, copy latest from server
		if not self.latestVersion():
			self.retrieveFile()
		# Open local copy of front end (locFE) in MSAccess runtime mode. (Runtime hides Access home ribbon etc.)
		# Get the filepath for the selected front end (i.e. master, clinical or admin) 
		filepath = self.locFE[self.selection]
		# Use subprocess to execute command to launch MSACCESS with runtime flag, replacing {filepath} with the 
		# filepath to local copy of front end
		subprocess.call(r'start MSACCESS /runtime "{filepath}"'.format(filepath=filepath), shell=True)
		exit() # Exit script so thread with messagebox also closes. 

if __name__ == "__main__":
	#Launch GUI
	s = SelectFE()



