# -*- coding: utf-8 -*-

from tkinter import Label, Button, ttk, StringVar, Tk, Frame
from tkinter import filedialog
from tkinter import messagebox
from pathlib import Path

import os
import pandas as pd
import numpy as np
pd.options.display.max_colwidth = None


class Win(Tk):
    def __init__(self, master=None):
        Tk.__init__(self, master)
        main_bg_color = "#fff"
        control_width = 20
        pad_x = 0
        pad_y = 2
        self.configure(bg=main_bg_color)
        self.geometry("1000x500")
        self.title("File Explorer")
        self.frame = Frame(self)
        self.frame.configure(background=main_bg_color)
        self.frame.pack(pady=5, padx=5)
        # Global variable
        self.filenames = []
        self.logF = None
        self.badDefects = 0

        # File Explorer label
        self.label_file_explorer = Label(self.frame,
                                         width=100,
                                         height=4,
                                         fg="blue",
                                         text="Click Browse Files and navigate to the Invision csv then select it.")
        self.label_file_explorer.grid(row=0, column=0, padx=pad_x, pady=5)
        # Version choosen Combobox
        self.version = StringVar()
        self.version_choosen = ttk.Combobox(self.frame,
                                            width=control_width + 1,
                                            textvariable=self.version)
        self.version_choosen['values'] = ('V7', 'V6')
        self.version_choosen.current(0)
        self.version_choosen.grid(row=2, column=0, padx=pad_x, pady=pad_y)
        # Button File Browser
        button_explore = Button(self.frame,
                                width=control_width,
                                text="Browse Files",
                                command=self.browseFiles)
        button_explore.grid(row=1, column=0, padx=pad_x, pady=pad_y)
        # Button Call Process
        button_exit = Button(self.frame,
                             width=control_width,
                             text="Create Rally File",
                             command=self.fileCheck)
        button_exit.grid(row=3, column=0, padx=pad_x, pady=pad_y)

    def browseFiles(self):
        self.filenames = filedialog.askopenfilenames(initialdir="/",
                                                     filetypes=[("Comma files (csv)", "csv")],
                                                     title="Select a File")
        correctExtension = True
        displayNames = ''
        for file in self.filenames:
            extension = file[-4:]
            if extension != '.csv':
                correctExtension = False
            displayNames += file + '\n'
            if correctExtension:
                self.label_file_explorer.configure(fg="blue")
                self.label_file_explorer.configure(text=f"Click Create Rally File to build from: {displayNames}")
            else:
                self.label_file_explorer.configure(fg="red")
                self.label_file_explorer.configure(text="File is not a csv (comma space delimited) file. Click Browse files to select a csv.")

    def fileCheck(self):
        for file in self.filenames:
            extension = file[-4:]
            if extension != '.csv':
                messagebox.showinfo(title="File Error",
                                    message=f"{file} is not valid as csv.")
                return
        version_selected = self.version.get()
        if version_selected == "V7":
            self.v7_process()
            messagebox.showinfo(title="program Complete",
                                message="The Rally import file has been created.")
            self.destroy()
        elif version_selected == "V6":
            self.v6_process()
            messagebox.showinfo(title="program Complete",
                                message="The Rally import file has been created.")
            self.destroy()
        else:
            messagebox.showinfo(title="Invalid Version",
                                message="Invalid Version, version should be V7 or V6 ")
            print("Version is not valid")

    def getTextValues(self, row):
        """
        the getTextvalues() functions is used in combination with the
        zipfunction in order to create multiple new columns in the DataFrame
        based on the text values that were extracted

        Parameters
        ----------
        row : TYPE
            DESCRIPTION.

        Returns
        -------
        None.

        """
        cmmentLines = row['Comment'].split("\n")
        commentDict = {'defect type': None,
                       'defect name': None,
                       'usaa guideline': None,
                       'defect decscription': None,
                       'impact': None,
                       'priority': None,
                       'severity': None,
                       'library': None}

        # what if the same header twice
        cleanDefect = True
        for elt in cmmentLines:
            if (len(elt.split(":", 1)) == 2):
                key, value = elt.split(":", 1)
                if key.strip().lower() in ['defect type', 'defect name', 'usaa guideline', 'defect description', 'impact', 'severity', 'library', 'priority']:
                    if not value.strip():
                        if cleanDefect:
                            self.logF.write("\n\n"+row['Screen Name'] + " " + str(row['Conversation']) + " " + row['Comment By'])
                        cleanDefect = False
                        self.logF.write("\n  Field present, no data: " + key.strip() + ": " + value.strinp()) #datamissing, all but guideline is required
                        continue

                    commentDict[key.strip().lower()] = value.strip()
                else:
                    key, value = elt.split(":", 1)
                    if cleanDefect:
                        self.logF.write("\n\n"+row['Screen Name'] + " " + str(row['Conversation']) + " " + row['Comment By'])
                    cleanDefect = False
                    self.logF.write("\n Invalid data field: " + key.strip())
                    continue
        for key, value in commentDict.items():
            if not value:
                if cleanDefect:
                    self.logF.write("\n\n"+row['Screen Name'] + " " + str(row['Conversation']) + " " + row['Comment By'])
                cleanDefect = False
                self.logF.write("\n No value for: "+key)
                continue

        if not cleanDefect:
            self.badDefects += 1
            return commentDict['defect type'], commentDict['defect name'], commentDict['usaa guideline'], commentDict['defect description'], commentDict['impact'], commentDict['priority'], commentDict['severity'], commentDict['library']

    def getTextValues_v7(self, row):
        """
        the getTextvalues() functions is used in combination with the
        zipfunction in order to create multiple new columns in the DataFrame
        based on the text values that were extracted

        Parameters
        ----------
        row : TYPE
            DESCRIPTION.

        Returns
        -------
        None.

        """
        cmmentLines = row['Comment'].split("\n")
        commentDict = {'defect type': None,
                       'defect name': None,
                       'usaa guideline': None,
                       'defect decscription': None,
                       'impact': None,
                       'priority': None,
                       'severity': None,
                       'library': None}

        # what if the same header twice
        cleanDefect = True
        for elt in cmmentLines:
            if (len(elt.split(":", 1)) == 2):
                key, value = elt.split(":", 1)
                if key.strip().lower() in ['defect type', 'defect name', 'usaa guideline', 'defect description', 'impact', 'severity', 'library', 'priority']:
                    if not value.strip():
                        if cleanDefect:
                            self.logF.write("\n\n"+row['Screen Name'] + " " +
                                            row['Comment By'])
                        cleanDefect = False
                        self.logF.write("\n  Field present, no data: " + key.strip() + ": " + value.strinp()) #datamissing, all but guideline is required
                        continue

                    commentDict[key.strip().lower()] = value.strip()
                else:
                    key, value = elt.split(":", 1)
                    if cleanDefect:
                        self.logF.write("\n\n"+row['Screen Name'] + " " +
                                        row['Commented By'])
                    cleanDefect = False
                    self.logF.write("\n Invalid data field: " + key.strip())
                    continue
        for key, value in commentDict.items():
            if not value:
                if cleanDefect:
                    self.logF.write("\n\n"+row['Screen Name'] + " " +
                                    row['Commented By'])
                cleanDefect = False
                self.logF.write("\n No value for: "+key)
                continue

        if not cleanDefect:
            self.badDefects += 1
            return commentDict['defect type'], commentDict['defect name'], commentDict['usaa guideline'], commentDict['defect description'], commentDict['impact'], commentDict['priority'], commentDict['severity'], commentDict['library']

    def v7_process(self):
        """
        This function generate for new version of program.
        Upgrade from v6 to v7.
        In the version 7:
            1. The column 'Conversation' was removed.
            2. Conversation Type replaced by Comment Type
            3. Comment By -> Commented By
        Returns
        -------
        None.

        """
        for filename in self.filenames:
            self.badDefects = 0
            print(filename)
            p = Path(filename)
            fileDir = p.parent
            print(fileDir)
            # filename = r"C:/Users/ThanhHome/Documents/GitHub/pandas-test/InVision_v7.csv"
            inVisionDF = pd.read_csv(filename)
            if inVisionDF.shape[0] == 0:
                messagebox.showinfo(title="CSV is empty",
                                    message=f"The CSV has no row {filename}")
                continue
            if "Conversation" in inVisionDF.columns:
                messagebox.showinfo(title="CSV Version",
                                    message=f"The CSV is formated in V6.\nThis file will be skipped.\n {filename}")
                continue
            prototypeNm = inVisionDF['Prototype'].values[0].replace(" ", "_")
            outputFile = os.path.join(fileDir, f"{prototypeNm}_Rally.csv")
            print(outputFile)
            # Open a file to use as log
            logFileName = os.path.join(fileDir, f"{prototypeNm}_Log.txt")
            self.logF = open(logFileName, 'a')
            self.logF.write("Start Log for " + outputFile)
            # Create unique id for each conversation
            inVisionDF['ConversationId'] = inVisionDF['Screen Name']
            defectList = inVisionDF['Comment'].str.slice(stop=6) == 'DEFECT'
            defectDF = inVisionDF[defectList]
            print(defectDF)
            # Here we are defining a function to extract text values from the "Comment" column
            # Have set requirements for format and content of this field in Invision when logging a defect
            defectDF['DefectType'], defectDF['DefectName'], defectDF ['DefectGuideline'], defectDF['DefectDescrip'], defectDF['DefectImpact'], defectDF['DefectPriority'], defectDF['DefectSeverity'], defectDF['DefectLibrary'] = zip(*defectDF.apply(self.getTextValues_v7, axis=1))
            # defectDF.head()
            # create DF from other DF
            rallyDF = defectDF.copy()
            rallyDF['Expedite'] = "FALSE"
            rallyDF['Ready'] = "TRUE"
            rallyDF['Name'] = rallyDF['DefectType'] + ' | ' + rallyDF['DefectName']
            rallyDF['Description'] = '<b>Description: </b>' + rallyDF['DefectDescrip'] + '<br><b>Impact: <b>' +rallyDF['DefectImpact']
            rallyDF['Notes'] = '<b>Prototype: </b>' + rallyDF['Prototype'] + '<br><b>Screen: </b>' +  rallyDF['Screen Name'] + '<br><b>Conversation: </b>' + rallyDF['ConversationId'].astype('str') + '<br><a href=' + rallyDF['Console URL']+'>Link to Comment</a>'
            rallyDF['Schedule State'] = "Defined"
            rallyDF['State'] = "Submitted"
            rallyDF['Submitted By'] = rallyDF['Commented By']
            rallyDF['Blocked'] = "FALSE"
            rallyDF['Environment'] = "None"
            rallyDF['Priority'] = rallyDF['DefectPriority']
            rallyDF['Severity'] = rallyDF['DefectSeverity']
            rallyDF['Affects Doc'] = "FALSE"
            rallyDF['Release Note'] = "FALSE"
            rallyDF['Resolution'] = "None"
            # color set based on priority – pink Must Resolve, yellow optional
            colorCondition = [
              (rallyDF['DefectPriority'] == 'Must Resolve'),
              (rallyDF['DefectPriority'] == 'Optional'),
            ]
            rallyDF['Display Color'] = np.select(colorCondition,
                                                 ['#df1a7b', '#fce205'],
                                                 default="None")

            print(rallyDF.head())

            rallyDF[['Display Color', 'Expedite', 'Ready', 'Name',
                     'Description', 'Notes',
                     'Schedule State', 'State', 'Submitted By', 'Blocked',
                     'Environment',
                     'Severity', 'Priority', 'Affects Doc', 'Release Note',
                     'Resolution']].to_csv(outputFile, index=False)

            # Check how many comments for each conversation
            nrConversations = len(pd.unique(inVisionDF['ConversationId']))
            nrComments = len(inVisionDF)
            nrDefectComments = len(rallyDF)
            nrNonDefectComments = nrComments - nrDefectComments

            self.logF.write("\n\nNumber of Unique Conversations: " + str(nrConversations))
            self.logF.write("\nNumber of Comments: " + str(nrComments))
            self.logF.write("\n\nTotal Non-Defect Comments: " + str(nrNonDefectComments))
            self.logF.write("\n\nTotal Defect Comments: " + str(nrDefectComments))
            self.logF.write("\n    Defect Comments with Issues: " + str(self.badDefects))
            self.logF.close()

    def v6_process(self):
        """
        In testing found logging will not wrk on Mac as requireds
        admin privileges.
        import logging
        logging.basisConfig(filenamme='InVisionExport.log'. level=logging.INFO
        Here we are defining a function to extract text values from the
        "Comment" column
        Have set requirements for format and content of this field in
        Invision when logging a defect

        Returns
        -------
        None.

        """
        for filename in self.filenames:
            self.badDefects = 0
            print(filename)
            p = Path(filename)
            fileDir = p.parent
            print(fileDir)
            inVisionDF = pd.read_csv(filename)
            if inVisionDF.shape[0] == 0:
                messagebox.showinfo(title="CSV is empty",
                                    message=f"The CSV has no row {filename}")
                continue
            if "Conversation" not in inVisionDF.columns:
                messagebox.showinfo(title="CSV Version",
                                    message=f"The CSV is formated in V7.\nThis file will be skipped.\n {filename}")
                continue
            prototypeNm = inVisionDF['Prototype'].values[0].replace(" ", "_")
            outputFile = os.path.join(fileDir, f"{prototypeNm}_Rally.csv")
            print(outputFile)
            # Open a file to use as log
            logFileName = os.path.join(fileDir, f"{prototypeNm}_Log.txt")
            self.logF = open(logFileName, 'a')
            self.logF.write("Start Log for " + outputFile)
            # Create unique id for each conversation
            inVisionDF['ConversationId'] = inVisionDF['Screen Name'] + inVisionDF['Conversation'].astype('str')
            screenDF = inVisionDF.groupby('ConversationId').count()[['Conversation']]
            screenDF.to_string(self.logF)
            defectList = inVisionDF['Comment'].str.slice(stop=6) == 'DEFECT'
            defectDF = inVisionDF[defectList]
            print(defectDF)
            # Here we are defining a function to extract text values from the "Comment" column
            # Have set requirements for format and content of this field in Invision when logging a defect
            defectDF['DefectType'], defectDF['DefectName'], defectDF ['DefectGuideline'], defectDF['DefectDescrip'], defectDF['DefectImpact'], defectDF['DefectPriority'], defectDF['DefectSeverity'], defectDF['DefectLibrary'] = zip(*defectDF.apply(self.getTextValues, axis=1))
            # defectDF.head()
            # create DF from other DF
            rallyDF = defectDF.copy()
            rallyDF['Expedite'] = "FALSE"
            rallyDF['Ready'] = "TRUE"
            rallyDF['Name'] = rallyDF['DefectType'] + ' | ' + rallyDF['DefectName']
            rallyDF['Description'] = '<b>Description: </b>' + rallyDF['DefectDescrip'] + '<br><b>Impact: <b>' +rallyDF['DefectImpact']
            rallyDF['Notes'] = '<b>Prototype: </b>' + rallyDF['Prototype'] + '<br><b>Screen: </b>' +  rallyDF['Screen Name'] + '<br><b>Conversation: </b>' + rallyDF['Conversation'].astype('str') + '<br><a href=' + rallyDF['Console URL']+'>Link to Comment</a>'
            rallyDF['Schedule State'] = "Defined"
            rallyDF['State'] = "Submitted"
            rallyDF['Submitted By'] = rallyDF['Comment By']
            rallyDF['Blocked'] = "FALSE"
            rallyDF['Environment'] = "None"
            rallyDF['Priority'] = rallyDF['DefectPriority']
            rallyDF['Severity'] = rallyDF['DefectSeverity']
            rallyDF['Affects Doc'] = "FALSE"
            rallyDF['Release Note'] = "FALSE"
            rallyDF['Resolution'] = "None"
            # color set based on priority – pink Must Resolve, yellow optional
            colorCondition = [
              (rallyDF['DefectPriority'] == 'Must Resolve'),
              (rallyDF['DefectPriority'] == 'Optional'),
            ]
            rallyDF['Display Color'] = np.select(colorCondition,
                                                 ['#df1a7b', '#fce205'],
                                                 default="None")

            print(rallyDF.head())

            rallyDF[['Display Color', 'Expedite', 'Ready', 'Name',
                     'Description', 'Notes',
                     'Schedule State', 'State', 'Submitted By', 'Blocked',
                     'Environment',
                     'Severity', 'Priority', 'Affects Doc', 'Release Note',
                     'Resolution']].to_csv(outputFile, index=False)

            # Check how many comments for each conversation
            nrConversations = len(pd.unique(inVisionDF['ConversationId']))
            nrComments = len(inVisionDF)
            nrDefectComments = len(rallyDF)
            nrNonDefectComments = nrComments - nrDefectComments

            self.logF.write("\n\nNumber of Unique Conversations: " + str(nrConversations))
            self.logF.write("\nNumber of Comments: " + str(nrComments))
            self.logF.write("\n\nTotal Non-Defect Comments: " + str(nrNonDefectComments))
            self.logF.write("\n\nTotal Defect Comments: " + str(nrDefectComments))
            self.logF.write("\n    Defect Comments with Issues: " + str(self.badDefects))
            self.logF.close()


if __name__ == '__main__':
    win = Win()
    win.mainloop()
    del win
