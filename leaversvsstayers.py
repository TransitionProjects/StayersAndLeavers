__author__ = "David Marienburg"

import pandas as pd
import numpy as np

from datetime import datetime
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfilename

class CreateLeaversReport:
    """
    This class will include methods that process the all entries provider,
    creating a processed report containing 5 seperate sheets:
    --leavers
    --in provider for more than 3 months to date
    --in provider for more than 6 months to date
    --in provider for more than 12 months to date
    --raw data

    Each sheet will contain the following data: Client Uid, First Name, Last
    Name, Entry Date, Exit Date, Provider Id, Months in provider
    """
    def __init__(self):
        self.entry_df = pd.read_excel(
            askopenfilename(title="Open the All Entrie Report")
        )
        self.today = datetime.today()

    def drop_extraneous_columns(self, data_frame):
        """
        Drop all columns that are not required by the output report and convert
        all date columns into datetime.date columns
        """
        # create a local copy of the data frame
        raw_data = data_frame

        # create datetime.date columns for the entry and exit dates
        raw_data["Entry Date"] = raw_data["Entry Exit Entry Date"].dt.date
        raw_data["Exit Date"] = raw_data["Entry Exit Exit Date"].dt.date

        # return the cleaned up client data
        return raw_data[[
            "Client Uid",
            "Entry Exit Provider Id",
            "Entry Date",
            "Exit Date",
            "Months Since Entry"
        ]]

    def calculate_months_since_entry(self, data_frame):
        """
        Use the np.timedelta64 method to calculate the months since entry
        """
        # create a local copy of the data frame
        raw_data = data_frame

        # fill the empty exit date fields with today's date
        raw_data["Months Since Entry"] = (
            raw_data["Entry Exit Entry Date"]-raw_data["Entry Exit Exit Date"]
        )/np.timedelta64(1, "M")

        # return the new data frame
        return raw_data

    def seperate_leavers(self, data_frame):
        """
        Slice the data_frame so that leavers and non-leavers are grouped into
        different data frames, the non-leavers have their exit date filled with
        today's date, then both data frames are passed to the
        calculate_months_since_entry and return both data frames as elements in
        a tuple.
        """
        # create a local copy of the data frame being passed
        raw_data = data_frame

        # create the leavers version of the data_frame
        leavers = raw_data[raw_data["Entry Exit Exit Date"].isna()]

        # create the stayers version of the data_frame
        stayers = raw_data[~(raw_data["Entry Exit Exit Date"].isna())]
        stayers["Entry Exit Exit Date"].fillna(self.today, axis=1, inplace=True)

        # return the leavers and stayers data frames after passing them through
        # the calculate_months_since_entry method
        return (
            self.calculate_months_since_entry(leavers),
            self.calculate_months_since_entry(stayers)
        )

    def save_data_frame(self):
        """
        Create seperate leavers and statyers data frames, initialize a writer
        object, then save the data frames to that object using slicers to create
        multiple stayers sheets.
        """
        # create the base stayers and leavers sheets
        leavers, stayers = seperate_leavers(self.entry_df)

        # strip extraneous columns using the same named method
        leavers = self.drop_extraneous_columns(leavers)
        stayers = self.drop_extraneous_columns(stayers)

        # initialize the writer object
        writer = pd.ExcelWriter(
            askopenfilename(title="Save the Processed Report"),
            engine="xlsxwriter"
        )
        leavers.to_excel(writer, sheet_name="Leavers", index=False)
        stayers[stayers["Months Since Entry"] > 3].to_excel(
            writer, sheet_name="Stayer More Than 3 Months", index=False
        )
        stayers[stayers["Months Since Entry"] > 6].to_excel(
            writer, sheet_name="Stayer More Than 6 Months", index=False
        )
        stayers[stayers["Months Since Entry"] > 9].to_excel(
            writer, sheet_name="Stayer More Than 9 Months", index=False
        )
        stayers[stayers["Months Since Entry"] > 12].to_excel(
            writer, sheet_name="Stayer More Than 12 Months", index=False
        )
        self.raw_data.to_excel(writer, sheet_name="Raw Data", index=False)
        writer.save()
