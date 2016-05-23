# Import all the required modules
import sys
import os
import xlrd
import xlwt
import pandas as pd
import datetime as dt
from hammock import Hammock as GendreAPI

gendre_api = GendreAPI("http://api.namsor.com/onomastics/api/json/gendre")
# function to return gender based on first + last name combo


def return_gender(first_name, last_name):
    try:
        response = gendre_api(first_name, last_name).GET()
        return response.json().get("gender")
    except ValueError:
        return "Manually Review"
# function to assign gender to specified column in dataframe when Unknown
# or blank


def assign_gender(df, col):
    for i in df.index:
        if str(df[col][i]).lower() in ("unknown", "nan"):
            fn = col.replace("GENDER", "FIRST")
            ln = col.replace("GENDER", "LAST")
            if str(df[ln][i]).lower() != "nan":
                gender = return_gender(str(df[fn][i]), str(df[ln][i]))
                new_gender_value = ""
                if gender == "male":
                    new_gender_value = "Male"
                elif gender == "female":
                    new_gender_value = "Female"
                else:
                    new_gender_value = "Manually Verify"
                df[col][i] = new_gender_value
            else:
                pass
        else:
            pass

# function to flag entries where api and dataframe disagree on gender


def gender_doublechecker(df, col):
    for i in df.index:
        if str(df[col][i]) != "Manually Review":
            api_gender = return_gender(str(df[col.replace('GENDER', 'FIRST')][
                                       i]), str(df[col.replace('GENDER', 'LAST')][i]))
            if str(api_gender).lower() != str(df[col][i]).lower():
                df[col][i] = "Manually Review"
            else:
                pass
        else:
            pass

# function to change "Miss" to "Ms."


def miss_to_ms(df, col):
    for i in df.index:
        if str(df[col][i]) == "Miss":
            df[col][i] = "Ms."
        else:
            pass

# function to create title for singles


def title_singles(df):
    for i in df.index:
        if str(df["HHM1 TITLE"][i]).lower() == "nan":
            if str(df["HHM2 LAST"][i]).lower() == "nan":
                title = ""
                if str(df["HHM1 GENDER"][i]) == "Male":
                    title = "Mr."
                elif str(df["HHM1 GENDER"][i]) == "Female":
                    title = "Ms."
                else:
                    pass
                df["HHM1 TITLE"][i] = title
            else:
                pass
        else:
            pass


# function to create titles for cohabitants
def title_cohab(df):
    for i in df.index:
        # couple defined by hhm2 last not null
        if str(df["HHM2 LAST"][i]).lower() != "nan":
            # first treat hhm1. if man: Mr. if woman: if same last name (and
            # different gender) Mrs. if same gender or different last name Ms.
            if str(df["HHM1 TITLE"][i]).lower() == "nan":
                m1_title = ""
                if str(df["HHM1 GENDER"][i]).lower() == "male":
                    m1_title = "Mr."
                elif str(df["HHM1 GENDER"][i]).lower() == "female":
                    if str(df["HHM1 LAST"][i]).lower() == str(df["HHM2 LAST"][i]).lower():
                        if str(df["HHM2 GENDER"][i]).lower() == "male":
                            m1_title = "Mrs."
                        elif str(df["HHM2 GENDER"][i]).lower() == "female":
                            m1_title = "Manually Review"
                        else:
                            pass
                    elif str(df["HHM1 LAST"][i]).lower() != str(df["HHM2 LAST"][i]).lower():
                        m1_title = "Ms."
                    else:
                        pass
                else:
                    pass
                df["HHM1 TITLE"][i] = m1_title
            # treat hhm2 the same
            if str(df["HHM2 TITLE"][i]).lower() == "nan":
                m2_title = ""
                if str(df["HHM2 GENDER"][i]).lower() == "male":
                    m2_title = "Mr."
                elif str(df["HHM2 GENDER"][i]).lower() == "female":
                    if str(df["HHM1 LAST"][i]).lower() == str(df["HHM2 LAST"][i]).lower():
                        if str(df["HHM1 GENDER"][i]).lower() == "male":
                            m2_title = "Mrs."
                        elif str(df["HHM1 GENDER"][i]).lower() == "female":
                            m2_title = "Manually Review"
                        else:
                            pass
                    elif str(df["HHM1 LAST"][i]).lower() != str(df["HHM2 LAST"][i]).lower():
                        m2_title = "Ms."
                    else:
                        pass
                else:
                    pass
                df["HHM2 TITLE"][i] = m2_title
            else:
                pass
        else:
            pass


def fix_wrong_ms_title(df):
    for i in df.index:
        fix = ""
        if str(df["HHM1 TITLE"][i]) == "Ms." and str(df["HHM1 LAST"][i]).lower() == str(df["HHM2 LAST"][i]).lower() and df["HHM1 GENDER"][i] != df["HHM2 GENDER"][i]:
            fix = "Mrs."
            df["HHM1 TITLE"][i] = fix
        elif str(df["HHM2 TITLE"][i]) == "Ms." and str(df["HHM2 LAST"][i]).lower() == str(df["HHM1 LAST"][i]).lower() and df["HHM1 GENDER"][i] != df["HHM2 GENDER"][i]:
            fix = "Mrs."
            df["HHM2 TITLE"][i] = fix
        else:
            pass

# function to populate 'FINAL SALUTATION' column


def add_salutation(df, fs):
    for i in df.index:
        sp = " "
        df[fs][i] == ""
        m1_first = str(df["HHM1 TITLE"][i]) + sp + str(df["HHM1 FIRST"][i]) + sp + str(df["HHM1 LAST"][i]) + \
            " and " + str(df["HHM2 TITLE"][i]) + sp + \
            str(df["HHM2 FIRST"][i]) + sp + str(df["HHM2 LAST"][i])
        m2_first = str(df["HHM2 TITLE"][i]) + sp + str(df["HHM2 FIRST"][i]) + sp + str(df["HHM2 LAST"][i]) + \
            " and " + str(df["HHM1 TITLE"][i]) + sp + \
            str(df["HHM1 FIRST"][i]) + sp + str(df["HHM1 LAST"][i])
        special = ('Dr.', 'Reverend', 'Professor',
                   'Pastor', 'Rabbi', 'The Honorable')
        # if single, simply add title to name (when title is known)
        if str(df["HHM2 LAST"][i]).lower() == "nan":
            if str(df["HHM1 TITLE"][i]).lower() not in ("nan", ""):
                sal = str(df["HHM1 TITLE"][i]) + sp + \
                    str(df["HHM1 FIRST"][i]) + sp + str(df["HHM1 LAST"][i])
            else:
                pass
        # if shared household, create title based on suggested logic
        elif str(df["HHM2 LAST"][i]).lower() != "nan":
            # if either gender is manual review, make that the address
            if str(df["HHM2 LAST"][i]) != "Manually Review" and str(df["HHM1 GENDER"][i]) != "Manually Review":
                # same gender gets manually reviewed
                if str(df["HHM1 GENDER"][i]) == str(df["HHM2 GENDER"][i]):
                    sal = "Manually Review"
                elif str(df["HHM1 GENDER"][i]) != str(df["HHM2 GENDER"][i]):  # different genders
                    # woman comes first, unless special title is present
                    if str(df["HHM1 TITLE"][i]) not in special and str(df["HHM2 TITLE"][i]) not in special:
                        if str(df["HHM1 GENDER"][i]) == "Female":
                            sal = m1_first
                        elif str(df["HHM2 GENDER"][i]) == "Female":
                            sal = m2_first
                        else:
                            pass
                    elif str(df["HHM1 TITLE"][i]) in special:
                        sal = m1_first
                    elif str(df["HHM2 TITLE"][i]) in special:
                        sal = m2_first
                    else:
                        pass
                else:
                    pass
            else:
                sal = "Manually Review"
        else:
            pass
        df[fs][i] = sal

# function to write transformed data frame to output Excel File


def write_excel(df, output):
    o = output + ".xlsx"
    writer = pd.ExcelWriter(o)
    df.to_excel(writer, "Sheet1")
    writer.save()
    print "All done!"

print dt.datetime.now()
#excel_sheet = pd.ExcelFile("/users/mnaveed/Personal/salutations/name_file.xlsx")
#sheet = "Sheet1"
excel_sheet = pd.ExcelFile(sys.argv[-3])
sheet = sys.argv[-2]
data_frame = excel_sheet.parse(
    sheet, skiprows=0, index_col=None, na_values=["NA"])
assign_gender(data_frame, "HHM1 GENDER")
assign_gender(data_frame, "HHM2 GENDER")
title_singles(data_frame)
miss_to_ms(data_frame, "HHM1 TITLE")
miss_to_ms(data_frame, "HHM2 TITLE")
title_cohab(data_frame)
fix_wrong_ms_title(data_frame)
# title_cohab_2(data_frame) MADE IT WORSE
if "FINAL SALUTATION" not in data_frame.columns:
    data_frame["FINAL SALUTATION"] = ""
else:
    pass
add_salutation(data_frame, "FINAL SALUTATION")
#output = "/users/mnaveed/Personal/salutations/modded_msfix_output"
output = sys.argv[-1]
write_excel(data_frame, output)
print dt.datetime.now()
