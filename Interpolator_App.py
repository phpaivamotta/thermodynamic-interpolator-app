from tkinter import *
from tkinter import ttk
import pandas as pd
from scipy import interpolate

root = Tk()

root.title("Interpolator")
root.geometry("600x500")

backgroung = Frame(root, bg='#ccffff')
backgroung.place(relwidth=1, relheight=1)

frame = Frame(root, bg='black')
frame.place(relx=0.1, rely=0.05, relwidth=0.8, relheight=0.45)

# Creating warning label outside of try/except ##########
labl_parameterValueErr = Label(frame)

def interpolation(pmtr_type,pmtr,df):
    # Assigning x vector value (value of rows in parameter type's column)
    x = df[pmtr_type]
    # Getting all headers in excel sheet for given table
    cols = df.columns
    # Initiating list for values
    vals = []
    # Looping through all columns to obtain interpoalted values for all parameter types at the given parameter value
    for y in cols:
        f = interpolate.interp1d(x, df[y])
        vals.append(f(pmtr))
    return vals

# Validating parameter value entry
def vali(ll,ul,parameterValue):
    global labl_parameterValueErr
    labl_parameterValueErr.destroy()
    try:
        parameterValue = float(parameterValue)
        if parameterValue < ll or parameterValue > ul:
            labl_parameterValueErr = Label(frame, text=f'*Value out of bounds. Value should be between {ll} and {ul}.', bg='black', fg='red')
            labl_parameterValueErr.grid(row=7, columnspan=3, padx=0, pady=0)
            return
    except ValueError:
        labl_parameterValueErr = Label(frame, text='*This is not a valid number, silly billy.', bg='black', fg='red')
        labl_parameterValueErr.grid(row=7, columnspan=3, padx=0, pady=0)
        return
    else:
        return parameterValue

def buttonPressed():
    # Getting inputs
    table = my_table.get()
    parameterType = my_pmtr_t.get()
    pressure = my_press.get()
    parameterValue = my_entry.get()
    # Assigning a dataframe according to inputs
    df = pd.read_excel(file, table)
    # List parameter type
    listParameterType = df.columns

    # If a "pressure table" is selected get new dataframe
    if not pressure == "--":
        ps = df.loc[df['T'].str.contains('P', na=False)].index
        ls_p = df.iloc[ps, 0]
        # p = ls_p.iloc[p_indx - 1]
        ind = df.loc[df['T'].str.contains(pressure, na=False)].index
        i = ind[0]
        # New dataframe for the given pressure
        df = df.iloc[(i + 1):(df.last_valid_index() + 1)] if ps[ps > i].empty else df.iloc[(i + 1):ps[ps > i][0]]

    # Assigning max. and min. values for givin parameter type
    ll = min(df[parameterType])
    ul = max(df[parameterType])
    # Validating parameter value entry
    parameterValue = vali(ll, ul, parameterValue)
    # Exit function if wrong input was given
    if not parameterValue:
        return
    # Interpolating
    listInterpValues = interpolation(parameterType, parameterValue, df)
    # Creating the text for display label
    text = []
    for ptype, pvalue in zip(listParameterType, listInterpValues):
        text.append(f'{ptype} = {pvalue}')
    # Printing values to display label
    labl_results.config(text=("\n".join(text)))

def runB(*parameterValue):
    x = entryvar.get()
    y = parametervar.get()
    if x and y:
        my_button.config(state=NORMAL)
    else:
        my_button.config(state=DISABLED)

def pmtrType(e):
    table_click = my_table.get()
    df = pd.read_excel(file, table_click)
    pt = df.columns
    my_pmtr_t.config(value=(list(pt)))
    my_pmtr_t.current(0)
    if (table_click == "Table A13-SI") or (table_click == "Table A13-E") or (table_click == "Table A6-SI") or (table_click == "Table A6-E") or (table_click == "Table A7-SI") or (table_click == "Table A7-E"):  #### NEEDS UPDATING AS TABLES ADDED/TABLE NAMES CHANGE
        my_press['state'] = 'readonly'
        ps = df.loc[df['T'].str.contains('P', na=False)].index
        ls_p = df.iloc[ps, 0]
        my_press.config(value=(list(ls_p)))
        my_press.current(0)
    else:
        my_press.config(value="--")
        my_press.current(0)
        my_press['state'] = 'disabled'


# Assign spreadsheet filename: file
file = 'Thermo Tables Python.xlsx'

# Load spreadsheet: xls
xls = pd.ExcelFile(file)

# Assign sheet names list: ls_t
ls_t = xls.sheet_names

# Tables
labl_table = Label(frame, text='Choose Table:', bg='black', fg='white')
labl_table.grid(row=0, column=0, padx=15, pady=5)
my_table = ttk.Combobox(frame, state='readonly', value=ls_t)
my_table.current(0)
my_table.grid(row=1, column=0, padx=15, pady=0)

# bind combobox
my_table.bind("<<ComboboxSelected>>", pmtrType)

# Parameter type
parametervar = StringVar()
parametervar.trace("w", runB)

labl_parameterType = Label(frame, text='Choose Parameter Type:', bg='black', fg='white')
labl_parameterType.grid(row=0, column=2, padx=15, pady=5)
my_pmtr_t = ttk.Combobox(frame, state='readonly', textvariable=parametervar ,value=[""]) ####### textvariable=parametervar
my_pmtr_t.grid(row=1, column=2, padx=15, pady=0)


# Pressure box (only for some tables)
labl_pressure = Label(frame, text='Choose Pressure:', bg='black', fg='white')
labl_pressure.grid(row=3, column=1, padx=0, pady=5)
my_press = ttk.Combobox(frame, value=["--"], state = ["disabled"], width=15)
my_press.current(0)
my_press.grid(row=4, column=1, padx=0, pady=0)

# Entry for parameter value
entryvar = StringVar()
entryvar.trace("w", runB)

labl_parameterValue = Label(frame, text='Enter Parameter Value:', bg='black', fg='white')
labl_parameterValue.grid(row=5, column=1, padx=0, pady=5)
my_entry = ttk.Entry(frame, textvariable=entryvar, width=18)
my_entry.grid(row=6, column=1, padx=0, pady=0)

# Button (Run) that gets selected values
my_button = Button(frame, text = "Interpolate", state=DISABLED, command = buttonPressed)
my_button.grid(row=8, column=1, padx=0, pady=10)    # row=8 to leave room for err label at row=7

# Lower frame frame.place(relx=0.1, rely=0.1, relwidth=0.8, relheight=0.4)
lower_frame = Frame(root, bg='black', bd=5)
lower_frame.place(relx=0.1, rely=0.52, relwidth=0.8, relheight=0.45)

# Results label
labl_results = Label(lower_frame, bg='white')
labl_results.place(relwidth=1, relheight=1)

root.mainloop()