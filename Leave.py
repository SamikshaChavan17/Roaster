import random 
import tkinter as tk
from tkinter import messagebox, ttk
from tkcalendar import Calendar
import pandas as pd
import os
from datetime import datetime

# Constants
LEAVE_CUTOFF_DATE = 25
MAX_EMERGENCY_LEAVE_PER_DAY = 2
MAX_CONTINUOUS_DAYS = 3
MAX_LEAVE_PER_DAY = 5

# File Paths
LEAVE_FILE = r"C:\Users\SC001062379\Desktop\roaster\leave_data.xlsx"
SHIFT_FILE = r"C:\Users\SC001062379\Desktop\roaster\shift_assignments_month2.xlsx"

# Constants for DataFrame columns
COL_USER = "User"
COL_LEAVE_DATE = "Leave Date"
COL_LEAVE_TYPE = "Leave Type"
COL_STATUS = "Status"

# User Lists
L1_USERS = [
    "Prasanna Gadamsetty-L1", "Gopi Gade-L1", "Mayur Gujarathi-L1", "Macharla Srinivas-L1", "Sunil-L1",
    "Aachal Thakare-L1", "Aakash Sinha-L1", "Hussain Dudekula-L1", "Paras Bharat-L1", "Mansi Arora-L1",
    "Anil-L1", "Shazeb-L1", "Bhargavi-L1"
]

L2_USERS = [
    "Ankadi Lokesh Manivarma-L2", "Vinod Kumar-L2", "Sourish Bhowmik-L2", "Shresht Jain-L2"
]

L3_USERS = [
    "Naveen Kumar-L3", "Vijay Bhashanpally-L3", "Vijay Kumar Sinha-L3"
]

# Combine all users
ALL_USERS = L1_USERS + L2_USERS + L3_USERS

# Shift Codes
SHIFTS = ["A2", "M1", "N3"] # Afternoon, Morning, Night

# Track assigned backups and users on leave
assigned_backups = {} # Format: {date: [backup_user1, backup_user2, ...]}
users_on_leave = {} # Format: {date: [user1, user2, ...]}

# Helper Functions
def is_date_in_current_month(date, current_month, current_year):
    return date.month == current_month and date.year == current_year

def is_date_in_next_month(date, next_month, next_year):
    return date.month == next_month and date.year == next_year

# Load or Create Leave Data
def load_leave_data():
    try:
        if os.path.exists(LEAVE_FILE):
            df = pd.read_excel(LEAVE_FILE)
            if COL_USER not in df.columns or COL_LEAVE_DATE not in df.columns:
                df = pd.DataFrame(columns=[COL_USER, COL_LEAVE_DATE, COL_LEAVE_TYPE, COL_STATUS])
            return df
        return pd.DataFrame(columns=[COL_USER, COL_LEAVE_DATE, COL_LEAVE_TYPE, COL_STATUS])
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load leave data: {e}")
        return pd.DataFrame(columns=[COL_USER, COL_LEAVE_DATE, COL_LEAVE_TYPE, COL_STATUS])

def save_leave_data(data):
    try:
        data.to_excel(LEAVE_FILE, index=False)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to save leave data: {e}")

# Shift Roster Generation
def assign_weekoff(user, day):
    """Assign weekoffs; Bhargavi & Paras Bharat get fixed weekends off."""
    if user in ["Bhargavi-L1", "Paras Bharat-L1"]:
        return "W" if day in [6, 7] else None # Saturday & Sunday off
    return "W" if day % 7 == 0 else None # Rotating weekoff

def create_roster(save_to_file=False):
    roster = {}
    for user in ALL_USERS:
        user_schedule = {}
        base_shift = random.choice(SHIFTS)

        for day in range(1, 31): # 30-day month
            weekoff = assign_weekoff(user, day)
            user_schedule[day] = weekoff if weekoff else base_shift

            # Ensure fixed shift for Bhargavi & Paras
            if user in ["Bhargavi-L1", "Paras Bharat-L1"]:
                user_schedule[day] = weekoff if weekoff else "M1"

        roster[user] = user_schedule

    # Save to file only if save_to_file is True
    if save_to_file:
        try:
            df = pd.DataFrame(roster).transpose()
            df.to_excel(SHIFT_FILE, index_label="Day")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save shift roster: {e}")

    return roster

# Backup for Leave
"""def get_backup_for_leave(user, leave_date, roster):
    Find an available backup for the user.
    # Determine the user category (L1, L2, or L3)
    if user.endswith("-L1"):
        # Backup must be an L1 user
        available_users = [u for u in L1_USERS if u != user and roster[u][leave_date] not in ["W", None]]
    elif user.endswith("-L2"):
        # Backup must be an L2 user
        available_users = [u for u in L2_USERS if u != user and roster[u][leave_date] not in ["W", None]]
    else:
        # For L3 users, any user can be a backup (or restrict as needed)
        available_users = [u for u in ALL_USERS if u != user and roster[u][leave_date] not in ["W", None]]

    # Filter out backups already assigned for the same date
    if leave_date in assigned_backups:
        available_users = [u for u in available_users if u not in assigned_backups[leave_date]]

    # Filter out users who are on leave on the same date
    if leave_date in users_on_leave:
        available_users = [u for u in available_users if u not in users_on_leave[leave_date]]

    return random.choice(available_users) if available_users else None

def adjust_for_leave(user, leave_date, roster):
    #Adjust shifts for leave and assign a backup.
    backup = get_backup_for_leave(user, leave_date, roster)
    if backup:
        # Update the roster
        roster[backup][leave_date] = roster[user][leave_date]
        roster[user][leave_date] = "W"

        # Track the assigned backup
        if leave_date not in assigned_backups:
            assigned_backups[leave_date] = []
        assigned_backups[leave_date].append(backup)

        # Track the user on leave
        if leave_date not in users_on_leave:
            users_on_leave[leave_date] = []
        users_on_leave[leave_date].append(user)

        messagebox.showinfo("Leave Adjustment", f"{backup} assigned as backup for {user} on {leave_date}.")
    else:
        messagebox.showerror("Error", f"No backup available for {user} on {leave_date}.")
"""
# Leave Validation
def validate_leave(user, leave_dates, leave_type): 
    """
    Validates the leave application based on the following rules:
    1. The selected leave date cannot be before today's date.
    2. A user cannot apply for leave more than once on the same date.
    3. After the 25th of the previous month, only emergency leave can be applied for.
    4. Emergency leave cannot exceed the maximum number of users per day.
    5. Users cannot apply for more than 4 continuous leave days.
    6. Only 3 L1 users and 2 L2 users can take leave on the same day.
    """
    df = load_leave_data()
    today = datetime.now().date() # Convert datetime to date for comparison

    # Rule 1: Check if the selected leave date is before today's date
    for date in leave_dates:
        if date < today:
            messagebox.showerror("Error", "Selected leave date cannot be before today's date!")
            return False

    # Rule 2: Check if the user has already applied for leave on the selected date
    for date in leave_dates:
        formatted_date = date.strftime("%d-%m-%Y")
        if not df.empty:
            existing_leave = df[(df[COL_USER] == user) & (df[COL_LEAVE_DATE] == formatted_date)]
            if not existing_leave.empty:
                messagebox.showerror("Error", f"You have already applied for leave on {formatted_date}!")
                return False

    # Get the current month and year
    current_month = today.month
    current_year = today.year

    # Determine the previous month's 25th date
    if current_month == 1: # If it's January, previous month is December of the previous year
        prev_month = 12
        prev_year = current_year - 1
    else:
        prev_month = current_month - 1
        prev_year = current_year

    prev_month_25th = datetime(prev_year, prev_month, 25).date() # 25th of the previous month

    # Rule 3: If today's date is greater than the previous month's 25th, only emergency leave is allowed
    if today > prev_month_25th:
        if leave_type == "Full Day":
            messagebox.showerror("Error", "After the 25th of the previous month, you can only apply for emergency leave.")
            return False

    # Rule 4: Emergency leave cannot exceed the maximum number of users per day
    if leave_type == "Emergency":
        for date in leave_dates:
            formatted_date = date.strftime("%d-%m-%Y")
            # Count the number of emergency leaves on the same date
            emergency_leave_count = df[(df["Leave Date"] == formatted_date) & (df["Leave Type"] == "Emergency")].shape[0]
            if emergency_leave_count >= MAX_EMERGENCY_LEAVE_PER_DAY:
                messagebox.showerror("Error", f"Max emergency leave limit reached for {formatted_date}!")
                return False

    # Rule 5: Users cannot apply for more than 3 continuous leave days
    if len(leave_dates) > MAX_CONTINUOUS_DAYS:
        messagebox.showerror("Error", "Cannot apply for more than 4 continuous leave days!")
        return False

    # Rule 6: Only 3 L1 users and 2 L2 users can take leave on the same day
    for date in leave_dates:
        formatted_date = date.strftime("%d-%m-%Y")

        # Count L1 user leaves on the same day
        l1_leave_count = df[(df["Leave Date"] == formatted_date) & (df["User"].str.endswith("-L1"))].shape[0]

        # Count L2 user leaves on the same day
        l2_leave_count = df[(df["Leave Date"] == formatted_date) & (df["User"].str.endswith("-L2"))].shape[0]

        # Rule 6a: Only 3 L1 users can take leave on the same day
        if user.endswith("-L1") and l1_leave_count >= 3:
            messagebox.showerror("Error", f"Max L1 user leave limit reached for {formatted_date}!")
            return False

        # Rule 6b: Only 2 L2 users can take leave on the same day
        if user.endswith("-L2") and l2_leave_count >= 2:
            messagebox.showerror("Error", f"Max L2 user leave limit reached for {formatted_date}!")
            return False
    
    # If all validations pass, return True
    return True

# Submit Leave
def submit_leave(): 
    user = user_dropdown.get()
    leave_type = leave_type_var.get()
    selected_date = cal.selection_get()

    if not user or not selected_date:
        messagebox.showerror("Error", "Please select a user and leave date!")
        return

    leave_dates = [selected_date] # selected_date is already a datetime.date object

    if validate_leave(user, leave_dates, leave_type):
        df = load_leave_data()
        new_entries = pd.DataFrame({
            COL_USER: [user] * len(leave_dates),
            COL_LEAVE_DATE: [date.strftime("%d-%m-%Y") for date in leave_dates],
            COL_LEAVE_TYPE: [leave_type] * len(leave_dates),
            COL_STATUS: ["Approved"] * len(leave_dates)
        })
        df = pd.concat([df, new_entries], ignore_index=True)
        save_leave_data(df)

        #roster = create_roster(save_to_file=False) # Generate roster without saving
        #adjust_for_leave(user, selected_date.day, roster) # Adjust shifts

        messagebox.showinfo("Success", "Leave application submitted successfully!")
    else:
        print("Leave validation failed.") # Debugging statement


# GUI Setup
root = tk.Tk()
root.title("Leave Management System")
root.geometry("550x500")
root.configure(bg="#e3f2fd")

frame = tk.Frame(root, padx=20, pady=20, bg="#ffffff", relief=tk.GROOVE, borderwidth=5)
frame.pack(pady=20)

tk.Label(frame, text="Leave Application", font=("Arial", 16, "bold"), fg="#1a237e", bg="#ffffff").pack(pady=10)

# Scrollable User Dropdown
scroll_frame = tk.Frame(frame, bg="#ffffff")
scroll_frame.pack()

canvas = tk.Canvas(scroll_frame, height=100, bg="#ffffff")
canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

scrollbar = ttk.Scrollbar(scroll_frame, orient=tk.VERTICAL, command=canvas.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

canvas.configure(yscrollcommand=scrollbar.set)
canvas.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

user_frame = tk.Frame(canvas, bg="#ffffff")
canvas.create_window((0, 0), window=user_frame, anchor="nw")

tk.Label(user_frame, text="Select User:", bg="#ffffff", font=("Arial", 12)).pack()
user_dropdown = ttk.Combobox(user_frame, values=ALL_USERS, font=("Arial", 12))
user_dropdown.pack()

tk.Label(frame, text="Select Leave Date:", bg="#ffffff", font=("Arial", 12)).pack()
cal = Calendar(frame, selectmode="day", year=2025, month=2, day=1, background="#81d4fa", foreground="black", borderwidth=2)
cal.pack()

tk.Label(frame, text="Leave Type:", bg="#ffffff", font=("Arial", 12)).pack()
leave_type_var = tk.StringVar(value="Full Day")
leave_type_dropdown = ttk.Combobox(frame, textvariable=leave_type_var, values=["Full Day", "Half Day", "Emergency"], font=("Arial", 12))
leave_type_dropdown.pack()

apply_btn = tk.Button(frame, text="Apply Leave", command=submit_leave, bg="#4CAF50", fg="white", font=("Arial", 12, "bold"))
apply_btn.pack(pady=10)

root.mainloop()
