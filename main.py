import pandas as pd

# Allowing user to choose between sheets within the file
print(" 1)-FacStaff 2)-MACS 3)-MASU ")
shNumber = int(input("Which Sheet you would like to use (1,2,3): "))

# input of computer identifier
mNumber = input("Enter M number: ")

# input of new password // which will be changed within the excel file
password = input("Enter New Password: ")


# Sheet number equals itself - 1 because index starts at 0
shNumber -= 1

# Assigning our data frame
df = pd.read_excel('FacStaff.xls', sheet_name=shNumber)

# Doing a try except statement
try:

    # iterating through the names of each computer
    for i in range(0, 3000):
        # the mNumbers of the computer is assigned to a variable
        _computer = df.at[i, "PCI"]

        # The assigned variable is compared with given mNumber input
        if mNumber == _computer:

            # Locating and outputting the old password
            currentVal = df.at[i, "Password"]
            print(f"The password of {mNumber} is {currentVal}")


            # The new password replaces old password
            df.at[i, "Password"] = password

            # If the program prints the inputted password below, that means it worked how it was supposed to
            # Essentially a way to doublecheck if password change worked or not
            newVal = df.at[i, "Password"]
            print(f"Password has been updated, new password: {newVal}")


# An error due to the range of the for loop
# Thats why except statement was used here, to avoid the "not in range" error
except:
    pass
