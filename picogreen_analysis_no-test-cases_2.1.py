# Processing Picogreen standard curve results measured by a plate reader
# Author Thomas Heigl
# python 3.9


import xlrd
from matplotlib import pyplot as plt
import numpy as np


def find_row(a_file, a_cell, a_dict, a_column=0):
    """
    Searches for the row index of the named input cell name in a given column.
    Standard column is set to index = 0 (first column).
    The targeted row is accessed and a list of the row's values, convertable to float, is returned.

    :param a_file: name of the Excel file with .xls ending
    :type a_file: str

    :param a_cell: name of a cell in the Excel sheet
    :type a_cell: str, int, float

    :param a_dict: dictionary to be updated with the name of the cell as key and the list as the value
    :type a_dict: dict

    :param a_column: column index of an Excel sheet. index=0 represents the first sheet
    :type a_column: int

    :return: updated dictionary with added item
    :rtype: dict
    """
    # argument correctness
    if type(a_file) != str or type(a_column) != int or a_column < 0:
        raise TypeError("The arguments for this function have the wrong type and/or wrong target index.")

    # opening the correct file and sheet
    try:
        wb = xlrd.open_workbook(a_file)
    except FileNotFoundError:
        raise FileNotFoundError("File name not found: '" + str(a_file) + "'")

    if wb.sheet_names()[0] != "Input":
        raise ValueError("'Input' sheet has to be at the first position.")

    try:
        sheet = wb.sheet_by_index(0)
    except IndexError:
        raise IndexError("Sheet not found.")

    # make a list of row names in selected column to access specific rows later.
    # if user adds one replicate (+1 row) program still works
    info = [sheet.cell_value(row, a_column) for row in range(sheet.nrows)]

    # making sure if user deleted name of a row or duplicated a name of a row to tell him something went wrong
    one_appearance = info.count(a_cell)
    if one_appearance < 1:
        raise ValueError("Searched cell name was not found in the selected column: '" + str(a_cell) + "'")
    elif one_appearance > 1:
        raise ValueError("Searched cell name appears more often in selected column: '" + str(a_cell)
                         + "' Should only appear once.")

    # access the wanted row and put its numbers as a list into a dictionary
    index_row = info.index(a_cell)
    lst = sheet.row_values(index_row)
    lst_clean = [i for i in lst if type(i) == float]
    if len(lst_clean) == 0:     # in case user forgot to put in the results in the excel sheet
        raise ValueError("No numbers were found in named row: '" + str(a_cell) + "'")
    a_dict[a_cell] = lst_clean
    return a_dict       # nice discovery. here, i do not need a return as dictionary is updated globally anyways.


def subtract_blank(a_dict, a_key):
    """
    Takes a dictionary and looks for the key to access its value (a list),
    which is then used to subtract the blank value of the measurement from.

    :param a_dict: a dictionary with items of which one item needs to have the key "Mean Blank"
    :type a_dict: dict

    :param a_key: name of a cell, which was saved as key in a dictionary
    :type a_key: str

    :return: updated dictionary
    :rtype: dict
    """
    # argument correctness
    if type(a_dict) != dict or (type(a_key) != str and type(a_key) != int and type(a_key) != float):
        raise TypeError("Wrong type of arguments used to do blank subtraction.")
    elif not(a_key in a_dict):
        raise ValueError("Your key was not found in the dictionary: '" + str(a_key) + "'")
    if len(a_dict[a_key]) == 0:
        raise ValueError("The value of the key has an empty list: '" + str(a_key) + "'")

    # updating dictionary by subtracting the blank value.
    # The dictionary key for the blank value is fixed here
    # so that I do not need to handle with it in the main part
    update = []
    for i in a_dict[a_key]:
        for j in a_dict["Mean Blank"]:
            update.append(i - j)
    a_dict[a_key] = update
    return a_dict


def clean_two_lsts(a_list, b_list):
    """
    Takes two lists as an argument to make sure they are equal in length and only have numbers as values,
    else an error is raised.
    A tuple of two list is returned if no error was raised.
    :param a_list: first list
    :type a_list: list

    :param b_list: second list
    :type b_list: list

    :return: slope and y value at x = 0
    :rtype: tuple
    """
    # correct arguments
    if type(a_list) != list or type(b_list) != list:
        raise TypeError("Input needs to be two lists.")

    if len(a_list) == 0 or len(b_list) == 0:
        raise ValueError("At least one list has no values: '" + str(a_list) + "', '" + str(b_list) + "'")

    if len(a_list) != len(b_list):
        raise IndexError("These two lists are not equal in length: " + str(a_list) + " and " + str(b_list))

    # clean and put list in tuple
    cln_a_list = []
    for i in a_list:
        try:
            cln_a_list.append(float(i))
        except ValueError:
            raise ValueError("List's values need to consist of numbers: " + str(a_list))
    cln_b_list = []
    for i in b_list:
        try:
            cln_b_list.append(float(i))
        except ValueError:
            raise ValueError("List's values need to consist of numbers: " + str(b_list))
    return cln_a_list, cln_b_list


def calculate_linear_equation(x, y):
    """
    In a cartesian coordinate system with x and y coordinates:
    calculates the the slope (=a1) and intercept (=a0) of the linear equation: y = a0 + a1 * x
    slope: a1 =(sum([xi*] * [yi*])) / sum([xi*]²)
    xi* = xi - xmean    -> x coordinate at index position i of a list subtracted by the mean of all x in that list
    yi* = yi - ymean    -> y coordinate at index position i of a list subtracted by the mean of all y in that list
    intercep: a0 = y - a1 * x

    :param x: a list of different standard concentrations [pg/µl] as int or float values
    :type x: list

    :param y: a list of different signal intensities for standard concentrations
                [artificial units, a.u.]) as int or float values
    :type y: list

    :return: slope and intercept
    :rtype: list
    """
    cleaned_tpl_xi_yi = clean_two_lsts(x, y)

    # calculation for x values
    diff_mean_x = []  # x* of slope-formula
    x_sum = sum(cleaned_tpl_xi_yi[0])
    x_mean = x_sum / len(cleaned_tpl_xi_yi[0])
    for i in cleaned_tpl_xi_yi[0]:  # new list for x* is made by calculating the difference of each
                                    # xi-value to the mean of all x-values
        diff_mean_x.append(i - x_mean)

    squared_diff_mean_x = []  # xi*^2
    for i in diff_mean_x:
        squared_diff_mean_x.append(i ** 2)

    # calculations for y values
    diff_mean_y = []  # y* from slope-formula
    y_sum = sum(cleaned_tpl_xi_yi[1])
    y_mean = y_sum / len(cleaned_tpl_xi_yi[1])
    for i in cleaned_tpl_xi_yi[1]:
        diff_mean_y.append(i - y_mean)

    # multiply x* with y*
    mult_diff_mean_x_y = [diff_mean_x[i] * diff_mean_y[i] for i in range(len(diff_mean_x))]  # (xi* * yi*)

    # calculate slope: a1 =(xi* * yi*)/((xi*)²)
    sum_mult_diff_mean_x_y = sum(mult_diff_mean_x_y)    # sum([xi*] * [yi*])
    sum_squared_diff_mean_x = sum(squared_diff_mean_x)  # sum([xi*]²)
    slope = sum_mult_diff_mean_x_y / sum_squared_diff_mean_x  # a1

    # calculate intercet a0:  f(x)=y=a0+a1x
    a0 = y_mean - x_mean * slope

    return [slope, a0]


if __name__ == '__main__':

    # finding correct rows and save data in a dictionary for further computation
    filtered_dic = {}
    important_rows = ["Mean Blank", "Standard Concentration [pg/µl]", "Mean Standard",
                      "Dilutions used for sample", "Mean Sample"]
    to_blank_subtr = ["Mean Standard", "Mean Sample"]
    for i in important_rows:
        find_row("test.xls", i, filtered_dic)

    # update dictionary by subtracting the fluorescence value of the blank from the
    # fluorescent of the standard and samples. needed because carrier substance
    # does fluoresce as well and increaes measured fluorescence
    for i in to_blank_subtr:
        subtract_blank(filtered_dic, i)

    # relationship between DNA concentration and fluorescent value is being calculated
    slope_intercept_lst = calculate_linear_equation(filtered_dic["Standard Concentration [pg/µl]"],
                                                    filtered_dic["Mean Standard"])

    sample_tpl = clean_two_lsts(filtered_dic["Dilutions used for sample"], filtered_dic["Mean Sample"])

    # using the DNA concentration and fluorescent relationship from the standard curve to
    # calculate the DNA concentration of my sample with the measured fluorescent values
    conc_sum = 0
    for n, i in enumerate(sample_tpl[1]):
        a = (i - slope_intercept_lst[1]) / slope_intercept_lst[0]   # x = (y - a0) / a1
        conc_sum += sample_tpl[0][n] * a                            # since the sample had different dilutions, we need use the dilution factor for each sample to get the original DNA concentration
    final_conc = conc_sum / len(sample_tpl[0])

    # calculating the statistical value R² (coefficient of determination) and the best fit for the standard curve
    r_sq = (np.corrcoef(filtered_dic["Standard Concentration [pg/µl]"], filtered_dic["Mean Standard"])[0, 1])**2
    y_regression_line = [x * slope_intercept_lst[0] + slope_intercept_lst[1] for x in filtered_dic["Standard Concentration [pg/µl]"]]

    # creating the graph
    plt.style.use('ggplot')
    plt.plot(filtered_dic["Standard Concentration [pg/µl]"], filtered_dic["Mean Standard"], '.k')
    plt.plot(filtered_dic["Standard Concentration [pg/µl]"], y_regression_line, 'k:')
    plt.xlabel("DNA concentration [pg/µl]")
    plt.ylabel("Fluorescence Value [a.u.]")
    plt.title("Standard Curve")
    plt.grid(True)
    plt.tight_layout()

    # OUTPUT
    print("The original concentration of your sample is:", 20*" ", final_conc, "pg/µl")
    if slope_intercept_lst[1] >= 0:
        print("Your linear equation and correlation coefficient are as follows:", 2*" ",
              "y = " + str(slope_intercept_lst[0]) + " * x + " + str(slope_intercept_lst[1]), "\n", 66*" ", "R² =", r_sq)
    else:
        print("Your linear equation and correlation coefficient are as follows:", 2*" ",
              "y = " + str(slope_intercept_lst[0]) + " * x - " + str(abs(slope_intercept_lst[1])), "\n", 66*" ", "R² =", r_sq)
    plt.show()
