import numpy as np
import matplotlib.pyplot as plt
import xlrd
import os
import csv
from scipy.stats import kurtosis
from scipy.stats import skew
from scipy.stats import jarque_bera
from scipy.stats import f_oneway


def estimate_coef(x, y):
    n = np.size(x)

    m_x, m_y = np.mean(x), np.mean(y)

    SS_xy = np.sum(y*x) - n*m_y*m_x
    SS_xx = np.sum(x*x) - n*m_x*m_x
    b_1 = SS_xy / SS_xx
    b_0 = m_y - b_1*m_x

    return(b_0, b_1)


def plot_regression_line(x, y, b):
    plt.scatter(x, y, color="m",
                marker="o", s=30)

    y_pred = b[0] + b[1]*x

    plt.plot(x, y_pred, color="g")

    plt.xlabel('RD in GDP')
    plt.ylabel('GDP per capita')

    plt.show()


def read_excel(loc, col_index):
    col_values = []
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, col_index)

    for i in range(1, sheet.nrows):
        col_values.append(sheet.cell_value(i, col_index))
    return col_values


def print_results(rd_in_gdp, gdp_per_capita, b):    
    
    #create a new file and print the results of the analysis
    with open('results.csv', mode='w', newline='') as results_file:
        results_writer = csv.writer(results_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)

        results_writer.writerow(['GPD per capita descriptive analysis'])
        results_writer.writerow(['Mean', np.mean(gdp_per_capita)])
        results_writer.writerow(['Meadian', np.median(gdp_per_capita)])
        results_writer.writerow(['Maximum', np.amax(gdp_per_capita)])
        results_writer.writerow(['Minimum' ,np.amin(gdp_per_capita)])
        results_writer.writerow(['Standard deviation', np.std(gdp_per_capita)])
        results_writer.writerow(['Jarques-Bera and Probability', jarque_bera(gdp_per_capita)[0], jarque_bera(gdp_per_capita)[1]])
        results_writer.writerow(['Skewness', skew(gdp_per_capita)])
        results_writer.writerow([''])

        results_writer.writerow(['RD IN GDP descriptive analysis'])
        results_writer.writerow(['Mean', np.mean(rd_in_gdp)])
        results_writer.writerow(['Meadian', np.median(rd_in_gdp)])
        results_writer.writerow(['Maximum', np.amax(rd_in_gdp)])
        results_writer.writerow(['Minimum' ,np.amin(rd_in_gdp)])
        results_writer.writerow(['Standard deviation', np.std(rd_in_gdp)])
        results_writer.writerow(['Jarques-Bera and Probability', jarque_bera(rd_in_gdp)[0], jarque_bera(rd_in_gdp)[1]])
        results_writer.writerow(['Skewness', skew(rd_in_gdp)])

        results_writer.writerow([''])
        results_writer.writerow(['Estimated coefficients', b[0], b[1]])

        results_writer.writerow([''])
        results_writer.writerow(['F Test', f_oneway(rd_in_gdp, gdp_per_capita)])

def main():
    #get location of the data excel
    loc = (os.getcwd() + "/Data - Simple Regression Model.xlsx")

    #read the data from the excel
    rd_in_gdp_values = read_excel(loc, 1)
    gdp_per_capita_values = read_excel(loc, 2)

    rd_in_gdp = np.array(rd_in_gdp_values)
    gdp_per_capita = np.array(gdp_per_capita_values)

    #estimate the coefficients for the linear regression
    b = estimate_coef(rd_in_gdp, gdp_per_capita)

    #print the results to a new file
    print_results(rd_in_gdp, gdp_per_capita, b)

    #plot the line of the regression
    plot_regression_line(rd_in_gdp, gdp_per_capita, b)

if __name__ == "__main__":
    main()
