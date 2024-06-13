import pandas as pd
import csv

# append 
sr = pd.read_csv(r'C:\Py_Project\output\Estate_output_200000to10000.csv')

sr.to_csv(r'C:\Py_Project\project\pytest\batch_testing.csv',encoding='utf-8-sig')