import pandas as pd
import pymssql

IDs = pd.read_excel("C:/Users/fhan/Desktop//High Throughput SQL Data Retriever Test/IDs.xlsx",header=None).iloc[:,0]
measurements = pd.read_excel("C:/Users/fhan/Desktop/High Throughput SQL Data Retriever Test/Measurements.xlsx",header=None).iloc[:,0]

conn = pymssql.connect(server='XXXXXX', port=00000)

def concatenate_string(SeriesData,low_index,high_index):
    string = "("
    for item in SeriesData[low_index:high_index]:
        string = string + "'" + item + "',"
    string = string + "'" + SeriesData[high_index] + "')"
    return string

measurements_string = concatenate_string(measurements,0,len(measurements)-1)

def retrieve_measurement_data(IDs_string,measurements_string):
    sqlCommand = "SELECT sample_ID, measurement, measured_value from table_for_sample_measurements WHERE sample_ID IN "+IDs_string+" AND measurement IN "+measurements_string
    cursor = conn.cursor()
    cursor.execute(sqlCommand)
    data = cursor.fetchall()
    return data

if len(IDs) <= 5000:
    IDs_string = concatenate_string(IDs,0,len(IDs)-1)
    data = retrieve_measurement_data(IDs_string,measurements_string)
else:
    data = []
    low = 0
    high = 4999
    while high <= len(IDs)-1:
        IDs_string = concatenate_string(IDs,low,high)
        data += retrieve_measurement_data(IDs_string,measurements_string)
        low += 5000
        high += 5000
    IDs_string = concatenate_string(IDs,low,len(IDs)-1)
    data += retrieve_measurement_data(IDs_string,measurements_string)

conn.close()

data = pd.DataFrame(data,columns=["Sample ID","Measurement","Measured Value"])
data = data.pivot(index=['Sample ID'], columns='Measurement', values='Measured Value')
data = data.reset_index()

data.to_excel("C:/Users/fhan/Desktop/High Throughput SQL Data Retriever Test/Retrieved data.xlsx", index=False)