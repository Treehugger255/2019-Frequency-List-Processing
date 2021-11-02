import numpy as np
import pandas as pd
import math

readSheet = input("Read File Location: ")
writeSheet = input("Write File Location")

df = pd.read_excel(readSheet)

def writeFPMW(defaultSeries):
    wordFRQ = defaultSeries.sum()
    print("FPMW has been called")
    fpmw = float(wordFRQ) * (10 ** 6) / 51000000
    return fpmw

def writeZipf(defaultSeries):
    wordFRQ = defaultSeries.sum()
    print("Zipf has been called")
    fpmw = float(wordFRQ) * (10 ** 6) / 51000000
    zipfV = math.log(fpmw * 1000, 10)
    return zipfV

grouped = df.groupby(["Word", "PoS"], sort=False)
grouped = grouped.agg(
    newFREQ = pd.NamedAgg(column="FREQcount", aggfunc=np.sum),
    FPMW = pd.NamedAgg(column="FREQcount", aggfunc=writeFPMW),
    ZIPF = pd.NamedAgg(column="FREQcount", aggfunc=writeZipf),
    )

#grouped.to_excel(writeSheet)
grouped.to_csv(writeSheet)