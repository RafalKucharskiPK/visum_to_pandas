# visum_to_pandas
python scripts to parse visum .net and .dmd file to pandas and store as .csv files 
(c) Rafal Kucharski, Politechnika Krakowska, Krakow, Poland

## usage

* save you .ver file to .dmd and .net (Visum->File->Save As -> Network/Demand Model)
* either modify paths inside the script `NETPATH`, `DMDPATH`, or overwrite files in [test](https://github.com/RafalKucharskiPK/visum_to_pandas/blob/master/test/MOMM_net.net)
* run the script (cmd/python, pywin, pycharm console, visum python console, etc.)
* enjoy the csv files to be analyzed with pandas

## Third party

We link [PTV Visum](http://vision-traffic.ptvgroup.com/en-us/products/ptv-visum/) with [pandas](https://pandas.pydata.org/)

## test

run `test.py` to see if the `links.csv` is as we expect and if `matrix entries.csv` are the same. Use only for the original files.

## possible modifications

* `pd.DataFrame(data, columns=cols).to_csv(OUTPATH+table_name+".csv")` can be easily modified to be used directly in pandas (e.g. dict of DataFrames, or h5 store).
* pipeline can be smoother if this is called from within Visum, you shall add `Visum.IO.SaveNet, Visum.IO.SaveDmd` to prepare the files under good paths and then use them in `parse`

