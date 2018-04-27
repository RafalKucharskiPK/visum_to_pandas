# visum_to_pandas
python scripts to parse visum .net and .dmd file to pandas and store as .csv files 
(c) Rafal Kucharski, Politechnika Krakowska, Krakow, Poland

## usage

open `ptv_visum_to_pandas.py` and go to `main()` where you can find the below pipeline:

   ### 0. Set your paths - your ver files and export paths
    NETPATH = "./test/MOMM_net.net"  # .net file to parse into csv tables
    DMDPATH = "./test/MOMM_full_dmd.dmd"  # .dmd file to parse into csv tables
    OUTPATH = "./test/data/"  # save the .csv's here
    VERPATH = "./test/test_matrices.ver"
    
   ### 1. Export .ver file to editable .net and .dmd
    export_net_dmd(path=VERPATH)

   ### 2. export network to csv
    parse(NETPATH, export_path=OUTPATH) # a) full export
    parse(NETPATH, export_path=OUTPATH, export_list=["Links", "Nodes"])  # b) just some tables

   ### 3. export demand files to csv
    parse(DMDPATH, export_path=OUTPATH)

   ### 4. Export matrices (export from .dmd is not very useful)
    matrices_export_via_com(VERPATH, export_path=OUTPATH)  # a) full export
    matrices_export_via_com(VERPATH, export_path=OUTPATH, export_list=[101, 102])   # b) just some matrices
    
  ### 5. see if everything went fine and understand how to use it further with pandas
    test_read("./test/data/Tripgeneration.csv")
    test_read("./test/data/Links.csv")
    test_read("./test/data/Mtx_10.csv")  


## Third party

We link [PTV Visum](http://vision-traffic.ptvgroup.com/en-us/products/ptv-visum/) with [pandas](https://pandas.pydata.org/)

## test

run `test.py` to see if the `links.csv` is as we expect and if `matrix entries.csv` are the same. Use only for the original files.

## contributors

https://github.com/PatrikHlobil

## possible modifications

* `pd.DataFrame(data, columns=cols).to_csv(OUTPATH+table_name+".csv")` can be easily modified to be used directly in pandas (e.g. dict of DataFrames, or h5 store).
* ~~pipeline can be smoother if this is called from within Visum, you shall add `Visum.IO.SaveNet, Visum.IO.SaveDmd` to prepare the files under good paths and then use them in `parse`~~ done

