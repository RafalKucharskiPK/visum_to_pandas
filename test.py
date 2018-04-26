from ptv_visum_to_pandas import *
import filecmp

parse(NETPATH)
assert filecmp.cmp("./test/data/links.csv", "./test/data/links_ref.csv")

parse(DMDPATH)
assert filecmp.cmp("./test/data/matrix entries.csv", "./test/data/matrix entries_ref.csv")

matrices_export_via_com(VERPATH, export_path=OUTPATH, export_list=[101, 102])   # b) just some matrices
assert filecmp.cmp("./test/data/mtx_10.csv", "./test/data/mtx_10_ref.csv")

