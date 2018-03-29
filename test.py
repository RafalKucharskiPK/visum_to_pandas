from ptv_visum_to_pandas import *
import filecmp

parse(NETPATH)
assert filecmp.cmp("./test/data/links.csv", "./test/data/links_ref.csv")

parse(DMDPATH)
assert filecmp.cmp("./test/data/matrix entries.csv", "./test/data/matrix_entries_ref.csv")
