import codecs
import pandas as pd

# semantics
COL_DELIMITER = ";"
ATTR_DELIMITER = ":"
TABLE_NAME_HEADER = "* Table: "
COL_DEF_HEADER = "$"
NEWLINE_WIN = "\r\n"
NEWLINE_MAC = "\n"
NEWLINES = [NEWLINE_WIN, NEWLINE_MAC]

#paths
NETPATH = "./test/MOMM_net.net"
DMDPATH = "./test/MOMM_full_dmd.dmd"
OUTPATH = "./test/data/"


def parse(path=None):

    _table_flag = False
    _line_count = 0
    cols = None

    with codecs.open(path, encoding='utf-8', errors='ignore') as net:

        for line in net:
            if line.startswith(TABLE_NAME_HEADER):
                # new table
                table_name = line.split(ATTR_DELIMITER)[-1].replace(" ","").replace(NEWLINE_WIN, "")
                print("Exporting table: ", table_name)
                _line_count = 0
            if line[0] == COL_DEF_HEADER:
                # line with column names
                _table_flag = True
                data = list()  # initialize data
                try:
                    cols = line.split(":")[1].split(COL_DELIMITER)
                except:
                    cols = line.split(COL_DELIMITER)
            if line in NEWLINES:
                _table_flag = False
                if len(data) > 0:
                    # only if table is not empty
                    if cols is not None:
                        pd.DataFrame(data, columns=cols).to_csv(OUTPATH+table_name.split(" ")[0]+".csv")
                        print("Exported {} objects in table: {} - saved to {}"
                              .format(_line_count, table_name, OUTPATH+table_name.split(" ")[0]+".csv"))

            if _table_flag and line[0] != COL_DEF_HEADER:
                data.append(line.split(COL_DELIMITER))
                _line_count += 1


if __name__ == "__main__":
    parse(DMDPATH)
    parse(NETPATH)








