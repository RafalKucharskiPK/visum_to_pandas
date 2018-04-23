import codecs
import pandas as pd
import os

# semantics
COL_DELIMITER = ";"
ATTR_DELIMITER = ":"
TABLE_NAME_HEADER = "* Table: "
COL_DEF_HEADER = "$"
NEWLINE_WIN = "\r\n"
NEWLINE_MAC = "\n"
NEWLINES = [NEWLINE_WIN, NEWLINE_MAC]

# params
LIMIT = 100000  # number of lines to break the csv into separate files

# paths
NETPATH = "./test/MOMM_net.net"
DMDPATH = "./test/MOMM_full_dmd.dmd"
OUTPATH = "./test/data/"  # save the .csv's here
VERPATH = "./test/test_matrices.ver"

VERPATH = "E:\\wbr.ver"
OUTPATH = "E:/csvs/"


def matrices(path):
    if 'Visum' in dir():
        pass
    else:
        import win32com.client
        Visum = win32com.client.Dispatch('Visum.Visum.16')
        Visum.LoadVersion(os.path.join(os.getcwd(), path))

    cols = Visum.Net.Zones.GetMultiAttValues('No', False)
    cols = [int(_[1]) for _ in cols]
    iterator = Visum.Net.Matrices.Iterator
    while iterator.Valid:
        mtx = iterator.Item
        no = str(int(mtx.AttValue("No")))
        vals = list(mtx.GetValuesFloat())
        df = pd.DataFrame(vals, cols, cols)
        df.to_csv(OUTPATH + "MTX_" + no + ".csv")
        print('Exported mtx {} to file'.format(no))
        iterator.Next()


def parse(path=None):
    """
    Main function to parse the editable PTV Visum export files (.net and .dmd).
    It parses line by line and listens for table header, column definitions, data and table end.
    It stores the data to csvs (one csv per Visum table)
    :param path:
    :return:
    """
    _table_flag = False
    _line_count = 0
    _cols = None
    _split_count = 0
    _split_flag = 0
    table_name = None

    with codecs.open(path, encoding='utf-8', errors='ignore') as net:

        for line in net:
            if line.startswith(TABLE_NAME_HEADER):
                # new table
                table_name = line.split(ATTR_DELIMITER)[-1].replace(" ", "").replace(NEWLINE_WIN, "")
                print("Exporting table: ", table_name)
                _line_count = 0
            if line[0] == COL_DEF_HEADER:
                # line with column names
                _table_flag = True
                data = list()  # initialize data
                _split_count = 0
                line = line.replace(NEWLINE_WIN, "").replace("\r", "")
                if len(line.split(":")) > 1:
                    _cols = line.split(":")[1].split(COL_DELIMITER)
                else:
                    _cols = line.split(COL_DELIMITER)
            if line in NEWLINES or _split_flag:
                if not _split_flag:
                    _table_flag = False
                if len(data) > 0:
                    # only if table is not empty
                    if _cols is not None:
                        _file_name = OUTPATH+table_name.split(" ")[0]
                        if _split_count > 0:
                            _file_name += "_" + str(_split_count)
                        _file_name += ".csv"
                        pd.DataFrame(data, columns=_cols).to_csv(_file_name)
                        print("Exported {} objects in table: {} - saved to {}"
                              .format(_line_count, table_name, _file_name))
                        data = list()  # initialize data
                        _line_count = 0
                        _split_flag = False

            if _table_flag and line[0] != COL_DEF_HEADER:
                data.append(line.replace(NEWLINE_WIN, "").split(COL_DELIMITER))
                _line_count += 1
                if _line_count >= LIMIT:
                    _split_count += 1
                    _split_flag = True


def test_read(_path):
    df = pd.read_csv(_path)
    df = df.set_index(df["Unnamed: 0"])
    del df["Unnamed: 0"]
    print(df.columns)
    print(df.head())
    print(df.describe().T)


if __name__ == "__main__":
    matrices(VERPATH)
    quit()
    parse(DMDPATH)
    parse(NETPATH)

    test_read("./test/data/Tripgeneration.csv")
    test_read("./test/data/Links.csv")
    test_read("./test/data/Mtx_10.csv")
