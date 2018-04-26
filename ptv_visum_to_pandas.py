import codecs
import pandas as pd
import os

# paths
NETPATH = "./test/MOMM_net.net"  # .net file to parse into csv tables
DMDPATH = "./test/MOMM_full_dmd.dmd"  # .dmd file to parse into csv tables
OUTPATH = "./test/data/"  # save the .csv's here
VERPATH = "./test/test_matrices.ver"


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

# my big files to test (Warsaw model) - Rafal
# NETPATH = "E:/wbr.net"
# DMDPATH = "E:/wbr.dmd"
VERPATH = "C://Users//Rafal//PycharmProjects//visum_to_pandas//test//test_matrices.ver"
# VERPATH = "E:/wbr.ver"
# OUTPATH = "E:/csvs/"


def _get_or_dispatch_visum(path= None):
    """
    internal procedure to handle Visum and run it if the script is not called from Visum
    :param path: version file path
    :return: Visum object
    """
    if 'Visum' in dir():
        pass
    else:
        import win32com.client
        Visum = win32com.client.Dispatch('Visum.Visum')
        Visum.LoadVersion(path)
    return Visum


def export_net_dmd(path):
    """
    Shortcut to export .ver into .net and .dmd
    :param path: version file path
    :return: None
    """
    Visum = _get_or_dispatch_visum(path)
    Visum.IO.SaveNet(os.path.join(os.path.dirname(path), "net_file.net"))
    Visum.IO.SaveDemandFile(os.path.join(os.path.dirname(path), "dmd_file.dmd"),True)


def matrices_export_via_com(path, export_path=None, export_list=[None]):
    """
    Procedure iterating over matrices inside Visum version file accessed over com and exporting
    matrix values to csv files.
    :param path: Visum version file to execute (not needed if run from Visum (script)
    :param export_path: path to export matrices to
    :param export_list: list of tables to export
    :return:
    """
    Visum = _get_or_dispatch_visum(path)

    cols = Visum.Net.Zones.GetMultiAttValues('No', False)
    cols = [int(_[1]) for _ in cols]
    iterator = Visum.Net.Matrices.Iterator
    while iterator.Valid:
        mtx = iterator.Item
        if (None in export_list) or (mtx.AttValue("No") in export_list):
            no = str(int(mtx.AttValue("No")))
            vals = list(mtx.GetValuesFloat())
            df = pd.DataFrame(vals, cols, cols)
            df.to_csv(export_path + "MTX_" + no + ".csv")
            print('Exported mtx {} to file'.format(no))
        iterator.Next()


def parse(path, export_path=None, export_list=[None]):
    """
    Main function to parse the editable PTV Visum export files (.net and .dmd).
    It parses line by line and listens for table header, column definitions, data and table end.
    It stores the data to csvs (one csv per Visum table)
    :param path: Path of the editable file (.net, .dmd) to process
    :param export_path: Path to folder, where to save the CSV files.
                               If not specified, the CSV files are save in
                               folder of the net files.
    :param export_list: List of network objects that should be exported. If
                          not specified, all network objects will be exported.
    :return: None
    """
    _table_flag = False
    _line_count = 0
    _cols = None
    _split_count = 0
    _split_flag = 0
    table_name = None

    # use path of net-file for export if no export_path is specified:
    if export_path is None:
        export_path = os.path.dirname(path)

    # Autodetect COL_DELIMITER:
    _col_delimiter = COL_DELIMITER
    _col_identified_flag = True
    with codecs.open(path, encoding='utf-8', errors='ignore') as net:
        for line in net:
            if _col_identified_flag:
                if "$VERSION:" in line:
                    tmp = line.replace("$VERSION:", "").strip()
                    for delimiter in [" ", ";", "\t"]:
                        if delimiter in tmp:
                            _col_delimiter = delimiter
                            print('Use delimiter "%s"' % _col_delimiter)
                            _col_identified_flag = False

            if line.startswith(TABLE_NAME_HEADER):
                # new table
                table_name = line.split(ATTR_DELIMITER)[-1].replace(" ", "").strip()
                if (None in export_list) or (table_name in export_list):
                    print("Parsing table: ", table_name)
                _line_count = 0
            if line[0] == COL_DEF_HEADER:
                # line with column names
                _table_flag = True
                data = list()  # initialize data
                _split_count = 0
                line = line.replace(NEWLINE_WIN, "").replace("\r", "")
                if len(line.split(":")) > 1:
                    _cols = line.split(":")[1].split(_col_delimiter)
                else:
                    _cols = line.split(_col_delimiter)
            if line in NEWLINES or _split_flag:
                if not _split_flag:
                    _table_flag = False
                if len(data) > 0:
                    # only if table is not empty
                    if _cols is not None:
                        _file_name = os.path.join(export_path, table_name)
                        if _split_count > 0:
                            _file_name += "_" + str(_split_count)
                        _file_name += ".csv"
                        if (None in export_list) or (table_name in export_list):
                            pd.DataFrame(data, columns=_cols).to_csv(_file_name, index=False)
                            print("Exported {} objects in table: {} - saved to {}"
                                  .format(_line_count, table_name, _file_name))
                        data = list()  # initialize data
                        _line_count = 0
                        _split_flag = False

            if _table_flag and line[0] != COL_DEF_HEADER:
                data.append(line.strip().split(_col_delimiter))
                _line_count += 1
                if _line_count >= LIMIT:
                    _split_count += 1
                    _split_flag = True


def test_read(_path):
    """
    Check if everything went fine with csv file (of _path)
    :param _path:
    :return: None
    """
    df = pd.read_csv(_path)
    print(df.columns)
    print(df.head())
    print(df.describe().T)

def main():
    """
    full pipeline and showcase
    :return: NOne
    """

    # 0. Set your paths - your ver files and export paths
    NETPATH = "./test/MOMM_net.net"  # .net file to parse into csv tables
    DMDPATH = "./test/MOMM_full_dmd.dmd"  # .dmd file to parse into csv tables
    OUTPATH = "./test/data/"  # save the .csv's here
    VERPATH = "./test/test_matrices.ver"

    # 1. Export .ver file to editable .net and .dmd
    export_net_dmd(path=VERPATH)

    # 2. export network to csv
    parse(NETPATH, export_path=OUTPATH) # a) full export
    parse(NETPATH, export_path=OUTPATH, export_list=["Links", "Nodes"])  # b) just some tables

    # 3. export demand files to csv
    parse(DMDPATH, export_path=OUTPATH)

    # 4. Export matrices (export from .dmd is not very useful)
    matrices_export_via_com(VERPATH, export_path=OUTPATH)  # a) full export
    matrices_export_via_com(VERPATH, export_path=OUTPATH, export_list=[101, 102])   # b) just some matrices

    # 5. see if everything went fine and understand how to use it further with pandas
    test_read("./test/data/Tripgeneration.csv")
    test_read("./test/data/Links.csv")
    test_read("./test/data/Mtx_10.csv")


if __name__ == "__main__":
    main()
