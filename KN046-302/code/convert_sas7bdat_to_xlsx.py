from sas7bdat import SAS7BDAT
import argparse

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--path", help="Must provide the sas7bdat file name")

    args = parser.parse_args()

    # file_path = os.path.join(r'..\sheets', args.cancer)    
    path = args.path
    # fbbz_sheet = args.fbbz
    # xbz_sheet = args.xbz
    # recist_sheet = args.recist

    convert_path = path.replace('sas7bdat', 'xlsx')

    f = SAS7BDAT(path, encoding="UTF-8").to_data_frame()

    f.to_excel(convert_path)

    print("Converted file path: {0}".format(convert_path))