import argparse, csv
import pandas as pd
import yaml


def export_students(advisor, config, sheet):
    students = []
    for index, row in df.iterrows():
        for prog in config['programs']:
            for pf in config['program-fields']:
                for af in config['advisor-fields']:
                    if prog == row[pf] and row[af] == advisor:
                        id = row[config['student-field']]
                        id = str(id).zfill(9)
                        print(advisor, "has", row['Last Name'], "("+id+")")
                        students.append(id)

    path = '_'.join(advisor.split()) + '_students.csv'
    with open(path, 'w', newline='') as f:
        # https://stackoverflow.com/questions/3191528/csv-in-python-adding-an-extra-carriage-return-on-windows
        records = csv.writer(f, lineterminator='\n')
        for id in students:
            records.writerow([id])


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--excel", help="path to file with assignments",
            default='2020.09.04.Advising-rebalancing.xlsx')
    parser.add_argument("--config", help="path to config",
            default='advisors.yml')
    args = parser.parse_args()

    f = open(args.config, 'r')
    config = yaml.load(f)

    sheet = None
    xl = pd.ExcelFile(args.excel)
    for name in xl.sheet_names:
        df = xl.parse(name)
        
        fields = config['program-fields'] + \
                 config['advisor-fields'] + \
                 [ config['student-field'] ]
        for x in fields:
            if not x in df.keys():
                continue
        sheet = df

    if sheet is None:
        print("Counld not find required fields in", args.excel, fields)

    for a in config['advisors']:
        export_students(a, config, sheet)
    
