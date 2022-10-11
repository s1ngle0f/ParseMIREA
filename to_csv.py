import pandas as pd


def fill_free_space(path):
    file = pd.read_csv(path)
    # print(type(file.iloc[0, 0]), file.iloc[0, 0])
    # print(file.iloc[:, 0])
    for i in range(0, len(file.columns), 15):
        for col in range(0, 4):
            last_day = None
            n_col = i + col
            if n_col < len(file.columns):
                for i_row in range(0, len(file.iloc[:, n_col])):
                    if type(file.iloc[i_row, n_col]) != float:
                        last_day = file.iloc[i_row, n_col]
                    file.iloc[i_row, n_col] = last_day
    # print(file.iloc[:, 0])
    file.to_csv(path, index=False)


# def set_FIO_column(file):
#     for i in range(7, len(file.columns), 15):


def create_csv_from_excel(name):
    path = f'input/{name}'
    read_file = pd.read_excel(path)
    csv_path = f'csv/{path[path.find("/")+1:].replace(".xlsx", ".csv")}'
    read_file.to_csv(csv_path, header=False, index=False)
    fill_free_space(csv_path)
create_csv_from_excel('IIT_3-kurs_22_23_osen_07.10.2022.xlsx')

