import pandas as pd

def collect_data():
    data = {}
    data['Customer'] = input("Enter Customer Name: ")
    data['Location'] = input("Enter Location: ")
    data['Job Number'] = input("Enter Job Number: ")
    data['Production Number'] = input("Enter Production Number: ")

    hi_pot_test = input("Did the Hi-Pot test pass? (Yes/No): ").lower()
    if hi_pot_test == 'yes':
        data['Hi-Pot Test Result'] = 'Pass'
    else:
        data['Hi-Pot Test Result'] = 'Fail'
        data['Reason for Failure'] = input("Enter the reason for failure: ")

    return data

def save_to_excel(data, file_name='Hussmann_unit_records.xlsx'):
    try:
        df = pd.read_excel(file_name)
    except FileNotFoundError:
        df = pd.DataFrame()

    # Creating a DataFrame from the new data and concatenating it
    new_df = pd.DataFrame([data])
    df = pd.concat([df, new_df], ignore_index=True)

    df.to_excel(file_name, index=False)


def main():
    while True:
        data = collect_data()
        save_to_excel(data)
        print("Data saved. Starting new entry...\n")

if __name__ == "__main__":
    main()
