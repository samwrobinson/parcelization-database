import pandas as pd
import matplotlib.pyplot as plt

# Dictionary to map town codes to town names
town_code_map = {
    23040: 'Fayston',
    23080: 'Waitsfield',
    23085: 'Warren',
    # Add more mappings as needed
}

# Function to filter and save town data to individual sheets
def save_town_data(df, town_codes, writer):
    for town_code in town_codes:
        # Replace town codes with town names
        town_name = town_code_map.get(town_code, f'Town_{town_code}')
        
        # Filter the data for the specific town
        df_filtered = df[df['Town'] == town_code]
        
        # Write the filtered data to a new sheet in the Excel workbook
        sheet_name = f'{town_name}_metric1'
        df_filtered.to_excel(writer, sheet_name=sheet_name, index=False)

# Function to generate graphs for each town
def generate_town_graphs(df, town_codes):
    for town_code in town_codes:
        # Replace town codes with town names
        town_name = town_code_map.get(town_code, f'Town_{town_code}')
        
        # Filter the data for the specific town
        df_filtered = df[df['Town'] == town_code]
        
        # Pivot table to analyze the number of parcels by size category over time for the selected town
        pivot_table_town = df_filtered.pivot_table(
            index='TaxYear',
            columns='Town',
            values=['Metric1_0To2Acres', 'Metric1_2To5Acres', 'Metric1_5To10Acres',
                    'Metric1_10To25Acres', 'Metric1_25To50Acres', 'Metric1_50To100Acres',
                    'Metric1_100To200Acres', 'Metric1_GT200Acres'],
            aggfunc='sum'
        )
        
        # Plotting the change over time for the town
        pivot_table_town.plot(kind='line', title=f'Parcel Count by Parcel Size, 2004 - 2020 - {town_name}')
        plt.ylabel('Number of Parcels')
        plt.xlabel('Year')
        plt.show()

def main():
    # Load the Excel file
    file_path = '/Users/MRVPDAir/Desktop/scripts/parcelization_db/Parcelization_Database.xlsx'
    xls = pd.ExcelFile(file_path)
    
    # Updated list of town codes to analyze
    town_codes = [23040, 23080, 23085]
    
    # Load the specific sheet for analysis
    sheet_name = 'tbl_Metric1_step2_Town'
    df = pd.read_excel(xls, sheet_name=sheet_name)
    
    # Initialize a new Excel writer object
    output_path = '/Users/MRVPDAir/Desktop/scripts/parcelization_db/Analysis/Parcelization_Analysis_Metric1.xlsx'
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
    
    # Save town data to individual sheets
    save_town_data(df, town_codes, writer)
    
    # Save the Excel workbook
    writer.close()
    
    # Generate graphs for each town
    generate_town_graphs(df, town_codes)

if __name__ == "__main__":
    main()
