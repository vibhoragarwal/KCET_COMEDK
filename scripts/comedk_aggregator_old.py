import pandas as pd

def merge_excel_sheets(file_path):
    """
    Merges two sheets from an Excel file, handling missing college codes.

    Args:
        file_path (str): The path to the Excel file.

    Returns:
        pandas.DataFrame: The merged DataFrame, or None on error.
    """
    try:
        # Read the four sheets into pandas DataFrames
        cs_round1_df = pd.read_excel(file_path, sheet_name='CS-Round 1 2024')
        cs_round2_df = pd.read_excel(file_path, sheet_name='CS-Round 2 2024')
        is_round1_df = pd.read_excel(file_path, sheet_name='IS-Round 1 2024')
        is_round2_df = pd.read_excel(file_path, sheet_name='IS-Round 2 2024')

        # Rename the third column in each DataFrame for clarity and to avoid duplicate column names after merging
        cs_round1_df = cs_round1_df.rename(columns={
            ' CS-Computer Science & Engineering ROUND 1 2024': 'CS-Computer Science & Engineering ROUND 1 2024'})
        cs_round2_df = cs_round2_df.rename(columns={
            ' CS-Computer Science & Engineering ROUND 2 2024': 'CS-Computer Science & Engineering ROUND 2 2024'})
        is_round1_df = is_round1_df.rename(columns={
            ' IS-Information Science & Engineering ROUND 1 2024': 'IS-Information Science & Engineering ROUND 1 2024'})
        is_round2_df = is_round2_df.rename(columns={
            ' IS-Information Science & Engineering ROUND 2 2024': 'IS-Information Science & Engineering ROUND 2 2024'})

        # Merge the DataFrames, starting with CS rounds, then adding IS rounds
        merged_df = pd.merge(cs_round1_df, cs_round2_df, on=['College Code', 'College Name'], how='outer')
        merged_df = pd.merge(merged_df, is_round1_df, on=['College Code', 'College Name'], how='outer')
        merged_df = pd.merge(merged_df, is_round2_df, on=['College Code', 'College Name'], how='outer')

        # Save the merged DataFrame to a new Excel file
        with pd.ExcelWriter(file_path, engine='openpyxl', mode='a') as writer:
            merged_df.to_excel(writer, sheet_name='IS CS 2024', index=False)  # Save to a new sheet

        print(f"Merged data successfully written to a new sheet 'IS CS 2024' in {file_path}")
        return merged_df

    except FileNotFoundError:
        print(f"Error: File not found at {file_path}")
        return None
    except KeyError as e:
        print(f"Error: Sheet name not found.  Check that 'Sheet1' and 'Sheet2' are correct.  KeyError: {e}")
        return None
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return None



if __name__ == "__main__":
    excel_file_path = '../comedk_files/COMEDK 2024.xlsx'  # Replace with the actual path to your Excel file
    merged_data = merge_excel_sheets(excel_file_path)

    if merged_data is not None:
        print("\nMerged Data:")
        print(merged_data)
        # You can save to a new file if needed
        # merged_data.to_excel("merged_data.xlsx", index=False)
    else:
        print("Failed to merge data.")