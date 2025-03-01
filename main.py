import pandas as pd

def analyze_files(source_file, target_file, mapping_file):
    try:
        # Read all files
        source_df = pd.read_excel(source_file)  # NovaTek Data
        target_df = pd.read_excel(target_file)  # Labware Data
        mapping_df = pd.read_excel(mapping_file)  # Mapping File

        # Initialize results storage
        matches = []
        mismatches = []
        
        # Group source tests by study and condition
        for index, source_row in source_df.iterrows():
            try:
                # Get test details from source file
                project_id = source_row.iloc[0]  # Test ID from Column A
                test_name = source_row.iloc[3]  # Test name from Column D
                
                if pd.isna(project_id) or pd.isna(test_name):
                    continue
                
                # Find corresponding mapping
                mapping = mapping_df[mapping_df.iloc[:, 4] == test_name]  # Match on Column E in mapping file
                
                if len(mapping) == 0:
                    mismatches.append(f"No mapping found for test: {test_name} (ID: {project_id})")
                    continue
                
                # Get mapped test name from Column C in mapping file
                mapped_test_name = mapping.iloc[0, 2]
                
                # Look for the mapped test in target file
                target_match = target_df[target_df.iloc[:, 9] == mapped_test_name]  # Column J in target file
                
                if len(target_match) == 0:
                    mismatches.append(f"Test not found in target file: {test_name} -> {mapped_test_name} (ID: {project_id})")
                else:
                    matches.append(f"Match found: {test_name} -> {mapped_test_name} (ID: {project_id})")
                
            except Exception as row_error:
                mismatches.append(f"Error processing row {index}: {str(row_error)}")
        
        # Print results with counts
        print("\nAnalysis Results:")
        print(f"\nTotal tests processed: {len(matches) + len(mismatches)}")
        print(f"Matches found: {len(matches)}")
        print(f"Mismatches found: {len(mismatches)}")
        
        print("\nMatched Tests:")
        for match in matches:
            print(match)
            
        print("\nMissing or Mismatched Tests:")
        for mismatch in mismatches:
            print(mismatch)
            
    except Exception as e:
        print(f"Error occurred: {str(e)}")

if __name__ == "__main__":
    # File paths
    source_file = "ns24-00001_nt.xlsx"  # Source file with tests
    target_file = "ns24-0001export.xlsx"  # Target file to check against
    mapping_file = "NVT_LW mapping.xlsx"  # Mapping between source and target
    
    analyze_files(source_file, target_file, mapping_file)
