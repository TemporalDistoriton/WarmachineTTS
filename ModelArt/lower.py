import os

def force_lowercase_rename():
    directory = os.getcwd()
    
    for filename in os.listdir(directory):
        # Skip the script itself
        if filename == os.path.basename(__file__):
            continue
            
        if os.path.isfile(filename):
            new_name = filename.lower()
            
            # Only proceed if the case is actually different
            if new_name != filename:
                temp_name = filename + "_temp"
                try:
                    # Step 1: Rename to a temporary name
                    os.rename(filename, temp_name)
                    # Step 2: Rename from temp to the lowercase version
                    os.rename(temp_name, new_name)
                    print(f"Fixed: {filename} -> {new_name}")
                except OSError as e:
                    print(f"Error renaming {filename}: {e}")

if __name__ == "__main__":
    force_lowercase_rename()
    print("Force rename complete!")