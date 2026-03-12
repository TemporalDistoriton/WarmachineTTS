import os
import re

def clean_filenames():
    directory = os.getcwd()
    script_name = os.path.basename(__file__)
    
    for filename in os.listdir(directory):
        # 1. Skip the script and directories
        if filename == script_name or not os.path.isfile(filename):
            continue
            
        # 2. Separate name and extension
        name, ext = os.path.splitext(filename)
        
        # 3. Clean the name:
        # - Convert to lowercase
        # - Remove anything that IS NOT a letter, number, or space
        # [^a-z0-9 ] means "match anything except a-z, 0-9, or a space"
        clean_name = name.lower()
        clean_name = re.sub(r'[^a-z0-9 ]', '', clean_name)
        
        # 4. Re-attach extension (making extension lowercase too)
        new_name = clean_name + ext.lower()
        
        # 5. Perform the two-step rename for Windows case-insensitivity
        if new_name != filename:
            temp_name = filename + "_temp"
            try:
                os.rename(filename, temp_name)
                os.rename(temp_name, new_name)
                print(f"Cleaned: '{filename}' -> '{new_name}'")
            except OSError as e:
                print(f"Error renaming {filename}: {e}")

if __name__ == "__main__":
    clean_filenames()
    print("\nFilename cleanup complete!")