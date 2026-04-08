import os
import shutil

script_name = os.path.basename(__file__)
folder = os.path.dirname(os.path.abspath(__file__))

for filename in os.listdir(folder):
    if filename == script_name:
        continue

    name, ext = os.path.splitext(filename)
    upper_name = name.upper()
    upper_ext = ext.upper()
    new_filename = upper_name + upper_ext

    if new_filename == filename:
        continue  # Already correct, skip

    original = os.path.join(folder, filename)
    temp = os.path.join(folder, "TEMP_" + new_filename)
    final = os.path.join(folder, new_filename)

    shutil.copy2(original, temp)   # Copy to TEMP_UPPERCASE
    os.remove(original)            # Delete original
    os.rename(temp, final)         # Rename TEMP_ -> final

    print(f"Renamed: {filename} -> {new_filename}")

print("Done.")
