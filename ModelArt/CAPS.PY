import os

script_name = os.path.basename(__file__)
folder = os.path.dirname(os.path.abspath(__file__))

for filename in os.listdir(folder):
    if filename == script_name:
        continue

    new_name = filename.upper()
    if new_name != filename:
        os.rename(
            os.path.join(folder, filename),
            os.path.join(folder, new_name)
        )
        print(f"Renamed: {filename} -> {new_name}")

print("Done.")