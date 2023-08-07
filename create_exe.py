import subprocess

def run_pyinstaller_onefile(script_file):
    try:
        # Run the PyInstaller command with the given script file
        subprocess.run(['pyinstaller', '--onefile', "--distpath", ".", script_file], check=True)
        print(f"Successfully created the executable for {script_file}")
    except subprocess.CalledProcessError as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    # Replace 'your_script.py' with the name of your Python script
    run_pyinstaller_onefile('auto_so.py')