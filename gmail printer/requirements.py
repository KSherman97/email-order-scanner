'''
Author: Kyle Sherman
Created: 10/06/2024
Updated: 10/06/2024

Automatically install any dependencies listed in requirements.txt
'''

import subprocess
import sys

def install_requirements():
    try:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', '-r', 'requirements.txt'])
        print('all dependencies have been installed')
    except subprocess.CalledProcessError as error:
        print(f"error occurred while installing dependencies: {error}")
        sys.exit(1)

if __name__ == "__main__":
    install_requirements()