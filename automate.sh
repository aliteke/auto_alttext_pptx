#!/bin/zsh

BLUE='\033[0;36m'         # Blue (CYAN->36m, BLUE->34m)
RED='\033[0;31m'          # Red
COFF='\033[0m'       # Text Reset

FULL_PATH_TO_MODULES_ROOT="/Users/tekeoglua/Desktop/Research/autonomic/jhu/course/slides/Spring24"
PYTHON_SCRIPT="auto_alt_text_for_pptx.py"

echo -e "${RED}=============================================================================${COFF}"
echo -e "${BLUE}"; figlet -f slant "JHU-EP Auto-Alt-Text"
echo -e "${RED}=============================================================================${COFF}"

# Ask the user for the path to the modules root
echo "${BLUE}[+] Default PATH to Root Folder: ${RED}$FULL_PATH_TO_MODULES_ROOT${COFF}"
echo -n "${BLUE}[+] Do you want to change the path to the modules root? ${RED}(y/n): ${COFF}"
read CHANGE_PATH
if [ "$CHANGE_PATH" = "y" ]; then
    echo "Enter the full path to the modules root: "
    read FULL_PATH_TO_MODULES_ROOT
    echo "FULL_PATH_TO_MODULES_ROOT: ($FULL_PATH_TO_MODULES_ROOT)"
elif [ "$CHANGE_PATH" = "n" ]; then
    echo "Using Default PATH to Root Folder: $FULL_PATH_TO_MODULES_ROOT"
else
    echo "Invalid input. Enter y OR n. Exiting..."
    exit 1
fi

tree -P '*.pptx' $FULL_PATH_TO_MODULES_ROOT

echo -e "${RED}===================================================${COFF}"
echo "${RED}(1)${COFF} List alt-text for all images in .pptx files"
echo "${RED}(2)${COFF} Extract all images from .pptx files."
echo "${RED}(3)${COFF} Get auto-captions from Vertex-AI for images."
echo "${RED}(4)${COFF} Update alt-text for each image in place."
echo "${RED}(5)${COFF} Delete the image files from each (sub)directory."
echo "${RED}(6)${COFF} Delete the record_captions.csv file from each (sub)directory."
echo "${RED}(7)${COFF} Reset the alt-text for each image to empty string."
echo -e "${RED}===================================================${COFF}"
echo -n "${BLUE}Choose 1-7: ${COFF}"
read OPTION

if [ "$OPTION" = "1" -o "$OPTION" = "2" -o "$OPTION" = "3" -o "$OPTION" = "4" -o "$OPTION" = "5" -o "$OPTION" = "6"  -o "$OPTION" = "7" ]; then
	echo "[+] Selected Option: $OPTION"
else
	echo "[-] Invalid Option: $OPTION"
    exit 1
fi 

# Get the list of all the directories in the modules root
for dir in $FULL_PATH_TO_MODULES_ROOT/*; do
    # Check if the directory is a directory
    if [ -d "$dir" ]; then
        # Get the list of all the files in the directory
        for file in $dir/*; do
            # Check if the file is a file
            if [ -f "$file" ]; then
                # Check if the file is a .pptx file
                if [[ $file == *.pptx ]]; then
                    # Get the file name without the path
                    file_name=$(basename $file)
                    if [[ $file_name =~ "^~\\$" ]]; then        # On MacOS, there are some temporary files that start with ~, skip them.
                        continue
                    fi
                    # Get the file name without the extension
                    file_name_no_extension="${file_name%.*}"
                    # Get the directory name without the path
                    dir_name=$(basename $dir)
                    echo "FILE: $file, FILENAME: $file_name, FILENAME NO EXTENSION: $file_name_no_extension, DIRNAME: $dir_name"
                    if [ "$OPTION" = 1 ]; then
                        python3 $PYTHON_SCRIPT --pptx $file --list      # List all images and their Alt-Text for each PPTX file.
                        echo "==================================================="
                    elif [ "$OPTION" = 2 ]; then
                        python3 $PYTHON_SCRIPT --pptx $file             # Extract all images from each PPTX file.
                        echo "==================================================="
                    elif [ "$OPTION" = 3 ]; then                        # Get auto-captions from Vertex-AI for images.
                        if [ -z "$BEARER" ]; then
                            echo "[===] Please enter an active BEARER TOKEN (cloudshell# gcloud auth print-access-token)"
                            read BEARER
                        fi
                        python3 $PYTHON_SCRIPT --dir "$FULL_PATH_TO_MODULES_ROOT/$dir_name" --output "$FULL_PATH_TO_MODULES_ROOT/$dir_name/record_captions.csv" --bearer $BEARER
                    elif [ "$OPTION" = 4 ]; then
                        echo "[+] Before Updating Alt-Text: $file"
                        python3 $PYTHON_SCRIPT --pptx $file --list      # List all images and their Alt-Text for each PPTX file.
                        echo "[+] Updating Alt-Text: $file"
                        python3 $PYTHON_SCRIPT --pptx $file --update "$FULL_PATH_TO_MODULES_ROOT/$dir_name/record_captions.csv"      # Both $file and record_captions.csv should be absolute paths.
                        echo "[+] After Updating Alt-Text: $file"
                        python3 $PYTHON_SCRIPT --pptx $file --list      # List all images and their Alt-Text for each PPTX file.
                    elif [ "$OPTION" = 5 ]; then
                        echo "[+] Cleaning the directories from .png, .tiff, .jpg and .gif files..."
                        find $FULL_PATH_TO_MODULES_ROOT -iname "*.tiff" -o -iname "*.png" -o -iname "*.jpg"  -o -iname "*.gif" | xargs -I {} rm -v {}
                        exit 0
                    elif [ "$OPTION" = 6 ]; then
                        echo "[+] Cleaning the directories from .csv files..."
                        find $FULL_PATH_TO_MODULES_ROOT -iname "*.csv" | xargs -I {} rm -v {}
                        exit 0

                    elif [ "$OPTION" = 7 ]; then
                        echo "[+] Resetting the alt-text for each image in place to empty string."
                        python3 $PYTHON_SCRIPT --pptx $file --reset
                        echo "==================================================="

                    fi

                fi
            fi
        done
    fi
done
