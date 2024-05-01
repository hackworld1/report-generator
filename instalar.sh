RED="\033[01;31m"      # Issues/Errors
GREEN="\033[01;32m"    # Success
YELLOW="\033[01;33m"   # Warnings/Information
BLUE="\033[01;34m"     # Heading
BOLD="\033[01;01m"     # Highlight
RESET="\033[00m"       # Normal


echo -e "${RED}[+]${BLUE} Copiando ejecutables ${RESET}"
cp generate-reporte-web.py /usr/bin/
cp generate-reporte.py /usr/bin/
cp mergeDB.py  /usr/bin/

chmod a+x /usr/bin/generate-reporte.py
cp vulnerabilidades-web.xml /usr/share/lanscanner/vulnerabilidades-web.xml
cp vulnerabilidades.xml /usr/share/lanscanner/vulnerabilidades.xml
cp image.png /usr/share/lanscanner 

cp linux-fonts/* /usr/share/fonts/truetype