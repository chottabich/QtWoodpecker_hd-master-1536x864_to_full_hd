## QtWoodpecker_hd
### Extra libraries are required in addition to those for the installation of QTVCP for additional functionality

#### For viewing job setup sheets and help pages formatted in HTML
sudo apt-get install python3-pyqt5.qtwebengine

#### To communicate with Huanyang GT series VFDs
sudo apt-get install python3-pymodbus

#### To create Z level height maps when using Z level compensation
sudo apt-get install python3-numpy

sudo apt-get install python3-scipy

sudo apt-get install python3-matplotlib

#### To view PDF files
sudo apt-get install python3-poppler-qt5

#### To Install
Unzip the archive to the /home/name/linuxcnc/configs folder. Place style files and woodpecker.qrc with replacement at usr/share/qtvcp/screen/woodpecker, you will be able to switch styles from settings tab. The size changes to F10(min), F11(max). In the woodpecker_xyza.ini file, remove -f from the DISPLAY section. 
