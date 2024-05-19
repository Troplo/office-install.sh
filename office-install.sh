#!/bin/bash

echo "Welcome to Troplo's Microsoft Office installation script for WINE (Arch Linux, yay required)"
echo ""
echo "Known issues:"
echo "- Broken Microsoft login (good thing!)"
echo "- Doesn't receive feature updates due to Windows 7 EoL."
echo "- OneNote, and Teams don't work."
echo "- Excel has a tendency to flicker when typing."
echo ""
install_prereq() {
    echo "Installing pre-requisite packages..."
    yay -Sy glibc libice libsm libx11 libxext libxi freetype2 libpng zlib lcms2 libgl libxcursor libxrandr glu alsa-lib fontconfig gnutls gsm libcups libdbus libexif libgphoto2 libldap libpulse libxcomposite libxinerama libxml2 libxslt libxxf86vm mpg123 nss-mdns ocl-icd openal openssl sane v4l-utils wine p7zip wget samba --needed
    yay -Sy lib32-glibc lib32-libice lib32-libsm lib32-libx11 lib32-libxext lib32-libxi lib32-freetype2 lib32-libpng lib32-zlib lib32-lcms2 lib32-libgl lib32-libxcursor lib32-libxrandr lib32-glu lib32-alsa-lib lib32-fontconfig lib32-libcups lib32-libdbus lib32-libexif lib32-libldap lib32-libpulse lib32-gnutls lib32-libxcomposite lib32-libxinerama lib32-libxml2 lib32-libxslt lib32-mpg123 lib32-nss-mdns lib32-openal lib32-openssl lib32-v4l-utils --needed
}
install_wine() {
    local wine_folder="/home/$USER/.wine-msoffice/wine"
    if check_reinstall "$wine_folder"; then
        if [[ -n "$wine_folder" ]]; then
            rm -rf "$wine_folder"
        fi
        echo $wine_folder
        echo "Downloading WINE 9.7..."
        wget https://i.troplo.com/i/3512d274fa74.zst -O /home/$USER/wine-9.7.zst
        echo "Extracting WINE 9.7..."
        mkdir -p "/home/$USER/.wine-msoffice/wine"
        tar --use-compress-program=unzstd -xf /home/$USER/wine-9.7.zst -C "$wine_folder"
        rm /home/$USER/wine-9.7.zst
        if check_kill_wineserver; then
            wineserver -k
        fi
        echo "WINE package has been installed to $wine_folder"
    else
        echo "Skipping reinstallation of WINE."
    fi
}

check_reinstall() {
    local folder=$1
    if [ -d "$folder" ]; then
        read -p "An existing install was detected in $folder. Do you want to reinstall? (y/n): " choice
        case "$choice" in
            y|Y|yes ) return 0 ;; # Proceed with reinstallation
            n|N|no ) return 1 ;; # Skip reinstallation
            * ) echo "Invalid choice. Please enter y or n." ; check_reinstall "$folder" ;;
        esac
    else
        return 0 # Folder doesn't exist, proceed with installation
    fi
}

check_kill_wineserver() {
    read -p "Would you like to end the current wineserver process? (recommended) (y/n): " choice
    case "$choice" in
        y|Y|yes ) return 0 ;; # Proceed with reinstallation
        n|N|no ) return 1 ;; # Skip reinstallation
        * ) echo "Invalid choice. Please enter y or n." ; check_reinstall "$folder" ;;
    esac
}


install_proplus() {
    if check_reinstall "/home/$USER/.wine-msoffice/ProPlus"; then
        rm -rf /home/$USER/.wine-msoffice/ProPlus
        rm -rf /home/$USER/.wine-msoffice/Microsoft_Office_365-4
        echo "Installing/Reinstalling Microsoft Office 2021 LTSC..."
        echo "Downloading Microsoft Office 365 ProPlus..."
        wget https://i.troplo.com/i/b22de9957c24.7z -O /home/$USER/msoffice.7z
        echo "Extracting Microsoft Office 365 ProPlus..."
        mkdir -p ~/.wine-msoffice
        7z x /home/$USER/msoffice.7z -o/home/$USER/.wine-msoffice
        rm /home/$USER/msoffice.7z
        mv /home/$USER/.wine-msoffice/Microsoft_Office_365-4 /home/$USER/.wine-msoffice/ProPlus
        echo "Microsoft Office 365 ProPlus has been installed to ~/.wine-msoffice/ProPlus"

        # Create file associations
        echo "Creating file associations..."
        mkdir -p ~/.local/share/applications
        cat > ~/.local/share/applications/word-proplus.desktop << EOF
[Desktop Entry]
Type=Application
Name=Microsoft Word [ProPlus]
Icon=/home/$USER/.wine-msoffice/ProPlus/drive_c/Program Files/Microsoft Office/root/Office16/WINWORD.EXE
Exec=sh -c 'PATH="/home/$USER/.wine-msoffice/wine/usr/bin:$PATH" WINEARCH=win32 WINEPREFIX=/home/$USER/.wine-msoffice/ProPlus /home/$USER/.wine-msoffice/wine/usr/bin/wine "/home/$USER/.wine-msoffice/ProPlus/drive_c/Program Files/Microsoft Office/root/Office16/WINWORD.EXE" "%U"'
Categories=Office;
MimeType=application/msword;application/vnd.openxmlformats-officedocument.wordprocessingml.document;
EOF

        cat > ~/.local/share/applications/access-proplus.desktop << EOF
[Desktop Entry]
Type=Application
Name=Microsoft Access [ProPlus]
Icon=/home/$USER/.wine-msoffice/ProPlus/drive_c/Program Files/Microsoft Office/root/Office16/MSACCESS.EXE
Exec=sh -c 'PATH="/home/$USER/.wine-msoffice/wine/usr/bin:$PATH" WINEARCH=win32 WINEPREFIX=/home/$USER/.wine-msoffice/ProPlus /home/$USER/.wine-msoffice/wine/usr/bin/wine "/home/$USER/.wine-msoffice/ProPlus/drive_c/Program Files/Microsoft Office/root/Office16/MSACCESS.EXE" "%U"'
Categories=Office;
MimeType=application/msaccess;
EOF

        cat > ~/.local/share/applications/excel-proplus.desktop << EOF
[Desktop Entry]
Type=Application
Name=Microsoft Excel [ProPlus]
Icon=/home/$USER/.wine-msoffice/ProPlus/drive_c/Program Files/Microsoft Office/root/Office16/EXCEL.EXE
Exec=sh -c 'PATH="/home/$USER/.wine-msoffice/wine/usr/bin:$PATH" WINEARCH=win32 WINEPREFIX=/home/$USER/.wine-msoffice/ProPlus /home/$USER/.wine-msoffice/wine/usr/bin/wine "/home/$USER/.wine-msoffice/ProPlus/drive_c/Program Files/Microsoft Office/root/Office16/EXCEL.EXE" "%U"'
Categories=Office;
MimeType=application/msexcel;application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;
EOF

        cat > ~/.local/share/applications/powerpoint-proplus.desktop << EOF
[Desktop Entry]
Type=Application
Name=Microsoft PowerPoint [ProPlus]
Icon=/home/$USER/.wine-msoffice/ProPlus/drive_c/Program Files/Microsoft Office/root/Office16/POWERPNT.EXE
Exec=sh -c 'PATH="/home/$USER/.wine-msoffice/wine/usr/bin:$PATH" WINEARCH=win32 WINEPREFIX=/home/$USER/.wine-msoffice/ProPlus /home/$USER/.wine-msoffice/wine/usr/bin/wine "/home/$USER/.wine-msoffice/ProPlus/drive_c/Program Files/Microsoft Office/root/Office16/POWERPNT.EXE" "%U"'
Categories=Office;
MimeType=application/vnd.ms-powerpoint;application/vnd.openxmlformats-officedocument.presentationml.presentation;
EOF

        cat > ~/.local/share/applications/publisher-proplus.desktop << EOF
[Desktop Entry]
Type=Application
Name=Microsoft Publisher [ProPlus]
Icon=/home/$USER/.wine-msoffice/ProPlus/drive_c/Program Files/Microsoft Office/root/Office16/MSPUB.EXE
Exec=sh -c 'PATH="/home/$USER/.wine-msoffice/wine/usr/bin:$PATH" WINEARCH=win32 WINEPREFIX=/home/$USER/.wine-msoffice/ProPlus /home/$USER/.wine-msoffice/wine/usr/bin/wine "/home/$USER/.wine-msoffice/ProPlus/drive_c/Program Files/Microsoft Office/root/Office16/MSPUB.EXE" "%U"'
Categories=Office;
MimeType=application/vnd.ms-publisher;
EOF

        cat > ~/.local/share/applications/outlook-proplus.desktop << EOF
[Desktop Entry]
Type=Application
Name=Microsoft Outlook [ProPlus]
Icon=/home/$USER/.wine-msoffice/ProPlus/drive_c/Program Files/Microsoft Office/root/Office16/OUTLOOK.EXE
Exec=sh -c 'PATH="/home/$USER/.wine-msoffice/wine/usr/bin:$PATH" WINEARCH=win32 WINEPREFIX=/home/$USER/.wine-msoffice/ProPlus /home/$USER/.wine-msoffice/wine/usr/bin/wine "/home/$USER/.wine-msoffice/ProPlus/drive_c/Program Files/Microsoft Office/root/Office16/OUTLOOK.EXE" "%U"'
Categories=Office;
MimeType=application/microsoft-outlook;
EOF

        echo "File associations created."
        echo "Updating menu items"
        xdg-desktop-menu forceupdate
    else
        echo "Skipping reinstallation of Microsoft Office 2021 LTSC."
    fi
}

install_ltsc() {
    if check_reinstall "/home/$USER/.wine-msoffice/LTSC"; then
        rm -rf /home/$USER/.wine-msoffice/LTSC
        rm -rf /home/$USER/.wine-msoffice/Microsoft_Office_365-3
        echo "Downloading Microsoft Office 2021 LTSC..."
        wget https://i.troplo.com/i/721f0242a2c0.7z -O /home/$USER/msoffice_ltsc.7z
        echo "Extracting Microsoft Office 365 LTSC..."
        mkdir -p ~/.wine-msoffice
        7z x /home/$USER/msoffice_ltsc.7z -o/home/$USER/.wine-msoffice
        rm /home/$USER/msoffice_ltsc.7z
        mv /home/$USER/.wine-msoffice/Microsoft_Office_365-3 /home/$USER/.wine-msoffice/LTSC
        echo "Microsoft Office 365 ProPlus has been installed to ~/.wine-msoffice/LTSC"

        # Create file associations
        echo "Creating file associations..."
        mkdir -p ~/.local/share/applications

        cat > ~/.local/share/applications/word-ltsc.desktop << EOF
[Desktop Entry]
Type=Application
Name=Microsoft Word [LTSC]
Icon=/home/$USER/.wine-msoffice/LTSC/drive_c/Program Files/Microsoft Office/root/Office16/WINWORD.EXE
Exec=sh -c 'PATH="/home/$USER/.wine-msoffice/wine/usr/bin:$PATH" WINEARCH=win32 WINEPREFIX=/home/$USER/.wine-msoffice/LTSC /home/$USER/.wine-msoffice/wine/usr/bin/wine "/home/$USER/.wine-msoffice/LTSC/drive_c/Program Files/Microsoft Office/root/Office16/WINWORD.EXE" "%U"'
Categories=Office;
MimeType=application/msword;application/vnd.openxmlformats-officedocument.wordprocessingml.document;
EOF

        cat > ~/.local/share/applications/access-ltsc.desktop << EOF
[Desktop Entry]
Type=Application
Name=Microsoft Access [LTSC]
Icon=/home/$USER/.wine-msoffice/LTSC/drive_c/Program Files/Microsoft Office/root/Office16/MSACCESS.EXE
Exec=sh -c 'PATH="/home/$USER/.wine-msoffice/wine/usr/bin:$PATH" WINEARCH=win32 WINEPREFIX=/home/$USER/.wine-msoffice/LTSC /home/$USER/.wine-msoffice/wine/usr/bin/wine "/home/$USER/.wine-msoffice/LTSC/drive_c/Program Files/Microsoft Office/root/Office16/MSACCESS.EXE" "%U"'
Categories=Office;
MimeType=application/msaccess;
EOF

        cat > ~/.local/share/applications/excel-ltsc.desktop << EOF
[Desktop Entry]
Type=Application
Name=Microsoft Excel [LTSC]
Icon=/home/$USER/.wine-msoffice/LTSC/drive_c/Program Files/Microsoft Office/root/Office16/EXCEL.EXE
Exec=sh -c 'PATH="/home/$USER/.wine-msoffice/wine/usr/bin:$PATH" WINEARCH=win32 WINEPREFIX=/home/$USER/.wine-msoffice/LTSC /home/$USER/.wine-msoffice/wine/usr/bin/wine "/home/$USER/.wine-msoffice/LTSC/drive_c/Program Files/Microsoft Office/root/Office16/EXCEL.EXE" "%U"'
Categories=Office;
MimeType=application/msexcel;application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;
EOF

        cat > ~/.local/share/applications/powerpoint-ltsc.desktop << EOF
[Desktop Entry]
Type=Application
Name=Microsoft PowerPoint [LTSC]
Icon=/home/$USER/.wine-msoffice/LTSC/drive_c/Program Files/Microsoft Office/root/Office16/POWERPNT.EXE
Exec=sh -c 'PATH="/home/$USER/.wine-msoffice/wine/usr/bin:$PATH" WINEARCH=win32 WINEPREFIX=/home/$USER/.wine-msoffice/LTSC /home/$USER/.wine-msoffice/wine/usr/bin/wine "/home/$USER/.wine-msoffice/LTSC/drive_c/Program Files/Microsoft Office/root/Office16/POWERPNT.EXE" "%U"'
Categories=Office;
MimeType=application/vnd.ms-powerpoint;application/vnd.openxmlformats-officedocument.presentationml.presentation;
EOF

        cat > ~/.local/share/applications/publisher-ltsc.desktop << EOF
[Desktop Entry]
Type=Application
Name=Microsoft Publisher [LTSC]
Icon=/home/$USER/.wine-msoffice/LTSC/drive_c/Program Files/Microsoft Office/root/Office16/MSPUB.EXE
Exec=sh -c 'PATH="/home/$USER/.wine-msoffice/wine/usr/bin:$PATH" WINEARCH=win32 WINEPREFIX=/home/$USER/.wine-msoffice/LTSC /home/$USER/.wine-msoffice/wine/usr/bin/wine "/home/$USER/.wine-msoffice/LTSC/drive_c/Program Files/Microsoft Office/root/Office16/MSPUB.EXE" "%U"'
Categories=Office;
MimeType=application/vnd.ms-publisher;
EOF

        cat > ~/.local/share/applications/outlook-ltsc.desktop << EOF
[Desktop Entry]
Type=Application
Name=Microsoft Outlook [LTSC]
Icon=/home/$USER/.wine-msoffice/LTSC/drive_c/Program Files/Microsoft Office/root/Office16/OUTLOOK.EXE
Exec=sh -c 'PATH="/home/$USER/.wine-msoffice/wine/usr/bin:$PATH" WINEARCH=win32 WINEPREFIX=/home/$USER/.wine-msoffice/LTSC /home/$USER/.wine-msoffice/wine/usr/bin/wine "/home/$USER/.wine-msoffice/LTSC/drive_c/Program Files/Microsoft Office/root/Office16/OUTLOOK.EXE" "%U"'
Categories=Office;
MimeType=application/microsoft-outlook;
EOF

        echo "File associations created."
        echo "Updating menu items"
        xdg-desktop-menu forceupdate
    else
        echo "Skipping reinstallation of Microsoft Office 2021 LTSC."
    fi
}

# Main script
echo "Which version of Microsoft Office do you want to install?"
echo "1. Microsoft Office 365 ProPlus [Light theme document, Animations work, Access works]"
echo "2. Microsoft Office 2021 LTSC [Dark theme document option, Broken animations, Broken Access, OneNote, Teams, Outlook, and Publisher.]"
echo "3. Install both"
echo "4. (Re)install WINE 9.7 (Does not replace system WINE, required for Office, amd64/x86)"
echo "5. (Re)install pacman/yay dependencies"
read -p "Enter your choice [1, 2 or 3]: " choice

case $choice in
    1)
        install_prereq
        install_wine
        install_proplus
        ;;
    2)
        install_prereq
        install_wine
        install_ltsc
        ;;
    3)
        install_prereq
        install_wine
        install_proplus
        install_ltsc
        ;;
    4)
        install_wine
        ;;
    5)
        install_prereq
        ;;
    *)
        echo "Invalid choice. Please enter 1, 2 or 3."
        ;;
esac
