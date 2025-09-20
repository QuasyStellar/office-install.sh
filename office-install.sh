#!/bin/bash

USER_HOME="/home/$USER"
WINE_PREFIX="$USER_HOME/.wine-msoffice"
WINE_FOLDER="$WINE_PREFIX/wine"
ICON_FOLDER="$WINE_PREFIX/icons"

echo "Welcome to Troplo's Microsoft Office installation script for WINE (Arch Linux, yay or paru required)"
echo ""
echo "Known issues:"
echo "- Broken Microsoft login (good thing!)"
echo "- Doesn't receive feature updates due to Windows 7 EoL."
echo "- OneNote and Teams don't work."
echo "- Excel may flicker when typing."
echo ""

# ----------------------------
# Helper Functions
# ----------------------------
detect_package_manager() {
    if command -v paru &>/dev/null; then
        echo "paru"
    elif command -v yay &>/dev/null; then
        echo "yay"
    else
        echo ""
    fi
}

confirm() {
    local prompt="$1"
    local default="$2"
    while true; do
        read -p "$prompt ($default): " choice
        case "$choice" in
            y|Y|yes ) return 0 ;;
            n|N|no ) return 1 ;;
            "" ) [[ "$default" == "y" ]] && return 0 || return 1 ;;
            * ) echo "Invalid choice. Enter y or n." ;;
        esac
    done
}

check_reinstall() {
    local folder="$1"
    if [ -d "$folder" ]; then
        confirm "Existing install detected in $folder. Reinstall?" "y"
        return $?
    else
        return 0
    fi
}

extract_archive() {
    local archive="$1"
    local dest="$2"
    if command -v bsdtar &>/dev/null; then
        bsdtar -xf "$archive" -C "$dest"
    else
        echo "bsdtar not found. Please install bsdtar."
        exit 1
    fi
}

# ----------------------------
# Installation Functions
# ----------------------------
install_prereq() {
    local pkg_manager
    pkg_manager=$(detect_package_manager)
    if [ -z "$pkg_manager" ]; then
        echo "Please install yay or paru first."
        exit 1
    fi
    echo "Installing prerequisites using $pkg_manager..."
    $pkg_manager -Sy --needed \
        glibc libice libsm libx11 libxext libxi freetype2 libpng zlib lcms2 libgl libxcursor libxrandr glu alsa-lib fontconfig gnutls gsm libcups libdbus libexif libgphoto2 libldap libpulse libxcomposite libxinerama libxml2 libxslt libxxf86vm mpg123 nss-mdns ocl-icd openal openssl sane v4l-utils wine p7zip wget samba automake autoconf fakeroot make gcc
    $pkg_manager -Sy --needed \
        lib32-glibc lib32-libice lib32-libsm lib32-libx11 lib32-libxext lib32-libxi lib32-freetype2 lib32-libpng lib32-zlib lib32-lcms2 lib32-libgl lib32-libxcursor lib32-libxrandr lib32-glu lib32-alsa-lib lib32-fontconfig lib32-libcups lib32-libdbus lib32-libexif lib32-libldap lib32-libpulse lib32-gnutls lib32-libxcomposite lib32-libxinerama lib32-libxml2 lib32-libxslt lib32-mpg123 lib32-nss-mdns lib32-openal lib32-openssl lib32-v4l-utils
}

install_wine() {
    if check_reinstall "$WINE_FOLDER"; then
        rm -rf "$WINE_FOLDER"
        echo "Downloading WINE 9.7..."
        wget -O "$USER_HOME/wine-9.7.zst" https://i.troplo.com/i/3512d274fa74.zst
        echo "Extracting WINE 9.7..."
        mkdir -p "$WINE_FOLDER"
        tar --use-compress-program=unzstd -xf "$USER_HOME/wine-9.7.zst" -C "$WINE_FOLDER"
        rm "$USER_HOME/wine-9.7.zst"
        if confirm "Kill current wineserver? (recommended)" "y"; then
            wineserver -k
        fi
        echo "WINE installed to $WINE_FOLDER"
    else
        echo "Skipping WINE installation"
    fi
}

download_icons() {
    mkdir -p "$ICON_FOLDER"
    wget -O "$USER_HOME/msoffice_script_icons.7z" https://i.troplo.com/i/0070f8a89f52.7z
    extract_archive "$USER_HOME/msoffice_script_icons.7z" "$ICON_FOLDER"
    rm "$USER_HOME/msoffice_script_icons.7z"
}

install_office() {
    local archive="$1"
    local dest="$2"
    local temp="$USER_HOME/temp_office.7z"

    if check_reinstall "$dest"; then
        rm -rf "$dest"
        echo "Downloading $archive..."
        wget -O "$temp" "$archive"
        echo "Extracting $archive..."
        mkdir -p "$WINE_PREFIX"
        extract_archive "$temp" "$WINE_PREFIX"
        rm "$temp"
        mv "$WINE_PREFIX/Microsoft_Office_365-"* "$dest"
        echo "Installed Office to $dest"
    else
        echo "Skipping Office installation for $dest"
    fi
}

register_office_items() {
    local prefix="$1"
    local type="$2"
    local desktop_dir="$USER_HOME/.local/share/applications"

    mkdir -p "$desktop_dir"

    declare -A apps=(
        [word]=WINWORD.EXE
        [excel]=EXCEL.EXE
        [powerpoint]=POWERPNT.EXE
        [access]=MSACCESS.EXE
        [publisher]=MSPUB.EXE
        [outlook]=OUTLOOK.EXE
    )

    for app in "${!apps[@]}"; do
        local exe="${apps[$app]}"
        local desktop="$desktop_dir/${app}-${type}.desktop"
        cat > "$desktop" << EOF
[Desktop Entry]
Type=Application
Name=Microsoft ${app^} [$type]
Icon=$ICON_FOLDER/${app}_48x1.png
Exec=sh -c 'PATH="$WINE_FOLDER/usr/bin:\$PATH" WINEARCH=win32 WINEPREFIX=$prefix $WINE_FOLDER/usr/bin/wine "$prefix/drive_c/Program Files/Microsoft Office/root/Office16/$exe" "%U"'
Categories=Office;
EOF
    done

    echo "File associations created for $type."
    xdg-desktop-menu forceupdate
}

# ----------------------------
# Main Menu
# ----------------------------
echo "Which version of Microsoft Office do you want to install?"
echo "1. Microsoft Office 365 ProPlus"
echo "2. Microsoft Office 2021 LTSC"
echo "3. Install both"
echo "4. (Re)install WINE 9.7"
echo "5. (Re)install dependencies"
echo "6. Launch winecfg for ProPlus"
echo "7. Launch winecfg for LTSC"
echo "8. Re-register file/menu associations"
read -p "Enter your choice: " choice

case $choice in
    1)
        install_prereq
        install_wine
        download_icons
        install_office "https://i.troplo.com/i/b22de9957c24.7z" "$WINE_PREFIX/ProPlus"
        register_office_items "$WINE_PREFIX/ProPlus" "ProPlus"
        ;;
    2)
        install_prereq
        install_wine
        download_icons
        install_office "https://i.troplo.com/i/721f0242a2c0.7z" "$WINE_PREFIX/LTSC"
        register_office_items "$WINE_PREFIX/LTSC" "LTSC"
        ;;
    3)
        install_prereq
        install_wine
        download_icons
        install_office "https://i.troplo.com/i/b22de9957c24.7z" "$WINE_PREFIX/ProPlus"
        install_office "https://i.troplo.com/i/721f0242a2c0.7z" "$WINE_PREFIX/LTSC"
        register_office_items "$WINE_PREFIX/ProPlus" "ProPlus"
        register_office_items "$WINE_PREFIX/LTSC" "LTSC"
        ;;
    4)
        install_wine
        ;;
    5)
        install_prereq
        ;;
    6)
        sh -c "PATH=\"$WINE_FOLDER/usr/bin:\$PATH\" WINEARCH=win32 WINEPREFIX=$WINE_PREFIX/ProPlus $WINE_FOLDER/usr/bin/winecfg"
        ;;
    7)
        sh -c "PATH=\"$WINE_FOLDER/usr/bin:\$PATH\" WINEARCH=win32 WINEPREFIX=$WINE_PREFIX/LTSC $WINE_FOLDER/usr/bin/winecfg"
        ;;
    8)
        download_icons
        [ -d "$WINE_PREFIX/LTSC" ] && register_office_items "$WINE_PREFIX/LTSC" "LTSC"
        [ -d "$WINE_PREFIX/ProPlus" ] && register_office_items "$WINE_PREFIX/ProPlus" "ProPlus"
        ;;
    *)
        echo "Invalid choice."
        ;;
esac
