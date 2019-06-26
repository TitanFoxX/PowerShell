# - Check if Visual Studio Community is already installed at the target location.
# - If located, execute the vs_installer.exe to update the current installation.

If(Test-Path -Path "C:\Program Files (x86)\Microsoft Visual Studio\Installer"){
    Set-Location -Path "C:\Program Files (x86)\Microsoft Visual Studio\Installer"
    .\vs_installer.exe update --quiet --norestart --installpath "C:\Program Files (x86)\Microsoft Visual Studio\2017\Community"
}
Else{
}