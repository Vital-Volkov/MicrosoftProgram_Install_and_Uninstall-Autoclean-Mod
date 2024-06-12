# MicrosoftProgram_Install_and_Uninstall-Autoclean-Mod
batch autoclean MSI products

# !!!Warning!!! #

**Running this program Mod will clean MSI packages from Windows registry so making a restore point first is recommended.**

Doubleclick on `Run_MF_WindowsInstaller_ElevatedToAdmin.ps1 - Shortcut` to start program and select clean mode:
- Just press `Enter` key for clean all MSI packages from Windows registry if it not found *.msi or *.json files in MSI install source folders located usually in `C:\ProgramData\Package Cache` and `C:\ProgramData\Microsoft\VisualStudio\Packages`
  it can take long time(~1hour for clean 100 MSI packages) and also can restart PC by some cause so you need start it again after PC restart to clean left over MSI.

- Select required range of MSI's to clean from-to by input their index numbers.

I run into trouble when decided to free some space on SSD and delete files `C:\ProgramData\Microsoft\VisualStudio\Packages` thinking that it is just cache after VS components installations but as result I get lost VS and installer no more see
any version of VS installed on my PC while I has VS2019 and 2022. And even after manually uninstalling all VS and cleaning with `InstallCleanup.exe -f` and installing VS2022 again it is not working correctly and throws errors in compilation like `fxc.exe not found` etc.
I found problem in registry about installed MSI have dublicates and now even VS components uninstall throw errors `account already exist` so I found <a href="https://support.microsoft.com/en-au/topic/fix-problems-that-block-programs-from-being-installed-or-removed-cca7d1b6-65a9-3d98-426b-e9f927e1eb4d"><b>useful program from microsoft</b></a> unpack with WinRar and modify it to auto batch clean many MSI's by condition instead of manual check existance of each MSI by ID and delete one by one and spend too many hours for this.

After 1.5 hours AFK autoclean does all the work to clean 144 MSI with no existing *.msi with only two PC restarts. Then I finally removed WindowsSDK after manually uninstalling 64 corresponding MSI which blocks `WindowsSDK` uninstall.
I ended up with only 34 MSI installed on my PC. Install VS2022 and all working fine this time so I'm happy to fix my PC registry and VS2022 functionality without reinstalling Windows 10 :) and share this solution to you.

P.S. VS have option `Keep download cache after installation` and default path to the folder `C:\ProgramData\Microsoft\VisualStudio\Packages` and this is why I delete files inside thinking it's safe but never delete this files manually to avoid such trouble
because when you untick the option VS keep *.json file in replace of *.msi and such give you free space on disk but when you remove components VS first download(using info from *.json) same *.msi which was used for install and use it to uninstall component.
