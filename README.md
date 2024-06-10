# MicrosoftProgram_Install_and_Uninstall-Autoclean-Mod
autoclean all not valid MSI

# !!!Warning!!! #

**Running this program Mod will clean all MSI packages from Windows registry if it not found *.msi or *.json files in MSI install source folders located usually in `C:\ProgramData\Package Cache` and `C:\ProgramData\Microsoft\VisualStudio\Packages`**
so making a restore point first is recommended.**

Doubleclick on `Run_MF_WindowsInstaller_ElevatedToAdmin.ps1 - Shortcut` to start autoclean and it can take long time(~1hour for clean 100 MSI packages) and also can restart PC by some cause so you need start it again after PC restart to clean left over MSI.
I'm happy to restore my VS2022 functionality without reinstalling Windows 10 :) by auto-cleaning 144 MSI packages from the registry in about 1.5 hours instead of removing one by one package manually with the original program from Microsoft.
But I also removed valid 63 MSI packages(which blocks Windows SDK uninstall) manually using the original MicrosoftProgram_Install_and_Uninstall.meta.diagcab program so I included it also.
