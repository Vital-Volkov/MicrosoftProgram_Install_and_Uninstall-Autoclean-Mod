#################################################################################
# Copyright (c) 2011, Microsoft Corporation. All rights reserved.
#
# You may use this code and information and create derivative works of it,
# provided that the following conditions are met:
# 1. This code and information and any derivative works may only be used for
# troubleshooting a) Windows and b) products for Windows, in either case using
# the Windows Troubleshooting Platform
# 2. Any copies of this code and information
# and any derivative works must retain the above copyright notice, this list of
# conditions and the following disclaimer.
# 3. THIS CODE AND INFORMATION IS PROVIDED `AS IS'' WITHOUT WARRANTY OF ANY 
# KIND, WHETHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED
# WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE. IF THIS
# CODE AND INFORMATION IS USED OR MODIFIED, THE ENTIRE RISK OF USE OR RESULTS IN
# CONNECTION WITH THE USE OF THIS CODE AND INFORMATION REMAINS WITH THE USER.
#################################################################################
. .\utils_SetupEnv.ps1
. .\utils_SdpExtension.ps1
Import-LocalizedData -BindingVariable LocalizedStrings -FileName Strings
function ProductListingBuild
{      
$MasterHash = New-Object System.Collections.ArrayList
$path = [MakeStringTest]::loop()
$CombinedListing = New-Object System.Collections.ArrayList
foreach($pathKey in $path)
{
    [void]$CombinedListing.Add($pathKey)    
}
$MasterHashs = New-Object System.Collections.ArrayList #creating first listing item
$MasterHash= @{}
$MasterHash.Add("Name",$LocalizedStrings.WindowsInstaller_IID_Not_Listed)
$MasterHash.Add("Value",$LocalizedStrings.WindowsInstaller_IID_Not_Listed)
$MasterHash.Add("Description",$LocalizedStrings.WindowsInstaller_IID_Not_Listed)
$MasterHashs+=$MasterHash
$SortArray= new-object 'object[]' $path.Count # create array for sorting 
 
#Sort Products for Dialog
$i=0
foreach ($Product in $path)
{
$SortArray[$i]=$Product.Name+";"+$Product.Value
$i=$i+1
}
$SortArray=$SortArray|Sort-Object #Sort temp array
foreach ($Product in $SortArray)
{
$temp=$product.indexof(";")
$ProductCode=$product.substring($temp+1)
$ProductName=$product.substring(0,$temp)
$MasterHash= @{}
    #Loc code
    switch ($ProductName)
    {
        "Not Listed"
        {
        $ProductName=$LocalizedStrings.WindowsInstaller_IID_Not_Listed
        break;
        }
        
        "Name not available"
        {
        $ProductName=$LocalizedStrings.WindowsInstaller_IID_Name_no_available
        break;
        }
    }
                 
    
 [void]$MasterHash.Add("Name",$ProductName)
[void]$MasterHash.Add("Value",$ProductCode)
[void]$MasterHash.Add("Description",$ProductCode)
$MasterHashs+=$MasterHash    
}
     
$MasterHashs 
}

function HCRRegistryCheckMultiString
{
$HCRRegistryCheckMultiString=
@"
using System.Collections;
using System;
using Microsoft.Win32;  
public class HCRRegistryCheckMultiString
{
        public static string[] HCRRegistryCheck(string KeyName)
        {           
            RegistryKey key = Registry.ClassesRoot.OpenSubKey(KeyName, false);
            if (key != null)
            {
                string[] MultiNames = (string[])key.GetValue("Patches");
                return MultiNames;
            }
            return null;
        }
}
"@
            
   if ([HCRRegistryCheckMultiString] -eq $null)
   {
   #Code here will never be run;  the if statement with throw an exception because the type doesn't exist
   }
   else
   {
   #type is loaded, do nothing
   }
Trap [Exception]
{
  #type doesn't exist.  Load it
  Add-Type -TypeDefinition $HCRRegistryCheckMultiString -PassThru 
  continue
}
  
}

function PatchInfo
{
$PatchInfo=
@"
using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.Runtime.InteropServices;
using System.Xml;
public class CollectPatchInformation
{
[DllImport("msi.dll")]
public static extern int  MsiExtractPatchXMLData(string szPatchPath,int dwReserved, [Out] StringBuilder  szXMLData,ref int pcchXMLData);
        
        public static string GetPatchInfoFromPath(string szPath,string szReturntype)
        {
            StringBuilder szXMLData = new StringBuilder(1024);
            int ipcchXMLData = 0, iError = 0;
            string szPatchGUID="";
            iError = MsiExtractPatchXMLData(szPath, 0, szXMLData, ref ipcchXMLData);
            if (iError == 234)
            {
                StringBuilder szLargeXMLData = new StringBuilder(ipcchXMLData + 1);
                iError = MsiExtractPatchXMLData(szPath, 0, szLargeXMLData, ref ipcchXMLData);
                szPatchGUID= (PatchExtendedInfo(szLargeXMLData.ToString (), szReturntype)).ToString();    
                return (szPatchGUID);        
            }
   szPatchGUID= (PatchExtendedInfo(szXMLData.ToString(), szReturntype));    
   return (szPatchGUID);  
        }
        public static string PatchExtendedInfo(string xmlData,string szReturnType)
        {
   XmlDocument doc = new XmlDocument();
   doc.InnerXml = xmlData.ToString ();
   XmlElement root = doc.DocumentElement;
   return root.GetAttribute(szReturnType).ToString();
        }
}
"@
            
   if ([CollectPatchInformation] -eq $null)
   {
   #Code here will never be run;  the if statement with throw an exception because the type doesn't exist
   }
   else
   {
   #type is loaded, do nothing
   }
Trap [Exception]
{
  #type doesn't exist.  Load it
  $Assem = @("System.Xml, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
  Add-Type -TypeDefinition $PatchInfo -PassThru -ReferencedAssemblies $Assem
  continue
}

}
function MSIGetProductCodeFromPackage
{
$MSIGetProductCodeFromPackage=
@"
using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.Runtime.InteropServices;
public class CollectMSIProductCode
{
        [DllImport("msi.dll")]
        public static extern Int32 MsiOpenDatabase(
            string szDatabasePath,
            string szPersist,
            ref int phDatabase
            );
        [DllImport("msi.dll")]
        public static extern Int32 MsiDatabaseOpenView(
                        int hDatabase,
                        string szQuery,
                        ref int phView
                        );
        [DllImport("msi.dll")]
        public static extern Int32 MsiViewExecute(
            int hView,
            int hRecord
            );
        [DllImport("msi.dll")]
        public static extern Int32 MsiViewFetch(
            int hView,
            ref int phRecord
            );
        [DllImport("msi.dll")]
        public static extern int MsiRecordGetString(
            int hRecord,
            int iField,
            [MarshalAs(UnmanagedType.LPStr)] StringBuilder szValueBuf,
            ref int pcchValueBuf
            );

        [DllImport("msi.dll")]
        public static extern int MsiCreateRecord(
            int cParams
            );
        [DllImport("msi.dll")]
        public static extern int MsiCloseHandle(
            int hAny    // Installer handle
            );
            
        [DllImport("msi.dll")]
        public static extern int MsiCloseAllHandles();
        
        
        public static string MSIGetProductCodeFromPackage(string szPackagePath)
        {
            string szSQLQuery = "", szValue="ProductCode";
            int iWhd = 0, phView = 0, iRecord, iError = 0, ipcchValueBuf=0;
            StringBuilder sbRecord = new StringBuilder(0);
            try
            {
                iRecord = MsiCreateRecord(1);
                szSQLQuery = "SELECT `Value` FROM `Property` WHERE `Property`= " + "'" + szValue + "'";
                iError=MsiOpenDatabase(szPackagePath, "MSIDBOPEN_READONLY",ref iWhd);
                iError=MsiDatabaseOpenView(iWhd, szSQLQuery, ref phView);
                iError=MsiViewExecute(phView, 0);
                iError=MsiViewFetch(phView, ref iRecord);
                iError = MsiRecordGetString(iRecord, 1, sbRecord, ref ipcchValueBuf);
                if (iError != 0)
                {
                    ipcchValueBuf += 1;
                    sbRecord = new StringBuilder(ipcchValueBuf);
                    iError = MsiRecordGetString(iRecord, 1, sbRecord, ref ipcchValueBuf);
                    MsiCloseHandle(iWhd);
                    MsiCloseHandle(iRecord);
                    MsiCloseHandle(phView);
                    
                    if (iError == 0)
                    {
                        return sbRecord.ToString();
                    }
                    else
                    {
                        return "NA";
                    }
                }
            }
            catch
            {
            }
            MsiCloseAllHandles();
            return "NA";          
        }
}
"@
             
   if ([CollectMSIProductCode] -eq $null)
   {
   #Code here will never be run;  the if statement with throw an exception because the type doesn't exist
   }
   else
   {
   #type is loaded, do nothing
   }
Trap [Exception]
{    
    #type doesn't exist.  Load it
    $type = Add-Type -TypeDefinition $MSIGetProductCodeFromPackage -PassThru 
 continue
}
}

function ProductListing
{
$ProductListing=
@"
using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.Runtime.InteropServices;
  public class MakeStringTest
        {
            [DllImport("msi.dll")]
            public static extern Int32 MsiEnumProducts(int iProductIndex, StringBuilder lpProductBuf);
            [DllImport("msi.dll")]
            public static extern Int32 MsiGetProductInfo(string szProduct, string szProperty, StringBuilder lpValueBuf, ref int pcchValueBuf);
            public static ArrayList AllProductCodes = new ArrayList();
            public static ArrayList HashArrary = new ArrayList();
            public static ArrayList loop()
            {
                AllProductCodes = MsiEnumProducts();
                HashArrary = MsiGetProductInfoExtended(AllProductCodes);
                return (HashArrary);
            }
            public static string GetMSIProductInformation(string szProductCode, string szInformationRequested)
            {
                string szReturnValue = "";
                StringBuilder sb = new StringBuilder(1024);
                int iret = 1024, iErrorReturn = 0;
                iErrorReturn = MsiGetProductInfo(szProductCode, szInformationRequested, sb, ref iret);
                if (iErrorReturn == 0)
                {
                    szReturnValue = sb.ToString();
                }
                else
                {
                    szReturnValue = "Unknown";
                }
                return szReturnValue;
            }
            public static ArrayList MsiGetProductInfoExtended(ArrayList alAllProductCodes)
            {
                ArrayList alProdList = new ArrayList();
                string szProductName = "";
                foreach (string szCheck in alAllProductCodes)
                {
                    szProductName = GetMSIProductInformation(szCheck, "ProductName");
                    if (szProductName == "")
                    {
                        szProductName = "Name not available";
                    }
                    Hashtable Hash = new Hashtable();
                    Hash.Add("Name", szProductName);
                    Hash.Add("Value", szCheck);
                    Hash.Add("Description", szCheck);
                    Hash.Add("ExtensionPoint", "<Default/><Icon>@resource.dll,-104</Icon>");
                    alProdList.Add(Hash);
                }
                return alProdList;
            }
            public static ArrayList MsiEnumProducts()
            {
                int iErrorReturn = 0, iProductIndex = 0;
                StringBuilder szProdCode = new StringBuilder(39);
                string szProductCode = "";
                ArrayList alProdCodes = new ArrayList();
                while (iErrorReturn != 259)
                {
                    iErrorReturn = MsiEnumProducts(iProductIndex, szProdCode);
                    if (iErrorReturn == 0)
                    {
                        szProductCode = szProdCode.ToString();
                        alProdCodes.Add(szProductCode);
                    }

                    iProductIndex++;
                }
                return (alProdCodes);
            }
        }
"@
             
   if ([MakeStringTest] -eq $null)
   {
   #Code here will never be run;  the if statement with throw an exception because the type doesn't exist
   }
   else
   {
   #type is loaded, do nothing
   }
Trap [Exception]
{
    #type doesn't exist.  Load it
    $type = Add-Type -TypeDefinition $ProductListing -PassThru
continue
}
}

function CleanUpPatches
{
param ($RegistryPatch, $PatchChildToDelete)
Import-LocalizedData -BindingVariable LocalizedStrings -FileName Strings
#check cache 
$PatchGUIDList=New-Object System.Collections.ArrayList
[string[]]$MultiStringHolder
foreach($Property in (Get-ItemProperty -Path $RegistryPatch).AllPatches)
{
    if ($Property -and ($Property -ne $PatchGUID))
    {
        $PatchGUIDList.add($Property)
    }         
}
[string[]]$MultiStringHolder=$PatchGUIDList
Writelog(($LocalizedStrings.WindowsInstaller_WriteLog_Delete) + " - " + $PatchChildToDelete) 
del $PatchChildToDelete -recurse -ErrorAction SilentlyContinue
remove-itemproperty $RegistryPatch "AllPatches"
New-ItemProperty -name "AllPatches" $RegistryPatch -value $MultiStringHolder -propertyType "MultiString"
}

function CheckCacheFiles
{
#Build list of patches
$PatchDirList=Get-ChildItem -path $Env:WinDir\installer -include *.msp -recurse
$PatchArray=@{}
if ($PatchDirList.length -gt 0)
{
    foreach ($PatchCache in $PatchDirList)
    {
        $PatchCompressedGUID=[CleanUpRegistry]::CompressGUID(([CollectPatchInformation]::GetPatchInfoFromPath($Env:WinDir+"\installer\"+$PatchCache.Name,"PatchGUID")))
        If (!$PatchArray.Contains($PatchCompressedGUID))
        {
            $PatchArray.add($PatchCompressedGUID,$PatchCache.Name) 
        }
    }
}
$PatchArray
}

function CheckCacheFilesMSI
{
#Build list of msi
$MSIDirList=Get-ChildItem -path $Env:WinDir\installer -include *.msi -recurse
$MSIArray=@{}
if ($MSIDirList.length -gt 0)
{
    foreach ($MSICache in $MSIDirList)
    {
        $MSIProductCode=[CollectMSIProductCode]::MSIGetProductCodeFromPackage($Env:WinDir+"\installer\"+$MSICache.name)
        If (!$MSIArray.Contains($MSIProductCode))
        {
            $MSIArray.add($MSIProductCode,$MSICache.Name) 
        }
    }
}
$MSIArray
}

function Compressed
{
$sourceDelete = @"
       using System;
       using System.Collections.Generic;
       using System.Collections;
       using System.Text;
       using Microsoft.Win32;
public class CleanUpRegistry
{
        public static string ReverseString(string szGUID)
        {
            char[] arr = szGUID.ToCharArray();
            Array.Reverse(arr);
            return new string(arr);
        }
       public static string CompressGUID(string szStandardGUID)
        {
            return (ReverseString(szStandardGUID.Substring(1, 8)) +
                ReverseString(szStandardGUID.Substring(10, 4)) +
                ReverseString(szStandardGUID.Substring(15, 4)) +
                ReverseString(szStandardGUID.Substring(20, 2)) +
                ReverseString(szStandardGUID.Substring(22, 2)) +
                ReverseString(szStandardGUID.Substring(25, 2)) +
                ReverseString(szStandardGUID.Substring(27, 2)) +
                ReverseString(szStandardGUID.Substring(29, 2)) +
                ReverseString(szStandardGUID.Substring(31, 2)) +
                ReverseString(szStandardGUID.Substring(33, 2)) +
                ReverseString(szStandardGUID.Substring(35, 2)));
        }
}
"@
           
   if ([CleanUpRegistry] -eq $null)
   {
   #Code here will never be run;  the if statement with throw an exception because the type doesn't exist
   }
   else
   {
   #type is loaded, do nothing
   }
Trap [Exception]
{
    #type doesn't exist.  Load it
    $type=Add-Type -TypeDefinition $sourceDelete -PassThru
continue
}
}

function GetRegistryType
{
$GetRegistryTypeCode=
@"
using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
public class Registry
{
        [DllImport("advapi32.dll", EntryPoint = "RegQueryValueExA")]
        public static extern int RegQueryValueEx(int hKey, string lpValueName, int lpReserved, out int lpType, StringBuilder lpData, ref int lpcbData);
        [DllImport("advapi32.dll", CharSet = CharSet.Auto)]
        public static extern int RegOpenKeyEx(Microsoft.Win32.RegistryHive hKey, string subKey, int ulOptions, int samDesired, out int phkResult);
        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern int RegCloseKey(int hKey);
  enum RegistryRights
        {
            ReadKey = 131097,
            WriteKey = 131078
        }
    static Microsoft.Win32.RegistryHive RegHive(string szRegHive)
        {
            Microsoft.Win32.RegistryHive rhRegHiveOut;
            
            switch (szRegHive)
            {
                case "HKLM":
                    rhRegHiveOut = Microsoft.Win32.RegistryHive.LocalMachine;
                    break;
                case "HKCR":
                    rhRegHiveOut = Microsoft.Win32.RegistryHive.ClassesRoot ;
                    break;
                case "HKCU":
                    rhRegHiveOut = Microsoft.Win32.RegistryHive.CurrentUser;
                    break;
                case "HKU":
                    rhRegHiveOut = Microsoft.Win32.RegistryHive.Users;
                    break;
                default:
                    rhRegHiveOut = Microsoft.Win32.RegistryHive .LocalMachine;
                    break;
            }
            return rhRegHiveOut;
        }
        static string RegValueType(int iRegType)
        {
            string szRegistryType = "REG_NONE";
            switch (iRegType)
            {
                case 0:
                    szRegistryType = "REG_NONE";
                    break;
                case 1:
                    szRegistryType = "String";
                    break;
                case 2:
                    szRegistryType = "ExpandString";
                    break;
                case 3:
                    szRegistryType = "Binary";//Binary
                    break;
                case 4:
                    szRegistryType = "DWord";
                    break;
                case 5:
                    szRegistryType = "DWord"; //"REG_DWORD_BIG_ENDIAN";
                    break;
                case 6:
                    szRegistryType = "REG_LINK";
                    break;
                case 7:
                    szRegistryType = "MultiString";//multistring
                    break;
            }
            
            return szRegistryType;
        }
  public static string GetType(string szHive,string szRegRoot, string szRegValue,int iWow)
  {
  
            //iwow 64bit =256 and 32 = 512
            int hKeyVal = 0, type = 0, hwndOpenKey=0;
            string szType = "NA", szValue="NA" ;
            int valueRet = RegOpenKeyEx(RegHive(szHive), @szRegRoot, 0, (int)RegistryRights.ReadKey | iWow, out hKeyVal);
            if (valueRet == 0)
            {
                int iBuffSize = 10,iError=0;
                bool blSize = false;
                while (!blSize)
                {
                    StringBuilder sb = new StringBuilder(iBuffSize);
                    hwndOpenKey = iBuffSize;
                    try
                    {
                        iError = RegQueryValueEx(hKeyVal, szRegValue, 0, out type, sb, ref hwndOpenKey);
                    }
                    catch { }
                    if (iError == 234)
                    {
                        iBuffSize = iBuffSize + 1;
                    }
                    else if (iError == 0)
                    {
                        blSize = true;
                        szValue = sb.ToString();
                        szType = RegValueType(type);
                    }
                    else
                    {
                        blSize = true;
                    }
                    sb.Remove(0, sb.Length);
                }
            }
            return szType;
            //RegCloseKey(hKeyVal);
  }
}
"@
              
   if ([Registry] -eq $null)
   {
   #Code here will never be run;  the if statement with throw an exception because the type doesn't exist
   }
   else
   {
   #type is loaded, do nothing
   }
Trap [Exception]
{
    #type doesn't exist.  Load it
    $type = Add-Type -TypeDefinition $GetRegistryTypeCode -PassThru
continue
}
}
function SystemRestore
{
$SystemRestoreCode=
@"
using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

public class Restore
    {
        internal const Int16 MaxDescW = 256;
        internal const Int16 BeginSystemChange = 100; // Start of operation 
         [DllImport("srclient.dll")]
         [return: MarshalAs(UnmanagedType.Bool)]
         internal static extern bool SRSetRestorePointW(ref RestorePointInfo pRestorePtSpec, out STATEMGRSTATUS pSMgrStatus);
       
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
           internal struct RestorePointInfo
           {
               public int dwEventType; // The type of event
               public int dwRestorePtType; // The type of restore point
               public Int64 llSequenceNumber; // The sequence number of the restore point
               [MarshalAs(UnmanagedType.ByValTStr, SizeConst = MaxDescW + 1)]
               public string szDescription; // The description to be displayed so the user can easily identify a restore point
          }
             [StructLayout(LayoutKind.Sequential)]
          internal struct STATEMGRSTATUS
           {
               public int nStatus; // The status code
               public Int64 llSequenceNumber; // The sequence number of the restore point
           }
          public static int StartRestore(string strDescription)
          {
              long lSeqNum = 0;
              RestorePointInfo rpInfo = new RestorePointInfo();
              STATEMGRSTATUS rpStatus = new STATEMGRSTATUS();
    
              try
              {
                  // Prepare Restore Point
                  rpInfo.dwEventType = BeginSystemChange;
                  // By default we create a verification system
                  rpInfo.dwRestorePtType = 1;
                  rpInfo.llSequenceNumber = 0;//llSequenceNumber must be 0 when creating a restore point.
                  rpInfo.szDescription = strDescription;
 
                bool blError=  SRSetRestorePointW(ref rpInfo, out rpStatus);
               
              }
              catch (DllNotFoundException)
              {
                  lSeqNum = 0;
                  return -1;
              }
  
              lSeqNum = rpStatus.llSequenceNumber;
  
              return rpStatus.nStatus;
          }
    }
"@
             
   if ([Restore] -eq $null)
   {
   #Code here will never be run;  the if statement with throw an exception because the type doesn't exist
   }
   else
   {
   #type is loaded, do nothing
   }
Trap [Exception]
{
    #type doesn't exist.  Load it
    $type = Add-Type -TypeDefinition $SystemRestoreCode -PassThru
continue
}
}
function UninstallProduct
{
$ProductUninstall=
@"
using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.Runtime.InteropServices;
public class UninstallProduct2
{
    public enum INSTALLUILEVEL
    {
        INSTALLUILEVEL_NOCHANGE = 0,    // UI level is unchanged
        INSTALLUILEVEL_DEFAULT = 1,    // default UI is used
        INSTALLUILEVEL_NONE = 2,    // completely silent installation
        INSTALLUILEVEL_BASIC = 3,    // simple progress and error handling
        INSTALLUILEVEL_REDUCED = 4,    // authored UI, wizard dialogs suppressed
        INSTALLUILEVEL_FULL = 5,    // authored UI with wizards, progress, errors
        INSTALLUILEVEL_ENDDIALOG = 0x80, // display success/failure dialog at end of install
        INSTALLUILEVEL_PROGRESSONLY = 0x40, // display only progress dialog
        INSTALLUILEVEL_HIDECANCEL = 0x20, // do not display the cancel button in basic UI
        INSTALLUILEVEL_SOURCERESONLY = 0x100, // force display of source resolution even if quiet
    }
        [DllImport("msi.dll")]
            public static extern int MsiConfigureProduct(
            string szProduct,            // product code
            int iInstallLevel,            // install level
            int eInstallState    // install state
            );
              [DllImport("msi.dll", SetLastError = true)]
       public static extern int MsiSetInternalUI(INSTALLUILEVEL dwUILevel, ref int phWnd);
              public static int LogandUninstallProduct(string szProductCode)
        {
                     int phWnd=0;
                     MsiSetInternalUI(INSTALLUILEVEL.INSTALLUILEVEL_NONE, ref phWnd);
            int iErrorCode = MsiConfigureProduct(szProductCode, 1, 2);
            return (iErrorCode);
        }
}
"@
             
   if ([UninstallProduct2] -eq $null)
   {
   #Code here will never be run;  the if statement with throw an exception because the type doesn't exist
   }
   else
   {
   #type is loaded, do nothing
   }
Trap [Exception]
{
    #type doesn't exist.  Load it
    $type = Add-Type -TypeDefinition $ProductUninstall -PassThru
continue
}
}

function ProductLPR
{
$ProductLPR=
@"
using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.Runtime.InteropServices;
public class ProductLPRMain
{
[DllImport("msi.dll")]
public static extern Int32 MsiEnumComponents(
                int iComponentIndex,
                StringBuilder lpComponentBuf
                );
[DllImport("msi.dll")]
public static extern int MsiGetComponentPath(
    string szProduct,   // product code for client product
    string szComponent, // component ID
    StringBuilder lpPathBuf,    // returned path
    ref int pcchBuf       // buffer character count
    );
[DllImport("msi.dll")]
public static extern int MsiEnumClients(
    string szComponent, // component code, string GUID
    int iProductIndex, // 0-based index into client products
    StringBuilder lpProductBuf  // buffer to receive GUID
    );
   
 public static ArrayList GetComponentList()
        {
            StringBuilder sb = new StringBuilder(40);
            ArrayList alComponents = new ArrayList();
            int iError = 0,iCounter=1;
            iError= MsiEnumComponents(0,sb);
            while (iError == 0)
            {
                iError = MsiEnumComponents(iCounter , sb);
                if (iError == 0)
                {
                    alComponents.Add(sb.ToString());
                    iCounter = iCounter + 1;
                }
            }
            return alComponents;
        }
        public static bool SharedComponents(string szComponent,string szProductCode)
        {
           
            bool blShared=false;
            StringBuilder sb = new StringBuilder(1024);
            int iError = 0,iCounter=0;
            while (blShared != true && iError !=259)
            {
                sb.Remove(0, sb.Length);
                iError = MsiEnumClients(szComponent, iCounter, sb);
                if (iError == 0  && (szProductCode !=sb.ToString ()))
                {
                    blShared = true;                    
                }
                iCounter = iCounter + 1;
            }
            return blShared;
        }
        public static ArrayList GetComponentPath(string szProductCode)
        {           
            int pathBuf = 1024,iError=0;
            StringBuilder sbPath = new StringBuilder(1024);
            ArrayList alItems = new ArrayList();
            ArrayList alComponents= new ArrayList();
            alComponents=GetComponentList();
            
            foreach (string szComponent in alComponents)
            {
               pathBuf = 1024;
               iError=  MsiGetComponentPath(szProductCode, szComponent, sbPath, ref pathBuf);
               string szCompPath = sbPath.ToString().ToLower ();
                if (szCompPath=="")szCompPath ="123NA123";
               
              //if ((iError != -1 && szCompPath != "123NA123") && (szCompPath.IndexOf("\\", szCompPath.Length - 1, 1)==-1) && !(szCompPath.Contains(Environment.GetEnvironmentVariable("WinDir").ToLower ())))
      if ((iError != -1 && szCompPath != "123NA123") && (szCompPath.IndexOf("\\", szCompPath.Length - 1, 1)==-1))
               {
                  if ( !(SharedComponents (szComponent,szProductCode)))
                  {
      alItems.Add(sbPath.ToString());
                  }                    
               }  
     }
            return alItems;
        }
    }
"@
             
   if ([ProductLPRMain] -eq $null)
   {
   #Code here will never be run;  the if statement with throw an exception because the type doesn't exist
   }
   else
   {
   #type is loaded, do nothing
   }
Trap [Exception]
{
    #type doesn't exist.  Load it
    $type = Add-Type -TypeDefinition $ProductLPR -PassThru
continue
}
}

Function BackupFiles
{
Param ($Items,$ProductCode)

$root=$Env:SystemDrive
$DirectoryPath= $root+"\MATS\$ProductCode"
[void](New-PSDrive -Name HKCR -PSProvider registry -root HKEY_CLASSES_ROOT ) # Allows access to registry HKEY_CLASSES_ROOT
if (!(Test-Path -path $DirectoryPath)) 
 { 
 [void]( New-Item  "$DirectoryPath\FileBackup" -type directory -force )#Create the backup folder for the files we will delete
}

foreach ($Item in $Items)
{
[void]$Item
    if (($Item -ne $Null) -and ($Item.length -gt 2))
    {
        if(!($Item.substring(0,1)-match '^\d+$')) #not a registry
        {
        $DuplicateDirector="$DirectoryPath\FileBackup\"+$item.Replace(":","")
              if (!(Test-Path -path $DuplicateDirector.substring(0,($DuplicateDirector.LastIndexOf("\"))))) 
              { 
                    [void](New-Item $DuplicateDirector.substring(0,($DuplicateDirector.LastIndexOf("\"))) -type directory) #Create the backup directories ready to copy the files into then
              }
        
              If(Test-Path $Item)
              {
        $iError=Copy-Item $Item -Destination ($DuplicateDirector) -ErrorAction SilentlyContinue #copy files to the backup folder
        WriteXMLFiles ([System.Diagnostics.FileVersionInfo]::GetVersionInfo($Item).InternalName) ([System.Diagnostics.FileVersionInfo]::GetVersionInfo($Item).FileVersion) $Item  $DuplicateDirector
              }
        }
        else
        {
            #We are dealing with the registry   
            BackUpRegistry $Item $False
        }
    }
}
}
Function BackUpRegistry
{
Param($Item,$DeleteParent)

switch ($Item.substring(0,2)) 
{ 
    "-1"{
        $QueryAssignmenttype=[MakeString]::GetProductInfo($ProductCode,"AssignmentType")
        switch($QueryAssignmenttype)
        {
             "0"{ $Item=$Item.Replace("21:","HKCU:")}
       "1"{ $Item=$Item.Replace("02:","HKLM:")}
       default{ $Item=$Item.Replace("02:","HKLM:")}
        }
        }   
   "00"{ $Item=$Item.Replace("00:","HKCR:")} 
   "01"{ $Item=$Item.Replace("01:","HKCU:")}
   "02"{ $Item=$Item.Replace("02:","HKLM:")} 
   "03"{ $Item=$Item.Replace("03:","HKU:")}
   "20"{ $Item=$Item.Replace("20:","HKCR:")} 
   "21"{ $Item=$Item.Replace("21:","HKCU:")}
   "22"{ $Item=$Item.Replace("22:","HKLM:")} 
   "23"{ $Item=$Item.Replace("23:","HKU:")}
}     
       
$RegRoot = $Item.substring(0,$Item.LastIndexOf("\")) #Get Registry root value
$RegValue =$Item.substring($Item.LastIndexOf("\")+1,$Item.length-$Item.LastIndexof("\")-1) #Get registry value 
$Exista=$false
$Existsa=CheckRegistryValueExists -RegRoot $Regroot -RegValue $RegValue #Does this key still exist in the registry? MSI thinks it does
if (!$Existsa) #Maybe a path on the reg info so search value for known good
{
$InLoop=$true
$iCounter=0
$Slash =$Item.IndexOf("\",$iCounter)+1
while($InLoop)
{
$TempRoot=$Item.substring(0,$Slash)
$TempValue=$Item.substring($Slash,($Item.length-$Slash))
    $Existsa=CheckRegistryValueExists  -RegRoot $TempRoot  -RegValue $TempValue
    if (!$Existsa)
    {
        $iCounter=$Slash
        $Slash =$Item.IndexOf("\",$iCounter)+1
     If ($Slash -eq 0 -or $TempRoot -eq "") #$Slash =-1 before DEBUG
        {
     $InLoop=$False
        }       
     }
  else
     {
  #FOUND IT return here
  $InLoop=$False
  $RegRoot=$TempRoot
  $RegValue=$TempValue
  $RegistryType=[Registry]::GetType($item.substring(0,($item.IndexOf(":"))),$RegRoot.substring($RegRoot.indexOf(":")+2),$RegValue,256)
   }
}
}

if ($Existsa) #If we have found the registry key lets continue
{
$RegistryType="NA"
$RegistryType= [Registry]::GetType($item.substring(0,($item.IndexOf(":"))),$RegRoot.substring($RegRoot.indexOf(":")+2),$RegValue,256)  #Wow32 to false for registry redirection
If ($RegistryType -ne "NA")
{
    WriteXMLRegistry $RegistryType $RegRoot $RegValue $DeleteParent
}
$RegistryType="NA"
$regroot=$regroot.tolower()
$regroot=$regroot.replace("software\","software\wow6432node\")
$regroot=$regroot.replace("HKCR:","HKCR:\WOW6432node\")
$regroot=$regroot.replace("hkcr:","HKCR:\WOW6432node\")
$RegistryType= [Registry]::GetType($item.substring(0,($item.IndexOf(":"))),$RegRoot.substring($RegRoot.indexOf(":")+2),$RegValue,256) #Wow32 to false for registry redirection
If ($RegistryType -ne "NA")
{
    WriteXMLRegistry $RegistryType $RegRoot $RegValue $DeleteParent
}
}
$RegistryCountFailures
}
Function WriteXMLFiles
{
PARAM ($FileNamePath,$FileVersion,$FileBackupLocation, $FileDestination)

$xml = New-Object xml
if (!(Test-Path -path "$DirectoryPath\FileBackupTemplate.xml")) 
{ 
 $root = $xml.CreateElement("FileBackup")
[void]$xml.AppendChild($root)
$Record = $xml.CreateElement("File")
#$Record.PSBase.InnerText = $RegRoot
[void]$root.AppendChild($Record)
$RecordName = $xml.CreateElement("FileName")
    $RecordName.PSBase.InnerText = ($FileNamePath)
[void]$Record.AppendChild($RecordName)
$RecordVersion= $xml.CreateElement("FileVersion")
$RecordVersion.PSBase.InnerText = ($FileVersion)
[void]$Record.AppendChild($RecordVersion)
$RecordType = $xml.CreateElement("FileBackupLocation")
$RecordType.PSBase.InnerText = $FileBackupLocation
[void]$Record.AppendChild($RecordType)
    $RecordBackup = $xml.CreateElement("FileDestination")
$RecordBackup.PSBase.InnerText = $FileDestination
[void]$Record.AppendChild($RecordBackup)
[void]$xml.Save("$DirectoryPath\FileBackupTemplate.xml")
}
else
{
$xml =[xml] (get-content "$DirectoryPath\FileBackupTemplate.xml")
    $root = $xml.FileBackup
$Record = $xml.CreateElement("File")
[void]$root.AppendChild($Record)
$RecordName = $xml.CreateElement("FileName")
    $RecordName.PSBase.InnerText = ($FileNamePath)
[void]$Record.AppendChild($RecordName)
$RecordVersion= $xml.CreateElement("FileVersion")
$RecordVersion.PSBase.InnerText = ($FileVersion)
[void]$Record.AppendChild($RecordVersion)
$RecordType = $xml.CreateElement("FileBackupLocation")
$RecordType.PSBase.InnerText = $FileBackupLocation
[void]$Record.AppendChild($RecordType)
    $RecordBackup = $xml.CreateElement("FileDestination")
$RecordBackup.PSBase.InnerText = $FileDestination
[void]$Record.AppendChild($RecordBackup)
    [void]$xml.Save("$DirectoryPath\FileBackupTemplate.xml")
}       
}
Function WriteXMLRegistry
{
PARAM ($RegistryType,$RegRoot,$RegValue, $DeleteParent)
$root=$Env:SystemDrive
$DirectoryPath= $root+"\MATS\$ProductCode"
if (!($RegContent=Get-ItemProperty $RegRoot $RegValue))
{
#If the registry key is missing or has issues fail out
    return $True
}
    
if ($Regcontent -eq "")
{
$Failure=$true
}
$OutcomeContent=""
[string]$RegName=""
    
switch ($RegistryType)
{
"String"
    {
  $RegName=$RegContent.$RegValue
    }
        
    "MultiString"
    {
  foreach ($Value in $RegContent.$RegValue)
        {
   $RegName=$RegName+$value+";"
        }
  $RegName=$RegName.substring(0,$RegName.length-1)
     }
     
  "Binary"
     {
  foreach ($Value in $RegContent.$RegValue)
        {
   $RegName=$RegName+$value+","
        }
        $RegName=$RegName.substring(0,$RegName.length-1)
     }
     
  default
     {
  Foreach ($Value in $RegContent.$RegValue)
        {
   $Outcome= "{0:x}" -f $Value
            $OutcomeContent=$OutcomeContent+$Outcome
        }
        $RegContent.$RegValue=$OutcomeContent
        $RegName=$RegContent.$RegValue
     }
}
    
$xml = New-Object xml
If (!(Test-Path -path $DirectoryPath\registryBackupTemplate.xml)) 
{ 
 $root = $xml.CreateElement("RegistryBackup")
[void]$xml.AppendChild($root)
$Record = $xml.CreateElement("Registry")
#$Record.PSBase.InnerText = $RegRoot
[void]$root.AppendChild($Record)
$RegistryHive= $xml.CreateElement("RegistryHive")
$RegistryHive.PSBase.InnerText = ($RegRoot)
[void]$Record.AppendChild($RegistryHive)
    $XMLDeleteParent = $xml.CreateElement("RegistryDeleteParent")
$XMLDeleteParent.PSBase.InnerText = [string]$DeleteParent
    [void]$Record.AppendChild($XMLDeleteParent)
$Name = $xml.CreateElement("RegistryType")
$Name.PSBase.InnerText = $RegistryType
[void]$Record.AppendChild($name)
$Value = $xml.CreateElement("RegistryName")
$Value.PSBase.InnerText = ($RegValue)
[void]$Record.AppendChild($Value)
$Date = $xml.CreateElement("RegistryValue")
$Date.PSBase.InnerText = ($RegName)
[void]$Record.AppendChild($date)
    if (!(Test-Path -path $DirectoryPath)) 
    { 
  [void]( New-Item  "$DirectoryPath" -type directory -force )#Create the backup folder for the files we will delete
    }

[void]$xml.Save("$DirectoryPath\registryBackupTemplate.xml")
}
else
{
$xml =[xml] (get-content "$DirectoryPath\registryBackupTemplate.xml")
$root = $xml.RegistryBackup
$Record = $xml.CreateElement("Registry")
[void]$root.AppendChild($Record)
$RegistryHive= $xml.CreateElement("RegistryHive")
$RegistryHive.PSBase.InnerText = ($RegRoot)
[void]$Record.AppendChild($RegistryHive)
    $XMLDeleteParent = $xml.CreateElement("RegistryDeleteParent")
$XMLDeleteParent.PSBase.InnerText = [string]$DeleteParent
    [void]$Record.AppendChild($XMLDeleteParent)
$Name = $xml.CreateElement("RegistryType")
$Name.PSBase.InnerText = $RegistryType
[void]$Record.AppendChild($name)
$Value = $xml.CreateElement("RegistryName")
$Value.PSBase.InnerText = ($RegValue)
[void]$Record.AppendChild($Value)
$Date = $xml.CreateElement("RegistryValue")
$Date.PSBase.InnerText = ($RegName)
[void]$Record.AppendChild($date)
[void]$xml.Save("$DirectoryPath\registryBackupTemplate.xml")
}       
}
Function CheckRegistryValueExists([string]$RegRoot,[string]$RegValue)
{
Get-ItemProperty -path $RegRoot -name $RegValue -ErrorAction SilentlyContinue | Out-Null  
if (!$?)
{
#check64bit
$RegRoot=$RegRoot.tolower()
$RegValue=$RegValue.tolower()
$Regroot=$Regroot.Replace("software\","software\wow6432node\")
$Regroot=$Regroot.Replace("hkcr:\","hkcr:\wow6432node\")
Get-ItemProperty $RegRoot $RegValue -ErrorAction SilentlyContinue | Out-Null  
}
$?
}
Function  EnableComputerRestore
{
Param ($DriveLetter)

enable-computerrestore -drive $DriveLetter
Trap [Exception]
{
  # Legacy restore for PowerShell v1.0
  LegacyEnableComputerRestore $DriveLetter
  continue
}
}

Function LegacyEnableComputerRestore
{
Param ($DriveLetter)
       $SysRestore = [wmiclass]"\\.\root\default:systemrestore"
    $InParams = $SysRestore.psbase.GetMethodParameters("Enable")
    $InParams.Drive = $DriveLetter + "\"
    $InParams.WaitTillEnabled = $True
    $SysRestore.PSBase.InvokeMethod("Enable", $InParams, $Null)
       Trap [Exception]
       {
          continue
       }
}

Function CheckRestorePoint
{
Param($Items,$ProductCode)
if($Items -ne $null)
{
    $DriveList=new-Object System.Collections.ArrayList
    $DRiveList.add($Env:SystemDrive +"\")
    foreach ($Item in $Items)
    {
       
        
        if($Item -ne $Null)
        {
            $Item=$Item.toupper() 
            if(!($Item.substring(0,1)-match '^\d+$')) #not a registry
            {
                if ($DriveList.contains($Item.substring(0,3)) -eq $False)
                {
                    $DriveList.Add($Item.Substring(0,3))
                }
            }
        }
        
  Trap [Exception]
  {
   $error=""
   continue
  }
    }
    
    if ($DriveList)
    {
  Trap [Exception]
  {
     continue
  }
        EnableComputerRestore $DriveList
    }
}
}


Function DeleteRegistryKeysFromXML
{
Param($ProductCode)
$root=$Env:SystemDrive
$DirectoryPath= $root+"\MATS\$ProductCode"
New-PSDrive -Name HKCR -PSProvider registry -root HKEY_CLASSES_ROOT # Allows access to registry HKEY_CLASSES_ROOT

   if ((Test-Path -path $DirectoryPath\registryBackupTemplate.xml)) 
   { 
       $xml =[xml] (get-content $DirectoryPath\registryBackupTemplate.xml)
       #$root =$xml.RegistryBackup.Registry
       foreach ($XMLItems in $xml.RegistryBackup.Registry)
       {   if ((($XMLItems.RegistryDeleteParent).toupper()) -eq "TRUE")
           {
                del $XMLItems.RegistryHive -recurse -ErrorAction SilentlyContinue
           }
           else
            {
                remove-itemproperty -path $XMLItems.RegistryHive -name $XMLItems.RegistryName -ErrorAction SilentlyContinue
            }
       }
   }
   else
   {
   #Failed to find registry backup so exit
   Return $False
   }
}
Function DeleteFilesFromXML
{
Param($ProductCode)
$root=$Env:SystemDrive
$DirectoryPath= $root+"\MATS\$ProductCode"
if ((Test-Path -path $DirectoryPath\FileBackupTemplate.xml)) 
{ 
    $xml =[xml] (get-content $DirectoryPath\FileBackupTemplate.xml)
    $root =$xml.FileBackup.File
    foreach ($XMLItems in $xml.FileBackup.File)
    {    
        remove-item $XMLItems.FileBackupLocation -ErrorAction SilentlyContinue
    }
}   
}

Function CreateRegistryFileRecoveryFile
{
Param($ProductCode)
$root=$Env:SystemDrive
$DirectoryPath= $root+"\MATS\$ProductCode"
if (!(Test-Path -path $DirectoryPath)) 
{ 
    [void]( New-Item  "$DirectoryPath" -type directory -force -ErrorAction SilentlyContinue) #Create the backup folder
}
$WriteFile=
@"
Function RestoreFilesFromXML
{
    [void](New-PSDrive -Name HKCR -PSProvider registry -root HKEY_CLASSES_ROOT) # Allows access to registry HKEY_CLASSES_ROOT
    if ((Test-Path -path "$DirectoryPath\\FileBackupTemplate.xml")) 
    { 
        `$xml =[xml] (get-content "$DirectoryPath\\FileBackupTemplate.xml")
        `$root =`$xml.FileBackup.File
        foreach (`$XMLItems in `$xml.FileBackup.File)
        {    
            if (!(Test-Path `$XMLItems.FileBackupLocation))
            {
                if (!(Test-Path -path `$XMLItems.FileBackupLocation.substring(0,(`$XMLItems.FileBackupLocation.LastIndexOf("\"))))) 
                    { 
                         [void](New-Item `$XMLItems.FileBackupLocation.substring(0,(`$XMLItems.FileBackupLocation.LastIndexOf("\"))) -type directory) #Create the backup directories ready to copy the files into then
                    }
                Copy-Item `$XMLItems.FileDestination -Destination (`$XMLItems.FileBackupLocation) -ErrorAction SilentlyContinue #copy files to the backup folder
            }
        }
    }
}

Function RestoreRegistryFiles
{
    [void](New-PSDrive -Name HKCR -PSProvider registry -root HKEY_CLASSES_ROOT) # Allows access to registry HKEY_CLASSES_ROOT
    if ((Test-Path -path "$DirectoryPath\registryBackupTemplate.xml")) 
    { 
        `$xml =[xml] (get-content "$DirectoryPath\registryBackupTemplate.xml")
        `$root =`$xml.RegistryBackup.Registry
        foreach (`$XMLItems in `$xml.RegistryBackup.Registry)
        {    
            if (!(CheckRegistryValueExists -RegRoot `$XMLItems.RegistryHive -RegValue `$XMLItems.RegistryName))
            {
            if (`$XMLItems.RegistryType -eq "MultiString")
            {
            [string[]]`$MultiString=`$XMLItems.RegistryValue.Split(";")
            new-item -path `$XMLItems.RegistryHive -ErrorAction SilentlyContinue | Out-Null 
            New-ItemProperty `$XMLItems.RegistryHive `$XMLItems.RegistryName -value `$MultiString -propertyType `$XMLItems.RegistryType
            }
            elseif(`$XMLItems.RegistryType -eq "Binary")
            {
            new-item -path `$XMLItems.RegistryHive                   
            [Byte[]]`$ByteArray = @()
            foreach(`$byte in `$XMLItems.RegistryValue.Split(",")){`$ByteArray += `$byte}
            New-ItemProperty `$XMLItems.RegistryHive `$XMLItems.RegistryName -value `$ByteArray -propertyType `$XMLItems.RegistryType
            }
            elseif(`$XMLItems.RegistryType -eq "Dword")
            {
            `$DwordValue =Hex2Dec `$XMLItems.RegistryValue
            new-item -path `$XMLItems.RegistryHive -ErrorAction SilentlyContinue | Out-Null  
            New-ItemProperty `$XMLItems.RegistryHive `$XMLItems.RegistryName -value `$DwordValue  -propertyType `$XMLItems.RegistryType
            }
            else
            {
            ##Main Registry key missing so create here
            new-item -path `$XMLItems.RegistryHive -ErrorAction SilentlyContinue | Out-Null  
            New-ItemProperty `$XMLItems.RegistryHive `$XMLItems.RegistryName -value `$XMLItems.RegistryValue -propertyType `$XMLItems.RegistryType
            }
            }
         }
    }
}
Function CheckRegistryValueExists
{
Param(`$RegRoot,`$RegValue)
Get-ItemProperty `$RegRoot `$RegValue -ErrorAction SilentlyContinue | Out-Null  
`$?
}
function Hex2Dec
{
param(`$HEX)
ForEach (`$value in `$HEX)
{
    [Convert]::ToInt32(`$value,16)
}
}

function Test-Administrator
{
`$user = [Security.Principal.WindowsIdentity]::GetCurrent() 
(New-Object Security.Principal.WindowsPrincipal `$user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}
`$UsersTemp= `$Env:temp
Copy-Item  "$DirectoryPath\\RestoreYourFilesAndRegistry.ps1" -Destination (`$UsersTemp+"\RestoreYourFilesAndRegistry.ps1") -ErrorAction SilentlyContinue
If(Test-Administrator)
{
RestoreRegistryFiles
RestoreFilesFromXML
}
else
{
cls
`$ElevatedProcess = new-object System.Diagnostics.ProcessStartInfo "PowerShell"    
`$ElevatedProcess.Arguments =  (`$UsersTemp+"\RestoreYourFilesAndRegistry.ps1")
`$ElevatedProcess.Verb = "runas"     
[System.Diagnostics.Process]::Start(`$ElevatedProcess)     
exit
}
"@
$writefile | Out-File $DirectoryPath\RestoreYourFilesAndRegistry.ps1 -ErrorAction SilentlyContinue | Out-Null  
 
}
Function RapidProductRegistryRemove
{
Param ($ProductCode)
CreateRegistryFileRecoveryFile($ProductCode)
New-PSDrive -Name HKCR -PSProvider registry -root HKEY_CLASSES_ROOT # Allows access to registry HKEY_CLASSES_ROOT
$CompressedGUID=[CleanupRegistry]::CompressGUID([string]$ProductCode)
if ($CompressedGUID)
{
    ##################HKLM
    $RPRHives=Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Userdata\ #IF need to delete component then add back recurse HKL,HKCR code here
    foreach($SubKey in $RPRHives) 
    { 
       $RegistryKey=RegistryHiveReplace $SubKey.name
  
       if (Test-Path "$RegistryKey\Products\$CompressedGUID")
       {      
       New-ItemProperty "$RegistryKey\Products\$CompressedGUID" -name "RPRTAG" -value "Backup" -propertyType "String" -ErrorAction SilentlyContinue | Out-Null  
       BackupRegistry ("$RegistryKey\Products\$CompressedGUID"+"\RPRTAG") $False
       $RegistryBackup=dir "$RegistryKey\Products\$CompressedGUID" -recurse
       if ($RegistryBackup)
       {
          foreach ($a in $RegistryBackup)
          {
              foreach($Property in $a.Property)
                 {
                     $a=RegistryHiveReplace $a
                      BackupRegistry ($a+"\"+$Property) $False
                 }
           }
        }
   
       del "$RegistryKey\Products\$CompressedGUID" -recurse #Special delete for parent
  
       }
     }

     ############HKCR
     if (Test-Path "HKCR:\Installer\Products\$CompressedGUID" )
     {
      $RegistryBackup=dir "HKCR:\Installer\Products\$CompressedGUID" -recurse
       if ($RegistryBackup)
       {
            $ParentValues= Get-ItemProperty -Path "HKCR:\Installer\Products\$CompressedGUID" 
            #[string]$rr=get-itemproperty $Parentvalues
            [string]$temp=[string]$ParentValues
            $temp=$temp.Replace("@{","")
            $temp=$temp.Replace("}","")
            foreach ($ParentItem in $temp.Split(";"))
            {
                #FN to backup the root elements
                $BackupSplit= $ParentItem.substring(0,$ParentItem.indexof("=")).trim()
                BackupRegistry("HKCR:\Installer\Products\$CompressedGUID\$BackupSplit") $False
            }
    
            foreach ($a in $RegistryBackup)
            {
                if ($a.Property -ne "")
                {
                    foreach($Property in $a.Property)
                    {
                        $a=RegistryHiveReplace $a
                        BackupRegistry($a+"\"+$Property) $False
                    }
                }
                else
                {
                   New-ItemProperty (RegistryHiveReplace $a.name) -name "RPRTAG" -value "Backup" -propertyType "String" -ErrorAction SilentlyContinue | Out-Null 
                   BackupRegistry((RegistryHiveReplace $a.name) +"\RPRTAG") $False
                }
            }
            del "HKCR:\Installer\Products\$CompressedGUID" -recurse #Special delete for parent
        }
     }


     ############HKCU
     if (Test-Path "HKCU:\Software\Microsoft\Installer\Products\$CompressedGUID" )
     {
      $RegistryBackup=dir "HKCU:\Software\Microsoft\Installer\Products\$CompressedGUID" -recurse
       if ($RegistryBackup)
       {
            $ParentValues= Get-ItemProperty -Path "HKCU:\Software\Microsoft\Installer\Products\$CompressedGUID"
            [string]$SplitString=[string]$ParentValues
            $SplitString=$SplitString.Replace("@{","")
            $SplitString=$SplitString.Replace("}","")
            foreach ($ParentItem in $SplitString.Split(";"))
            {
                #FN to backup the root elements
                $BackupSplit= $ParentItem.substring(0,$ParentItem.indexof("=")).trim()
                BackupRegistry("HKCU:\Software\Microsoft\Installer\Products\$CompressedGUID\$BackupSplit") $False
            }
    
            foreach ($a in $RegistryBackup)
            {
                if ($a.Property -ne "")
                {
                    foreach($Property in $a.Property)
                    {
                        $a=RegistryHiveReplace $a
                        BackupRegistry($a+"\"+$Property) $False
                    }
                }
                else
                {
                   New-ItemProperty (RegistryHiveReplace $a.name) -name "RPRTAG" -value "Backup" -propertyType "String" -ErrorAction SilentlyContinue | Out-Null 
                   BackupRegistry((RegistryHiveReplace $a.name) +"\RPRTAG") $False
                }
            }
            del "HKCU:\Software\Microsoft\Installer\Products\$CompressedGUID" -recurse #Special delete for parent
        }
     }
}

}

Function SortArray
{
$SortArray = @()
[MakeStringTest]::loop()| %{$SortArray += $_.Value} #foreach
$SortArray
}
Function RegistryHiveReplace
{
Param ([string]$RegistryKey)
$RegistryKey=$RegistryKey.toupper()
$RegistryKey=$RegistryKey -Replace("HKEY_LOCAL_MACHINE","HKLM:") 
$RegistryKey=$RegistryKey -Replace("HKEY_CLASSES_ROOT","HKCR:")     
$RegistryKey=$RegistryKey -Replace("HKEY_CURRENT_USER","HKCU:")  
$RegistryKey=$RegistryKey -Replace("HKEY_USERS","HKU:")  
$RegistryKey
}
Function RegistryHiveReplaceShim
{
Param ([string]$RegistryKey)
$RegistryKey=$RegistryKey.toupper()
$RegistryKey=$RegistryKey.Replace("HKEY_CLASSES_ROOT","00:")
$RegistryKey=$RegistryKey.Replace("HKEY_CURRENT_USER","01:")
$RegistryKey=$RegistryKey.Replace("HKEY_LOCAL_MACHINE","02:")
$RegistryKey=$RegistryKey.Replace("HKEY_USERS","03:")
$RegistryKey=$RegistryKey.Replace("HKCR:","00:")
$RegistryKey=$RegistryKey.Replace("HKCU:","01:")
$RegistryKey=$RegistryKey.Replace("HKLM:","02:")
$RegistryKey=$RegistryKey.Replace("HKU:","03:")

$RegistryKey  
}
Function MATSFingerPrint
{
Param ($ProductCode,$Value,$Type,$DateTimeRun)
if (!(test-path "HKLM:\Software\Microsoft\MATS\WindowsInstaller\$ProductCode\$DateTimeRun"))
{
   New-Item -Path hklm:\software\Microsoft\MATS\ -ErrorAction SilentlyContinue
   New-Item -Path hklm:\software\Microsoft\MATS\WindowsInstaller -ErrorAction SilentlyContinue
   New-Item -Path hklm:\software\Microsoft\MATS\WindowsInstaller\$ProductCode -ErrorAction SilentlyContinue
   New-Item -Path hklm:\software\Microsoft\MATS\WindowsInstaller\$ProductCode\$DateTimeRun -ErrorAction SilentlyContinue
   
}
New-ItemProperty hklm:\software\Microsoft\MATS\WindowsInstaller\$ProductCode\$DateTimeRun -name $Value -value $Type -propertyType String
}
Function StopService
{
param ($ServiceName)
if ($ServiceName -ne $null)
{
    stop-service $ServiceName -force -ErrorAction SilentlyContinue
}
}

Function ShimXML
{
param ($ProductCode)
$Win32_OS=Get-WmiObject Win32_OperatingSystem | select BuildNumber,OSLanguage,version
$Build=$Win32_OS.Version
$Language=$Win32_OS.OSLanguage
# Valid directory/file tokens are the following:
# %5% - Root drive
# %10% - Windows directory; example – c:\windows
# %11% - system directory; example – c:\windows\system32
# %16422% - Program Files directory – c:\Program Files
# %16427% - Common Files directory – c:\Program Files\Common Files
# %16527% - USERPROFILE C:\Documents and Settings\someone
# %16627% - APPDATA     C:\Documents and Settings\kmyer\Application Data
$ShimItems= New-Object System.Collections.ArrayList($null)
if ((Test-Path -path .\Shim.xml)) 
{ 
    $xml =[xml] (get-content ".\shim.xml")
    $XMLProductCode=$xml.shim.Product |Where-Object {$_.ProductCode -match $ProductCode} 
    if ($XMLProductCode -ne $null)
    {
        $XMLBuild=$XMLProductCode.os |Where-Object {($_.Build -match $Build) -or ($_.Build -match "ALL")}
        if ($XMLBuild-ne $null)
        {
   Foreach ($Build in $XMLBuild)
        {        
        
  $XMLLanguage=$Build.Language |Where-Object {($_.Code -match $Language) -or ($_.Code -match "ALL") }
        if ($XMLLanguage-ne $null)
        { 
   Foreach ($Language in $XMLLanguage)
            { 
    foreach ($XMLRegistryKey in $Language.Registry) #Registry shims
    {  
     if ($XMLRegistryKey-ne $null)
     {
      $RegistryContainer=$XMLRegistryKey.RegistryHive
      if ($RegistryContainer.length  -gt 0)
      {
       foreach ($Action in $RegistryContainer) #Registry shims
       { 
        switch ($Action.Action)
        {
         "*"
         {
          $EnumShim=ShimRegistryCollection $Action.registryvalue
          if ($EnumShim.length -gt 0)
          {
           $ShimItems=$ShimItems+$EnumShim
          }
          break;
         }
         "-"
         {
          $ShimItems+=((RegistryHiveReplaceShim($Action.registryvalue))) #Registry value Additions
          break;
         }
        }
      }
     }
                }
            }
            foreach ($XMLFileKey in $Language.File) #File Shims
            {   
                if (($XMLFileKey.FileName).length -gt 0)
                {       
     $FilePathType=$XMLFileKey.FileName.substring($XMLFileKey.FileName.indexof("%"),$XMLFileKey.FileName.lastindexof("%")+1)
                
                    switch ($FilePathType)
                    {
                        "%5%"
                        {
       $FilePath=$XMLFileKey.FileName
       $root=$Env:SystemDrive
       $FilePath=$FilePath.replace("%5%",$root)
       break;
                        }
                        "%10%"
                        {
       $FilePath=$XMLFileKey.FileName
       $root=$Env:WinDir
       $FilePath=$FilePath.replace("%10%",$root)
       break;
                        }
                        "%16627%"
                        {
       $FilePath=$XMLFileKey.FileName
       $root=$Env:AppData
       $FilePath=$FilePath.replace("%16627%",$root)
       break;
                        }
                    }
                    $ShimItems+=($FilePath) #File Additions
                }                
            }
            
            foreach ($XMLService in $Language.Services) #Services
            {   
                if (($XMLService.Service).length -gt 0)
                {
                    StopService $XMLService.Service #Stop Services
                }
            }
        }
    }
}
}
}
}
$ShimItems
}
Function ARPEntries
{
param($ProductCode,$ProductName)
$Itemstoadd= @()
$ARPUninstall="HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
if($ProductName.length -gt 0)
{
    $ARPRegistryKeyName=$ARPUninstall+$ProductName
if (!(Test-Path -path $ARPRegistryKeyName))
{
        $ReturnedARP=(ARPFindDisplayName $ARPUninstall $ProductName)
        if ($ReturnedARP)
        {
            $Itemstoadd+=($ReturnedARP)
        }
}
    else
    {
        $Itemstoadd+=(ShimRegistryCollection($ARPRegistryKeyName))
    }
    
    if (Test-Path -path ($ARPUninstall+$ProductCode))
   {
        $Itemstoadd+=(ShimRegistryCollection($ARPUninstall+$ProductCode))
    }
}
$Itemstoadd
}
Function ARPFindDisplayName
{
param ($ARPUninstall, $ProductName)
#Enum for DisplayName
$ARPUninstallKey = dir $ARPUninstall -recurse
       
if ($ARPUninstallKey)
{
    foreach ($ARPEntry in $ARPUninstallKey)
    {
        foreach ($Property in $ARPEntry.property)
        {
            if ($Property -eq "DisplayName")
            {
                $ARPDisplayName =(Get-ItemProperty -Path (RegistryHiveReplace($ARPEntry).Name))
                if ($ARPDisplayName.DisplayName -eq $ProductName)
                {
                    return ShimRegistryCollection($ARPEntry.Name)
                }
                 
            }
        }
    }
}
$false
}
Function ShimLoop
{
Param($RegistryRoot)
$Itemstoadd= @()
if((Test-Path $RegistryRoot))
    {
    $RegistryRoot=RegistryHiveReplace $RegistryRoot
    [string]$ParentValues= get-itemproperty -path $RegistryRoot
    [string]$SplitString=[string]$ParentValues
    $SplitString=$SplitString.Replace("@{","")
    $SplitString=$SplitString.Replace("}","")
    
    foreach ($ParentItem in $SplitString.Split(";"))
    {
        if ($ParentItem.length -gt 0)
        {
            $Itemstoadd+=((RegistryHiveReplaceShim($RegistryRoot))+"\"+($ParentItem.substring(0,$ParentItem.indexof("=")).trim()))
        }
    }
                
    $RegistryPatchList = dir $RegistryRoot -recurse
    if ($RegistryPatchList)
    {
        foreach ($Patch in $RegistryPatchList)
        {
            foreach ($Property in $Patch.property)
            {
               if ($Property.length  -gt 0)
               {
                    $Itemstoadd+=((RegistryHiveReplaceShim($Patch.name))+"\"+ $Property)
               }
            }
        }
     }
    
        New-ItemProperty "$RegistryRoot" -name "RPRTAG" -value "Backup" -propertyType "String" -ErrorAction SilentlyContinue | Out-Null  
        BackupRegistry("$RegistryRoot"+"\RPRTAG") $True
   
  }
$Itemstoadd
}
Function ShimRegistryCollection
{
param($Registrykey)
New-PSDrive -Name HKCR -PSProvider registry -root HKEY_CLASSES_ROOT |out-null # Allows access to registry HKEY_CLASSES_ROOT
$ShimArray=@()
foreach ($RegistryRoot in $RegistryKey)
{
$RegistryRoot=RegistryHiveReplace($RegistryRoot)
  If (ShimLoop ($RegistryRoot))
  {
     $ShimArray= $ShimArray+(ShimLoop ($RegistryRoot))
  }
   #IS WOW ALSO
   $RegistryRoot=$RegistryRoot.tolower()
   $RegistryRoot=$RegistryRoot.replace("software\","software\wow6432node\")
   $RegistryRoot=$RegistryRoot.replace("HKCR:","HKCR:\WOW6432node\")
   $RegistryRoot=$RegistryRoot.replace("hkcr:","HKCR:\WOW6432node\")
  
  $RegistryRoot=RegistryHiveReplace($RegistryRoot)
  If (ShimLoop ($RegistryRoot))
  {
     $ShimArray= $ShimArray+(ShimLoop ($RegistryRoot))
  }
}
return $ShimArray 
}
Function LPR
{
Param($ProductCode,$ProductName,$DateTimeRun)
MATSFingerPrint $ProductCode "LPR" $true $DateTimeRun
Write-DiagProgress -Activity ($LocalizedStrings.WindowsInstaller_IID_Attempting_to_resolve_problems_with+" " + $ProductName) # -Status $LocalizedStrings.WindowsInstaller_IID_Examining_registry_and_files
$alItemsToDelete= New-Object System.Collections.ArrayList($null)
$alItemsToDelete= [ProductLPRMain]::GetComponentPath($ProductCode) #generate list of installed items
$alItemsToDelete= $alItemsToDelete+(ShimXML($ProductCode)) #Do we have any shims If so add
$alItemsToDelete= $alItemsToDelete+(ARPEntries $ProductCode $ProductName ) #Do we have any extra ARP entry
if ($alItemsToDelete -ne $null)#If we found files and registry keys proceed 
{
                CheckRestorePoint $alItemsToDelete $ProductCode #If we have items to delete we will backup files and registry and create a system restore
    BackupFiles $alItemsToDelete $ProductCode #Attempting to backup all files and registry keys for selected product
                MATSFingerPrint $ProductCode "BackupFiles" $true $DateTimeRun
                Write-DiagProgress -Activity ($LocalizedStrings.WindowsInstaller_IID_Attempting_to_resolve_problems_with+" " + $ProductName) # -Status $LocalizedStrings.WindowsInstaller_IID_Creating_System_Restore_Checkpoint
                [Restore]::StartRestore((($LocalizedStrings.WindowsInstaller_IID_Final_Restore_Point_for+" "+ $ProductName+" "+$LocalizedStrings.WindowsInstaller_IID_using_Program_Install_and_Uninstall_troubleshooter))) #System Restore
                Write-DiagProgress -Activity ($LocalizedStrings.WindowsInstaller_IID_Attempting_to_resolve_problems_with+" " + $ProductName) # -Status $LocalizedStrings.WindowsInstaller_IID_Performing_Full_Product_Removal
    DeleteFilesFromXML($ProductCode)
                MATSFingerPrint $ProductCode "DeleteFiles" $True $DateTimeRun
                DeleteRegistryKeysFromXML($ProductCode)
}
}

Function BuildHKCRPatchList
{
param($Registrykey)
New-PSDrive -Name HKCR -PSProvider registry -root HKEY_CLASSES_ROOT |out-null # Allows access to registry HKEY_CLASSES_ROOT
$HKCRPatchListArray=@()
Get-ChildItem $Registrykey | foreach-object {
[string]$PatchGUID = ($_.name).substring(($_.name).lastindexof("\")+1,(($_.name).length-($_.name).lastindexof("\")-1)) 
$HKCRPatchListArray+=$PatchGUID
#test-path $Registrykey\$PatchGUID\patches
}
$HKCRPatchListArray
}

Function BuildHKLMPatchList
{
param($Registrykey)
$HKLMPatchListArray=@()
$Registrykey= RegistryHiveReplace($Registrykey)
Get-ChildItem $Registrykey | foreach-object {
[string]$PatchGUID = ($_.name).substring(($_.name).lastindexof("\")+1,(($_.name).length-($_.name).lastindexof("\")-1)) 
$HKLMPatchListArray+=$PatchGUID
}
$HKLMPatchListArray
}

Function CheckandFixHKCR
{
param ($CurrentPatchHKLMList, $CurrentPatchHKCRList, $Phase, $Deep)
Import-LocalizedData -BindingVariable LocalizedStrings -FileName Strings
New-PSDrive -Name HKCR -PSProvider registry -root HKEY_CLASSES_ROOT |out-null # Allows access to registry HKEY_CLASSES_ROOT

foreach ($HKCRPatchGUID in $CurrentPatchHKCRList)
{ 
    if ($Deep)
    {
         $PatchGUIDList=New-Object System.Collections.ArrayList
         $ProblemFound=$False
         $ProblemGUID=""
         $HKCRRegistryPatches="HKCR:\Installer\Products\$HKCRPatchGUID\Patches"  
         $HKCRRegistryPatchesManaged="Installer\Products\$HKCRPatchGUID\Patches"   
         $MultiStringToCheck=[HCRRegistryCheckMultiString]::HCRRegistryCheck($HKCRRegistryPatchesManaged)
         
         if ($MultiStringToCheck.length -gt 0)
         {
            foreach($Property in $MultiStringToCheck)
            {
                if ($Property.length -gt 0)
                {
                     if(!($CurrentPatchHKLMList -contains $Property))
                     {
                        if ($Phase -eq "TS")
                        {
                            Return $True                           
                        }
                        else
                        {
                            $ProblemFound=$true
                            $ProblemGUID=$Property
                            remove-itemproperty $HKCRRegistryPatches $ProblemGUID -ErrorAction SilentlyContinue
                        }                     
                     }
                     else
                     {
                        [void]$PatchGUIDList.add($Property)     
                     }
                }            
            }
            if ($ProblemFound)
            {
             [string[]]$MultiStringHolder=$PatchGUIDList
             remove-itemproperty $HKCRRegistryPatches "Patches"             
             New-ItemProperty -name "Patches" $HKCRRegistryPatches -value $MultiStringHolder -propertyType "MultiString"  
             }
            
         }
        
    }
    else
    {    
    #This handles the HKCR patches hive
    if(!($CurrentPatchHKLMList -contains $HKCRPatchGUID))
    {
        if ($Phase -eq "RS")
        {
            if((Test-Path "HKCR:\Installer\Patches\$HKCRPatchGUID") -and ($HKCRPatchGUID.length -gt 0))
            {
                WriteLog(($LocalizedStrings.WindowsInstaller_WriteLog_Delete)+ " - HKCR:\Installer\Patches\$HKCRPatchGUID")
                del "HKCR:\Installer\Patches\$HKCRPatchGUID" -recurse  #Delete HKCR
            }
        }
        else
        {
            Return $True
        }
    }
}
}
$False
}
Function FixMSITSRS
{
param ($SortArray, $Phase)
Import-LocalizedData -BindingVariable LocalizedStrings -FileName Strings
$DateTimeRun="{0:yyyy.MM.dd.HH.mm.ss}" -f (get-date) #Run Date\Time
WriteLog(($LocalizedStrings.WindowsInstaller_WriteLog_Start_Logging_MSI_Cache_Repair) + " " + ($Phase))  
$MSIMasterList=CheckCacheFilesMSI
$InstallerHive="HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Userdata\"
$RPRHives=Get-ChildItem $InstallerHive #IF need to delete component then add back recurse HKL,HKCR code here
if($Phase -eq "TS")
{
if (!((get-childitem -path $InstallerHive -ErrorAction SilentlyContinue).count -gt 0))
    {
       #Installer Hive MT Detect condition
    return $True
    }
}
foreach($SID in $RPRHives) 
{    
    foreach ($ProductCode in $SortArray)
    {
        $ProductCodeCompress=[CleanUpRegistry]::CompressGUID([string]$ProductCode)
        $RegistryKey=RegistryHiveReplace $SID.name
        $RegistryInstallProperty= $RegistryKey+"\Products\"+ $ProductCodeCompress +"\InstallProperties"
        if(test-path $RegistryInstallProperty)
        {
   $LocalPackage=get-itemproperty -path $RegistryInstallProperty -name LocalPackage -ErrorAction SilentlyContinue 
 
            if (($LocalPackage.LocalPackage).length -gt 0)
            {
                if (!((Test-Path -Path $LocalPackage.LocalPackage -PathType Leaf) -and ($LocalPackage.LocalPackage -like “*.msi”)))
                {
                    #no msi see if we can repair
                    if($Phase -eq "TS")
                    {
      WriteLog(($LocalizedStrings.WindowsInstaller_WriteLog_End_Logging_MSI_Cache_Repair) + " " + ($Phase))   
                        return $True
                    }
                    #in RS mode
                    RebuildInstallerCacheOnDrive $MSIMasterList $ProductCode $Phase $RegistryKey $ProductCodeCompress                        
                 }
            }
            else
            {
                #Local Package value is MT
                if($Phase -eq "TS")
                {
        return $True
                }
                #in RS mode
                RebuildInstallerCacheOnDrive $MSIMasterList $ProductCode $Phase $RegistryKey $ProductCodeCompress
            }
        }
    }
}
WriteLog(($LocalizedStrings.WindowsInstaller_WriteLog_End_Logging_MSI_Cache_Repair) + " " + ($Phase))   
return $false   
}
Function RebuildInstallerCacheOnDrive
{
param ($MSIMasterList, $ProductCode, $Phase, $RegistryKey, $ProductCodeCompress)
If ($MSIMasterList.ContainsKey($ProductCode))
                    {    
                        if($Phase -eq "TS")
                        {
       WriteLog(($LocalizedStrings.WindowsInstaller_WriteLog_End_Logging_MSI_Cache_Repair) + " " + ($Phase))   
                            return $True
                        }
                        else
                        {   
                            $value=$Env:WinDir+"\installer\"+[string]($MSIMasterList.Get_Item($ProductCode))
                            remove-itemproperty -path $RegistryKey\Products\$ProductCodeCompress\InstallProperties -name "LocalPackage" -ErrorAction SilentlyContinue
                            New-ItemProperty -path $RegistryKey\Products\$ProductCodeCompress\InstallProperties -name "LocalPackage"-value $value
                            WriteLog($LocalizedStrings.WindowsInstaller_WriteLog_Recreate + " -   $RegistryInstallProperty $value")
                        }
      
                    }
                    else
                    {
                        WriteLog($LocalizedStrings.WindowsInstaller_WriteLog_Cannot_recreate +" - $RegistryInstallProperty\$LocalPackage")
                    }
}
Function FixPatchTSRS
{
param ($SortArray, $Phase)
Import-LocalizedData -BindingVariable LocalizedStrings -FileName Strings
New-PSDrive -Name HKCR -PSProvider registry -root HKEY_CLASSES_ROOT |out-null # Allows access to registry HKEY_CLASSES_ROOT
WriteLog(($LocalizedStrings.WindowsInstaller_WriteLog_Start_Logging_Patch_Repair) + " " + ($Phase))  
$PatchHKLMList=@()
$PatchMasterList=CheckCacheFiles
$RPRHives=Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Userdata\ 
foreach($SID in $RPRHives) 
{     
    foreach ($ProductCode in $SortArray)
    {
        $ProductCode=[CleanUpRegistry]::CompressGUID([string]$ProductCode)
        $RegistryKey=RegistryHiveReplace $SID.name
        $RegistryPatch= $RegistryKey+"\Products\"+ $ProductCode +"\Patches"
        if(test-path $RegistryPatch)
        {
            $RegistryPatchList = dir $RegistryPatch -recurse
            if ($RegistryPatchList)
            {
                foreach ($Patch in $RegistryPatchList)
                {
                    $Patch=RegistryHiveReplace $Patch.name
                    $PatchGUID = $patch.substring($patch.lastindexof("\")+1,($patch.length-$patch.lastindexof("\")-1))
                    
                    if (test-path ($RegistryKey+"\Patches\"+$PatchGUID))
                    {
                         #part 2 check if msp exists
                         $LocalPatchPackage =(Get-ItemProperty -Path ($RegistryKey+"\Patches\"+$PatchGUID)).LocalPackage
                         If (!$LocalPatchPackage)
                         {
                            if($Phase -eq "TS")
                            {
        WriteLog(($LocalizedStrings.WindowsInstaller_WriteLog_End_Logging_MSP_Cache_Repair) + " " + ($Phase))   
                                return $True
                            }
                            else
                            {
                                #Resolver Phase
                                WriteLog(($LocalizedStrings.WindowsInstaller_WriteLog_Delete) + " - $RegistryKey\Patches\$PatchGUID")
                                del $RegistryKey\Patches\$PatchGUID
                                CleanUpPatches $RegistryPatch $Patch
                            }
                         }
                         elseif (!(Test-Path -path $LocalPatchPackage))
                         {         
                            if($Phase -eq "TS")
                            {
        WriteLog(($LocalizedStrings.WindowsInstaller_WriteLog_End_Logging_MSP_Cache_Repair) + " " + ($Phase))   
                                return $True
                            }                  
                            else
                            {
                                 #Resolver Phase

                                If ($PatchMasterList.ContainsKey($PatchGUID))
                                {
                                    WriteLog(($LocalizedStrings.WindowsInstaller_WriteLog_Delete) + " - $RegistryKey\Patches\$PatchGUID $LocalPatchPackage")
                                    del $RegistryKey\Patches\$PatchGUID
                                    new-item -path $RegistryKey\Patches\$PatchGUID  -ErrorAction SilentlyContinue | Out-Null  
                                    $value=$Env:WinDir+"\installer\"+[string]($PatchMasterList.Get_Item($PatchGUID))
                                    New-ItemProperty -path $RegistryKey\Patches\$PatchGUID -name "LocalPackage" -value $value
                                    WriteLog($LocalizedStrings.WindowsInstaller_WriteLog_Recreate +"- $RegistryKey\Patches\$PatchGUID $value")
                                 }
                                 else
                                 {
                                    WriteLog(($LocalizedStrings.WindowsInstaller_WriteLog_Delete) + " - $RegistryKey\Patches\$PatchGUID")
                                    del $RegistryKey\Patches\$PatchGUID
                                    CleanUpPatches $RegistryPatch $Patch
                                 }
                                 
                            }
                         } 
                    }
                    else
                    {
                            #no patch in expected location 
                            #Now check if file is on drive but registry key is not linked.
                            If ($PatchMasterList.ContainsKey($PatchGUID))
                            {
                                if($Phase -eq "TS")
                                {
         WriteLog(($LocalizedStrings.WindowsInstaller_WriteLog_End_Logging_MSP_Cache_Repair)+ " " +($Phase))   
                                    Return $True
                                }
                                else
                                {
                                    new-item -path $RegistryKey\Patches\$PatchGUID  -ErrorAction SilentlyContinue | Out-Null  
                                    $value=$Env:WinDir+"\installer\"+[string]($PatchMasterList.Get_Item($PatchGUID))
                                    if (!(Test-Path -path RegistryKey\Patches))
                                    {
                                         new-item -path $RegistryKey\Patches -ErrorAction SilentlyContinue | Out-Null 
                                         new-item -path $RegistryKey\Patches\$PatchGUID  -ErrorAction SilentlyContinue | Out-Null 
                                    }
                                    
                                    New-ItemProperty -path $RegistryKey\Patches\$PatchGUID -name "LocalPackage" -value $value
                                    WriteLog($LocalizedStrings.WindowsInstaller_WriteLog_Recreate + " - $RegistryKey\Patches\$PatchGUID $value")
                                }
                           }
                           else
                           {
                                if($Phase -eq "TS")
                                {
         WriteLog(($LocalizedStrings.WindowsInstaller_WriteLog_End_Logging_MSP_Cache_Repair) + " " + ($Phase))   
                                    Return $True
                                }
                                else
                                {
                                    CleanUpPatches $RegistryPatch $Patch
                                }
                            }                
                    }
                    if (test-path ($RegistryKey+"\Patches\"+$PatchGUID))
                    {
                        $PatchHKLMList+=($PatchGUID)
                    }
                }
            }
       }
      
     }
}
#Check HKCR fix
if (Test-Path "HKCR:\Installer\Patches")
{
    $PatchHKCRList=BuildHKCRPatchList("HKCR:\Installer\Patches") 
}
    if ($PatchHKCRList.length -gt 0)
    {
        $HCRReturn= CheckandFixHKCR $PatchHKLMList $PatchHKCRList $Phase $False
        if ($HCRReturn)
        {
            WriteLog(($LocalizedStrings.WindowsInstaller_WriteLog_End_Logging_MSP_Cache_Repair) + " " + ($Phase))   
            Return $true
        }
    }
    
    $ProductHKCRList=BuildHKCRPatchList("HKCR:\Installer\Products") 
    if ($ProductHKCRList.length -gt 0)
    {
        $HCRReturn= CheckandFixHKCR $PatchHKLMList $ProductHKCRList $Phase $True
		if ($HCRReturn)
        {
            WriteLog(($LocalizedStrings.WindowsInstaller_WriteLog_End_Logging_MSP_Cache_Repair) + " " + ($Phase))   
            Return $true
        }
    }

    WriteLog(($LocalizedStrings.WindowsInstaller_WriteLog_End_Logging_MSP_Cache_Repair) + " " + ($Phase))   
 
return $False
}


Function WriteLog
{
param ($TextOutput)
$DateTimeRun="{0:yyyy.MM.dd.HH.mm.ss}" -f (get-date) #Run Date\Time
$DateTimeRun+": "+ $TextOutput | out-File $env:Temp"\RPR_Patch_Repair.txt" -append  -ErrorAction SilentlyContinue | Out-Null  
}

function Setup-NamedEvent()
{
    switch($script:___Setup_NamedEvent)
    {
        1 {return $true}
        -1 {return $false}
    }
$sourceCode = @"
using System;
using System.Runtime.InteropServices;
public class NamedEvent : IDisposable
{
  [DllImport("kernel32.dll", SetLastError=true)]
  static extern IntPtr CreateEvent(IntPtr lpEventAttributes, bool bManualReset, bool bInitialState, string lpName);
  [DllImport("kernel32.dll", SetLastError = true)]
  [return: MarshalAs(UnmanagedType.Bool)]
  static extern bool CloseHandle(IntPtr hObject);
  public NamedEvent(string eventName)
            : this(eventName, false)
        {
        }
        public NamedEvent(string eventName, bool manualReset)
        {
            m_hEvent = CreateEvent(IntPtr.Zero, manualReset, false, eventName);
            m_alreadyExisted = (Marshal.GetLastWin32Error() == c_ERROR_ALREADY_EXISTS);
            GC.KeepAlive(m_hEvent);
        }
  public bool AlreadyExisted
  {
   get { return m_alreadyExisted; }
   set { m_alreadyExisted = value; }
  }
  private IntPtr m_hEvent;
  private bool m_alreadyExisted;
  private const int c_ERROR_ALREADY_EXISTS = 183;
  #region IDisposable Members
  public void Dispose()
  {
   CloseHandle(m_hEvent);
  }
  #endregion
}
"@
Add-Type -TypeDefinition $sourceCode -PassThru | Out-Null
$script:___Setup_NamedEvent = $true

    if($?)  { $script:___Setup_NamedEvent = 1; return $true  }
    else    { $script:___Setup_NamedEvent = -1; return $false }
}

function Mark($markName)
{
    if(Setup-NamedEvent)
    {
        $namedEvent = New-Object "NamedEvent" @($markName)
        return (-not $namedEvent.AlreadyExisted)
    }
}
#.NET Functions
#========================
PatchInfo
ProductListing
Compressed
GetRegistryType
SystemRestore
UninstallProduct
ProductLPR
MSIGetProductCodeFromPackage
Setup-NamedEvent
HCRRegistryCheckMultiString

# SIG # Begin signature block
# MIIarQYJKoZIhvcNAQcCoIIanjCCGpoCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUziAnyg/nGucxCK2F8XqRraPl
# YtOgghWCMIIEwzCCA6ugAwIBAgITMwAAAI1AOzwvrcZ7uAAAAAAAjTANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTUxMDA3MTgxNDA0
# WhcNMTcwMTA2MTgxNDA0WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OjE0OEMtQzRCOS0yMDY2MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAmlIcRUD3lFDw
# EK27FJ0L7nJcRvoCDIV9j+iH+phgZsI1T1UGjG7uVQs4h9Mb2JBiK5aScts3g08g
# LnHe5pCo2MWegFpCGe1iEOirtCDXABUGrJ9ZiiCGKfsEXsNPVsY++5zSMY/12JF6
# l7bpGQJelF6hlPrL2XDrOQZtg2MCfZB0FJXRsFR/IxkLriSCe6GrmTNlgTubfcrW
# YFtYRXYJ+SNfM4YCFuygR4Tajdl8gKdJPgtbSls1CS6q+XUgCxusXijAemjvDnd9
# HoFms2vuNAhkbfL1h8TF5Sgt3ZkOqgtrvw3c3d7hTiE2XhnokXxz3g7KwVRzwapf
# YhWha1xNPwIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFO+w/NAncAPOnsWMxNE+EIfv
# sAbaMB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBACJRuAWCXkNgGnPI5wLB03Kk2xYLTFzuC0uL6Pa4kQ6OlZ5O
# lJ8D1ISf29/eec92Gw/0oSi4XmgLE6yICiYYJUI+0esJjKj79iMyb3+Qo3lwgLBD
# ZPqg4rUXT32fZFHURjPJ0DmCV8JLhIllwTHp1tjrnpQwTq0oqeO8rbwrU0hvsgSv
# esdvxnu54IzsJnSWOsmpE3RCC0fTrnHIsU382SoqUnB5AlPK64q3ImJcQz0Ns5O5
# yUsWy0ef50ExRbA9tvBlvCZtVfCwLH/6U7MLoJcSkwpB7+VPlklyOxDBM8TXvRcE
# rgdbRaiScYxbKZGkb8JwJxLKt3KO4D8JT5MxVxYwggTsMIID1KADAgECAhMzAAAB
# Cix5rtd5e6asAAEAAAEKMA0GCSqGSIb3DQEBBQUAMHkxCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xIzAhBgNVBAMTGk1pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBMB4XDTE1MDYwNDE3NDI0NVoXDTE2MDkwNDE3NDI0NVowgYMxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
# MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xDTALBgNVBAsTBE1PUFIx
# HjAcBgNVBAMTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjCCASIwDQYJKoZIhvcNAQEB
# BQADggEPADCCAQoCggEBAJL8bza74QO5KNZG0aJhuqVG+2MWPi75R9LH7O3HmbEm
# UXW92swPBhQRpGwZnsBfTVSJ5E1Q2I3NoWGldxOaHKftDXT3p1Z56Cj3U9KxemPg
# 9ZSXt+zZR/hsPfMliLO8CsUEp458hUh2HGFGqhnEemKLwcI1qvtYb8VjC5NJMIEb
# e99/fE+0R21feByvtveWE1LvudFNOeVz3khOPBSqlw05zItR4VzRO/COZ+owYKlN
# Wp1DvdsjusAP10sQnZxN8FGihKrknKc91qPvChhIqPqxTqWYDku/8BTzAMiwSNZb
# /jjXiREtBbpDAk8iAJYlrX01boRoqyAYOCj+HKIQsaUCAwEAAaOCAWAwggFcMBMG
# A1UdJQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBSJ/gox6ibN5m3HkZG5lIyiGGE3
# NDBRBgNVHREESjBIpEYwRDENMAsGA1UECxMETU9QUjEzMDEGA1UEBRMqMzE1OTUr
# MDQwNzkzNTAtMTZmYS00YzYwLWI2YmYtOWQyYjFjZDA1OTg0MB8GA1UdIwQYMBaA
# FMsR6MrStBZYAck3LjMWFrlMmgofMFYGA1UdHwRPME0wS6BJoEeGRWh0dHA6Ly9j
# cmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY0NvZFNpZ1BDQV8w
# OC0zMS0yMDEwLmNybDBaBggrBgEFBQcBAQROMEwwSgYIKwYBBQUHMAKGPmh0dHA6
# Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljQ29kU2lnUENBXzA4LTMx
# LTIwMTAuY3J0MA0GCSqGSIb3DQEBBQUAA4IBAQCmqFOR3zsB/mFdBlrrZvAM2PfZ
# hNMAUQ4Q0aTRFyjnjDM4K9hDxgOLdeszkvSp4mf9AtulHU5DRV0bSePgTxbwfo/w
# iBHKgq2k+6apX/WXYMh7xL98m2ntH4LB8c2OeEti9dcNHNdTEtaWUu81vRmOoECT
# oQqlLRacwkZ0COvb9NilSTZUEhFVA7N7FvtH/vto/MBFXOI/Enkzou+Cxd5AGQfu
# FcUKm1kFQanQl56BngNb/ErjGi4FrFBHL4z6edgeIPgF+ylrGBT6cgS3C6eaZOwR
# XU9FSY0pGi370LYJU180lOAWxLnqczXoV+/h6xbDGMcGszvPYYTitkSJlKOGMIIF
# vDCCA6SgAwIBAgIKYTMmGgAAAAAAMTANBgkqhkiG9w0BAQUFADBfMRMwEQYKCZIm
# iZPyLGQBGRYDY29tMRkwFwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0MS0wKwYDVQQD
# EyRNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkwHhcNMTAwODMx
# MjIxOTMyWhcNMjAwODMxMjIyOTMyWjB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBD
# QTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBALJyWVwZMGS/HZpgICBC
# mXZTbD4b1m/My/Hqa/6XFhDg3zp0gxq3L6Ay7P/ewkJOI9VyANs1VwqJyq4gSfTw
# aKxNS42lvXlLcZtHB9r9Jd+ddYjPqnNEf9eB2/O98jakyVxF3K+tPeAoaJcap6Vy
# c1bxF5Tk/TWUcqDWdl8ed0WDhTgW0HNbBbpnUo2lsmkv2hkL/pJ0KeJ2L1TdFDBZ
# +NKNYv3LyV9GMVC5JxPkQDDPcikQKCLHN049oDI9kM2hOAaFXE5WgigqBTK3S9dP
# Y+fSLWLxRT3nrAgA9kahntFbjCZT6HqqSvJGzzc8OJ60d1ylF56NyxGPVjzBrAlf
# A9MCAwEAAaOCAV4wggFaMA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFMsR6MrS
# tBZYAck3LjMWFrlMmgofMAsGA1UdDwQEAwIBhjASBgkrBgEEAYI3FQEEBQIDAQAB
# MCMGCSsGAQQBgjcVAgQWBBT90TFO0yaKleGYYDuoMW+mPLzYLTAZBgkrBgEEAYI3
# FAIEDB4KAFMAdQBiAEMAQTAfBgNVHSMEGDAWgBQOrIJgQFYnl+UlE/wq4QpTlVnk
# pDBQBgNVHR8ESTBHMEWgQ6BBhj9odHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtp
# L2NybC9wcm9kdWN0cy9taWNyb3NvZnRyb290Y2VydC5jcmwwVAYIKwYBBQUHAQEE
# SDBGMEQGCCsGAQUFBzAChjhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2Nl
# cnRzL01pY3Jvc29mdFJvb3RDZXJ0LmNydDANBgkqhkiG9w0BAQUFAAOCAgEAWTk+
# fyZGr+tvQLEytWrrDi9uqEn361917Uw7LddDrQv+y+ktMaMjzHxQmIAhXaw9L0y6
# oqhWnONwu7i0+Hm1SXL3PupBf8rhDBdpy6WcIC36C1DEVs0t40rSvHDnqA2iA6VW
# 4LiKS1fylUKc8fPv7uOGHzQ8uFaa8FMjhSqkghyT4pQHHfLiTviMocroE6WRTsgb
# 0o9ylSpxbZsa+BzwU9ZnzCL/XB3Nooy9J7J5Y1ZEolHN+emjWFbdmwJFRC9f9Nqu
# 1IIybvyklRPk62nnqaIsvsgrEA5ljpnb9aL6EiYJZTiU8XofSrvR4Vbo0HiWGFzJ
# NRZf3ZMdSY4tvq00RBzuEBUaAF3dNVshzpjHCe6FDoxPbQ4TTj18KUicctHzbMrB
# 7HCjV5JXfZSNoBtIA1r3z6NnCnSlNu0tLxfI5nI3EvRvsTxngvlSso0zFmUeDord
# EN5k9G/ORtTTF+l5xAS00/ss3x+KnqwK+xMnQK3k+eGpf0a7B2BHZWBATrBC7E7t
# s3Z52Ao0CW0cgDEf4g5U3eWh++VHEK1kmP9QFi58vwUheuKVQSdpw5OPlcmN2Jsh
# rg1cnPCiroZogwxqLbt2awAdlq3yFnv2FoMkuYjPaqhHMS+a3ONxPdcAfmJH0c6I
# ybgY+g5yjcGjPa8CQGr/aZuW4hCoELQ3UAjWwz0wggYHMIID76ADAgECAgphFmg0
# AAAAAAAcMA0GCSqGSIb3DQEBBQUAMF8xEzARBgoJkiaJk/IsZAEZFgNjb20xGTAX
# BgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1pY3Jvc29mdCBSb290
# IENlcnRpZmljYXRlIEF1dGhvcml0eTAeFw0wNzA0MDMxMjUzMDlaFw0yMTA0MDMx
# MzAzMDlaMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAf
# BgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQTCCASIwDQYJKoZIhvcNAQEB
# BQADggEPADCCAQoCggEBAJ+hbLHf20iSKnxrLhnhveLjxZlRI1Ctzt0YTiQP7tGn
# 0UytdDAgEesH1VSVFUmUG0KSrphcMCbaAGvoe73siQcP9w4EmPCJzB/LMySHnfL0
# Zxws/HvniB3q506jocEjU8qN+kXPCdBer9CwQgSi+aZsk2fXKNxGU7CG0OUoRi4n
# rIZPVVIM5AMs+2qQkDBuh/NZMJ36ftaXs+ghl3740hPzCLdTbVK0RZCfSABKR2YR
# JylmqJfk0waBSqL5hKcRRxQJgp+E7VV4/gGaHVAIhQAQMEbtt94jRrvELVSfrx54
# QTF3zJvfO4OToWECtR0Nsfz3m7IBziJLVP/5BcPCIAsCAwEAAaOCAaswggGnMA8G
# A1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFCM0+NlSRnAK7UD7dvuzK7DDNbMPMAsG
# A1UdDwQEAwIBhjAQBgkrBgEEAYI3FQEEAwIBADCBmAYDVR0jBIGQMIGNgBQOrIJg
# QFYnl+UlE/wq4QpTlVnkpKFjpGEwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcG
# CgmSJomT8ixkARkWCW1pY3Jvc29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3Qg
# Q2VydGlmaWNhdGUgQXV0aG9yaXR5ghB5rRahSqClrUxzWPQHEy5lMFAGA1UdHwRJ
# MEcwRaBDoEGGP2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1
# Y3RzL21pY3Jvc29mdHJvb3RjZXJ0LmNybDBUBggrBgEFBQcBAQRIMEYwRAYIKwYB
# BQUHMAKGOGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9z
# b2Z0Um9vdENlcnQuY3J0MBMGA1UdJQQMMAoGCCsGAQUFBwMIMA0GCSqGSIb3DQEB
# BQUAA4ICAQAQl4rDXANENt3ptK132855UU0BsS50cVttDBOrzr57j7gu1BKijG1i
# uFcCy04gE1CZ3XpA4le7r1iaHOEdAYasu3jyi9DsOwHu4r6PCgXIjUji8FMV3U+r
# kuTnjWrVgMHmlPIGL4UD6ZEqJCJw+/b85HiZLg33B+JwvBhOnY5rCnKVuKE5nGct
# xVEO6mJcPxaYiyA/4gcaMvnMMUp2MT0rcgvI6nA9/4UKE9/CCmGO8Ne4F+tOi3/F
# NSteo7/rvH0LQnvUU3Ih7jDKu3hlXFsBFwoUDtLaFJj1PLlmWLMtL+f5hYbMUVbo
# nXCUbKw5TNT2eb+qGHpiKe+imyk0BncaYsk9Hm0fgvALxyy7z0Oz5fnsfbXjpKh0
# NbhOxXEjEiZ2CzxSjHFaRkMUvLOzsE1nyJ9C/4B5IYCeFTBm6EISXhrIniIh0EPp
# K+m79EjMLNTYMoBMJipIJF9a6lbvpt6Znco6b72BJ3QGEe52Ib+bgsEnVLaxaj2J
# oXZhtG6hE6a/qkfwEm/9ijJssv7fUciMI8lmvZ0dhxJkAj0tr1mPuOQh5bWwymO0
# eFQF1EEuUKyUsKV4q7OglnUa2ZKHE3UiLzKoCG6gW4wlv6DvhMoh1useT8ma7kng
# 9wFlb4kLfchpyOZu6qeXzjEp/w7FW1zYTRuh2Povnj8uVRZryROj/TGCBJUwggSR
# AgEBMIGQMHkxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xIzAh
# BgNVBAMTGk1pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBAhMzAAABCix5rtd5e6as
# AAEAAAEKMAkGBSsOAwIaBQCgga4wGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQw
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFBbj
# J5fUvJ2/qKerzefrWtbtaUJOME4GCisGAQQBgjcCAQwxQDA+oByAGgBNAFMASQBN
# AEEAVABTAEYATgAuAHAAcwAxoR6AHGh0dHA6Ly9zdXBwb3J0Lm1pY3Jvc29mdC5j
# b20wDQYJKoZIhvcNAQEBBQAEggEAWf63T0NS+jyppS0mYmrfi3b4C4G10pj2aiIv
# ju4PaXYuLwGxJSPUIwNBebGKW7unV4gLfyYr6djVyxvDybR/i/jOdTSH/jOh8vxg
# rgnDlOpOngg7XA9duIyQ4nhbE13xTPIJC2F9LyachBfSMGNpIAL+4ES/NYyhPPQc
# oNGkoP3hYHEWuNJteRRHCUxSvX4M/FJxqjgI/1w7arK8E0DUsnJs+pc0+iyodYZ2
# 9OCJETMwS4lJBuOSgCDYJk8yQRlSH0FlGLUV5euEefVE8tav4eqYAavAmL7sTL7d
# h2Iv4RqP6GZJNIxeEWFoIg1r1MULgN2i3jPH3/fc9+kkWOG2pqGCAigwggIkBgkq
# hkiG9w0BCQYxggIVMIICEQIBATCBjjB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMSEwHwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EC
# EzMAAACNQDs8L63Ge7gAAAAAAI0wCQYFKw4DAhoFAKBdMBgGCSqGSIb3DQEJAzEL
# BgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkFMQ8XDTE1MTIxNzA4NTMwN1owIwYJKoZI
# hvcNAQkEMRYEFGcd/xmKWz3L9UY350+9f5ppunaBMA0GCSqGSIb3DQEBBQUABIIB
# ADKU4u1j1zkV8NCePXbgc0jH61c07OWZkhDvxiObgCDJmd9srfEiaIGmgokq+05f
# F3oR4KHyIL/fDmb4wY/Z6YkcWRMB3ukArdNdk9izvWtRipjmaB48tSsNHEQRbDP/
# J/2P/tx8i9W0qMEDMksW7/GfGkHvpHiUb5JWElUDpgEDaHOEnDzwyxE6fvVM4XHs
# 6YjIspKUuYL1tI9iLuAKyaP/FutHzzb3C2QZeq3G5cQIysWgvEmZboKy1sAnfVN5
# VO8IhjiULBILbOqv1ZDbEpN9ckyz/ti7Gco6qQl29vSc9zld1Bdnh7IHNaJYlgKY
# BIofY06hybGTu9siGXe4iiM=
# SIG # End signature block
