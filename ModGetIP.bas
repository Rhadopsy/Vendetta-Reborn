Attribute VB_Name = "ModGetIP"
' Vendetta - Standalone Role Playing game, simulate a MMORPG with AI
' Copyright (C) 2003-2007  S. HARLAUT (sharlaut@netcourrier.com)
'
' This program is free software; you can redistribute it and/or
' modify it under the terms of the GNU General Public License
' as published by the Free Software Foundation; either version 2
' of the License, or (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

Public Type RESUL_IP
    IPAdr As String
    IPMask As String
    Interface_Nbre As String
End Type

Public Type DBYTE
    unused1 As Byte
    unused2 As Byte
End Type

Public Type MIB_IPADDRROW
    dwAddr As Long ' IP address
    dwIndex As Long ' interface index
    dwMask As Long ' subnet mask
    dwBCastAddr As Long ' broadcast address
    dwReasmSize As Long ' rassembly size
    unused1 As DBYTE ' not currently used
    unused2 As DBYTE ' not currently used
End Type

Public Type MIB_IPADDRTABLE
    dwNumEntries As Long
    table(400) As MIB_IPADDRROW
End Type

' recupere le nombre d'interface sur le pc
Public Declare Function GetNumberOfInterfaces Lib "iphlpapi.dll" _
   (ByRef PDWORD As Long) As Long

' recupere les adresses IP de la machine
Public Declare Function GetIpAddrTable Lib "iphlpapi.dll" _
     (ByRef pIpAddrTable As MIB_IPADDRTABLE, _
      ByRef pdwSize As Long, _
      bOrder As Boolean) As Long
Public Function FIRST_IPADDRESS(ByVal ipAddress As Long) As Long
    FIRST_IPADDRESS = Val("&H" & Left(Right("00000000" & Hex(ipAddress), 8), 2))
End Function
Public Function SECOND_IPADDRESS(ByVal ipAddress As Long) As Long
    SECOND_IPADDRESS = Val("&H" & Mid(Right("00000000" & Hex(ipAddress), 8), 3, 2))
End Function
Public Function THIRD_IPADDRESS(ByVal ipAddress As Long) As Long
    THIRD_IPADDRESS = Val("&H" & Mid(Right("00000000" & Hex(ipAddress), 8), 5, 2))
End Function
Public Function FOURTH_IPADDRESS(ByVal ipAddress As Long) As Long
    FOURTH_IPADDRESS = Val("&H" & Right("00" & Hex(ipAddress), 2))
End Function
Function recup_ip(Num_interface As Long) As RESUL_IP
    'recupere l'adresse ip de l'interface passée en paramettre
    Dim toto1 As MIB_IPADDRTABLE
    Dim toto2 As Long
    toto2 = 400
    yop1 = GetIpAddrTable(toto1, toto2, True)
    recup_ip.IPAdr = FOURTH_IPADDRESS(toto1.table(Num_interface).dwAddr) & "." & _
             THIRD_IPADDRESS(toto1.table(Num_interface).dwAddr) & "." & _
             SECOND_IPADDRESS(toto1.table(Num_interface).dwAddr) & "." & _
             FIRST_IPADDRESS(toto1.table(Num_interface).dwAddr)
    recup_ip.IPMask = FOURTH_IPADDRESS(toto1.table(Num_interface).dwMask) & "." & _
             THIRD_IPADDRESS(toto1.table(Num_interface).dwMask) & "." & _
             SECOND_IPADDRESS(toto1.table(Num_interface).dwMask) & "." & _
             FIRST_IPADDRESS(toto1.table(Num_interface).dwMask)
    recup_ip.Interface_Nbre = toto1.table(Num_interface).dwIndex
End Function

