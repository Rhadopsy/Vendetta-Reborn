Attribute VB_Name = "ModDirectXAliases"
Option Explicit

' twinBASIC/WinDevLib exposes the native DirectX COM interfaces with their
' canonical I-prefixed names. The historical project used the friendly
' aliases exported by dx7vb.dll. Keeping the aliases here lets the port move
' one subsystem at a time without touching every public method signature.
#If TWINBASIC Then
Public Alias DirectDraw7 As IDirectDraw7
Public Alias DirectDrawSurface7 As IDirectDrawSurface7
Public Alias DirectDrawClipper As IDirectDrawClipper

Public Alias DirectInput As IDirectInput7

Public Alias DirectSoundBuffer As IDirectSoundBuffer

Public Alias DirectMusicSegment As IDirectMusicSegment
Public Alias DirectMusicLoader As IDirectMusicLoader
#End If
