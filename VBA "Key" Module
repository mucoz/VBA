Option Explicit

'Created by Mustafa Can Ozturk on 13.02.2021
'===========================================
'The functions below will be used for keyboard operations
'It includes the most common functions
'It needs to be used in "Key" module
'Functions:
'Press, Special
'===========================================

Enum SpecialKeys

    BACKSPACE
    BREAK
    CAPSLOCK
    DELETEKey
    DOWNARROW
    ENDKey
    ENTER
    ESC
    HELPKey
    HOME
    INSERTKey
    LEFTARROW
    NUMLOCK
    PAGEDOWN
    PAGEUP
    PRINTSCREEN
    RIGHTARROW
    SCROLLLOCK
    TABKey
    UPARROW
    F1
    F2
    F3
    F4
    F5
    F6
    F7
    F8
    F9
    F10
    F11
    F12
    F13
    F14
    F15
    F16
    SHIFT
    CTRL
    ALT

End Enum

Public Function Special(SpecialKey As SpecialKeys, Optional times As Variant) As String
    
    Dim result As String
    

 
     Select Case SpecialKey
     
         Case SpecialKeys.BACKSPACE
         result = "{BACKSPACE}"
         Case SpecialKeys.BREAK
         result = "{BREAK}"
         Case SpecialKeys.CAPSLOCK
         result = "{CAPSLOCK}"
         Case SpecialKeys.DELETEKey
         result = "{DELETE}"
         Case SpecialKeys.DOWNARROW
         result = "{DOWN}"
         Case SpecialKeys.ENDKey
         result = "{END}"
         Case SpecialKeys.ENTER
         result = "{ENTER}"
         Case SpecialKeys.ESC
         result = "{ESC}"
         Case SpecialKeys.HELPKey
         result = "{HELP}"
         Case SpecialKeys.HOME
         result = "{HOME}"
         Case SpecialKeys.INSERTKey
         result = "{INSERT}"
         Case SpecialKeys.LEFTARROW
         result = "{LEFT}"
         Case SpecialKeys.NUMLOCK
         result = "{NUMLOCK}"
         Case SpecialKeys.PAGEDOWN
         result = "{PGDN}"
         Case SpecialKeys.PAGEUP
         result = "{PGUP}"
         Case SpecialKeys.PRINTSCREEN
         result = "{PRTSC}"
         Case SpecialKeys.RIGHTARROW
         result = "{RIGHT}"
         Case SpecialKeys.SCROLLLOCK
         result = "{SCROLLLOCK}"
         Case SpecialKeys.TABKey
         result = "{TAB}"
         Case SpecialKeys.UPARROW
         result = "{UP}"
         Case SpecialKeys.F1
         result = "{F1}"
         Case SpecialKeys.F2
         result = "{F2}"
         Case SpecialKeys.F3
         result = "{F3}"
         Case SpecialKeys.F4
         result = "{F4}"
         Case SpecialKeys.F5
         result = "{F5}"
         Case SpecialKeys.F6
         result = "{F6}"
         Case SpecialKeys.F7
         result = "{F7}"
         Case SpecialKeys.F8
         result = "{F8}"
         Case SpecialKeys.F9
         result = "{F9}"
         Case SpecialKeys.F10
         result = "{F10}"
         Case SpecialKeys.F11
         result = "{F11}"
         Case SpecialKeys.F12
         result = "{F12}"
         Case SpecialKeys.F13
         result = "{F13}"
         Case SpecialKeys.F14
         result = "{F14}"
         Case SpecialKeys.F15
         result = "{F15}"
         Case SpecialKeys.F16
         result = "{F16}"
         Case SpecialKeys.SHIFT
         result = "+"
         Case SpecialKeys.CTRL
         result = "^"
         Case SpecialKeys.ALT
         result = "%"
     
     End Select
     
    If IsMissing(times) = True Then
     
        Special = result
     
    Else
        
        If IsNumeric(times) = False Then
        
            Debug.Print "Function Error : Key.Special -> Number of Times is not numeric"
        
        Else
        
            Special = JoinKeyAndNumber(result, CInt(times))
        
        End If
        
    End If
    
End Function

Private Function JoinKeyAndNumber(key As String, times As Integer) As String

    Dim specKey As String
    
    If Right(key, 1) = "}" Then
    
        specKey = Left(key, Len(key) - 1) + " " + CStr(times) + "}"
    
    Else
    
        specKey = key
    
    End If
    
    JoinKeyAndNumber = specKey

End Function
Public Sub Press(ByVal TextOrSingleKey As String, Optional NumberofTimes As Variant)

    Dim times As String
    Dim i As Long
    
    If IsMissing(NumberofTimes) = True Then
    
        SendKeys TextOrSingleKey
    
    Else
    
        If IsNumeric(NumberofTimes) = False Then
            
            Debug.Print "Enter a valid number (times) for Key Press function"
            Exit Sub
    
        Else
        
            times = NumberofTimes
              
            For i = 1 To times
                
                SendKeys TextOrSingleKey
            
            Next i
            
        End If
        
    End If
      
End Sub
