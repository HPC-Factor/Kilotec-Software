Attribute VB_Name = "Module1"
Option Explicit
Public CommandLine(100)

Public Hi As Integer
Public Lo As Integer

Public CHAR As Integer
Public LINE As Integer
Public VorNull As String
Public VorNull2 As String
Public COMMAND As String
Public JUMP As Boolean
Public HELP As String
Public LABEL As Integer


Public Hints As Integer
Public TimerSpeed As Integer
Public Layout As Integer
Public Scroll As Integer

Public TOGGLE As Boolean

Public Indexx As Integer
Public SaveFlag As Boolean

Public TempZeile As String
Public TempAnfang As String
Public TempMitte As String
Public TempEnde As String

Public Selected As Boolean

Public Const MemorySize = 64
Public Const StackSize = 72

Public stk(72) As Integer
Public mem(64) As Integer
Public reg(8) As Integer
Public A(8) As String
Public pc As Long
Public sp As Integer
Public err As Boolean
Public zero As Boolean
Public carry As Boolean
Public Bbyte As Integer
Public r As Integer
Public Enter As Integer
Public Sprung As Integer
Public Compare As Boolean
Public equal As Boolean
Public lower As Boolean
Public higher As Boolean

'NEUE VERSION
Public OPCODE As String




Public Sub ErrorMsg(ByVal s As String)
MsgBox "Error: " & s, vbOKOnly, "Error"
err = True
frmOutput.Show
End Sub







Public Sub Run()
Dim Leer As Integer
Dim ersterAbsatz As String
Dim zweiterAbsatz As String


On Error Resume Next



pc = pc + 1

If pc > frmMain.List1.ListCount Then
    pc = 0
    frmMain.Timer1.Enabled = False
    frmOutput.Show
    Exit Sub
End If


If InStr(9, frmMain.List1.List(pc - 1), " ") = 12 Then Leer = 3
If InStr(9, frmMain.List1.List(pc - 1), " ") = 11 Then Leer = 2
If InStr(9, frmMain.List1.List(pc - 1), " ") = 13 Then Leer = 4

OPCODE = Mid(frmMain.List1.List(pc - 1), 9, Leer)


'#####
If InStr(9, frmMain.List1.List(pc - 1), " ") = 11 Then Leer = 2
If InStr(9, frmMain.List1.List(pc - 1), " ") = 12 Then Leer = 3
If InStr(9, frmMain.List1.List(pc - 1), " ") = 13 Then Leer = 4
ersterAbsatz = (InStr(9, frmMain.List1.List(pc - 1), " "))
Hi = Mid(frmMain.List1.List(pc - 1), 9 + Leer + 2, 1)
If OPCODE = "MOV" Or OPCODE = "CMP" Then Lo = Mid(frmMain.List1.List(pc - 1), 9 + Leer + 5, 1)
If OPCODE = "LDI" Or OPCODE = "RND" Then Lo = Mid(frmMain.List1.List(pc - 1), 9 + Leer + 4, 3)
If OPCODE = "JNZ" Then Hi = Mid(frmMain.List1.List(pc - 1), 9 + Leer + 3, 2)
If OPCODE = "JNE" Then Hi = Mid(frmMain.List1.List(pc - 1), 9 + Leer + 3, 2)
If OPCODE = "JSR" Then Hi = Mid(frmMain.List1.List(pc - 1), 9 + Leer + 3, 2)
If OPCODE = "JZ" Then Hi = Mid(frmMain.List1.List(pc - 1), 9 + Leer + 3, 2)
If OPCODE = "JC" Then Hi = Mid(frmMain.List1.List(pc - 1), 9 + Leer + 3, 2)
If OPCODE = "JE" Then Hi = Mid(frmMain.List1.List(pc - 1), 9 + Leer + 3, 2)
If OPCODE = "JL" Then Hi = Mid(frmMain.List1.List(pc - 1), 9 + Leer + 3, 2)
If OPCODE = "JLE" Then Hi = Mid(frmMain.List1.List(pc - 1), 9 + Leer + 3, 2)
If OPCODE = "JA" Then Hi = Mid(frmMain.List1.List(pc - 1), 9 + Leer + 3, 2)
If OPCODE = "JAE" Then Hi = Mid(frmMain.List1.List(pc - 1), 9 + Leer + 3, 2)
If OPCODE = "JMP" Then Hi = Mid(frmMain.List1.List(pc - 1), 9 + Leer + 3, 2)

If OPCODE = "MSG" Then
    Dim EndeString As Integer
    Dim Laenge As Integer
    EndeString = (InStr(9 + Leer + 5, frmMain.List1.List(pc - 1), "'"))
    Laenge = Len(Mid(frmMain.List1.List(pc - 1), 9 + Leer + 5, EndeString))
    Lo = Mid(frmMain.List1.List(pc - 1), 9 + Leer + 5, Laenge - 9 - Leer - 5)
End If
'#####
          
     
          
          
Bbyte = mem(pc)
Sprung = Bbyte

If OPCODE = ";" Then pc = pc + 1


'BRK
If OPCODE = "BRK" Then
    frmMain.Timer1.Enabled = False
    ErrorMsg ("BREAK")
ElseIf OPCODE = "TOG" Then
    If TOGGLE = True Then
        TOGGLE = False
    Else
        TOGGLE = True
    End If

'RET
ElseIf OPCODE = "RET" Then
    If sp = 0 Then
        ErrorMsg ("STACK <")
    Else
        pc = stk(sp)
        sp = sp - 1
    End If

'MOV
ElseIf OPCODE = "MOV" Then
    reg(Hi) = reg(Lo)
    
'RND
ElseIf OPCODE = "RND" Then
    Randomize Timer
    reg(Hi) = Int(Rnd() * Lo)
    
'MSG
ElseIf OPCODE = "MSG" Then
    A(Hi) = Lo

'OUT
ElseIf OPCODE = "OUT" Then
    frmOutput.txtOutput.Text = A(Hi)

'PUSH
ElseIf OPCODE = "PUSH" Then
        If sp >= (StackSize - 1) Then
            ErrorMsg ("STACK >")
        Else
            sp = sp + 1
            stk(sp) = reg(Hi)
        End If
'POP
ElseIf OPCODE = "POP" Then
        If sp = 0 Then
            ErrorMsg ("STACK <")
        Else
            sp = sp - 1
            reg(Hi) = stk(sp)
        End If


'ADD
ElseIf OPCODE = "ADD" Then
        If (reg(0) + CInt(reg(Hi))) > 9999 Then
            carry = True
        Else
            carry = False
        End If
        reg(0) = reg(0) + CInt(reg(Hi))
        If carry = True Then reg(0) = reg(0) - 10000
        If reg(0) = 0 Then
            zero = True
        Else
            zero = False
        End If
    
'SUB
ElseIf OPCODE = "SUB" Then
        If CInt(reg(Hi)) > reg(0) Then
            carry = True
        Else
            carry = False
        End If
        reg(0) = reg(0) - CInt(reg(Hi))
        If carry = True Then reg(0) = reg(0) + 10000
        If reg(0) = 0 Then
            zero = True
        Else
            zero = False
        End If
    
'MUL
ElseIf OPCODE = "MUL" Then
        If (reg(0) * CInt(reg(Hi))) > 9999 Then
            carry = True
        Else
            carry = False
        End If
        reg(0) = reg(0) * CInt(reg(Hi))
        If carry = True Then reg(0) = reg(0) - 10000
        If reg(0) = 0 Then
            zero = True
        Else
            zero = False
        End If
    
'DIV
ElseIf OPCODE = "DIV" Then
        If CInt(reg(0)) / CInt(reg(Hi)) < 0 Then
            carry = True
        Else
            carry = False
        End If
        reg(0) = reg(0) / CInt(reg(Hi))
        If carry = True Then reg(0) = reg(0) + 10000
        If reg(0) = 0 Then
            zero = True
        Else
            zero = False
        End If
    
    
    
'LDI
ElseIf OPCODE = "LDI" Then
        reg(Hi) = Lo


'DEC
ElseIf OPCODE = "DEC" Then
        If reg(Hi) = 0 Then
            carry = True
        Else
            carry = False
        End If
        reg(Hi) = reg(Hi) - 1
        If carry = True Then reg(Hi) = 9999
        If reg(Hi) = 0 Then
            zero = True
        Else
            zero = False
        End If

'INC
ElseIf OPCODE = "INC" Then
        If reg(Hi) = 9999 Then
            carry = True
        Else
            carry = False
        End If
        reg(Hi) = reg(Hi) + 1
        If carry = True Then reg(Hi) = 0
        If reg(Hi) = 0 Then
            zero = True
        Else
            zero = False
        End If

'IN
ElseIf OPCODE = "IN" Then

            frmMain.Timer1.Enabled = False
            frmEnter.Label1.Caption = "R" & Hi & "= "
            frmEnter.Show

'CMP
ElseIf OPCODE = "CMP" Then
            If CInt(reg(Hi)) = CInt(reg(Lo)) Then
                equal = True
            Else
                equal = False
            End If
            
            If CInt(reg(Hi)) < CInt(reg(Lo)) Then
                higher = True
            Else
                higher = False
            End If
            
            If CInt(reg(Hi)) > CInt(reg(Lo)) Then
                lower = True
            Else
                lower = False
            End If
            
            

'########## JUMPS
'JE
ElseIf OPCODE = "JE" Then
            If equal = True Then pc = CLng("&H" & Hi)

'JNE
ElseIf OPCODE = "JNE" Then
            If equal = False Then pc = CLng("&H" & Hi)

'JL
ElseIf OPCODE = "JL" Then
            If lower = True Then pc = CLng("&H" & Hi)

'JLE
ElseIf OPCODE = "JLE" Then
            If lower = True Or equal = True Then pc = CLng("&H" & Hi)


'JA
ElseIf OPCODE = "JA" Then
            If higher = True Then pc = CLng("&H" & Hi)
            

'JAE
ElseIf OPCODE = "JAE" Then
            If higher = True Or equal = True Then pc = CLng("&H" & Hi)


'JZ
ElseIf OPCODE = "JZ" Then
            If zero = True Then pc = CLng("&H" & Hi)

'JNZ
ElseIf OPCODE = "JNZ" Then
            If zero = False Then
                pc = CLng("&H" & Hi)
            End If

'JC
ElseIf OPCODE = "JC" Then
            If carry = True Then
                pc = CLng("&H" & Hi)
            End If
'JSR
ElseIf OPCODE = "JSR" Then
            
            If sp >= StackSize - 1 Then
                ErrorMsg ("STACK >")
            Else
                sp = sp + 1
                stk(sp) = pc
                pc = CLng("&H" & Hi)
            End If
'JMP
ElseIf OPCODE = "JMP" Then
                pc = CLng("&H" & Hi)


ElseIf OPCODE = "NOP" Or OPCODE = ";" Then
        ' do nothing

End If


' DEBUG MONITOR
frmOutput.lstOutput.Clear
frmOutput.lstOutput.AddItem "R0 = " & reg(0)
frmOutput.lstOutput.AddItem "R1 = " & reg(1)
frmOutput.lstOutput.AddItem "R2 = " & reg(2)
frmOutput.lstOutput.AddItem "R3 = " & reg(3)
frmOutput.lstOutput.AddItem "R4 = " & reg(4)
frmOutput.lstOutput.AddItem "R5 = " & reg(5)
frmOutput.lstOutput.AddItem "R6 = " & reg(6)
frmOutput.lstOutput.AddItem "R7 = " & reg(7)

Dim vnull As String
vnull = ""
If pc < 16 Then vnull = "0"

frmOutput.List1.Clear
frmOutput.List1.AddItem "Line (adr):." & pc & " (" & vnull & Hex(pc) & ")"
frmOutput.List1.AddItem "Stack:......" & sp
frmOutput.List1.AddItem "Carry flag:." & carry
frmOutput.List1.AddItem "Zero flag:.." & zero
End Sub



Public Sub CheckInput()
Dim ersterAbsatz As Integer
Dim zweiterAbsatz As Integer
Dim i As Integer
Dim Leer As Integer
Dim Kommentar As Boolean

JUMP = False
VorNull = ""
VorNull2 = ""
COMMAND = ""
HELP = ""

LINE = frmMain.List1.ListCount

If Mid(frmMain.txtEingabe.Text, 1, 1) = ";" Then Kommentar = True

If InStr(1, frmMain.txtEingabe.Text, " ") = 0 Then Leer = 3
If InStr(1, frmMain.txtEingabe.Text, " ") = 3 Then Leer = 2
If InStr(1, frmMain.txtEingabe.Text, " ") = 4 Then Leer = 3
If InStr(1, frmMain.txtEingabe.Text, " ") = 5 Then Leer = 4

On Error Resume Next
ersterAbsatz = (InStr(1, frmMain.txtEingabe.Text, " "))
zweiterAbsatz = (InStr(1, frmMain.txtEingabe.Text, ","))

If Kommentar = False Then
    OPCODE = Mid(frmMain.txtEingabe.Text, 1, Leer)
    OPCODE = UCase(OPCODE)
    
    
    Hi = Mid(frmMain.txtEingabe.Text, ersterAbsatz + 1, 1)
    
    'If OPCODE <> "BRK" And OPCODE <> "RTS" And OPCODE <> "NOP" And OPCODE <> "TOG" Then
        'If Hi > 8 Then
        '    MsgBox "Register higher than 8", vbOKOnly, "Error"
        '    frmMain.txtEingabe.Text = ""
        '    Exit Sub
        'End If
    'End If
    
    Lo = Mid(frmMain.txtEingabe.Text, zweiterAbsatz + 1, 1)
    If OPCODE = "LDI" Or OPCODE = "RND" Then Lo = Mid(frmMain.txtEingabe.Text, zweiterAbsatz + 1)

Else
    OPCODE = ";"
End If

If OPCODE = "JNZ" Or OPCODE = "JZ" Or OPCODE = "JE" Or OPCODE = "JLE" Or OPCODE = "JA" Or OPCODE = "JAE" Or OPCODE = "JMP" Or OPCODE = "JC" Or OPCODE = "JSR" Or OPCODE = "JNE" Or OPCODE = "JL" Then
    Hi = CLng("&H" & Mid(frmMain.txtEingabe.Text, ersterAbsatz + 1, 2))
    If Len(Hex(Hi)) < 2 Then VorNull2 = "0"
End If

If OPCODE = "MSG" Then
    Lo = Mid(frmMain.txtEingabe.Text, zweiterAbsatz + 1)
End If



If LINE < 16 Then VorNull = "0"
   
LABEL = "     "
If LINE = 0 Then LABEL = "L" & VorNull & Hex(LINE) & ": "

COMMAND = ""

Select Case OPCODE
    Case ";"
        COMMAND = frmMain.txtEingabe.Text
        LABEL = ""
    
    Case "BRK"
        COMMAND = "BRK      "
        HELP = "; Break program"
    
    Case "RND"
        COMMAND = "RND R" & Hi & "," & Lo
        HELP = "; Randomize number Rx=Ry"
    
    Case "MOV"
        COMMAND = "MOV R" & Hi & ",R" & Lo
        HELP = "; MOVE Rx=Ry"
        
    Case "ADD"
        COMMAND = "ADD R" & Hi & "   "
        HELP = "; Add Rx to R0"

    Case "MUL"
        COMMAND = "MUL R" & Hi & "   "
        HELP = "; Multiply Rx and R0"
        
    Case "MSG"
        COMMAND = "MSG A" & Hi & ",'" & Lo & "'"
        HELP = "; Message in Ax"
    
    Case "OUT"
        COMMAND = "OUT A" & Hi
        HELP = "; Display a string in Ax"
        
    Case "DIV"
        COMMAND = "DIV R" & Hi & "   "
        HELP = "; Divide Rx and R0"
    
     Case "LDI"
        COMMAND = "LDI R" & Hi & "," & Lo & " "
        HELP = "; Load immediate"

     Case "DEC"
        COMMAND = "DEC R" & Hi & "   "
        HELP = "; Decrement Rx"

     Case "INC"
        COMMAND = "INC R" & Hi & "   "
        HELP = "; Increment Rx"

    
     Case "IN"
        COMMAND = "IN R" & Hi & "    "
        HELP = "; Input Rx"


    
     Case "NOP"
        COMMAND = "NOP" & "      "
        HELP = "; No operation"

       
      Case "RET"
        COMMAND = "RET" & "      "
        HELP = "; Return from subroutine"
        
     Case "SUB"
        COMMAND = "SUB R" & Hi & "   "
        HELP = "; Subtract Rx from R0"

     Case "TOG"
        COMMAND = "TOG" & "      "
        HELP = "; Toggle LED"
      
    Case "PUSH"
        COMMAND = "PUSH R" & Hi & "  "
        HELP = "; Push Rx to stack"
    
    Case "POP"
        COMMAND = "POP R" & Hi
        HELP = "; Pop Rx from stack"
        
     Case "CMP"
        COMMAND = "CMP R" & Hi & ",R" & Lo
        HELP = "; Compare Rx=Ry"
    
     Case "JE"
        COMMAND = "JE L" & VorNull2 & Hex(Hi) & "   "
        HELP = "; Jump if equal"
        JUMP = True

     Case "JNE"
        COMMAND = "JNE L" & VorNull2 & Hex(Hi) & "  "
        HELP = "; Jump if not equal"
        JUMP = True
     
    Case "JL"
        COMMAND = "JL L" & VorNull2 & Hex(Hi) & "  "
        HELP = "; Jump if lower"
        JUMP = True
     
    Case "JLE"
        COMMAND = "JLE L" & VorNull2 & Hex(Hi) & "  "
        HELP = "; Jump if lower or equal"
        JUMP = True
     
    Case "JA"
        COMMAND = "JA L" & VorNull2 & Hex(Hi) & "  "
        HELP = "; Jump if above"
        JUMP = True
     
    Case "JAE"
        COMMAND = "JAE L" & VorNull2 & Hex(Hi) & "  "
        HELP = "; Jump if above or equal"
        JUMP = True
     
     
     
     Case "JZ"
        COMMAND = "JZ L" & VorNull2 & Hex(Hi) & "   "
        HELP = "; Jump if zero"
        JUMP = True

     Case "JC"
        COMMAND = "JC L" & VorNull2 & Hex(Hi) & "   "
        HELP = "; Jump if carry set"
        JUMP = True
    
     Case "JSR"
        COMMAND = "JSR L" & VorNull2 & Hex(Hi)
        HELP = "; Jump to subroutine"
        JUMP = True
    
    
    Case "JNZ"
        COMMAND = "JNZ L" & VorNull2 & Hex(Hi) & "  "
        HELP = "; Jump if not zero"
        JUMP = True
    
    Case "JMP"
        COMMAND = "JMP L" & VorNull2 & Hex(Hi) & "  "
        HELP = "; Jump"
        JUMP = True
    
    
    
    Case Else
        If Kommentar = False Then
            MsgBox "Undefined opcode!", vbOKOnly, "Error"
            frmMain.txtEingabe.Text = ""
            Exit Sub
        End If
End Select




           If Selected = False Then
                If ViewHelp = True Then
                    frmMain.List1.AddItem VorNull & Hex(LINE) & " " & LABEL & COMMAND & "     " & HELP
                Else
                    frmMain.List1.AddItem VorNull & Hex(LINE) & " " & LABEL & COMMAND & " "
                End If
                
                'LINE = LINE + 1

            Else
                frmMain.List1.RemoveItem (Indexx)
                If Indexx < 16 Then VorNull = "0"
                If ViewHelp = True Then
                    frmMain.List1.AddItem VorNull & Hex(Indexx) & " " & LABEL & COMMAND & "     " & HELP, Indexx
                Else
                    frmMain.List1.AddItem VorNull & Hex(Indexx) & " " & LABEL & COMMAND & " ", Indexx
                End If
                Selected = False
                Indexx = 0
            End If
     

   
If JUMP = True Then
    'If (Hi) < 10 Then VorNull2 = "0"
    If Len(Hex(Hi)) < 2 Then VorNull2 = "0"
    TempZeile = frmMain.List1.List(Hi)
   
    TempAnfang = Mid(TempZeile, 1, 3)
    TempMitte = "L" & VorNull2 & Hex(Hi) & ":"
    TempEnde = Mid(TempZeile, 8)

    frmMain.List1.RemoveItem (Hi)
    frmMain.List1.AddItem TempAnfang & TempMitte & TempEnde, Int(Hi)

    JUMP = False
End If



frmMain.txtEingabe = ""
End Sub
