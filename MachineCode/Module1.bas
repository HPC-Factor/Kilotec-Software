Attribute VB_Name = "Module1"
Option Explicit


Public Hi As Integer
Public Lo As Integer

Public CHAR As Integer
Public LINE As Integer
Public VorNull As String
Public COMMAND As String
Public JUMP As Boolean
Public HELP As String
Public LABEL As Integer


Public Hints As Integer
Public TimerSpeed As Integer
Public Layout As Integer

Public TOGGLE As Boolean

Public Indexx As Integer

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
Public pc As Integer
Public sp As Integer
Public err As Boolean
Public zero As Boolean
Public carry As Boolean
Public Bbyte As Integer
Public r As Integer
Public Enter As Integer
Public Sprung As Integer




Public Sub ErrorMsg(ByVal s As String)
MsgBox "Error: " & s, vbOKOnly, "Error"
err = True
End Sub


Public Sub Eingabe()
'48 = 0     49 = 1
'50 = 2     51 = 3
'52 = 4     53 = 5
'54 = 6     55 = 7
'56 = 8     57 = 9
'65 = A     66 = B
'67 = C     68 = D
'69 = E     70 = F

    Dim i As Integer

    JUMP = False
    VorNull = ""
    COMMAND = ""
    HELP = ""
    If LINE < 16 Then VorNull = "0"
    
    LABEL = "     "
    If LINE = 0 Then LABEL = "L" & VorNull & Hex(LINE) & ": "
    
    On Error Resume Next
    
        If Asc(Hi) = 48 And Asc(Lo) = 48 Then
            COMMAND = "BRK      "
            HELP = "; Break program"
        ElseIf Asc(Hi) = 49 And Asc(Lo) = 49 Then
            COMMAND = "NOP      "
            HELP = "; No operation"
        ElseIf Asc(Hi) = 54 And Asc(Lo) = 54 Then
            COMMAND = "TOG      "
            HELP = "(Toggle LED)"
        ElseIf Asc(Hi) = 55 And Asc(Lo) = 55 Then
            COMMAND = "RTS      "
            HELP = "; Return from sub"
        ElseIf Asc(Hi) < 56 And Asc(Lo) < 56 Then
            COMMAND = "MOV R" & Hex(Hi) & ", R" & Hex(Lo)
            HELP = "; MOVE Rx=Ry"
        ElseIf Asc(Hi) < 56 Then
            If Asc(Lo) = 56 Then
                COMMAND = "PSH R"
                HELP = "; PUSH Rx onto stack"
            ElseIf Asc(Lo) = 57 Then
                COMMAND = "POP R"
                HELP = "; POP Rx from stack"
            ElseIf Asc(Lo) = 65 Then
                COMMAND = "ADD R"
                HELP = "; Add Rx to R0"
            ElseIf Asc(Lo) = 66 Then
                COMMAND = "SUB R"
                HELP = "; Subtract Rx from R0"
            ElseIf Asc(Lo) = 67 Then
                COMMAND = "CPY R"
                HELP = "; Copy const x into R0"
            ElseIf Asc(Lo) = 68 Then
                COMMAND = "DEC R"
                HELP = "; Decrement Rx"
            ElseIf Asc(Lo) = 69 Then
                COMMAND = "ENT R"
                HELP = "; Enter Rx"
            ElseIf Asc(Lo) = 70 Then
                COMMAND = "??? R"
            End If
            COMMAND = COMMAND & Hex(Hi) & "   "
        ElseIf Asc(Hi) = 56 Or Asc(Hi) = 57 Then
            COMMAND = "JZ "
            HELP = "; Jump if zero"
            COMMAND = COMMAND & " " & "L" & VorNull & Hex(Lo) & "  "
            JUMP = True
        ElseIf Asc(Hi) = 65 Or Asc(Hi) = 66 Then
            COMMAND = "JNZ"
            HELP = "; Jump if not zero"
            COMMAND = COMMAND & " " & "L" & VorNull & Hex(Lo) & "  "
            JUMP = True
        ElseIf Asc(Hi) = 67 Or Asc(Hi) = 68 Then
            COMMAND = "JC "
            HELP = "; Jump if carry set"
            COMMAND = COMMAND & " " & "L" & VorNull & Hex(Lo) & "  "
            JUMP = True
        Else
            COMMAND = "JSR"
            COMMAND = COMMAND & " " & "L" & VorNull & Hex(Lo) & "  "
            HELP = "; Jump to sub"
            JUMP = True
        End If
    
        
            If Selected = False Then
                If ViewHelp = True Then
                    frmMain.List1.AddItem VorNull & Hex(LINE) & " " & Hi & Lo & " " & LABEL & COMMAND & "     " & HELP
                Else
                    frmMain.List1.AddItem VorNull & Hex(LINE) & " " & Hi & Lo & " " & LABEL & COMMAND & "     "
                End If
                
                LINE = LINE + 1
            Else
                frmMain.List1.RemoveItem (Indexx)
                If ViewHelp = True Then
                    frmMain.List1.AddItem VorNull & Hex(Indexx) & " " & Hi & Lo & " " & LABEL & COMMAND & "     " & HELP, Indexx
                Else
                    frmMain.List1.AddItem VorNull & Hex(Indexx) & " " & Hi & Lo & " " & LABEL & COMMAND & "     ", Indexx
                End If
                Selected = False
                Indexx = 0
            End If
     

   
If JUMP = True Then

    TempZeile = frmMain.List1.List(Int(VorNull & Hex(Lo)))
    TempAnfang = Mid(TempZeile, 1, 6)
    TempMitte = "L" & VorNull & Hex(Lo) & ":"
    TempEnde = Mid(TempZeile, 11)

    frmMain.List1.RemoveItem (Int(VorNull & Hex(Lo)))
    frmMain.List1.AddItem TempAnfang & TempMitte & TempEnde, Int(VorNull & Hex(Lo))

    JUMP = False
End If

End Sub




Public Sub Run()
'48 = 0     49 = 1
'50 = 2     51 = 3
'52 = 4     53 = 5
'54 = 6     55 = 7
'56 = 8     57 = 9
'65 = A     66 = B
'67 = C     68 = D
'69 = E     70 = F
pc = pc + 1

If pc > frmMain.List1.ListCount Then
    pc = 0
    frmMain.Timer1.Enabled = False
    frmOutput.Show
    Exit Sub
End If
           
Hi = Mid(frmMain.List1.List(pc - 1), 4, 1)
Lo = Mid(frmMain.List1.List(pc - 1), 5, 1)

Bbyte = mem(pc)
Sprung = Bbyte
If Asc(Hi) = 48 And Asc(Lo) = 48 Then
    ErrorMsg ("BREAK")
ElseIf Asc(Hi) = 54 And Asc(Lo) = 54 Then
    If TOGGLE = True Then
        TOGGLE = False
    Else
        TOGGLE = True
    End If
ElseIf Asc(Hi) = 55 And Asc(Lo) = 55 Then
    If sp = 0 Then
        ErrorMsg ("STACK <")
    Else
        sp = sp - 1
        pc = stk(sp)
    End If
ElseIf Asc(Hi) < 56 And Asc(Lo) < 56 Then
    reg(Hi) = reg(Lo)
ElseIf Asc(Hi) < 56 Then
    If Asc(Lo) = 56 Then
        If sp >= (StackSize - 1) Then
            ErrorMsg ("STACK >")
        Else
            sp = sp + 1
            stk(sp) = reg(Hi)
        End If
    ElseIf Asc(Lo) = 57 Then
        If sp = 0 Then
            ErrorMsg ("STACK <")
        Else
            sp = sp - 1
            reg(Hi) = stk(sp)
        End If
    ElseIf Asc(Lo) = 65 Then                            'ADD
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
    
    ElseIf Asc(Lo) = 66 Then                            'SUB
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
    
    
    ElseIf Asc(Lo) = 67 Then
        reg(0) = Hi
    ElseIf Asc(Lo) = 68 Then
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
    ElseIf Asc(Lo) = 69 Then

            frmMain.Timer1.Enabled = False
            frmEnter.Label1.Caption = "R" & Hi & "= "
            frmEnter.Show

    End If

' JUMPS
ElseIf Asc(Hi) = 56 Or Asc(Hi) = 57 Then
            If Asc(Hi) = 57 Then
                If zero = True Then pc = Lo
            End If
        ElseIf Asc(Hi) = 65 Or Asc(Hi) = 66 Then        'JNZ
            If zero = False Then
                pc = Lo
            End If
        ElseIf Asc(Hi) = 67 Or Asc(Hi) = 68 Then        'JC
            If carry = True Then
                pc = Lo
            End If
        Else
            If sp >= StackSize - 1 Then                 'JSR
                ErrorMsg ("STACK >")
                sp = sp + 1
                stk(sp) = pc
                pc = Lo
            End If
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
frmOutput.lstOutput.AddItem ""
frmOutput.lstOutput.AddItem "Line (adr):." & pc & " (" & Hex(pc) & ")"
frmOutput.lstOutput.AddItem "Stack:......" & sp
frmOutput.lstOutput.AddItem "Carry flag:." & carry
frmOutput.lstOutput.AddItem "Zero flag:.." & zero

'If carry = True Then frmMain.Timer1.Enabled = False
End Sub



