Option Strict Off
Option Explicit On
Module modCout
	
	Sub Cout(ByRef s As String)
		'Display output on screen
		Const MAXTEXTBOX As Short = 2048
        Dim CForm As frmCout
		
		On Error GoTo errCout
		CForm = frmCout.DefInstance
        If s = "CLS" Then
            CForm.txtDebug.Text = ""
        Else
            If CForm.txtDebug.Text.Length + s.Length > MAXTEXTBOX Then
                CForm.txtDebug.Text = Right(CForm.txtDebug.Text & s, MAXTEXTBOX)
            Else
                CForm.txtDebug.Text = CForm.txtDebug.Text & s
            End If
            System.Windows.Forms.SendKeys.Send("^{END}")
        End If
exitCout:
        Exit Sub
errCout:
        MsgBox("Cout error: " & Err.Description) ' & " " & Err.Number)
        Stop 'Resume exitCout
	End Sub
End Module