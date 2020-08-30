Option Strict Off
Option Explicit On
Friend Class frmCout
	Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
	Public Sub New()
		MyBase.New()
		If m_vb6FormDefInstance Is Nothing Then
			If m_InitializingDefInstance Then
				m_vb6FormDefInstance = Me
			Else
				Try 
					'For the start-up form, the first instance created is the default instance.
					If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
						m_vb6FormDefInstance = Me
					End If
				Catch
				End Try
			End If
		End If
		'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents txtDebug As System.Windows.Forms.TextBox
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCout))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtDebug = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'txtDebug
        '
        Me.txtDebug.AcceptsReturn = True
        Me.txtDebug.BackColor = System.Drawing.SystemColors.Window
        Me.txtDebug.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDebug.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtDebug.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDebug.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDebug.Location = New System.Drawing.Point(0, 0)
        Me.txtDebug.MaxLength = 0
        Me.txtDebug.Multiline = True
        Me.txtDebug.Name = "txtDebug"
        Me.txtDebug.ReadOnly = True
        Me.txtDebug.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDebug.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtDebug.Size = New System.Drawing.Size(312, 213)
        Me.txtDebug.TabIndex = 0
        Me.txtDebug.WordWrap = False
        '
        'frmCout
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(312, 213)
        Me.Controls.Add(Me.txtDebug)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(4, 23)
        Me.Name = "frmCout"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Console"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
#Region "Upgrade Support "
    Private Shared m_vb6FormDefInstance As frmCout
	Private Shared m_InitializingDefInstance As Boolean
	Public Shared Property DefInstance() As frmCout
		Get
			If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
				m_InitializingDefInstance = True
				m_vb6FormDefInstance = New frmCout()
				m_InitializingDefInstance = False
			End If
			DefInstance = m_vb6FormDefInstance
		End Get
		Set
			m_vb6FormDefInstance = Value
		End Set
	End Property
#End Region 

End Class