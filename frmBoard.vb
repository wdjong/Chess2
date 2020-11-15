Option Strict Off
Option Explicit On
'(Dev notes are in the Chess.vb file)

Friend Class FrmBoard
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
    Public WithEvents Piece_25 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_32 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_31 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_30 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_29 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_28 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_27 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_26 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_24 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_23 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_22 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_21 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_20 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_19 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_18 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_17 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_16 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_15 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_14 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_13 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_12 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_11 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_10 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_9 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_8 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_1 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_5 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_4 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_6 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_3 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_7 As System.Windows.Forms.PictureBox
    Public WithEvents Piece_2 As System.Windows.Forms.PictureBox
    Public WithEvents Piece As System.Windows.Forms.PictureBox 'Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Public WithEvents Dragger As System.Windows.Forms.PictureBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents StatFrom As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatTo As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents MnuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents MnuFile As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnuFileNew As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnuFileExit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnuEdit As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnuTools As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnuHelp As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnuEditUndo As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnuToolsHistory As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnuToolsMove As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnuToolsTopHuman As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnuToolsBottomHuman As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MnuHelpInstructions As ToolStripMenuItem
    Friend WithEvents OptionsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmBoard))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.StatFrom = New System.Windows.Forms.ToolStripStatusLabel()
        Me.StatTo = New System.Windows.Forms.ToolStripStatusLabel()
        Me.MnuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.MnuFile = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnuFileNew = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnuFileExit = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnuEdit = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnuEditUndo = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnuTools = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnuToolsHistory = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnuToolsMove = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnuToolsTopHuman = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnuToolsBottomHuman = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnuHelp = New System.Windows.Forms.ToolStripMenuItem()
        Me.MnuHelpInstructions = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Dragger = New System.Windows.Forms.PictureBox()
        Me.Piece_25 = New System.Windows.Forms.PictureBox()
        Me.Piece_32 = New System.Windows.Forms.PictureBox()
        Me.Piece_31 = New System.Windows.Forms.PictureBox()
        Me.Piece_30 = New System.Windows.Forms.PictureBox()
        Me.Piece_29 = New System.Windows.Forms.PictureBox()
        Me.Piece_28 = New System.Windows.Forms.PictureBox()
        Me.Piece_27 = New System.Windows.Forms.PictureBox()
        Me.Piece_26 = New System.Windows.Forms.PictureBox()
        Me.Piece_24 = New System.Windows.Forms.PictureBox()
        Me.Piece_23 = New System.Windows.Forms.PictureBox()
        Me.Piece_22 = New System.Windows.Forms.PictureBox()
        Me.Piece_21 = New System.Windows.Forms.PictureBox()
        Me.Piece_20 = New System.Windows.Forms.PictureBox()
        Me.Piece_19 = New System.Windows.Forms.PictureBox()
        Me.Piece_18 = New System.Windows.Forms.PictureBox()
        Me.Piece_17 = New System.Windows.Forms.PictureBox()
        Me.Piece_16 = New System.Windows.Forms.PictureBox()
        Me.Piece_15 = New System.Windows.Forms.PictureBox()
        Me.Piece_14 = New System.Windows.Forms.PictureBox()
        Me.Piece_13 = New System.Windows.Forms.PictureBox()
        Me.Piece_12 = New System.Windows.Forms.PictureBox()
        Me.Piece_11 = New System.Windows.Forms.PictureBox()
        Me.Piece_10 = New System.Windows.Forms.PictureBox()
        Me.Piece_9 = New System.Windows.Forms.PictureBox()
        Me.Piece_8 = New System.Windows.Forms.PictureBox()
        Me.Piece_1 = New System.Windows.Forms.PictureBox()
        Me.Piece_5 = New System.Windows.Forms.PictureBox()
        Me.Piece_4 = New System.Windows.Forms.PictureBox()
        Me.Piece_6 = New System.Windows.Forms.PictureBox()
        Me.Piece_3 = New System.Windows.Forms.PictureBox()
        Me.Piece_7 = New System.Windows.Forms.PictureBox()
        Me.Piece_2 = New System.Windows.Forms.PictureBox()
        Me.OptionsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.MnuStrip1.SuspendLayout()
        CType(Me.Dragger, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_25, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_32, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_31, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_30, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_29, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_28, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_27, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_26, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_24, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_23, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_22, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_21, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_20, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_19, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_18, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_17, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_16, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_15, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_14, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_13, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_12, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_11, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_10, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_9, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Piece_2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem1})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(132, 36)
        Me.ContextMenuStrip1.Tag = ""
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(131, 32)
        Me.ToolStripMenuItem1.Text = "Castle"
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'StatusStrip1
        '
        Me.StatusStrip1.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.StatusStrip1.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.StatFrom, Me.StatTo})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 298)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(318, 32)
        Me.StatusStrip1.TabIndex = 36
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'StatFrom
        '
        Me.StatFrom.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.StatFrom.Name = "StatFrom"
        Me.StatFrom.Size = New System.Drawing.Size(54, 25)
        Me.StatFrom.Text = "From"
        '
        'StatTo
        '
        Me.StatTo.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.StatTo.Name = "StatTo"
        Me.StatTo.Size = New System.Drawing.Size(30, 25)
        Me.StatTo.Text = "To"
        '
        'MnuStrip1
        '
        Me.MnuStrip1.GripMargin = New System.Windows.Forms.Padding(2, 2, 0, 2)
        Me.MnuStrip1.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.MnuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnuFile, Me.MnuEdit, Me.MnuTools, Me.MnuHelp})
        Me.MnuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MnuStrip1.Name = "MnuStrip1"
        Me.MnuStrip1.Size = New System.Drawing.Size(318, 33)
        Me.MnuStrip1.TabIndex = 37
        Me.MnuStrip1.Text = "MenuStrip1"
        '
        'MnuFile
        '
        Me.MnuFile.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnuFileNew, Me.MnuFileExit})
        Me.MnuFile.Name = "MnuFile"
        Me.MnuFile.Size = New System.Drawing.Size(54, 29)
        Me.MnuFile.Text = "&File"
        '
        'MnuFileNew
        '
        Me.MnuFileNew.Name = "MnuFileNew"
        Me.MnuFileNew.Size = New System.Drawing.Size(270, 34)
        Me.MnuFileNew.Text = "&New"
        '
        'MnuFileExit
        '
        Me.MnuFileExit.Name = "MnuFileExit"
        Me.MnuFileExit.Size = New System.Drawing.Size(270, 34)
        Me.MnuFileExit.Text = "E&xit"
        '
        'MnuEdit
        '
        Me.MnuEdit.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnuEditUndo})
        Me.MnuEdit.Name = "MnuEdit"
        Me.MnuEdit.Size = New System.Drawing.Size(58, 29)
        Me.MnuEdit.Text = "&Edit"
        '
        'MnuEditUndo
        '
        Me.MnuEditUndo.Name = "MnuEditUndo"
        Me.MnuEditUndo.Size = New System.Drawing.Size(270, 34)
        Me.MnuEditUndo.Text = "&Undo"
        '
        'MnuTools
        '
        Me.MnuTools.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnuToolsHistory, Me.MnuToolsMove, Me.MnuToolsTopHuman, Me.MnuToolsBottomHuman, Me.OptionsToolStripMenuItem})
        Me.MnuTools.Name = "MnuTools"
        Me.MnuTools.Size = New System.Drawing.Size(69, 29)
        Me.MnuTools.Text = "&Tools"
        '
        'MnuToolsHistory
        '
        Me.MnuToolsHistory.Name = "MnuToolsHistory"
        Me.MnuToolsHistory.Size = New System.Drawing.Size(270, 34)
        Me.MnuToolsHistory.Text = "&History"
        '
        'MnuToolsMove
        '
        Me.MnuToolsMove.Name = "MnuToolsMove"
        Me.MnuToolsMove.Size = New System.Drawing.Size(270, 34)
        Me.MnuToolsMove.Text = "&Move"
        '
        'MnuToolsTopHuman
        '
        Me.MnuToolsTopHuman.Name = "MnuToolsTopHuman"
        Me.MnuToolsTopHuman.Size = New System.Drawing.Size(270, 34)
        Me.MnuToolsTopHuman.Text = "&Top Human"
        '
        'MnuToolsBottomHuman
        '
        Me.MnuToolsBottomHuman.Name = "MnuToolsBottomHuman"
        Me.MnuToolsBottomHuman.Size = New System.Drawing.Size(270, 34)
        Me.MnuToolsBottomHuman.Text = "&Bottom Human"
        '
        'MnuHelp
        '
        Me.MnuHelp.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MnuHelpInstructions, Me.AboutToolStripMenuItem})
        Me.MnuHelp.Name = "MnuHelp"
        Me.MnuHelp.Size = New System.Drawing.Size(65, 29)
        Me.MnuHelp.Text = "&Help"
        '
        'MnuHelpInstructions
        '
        Me.MnuHelpInstructions.Name = "MnuHelpInstructions"
        Me.MnuHelpInstructions.Size = New System.Drawing.Size(270, 34)
        Me.MnuHelpInstructions.Text = "&Instructions"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(270, 34)
        Me.AboutToolStripMenuItem.Text = "&About"
        '
        'Dragger
        '
        Me.Dragger.BackColor = System.Drawing.Color.Transparent
        Me.Dragger.Cursor = System.Windows.Forms.Cursors.Default
        Me.Dragger.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Dragger.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Dragger.Image = Global.Chess2.My.Resources.Resources.pawn
        Me.Dragger.Location = New System.Drawing.Point(229, 215)
        Me.Dragger.Name = "Dragger"
        Me.Dragger.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Dragger.Size = New System.Drawing.Size(32, 31)
        Me.Dragger.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.Dragger.TabIndex = 35
        Me.Dragger.TabStop = False
        Me.Dragger.Tag = "20"
        Me.Dragger.Visible = False
        '
        'Piece_25
        '
        Me.Piece_25.BackColor = System.Drawing.Color.Transparent
        Me.Piece_25.ContextMenuStrip = Me.ContextMenuStrip1
        Me.Piece_25.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_25.ErrorImage = CType(resources.GetObject("Piece_25.ErrorImage"), System.Drawing.Image)
        Me.Piece_25.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_25.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_25.Image = Global.Chess2.My.Resources.Resources.rook1
        Me.Piece_25.InitialImage = CType(resources.GetObject("Piece_25.InitialImage"), System.Drawing.Image)
        Me.Piece_25.Location = New System.Drawing.Point(51, 35)
        Me.Piece_25.Name = "Piece_25"
        Me.Piece_25.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_25.Size = New System.Drawing.Size(51, 45)
        Me.Piece_25.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_25.TabIndex = 25
        Me.Piece_25.TabStop = False
        Me.Piece_25.Tag = "25"
        '
        'Piece_32
        '
        Me.Piece_32.BackColor = System.Drawing.Color.Transparent
        Me.Piece_32.ContextMenuStrip = Me.ContextMenuStrip1
        Me.Piece_32.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_32.ErrorImage = CType(resources.GetObject("Piece_32.ErrorImage"), System.Drawing.Image)
        Me.Piece_32.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_32.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_32.Image = Global.Chess2.My.Resources.Resources.rook1
        Me.Piece_32.InitialImage = CType(resources.GetObject("Piece_32.InitialImage"), System.Drawing.Image)
        Me.Piece_32.Location = New System.Drawing.Point(410, 35)
        Me.Piece_32.Name = "Piece_32"
        Me.Piece_32.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_32.Size = New System.Drawing.Size(51, 45)
        Me.Piece_32.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_32.TabIndex = 32
        Me.Piece_32.TabStop = False
        Me.Piece_32.Tag = "32"
        '
        'Piece_31
        '
        Me.Piece_31.BackColor = System.Drawing.Color.Transparent
        Me.Piece_31.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_31.ErrorImage = CType(resources.GetObject("Piece_31.ErrorImage"), System.Drawing.Image)
        Me.Piece_31.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_31.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_31.Image = Global.Chess2.My.Resources.Resources.knight1
        Me.Piece_31.InitialImage = CType(resources.GetObject("Piece_31.InitialImage"), System.Drawing.Image)
        Me.Piece_31.Location = New System.Drawing.Point(358, 35)
        Me.Piece_31.Name = "Piece_31"
        Me.Piece_31.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_31.Size = New System.Drawing.Size(52, 45)
        Me.Piece_31.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_31.TabIndex = 31
        Me.Piece_31.TabStop = False
        Me.Piece_31.Tag = "31"
        '
        'Piece_30
        '
        Me.Piece_30.BackColor = System.Drawing.Color.Transparent
        Me.Piece_30.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_30.ErrorImage = CType(resources.GetObject("Piece_30.ErrorImage"), System.Drawing.Image)
        Me.Piece_30.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_30.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_30.Image = Global.Chess2.My.Resources.Resources.bishop1
        Me.Piece_30.InitialImage = CType(resources.GetObject("Piece_30.InitialImage"), System.Drawing.Image)
        Me.Piece_30.Location = New System.Drawing.Point(307, 35)
        Me.Piece_30.Name = "Piece_30"
        Me.Piece_30.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_30.Size = New System.Drawing.Size(51, 45)
        Me.Piece_30.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_30.TabIndex = 30
        Me.Piece_30.TabStop = False
        Me.Piece_30.Tag = "30"
        '
        'Piece_29
        '
        Me.Piece_29.BackColor = System.Drawing.Color.Transparent
        Me.Piece_29.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_29.ErrorImage = CType(resources.GetObject("Piece_29.ErrorImage"), System.Drawing.Image)
        Me.Piece_29.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_29.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_29.Image = Global.Chess2.My.Resources.Resources.king1
        Me.Piece_29.InitialImage = CType(resources.GetObject("Piece_29.InitialImage"), System.Drawing.Image)
        Me.Piece_29.Location = New System.Drawing.Point(256, 35)
        Me.Piece_29.Name = "Piece_29"
        Me.Piece_29.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_29.Size = New System.Drawing.Size(51, 45)
        Me.Piece_29.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_29.TabIndex = 29
        Me.Piece_29.TabStop = False
        Me.Piece_29.Tag = "29"
        '
        'Piece_28
        '
        Me.Piece_28.BackColor = System.Drawing.Color.Transparent
        Me.Piece_28.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_28.ErrorImage = CType(resources.GetObject("Piece_28.ErrorImage"), System.Drawing.Image)
        Me.Piece_28.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_28.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_28.Image = Global.Chess2.My.Resources.Resources.queen1
        Me.Piece_28.InitialImage = CType(resources.GetObject("Piece_28.InitialImage"), System.Drawing.Image)
        Me.Piece_28.Location = New System.Drawing.Point(205, 35)
        Me.Piece_28.Name = "Piece_28"
        Me.Piece_28.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_28.Size = New System.Drawing.Size(51, 45)
        Me.Piece_28.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_28.TabIndex = 28
        Me.Piece_28.TabStop = False
        Me.Piece_28.Tag = "28"
        '
        'Piece_27
        '
        Me.Piece_27.BackColor = System.Drawing.Color.Transparent
        Me.Piece_27.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_27.ErrorImage = CType(resources.GetObject("Piece_27.ErrorImage"), System.Drawing.Image)
        Me.Piece_27.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_27.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_27.Image = Global.Chess2.My.Resources.Resources.bishop1
        Me.Piece_27.InitialImage = CType(resources.GetObject("Piece_27.InitialImage"), System.Drawing.Image)
        Me.Piece_27.Location = New System.Drawing.Point(154, 35)
        Me.Piece_27.Name = "Piece_27"
        Me.Piece_27.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_27.Size = New System.Drawing.Size(51, 45)
        Me.Piece_27.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_27.TabIndex = 27
        Me.Piece_27.TabStop = False
        Me.Piece_27.Tag = "27"
        '
        'Piece_26
        '
        Me.Piece_26.BackColor = System.Drawing.Color.Transparent
        Me.Piece_26.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_26.ErrorImage = CType(resources.GetObject("Piece_26.ErrorImage"), System.Drawing.Image)
        Me.Piece_26.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_26.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_26.Image = Global.Chess2.My.Resources.Resources.knight1
        Me.Piece_26.InitialImage = CType(resources.GetObject("Piece_26.InitialImage"), System.Drawing.Image)
        Me.Piece_26.Location = New System.Drawing.Point(102, 35)
        Me.Piece_26.Name = "Piece_26"
        Me.Piece_26.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_26.Size = New System.Drawing.Size(52, 45)
        Me.Piece_26.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_26.TabIndex = 26
        Me.Piece_26.TabStop = False
        Me.Piece_26.Tag = "26"
        '
        'Piece_24
        '
        Me.Piece_24.BackColor = System.Drawing.Color.Transparent
        Me.Piece_24.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_24.ErrorImage = CType(resources.GetObject("Piece_24.ErrorImage"), System.Drawing.Image)
        Me.Piece_24.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_24.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_24.Image = Global.Chess2.My.Resources.Resources.pawn1
        Me.Piece_24.InitialImage = CType(resources.GetObject("Piece_24.InitialImage"), System.Drawing.Image)
        Me.Piece_24.Location = New System.Drawing.Point(410, 82)
        Me.Piece_24.Name = "Piece_24"
        Me.Piece_24.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_24.Size = New System.Drawing.Size(51, 45)
        Me.Piece_24.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_24.TabIndex = 24
        Me.Piece_24.TabStop = False
        Me.Piece_24.Tag = "24"
        Me.Piece_24.Visible = False
        '
        'Piece_23
        '
        Me.Piece_23.BackColor = System.Drawing.Color.Transparent
        Me.Piece_23.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_23.ErrorImage = CType(resources.GetObject("Piece_23.ErrorImage"), System.Drawing.Image)
        Me.Piece_23.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_23.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_23.Image = Global.Chess2.My.Resources.Resources.pawn1
        Me.Piece_23.InitialImage = CType(resources.GetObject("Piece_23.InitialImage"), System.Drawing.Image)
        Me.Piece_23.Location = New System.Drawing.Point(358, 82)
        Me.Piece_23.Name = "Piece_23"
        Me.Piece_23.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_23.Size = New System.Drawing.Size(52, 45)
        Me.Piece_23.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_23.TabIndex = 23
        Me.Piece_23.TabStop = False
        Me.Piece_23.Tag = "23"
        Me.Piece_23.Visible = False
        '
        'Piece_22
        '
        Me.Piece_22.BackColor = System.Drawing.Color.Transparent
        Me.Piece_22.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_22.ErrorImage = CType(resources.GetObject("Piece_22.ErrorImage"), System.Drawing.Image)
        Me.Piece_22.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_22.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_22.Image = Global.Chess2.My.Resources.Resources.pawn1
        Me.Piece_22.InitialImage = CType(resources.GetObject("Piece_22.InitialImage"), System.Drawing.Image)
        Me.Piece_22.Location = New System.Drawing.Point(307, 82)
        Me.Piece_22.Name = "Piece_22"
        Me.Piece_22.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_22.Size = New System.Drawing.Size(51, 45)
        Me.Piece_22.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_22.TabIndex = 22
        Me.Piece_22.TabStop = False
        Me.Piece_22.Tag = "22"
        Me.Piece_22.Visible = False
        '
        'Piece_21
        '
        Me.Piece_21.BackColor = System.Drawing.Color.Transparent
        Me.Piece_21.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_21.ErrorImage = CType(resources.GetObject("Piece_21.ErrorImage"), System.Drawing.Image)
        Me.Piece_21.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_21.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_21.Image = Global.Chess2.My.Resources.Resources.pawn1
        Me.Piece_21.InitialImage = CType(resources.GetObject("Piece_21.InitialImage"), System.Drawing.Image)
        Me.Piece_21.Location = New System.Drawing.Point(256, 82)
        Me.Piece_21.Name = "Piece_21"
        Me.Piece_21.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_21.Size = New System.Drawing.Size(51, 45)
        Me.Piece_21.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_21.TabIndex = 21
        Me.Piece_21.TabStop = False
        Me.Piece_21.Tag = "21"
        Me.Piece_21.Visible = False
        '
        'Piece_20
        '
        Me.Piece_20.BackColor = System.Drawing.Color.Transparent
        Me.Piece_20.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_20.ErrorImage = CType(resources.GetObject("Piece_20.ErrorImage"), System.Drawing.Image)
        Me.Piece_20.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_20.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_20.Image = Global.Chess2.My.Resources.Resources.pawn1
        Me.Piece_20.InitialImage = CType(resources.GetObject("Piece_20.InitialImage"), System.Drawing.Image)
        Me.Piece_20.Location = New System.Drawing.Point(205, 82)
        Me.Piece_20.Name = "Piece_20"
        Me.Piece_20.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_20.Size = New System.Drawing.Size(51, 45)
        Me.Piece_20.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_20.TabIndex = 20
        Me.Piece_20.TabStop = False
        Me.Piece_20.Tag = "20"
        Me.Piece_20.Visible = False
        '
        'Piece_19
        '
        Me.Piece_19.BackColor = System.Drawing.Color.Transparent
        Me.Piece_19.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_19.ErrorImage = CType(resources.GetObject("Piece_19.ErrorImage"), System.Drawing.Image)
        Me.Piece_19.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_19.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_19.Image = Global.Chess2.My.Resources.Resources.pawn1
        Me.Piece_19.InitialImage = CType(resources.GetObject("Piece_19.InitialImage"), System.Drawing.Image)
        Me.Piece_19.Location = New System.Drawing.Point(154, 82)
        Me.Piece_19.Name = "Piece_19"
        Me.Piece_19.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_19.Size = New System.Drawing.Size(51, 45)
        Me.Piece_19.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_19.TabIndex = 19
        Me.Piece_19.TabStop = False
        Me.Piece_19.Tag = "19"
        Me.Piece_19.Visible = False
        '
        'Piece_18
        '
        Me.Piece_18.BackColor = System.Drawing.Color.Transparent
        Me.Piece_18.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_18.ErrorImage = CType(resources.GetObject("Piece_18.ErrorImage"), System.Drawing.Image)
        Me.Piece_18.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_18.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_18.Image = Global.Chess2.My.Resources.Resources.pawn1
        Me.Piece_18.InitialImage = CType(resources.GetObject("Piece_18.InitialImage"), System.Drawing.Image)
        Me.Piece_18.Location = New System.Drawing.Point(102, 82)
        Me.Piece_18.Name = "Piece_18"
        Me.Piece_18.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_18.Size = New System.Drawing.Size(52, 45)
        Me.Piece_18.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_18.TabIndex = 18
        Me.Piece_18.TabStop = False
        Me.Piece_18.Tag = "18"
        Me.Piece_18.Visible = False
        '
        'Piece_17
        '
        Me.Piece_17.BackColor = System.Drawing.Color.Transparent
        Me.Piece_17.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_17.ErrorImage = CType(resources.GetObject("Piece_17.ErrorImage"), System.Drawing.Image)
        Me.Piece_17.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_17.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_17.Image = Global.Chess2.My.Resources.Resources.pawn1
        Me.Piece_17.InitialImage = CType(resources.GetObject("Piece_17.InitialImage"), System.Drawing.Image)
        Me.Piece_17.Location = New System.Drawing.Point(51, 82)
        Me.Piece_17.Name = "Piece_17"
        Me.Piece_17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_17.Size = New System.Drawing.Size(51, 45)
        Me.Piece_17.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_17.TabIndex = 17
        Me.Piece_17.TabStop = False
        Me.Piece_17.Tag = "17"
        Me.Piece_17.Visible = False
        '
        'Piece_16
        '
        Me.Piece_16.BackColor = System.Drawing.Color.Transparent
        Me.Piece_16.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_16.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_16.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_16.Image = Global.Chess2.My.Resources.Resources.pawn
        Me.Piece_16.Location = New System.Drawing.Point(410, 327)
        Me.Piece_16.Name = "Piece_16"
        Me.Piece_16.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_16.Size = New System.Drawing.Size(51, 46)
        Me.Piece_16.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_16.TabIndex = 16
        Me.Piece_16.TabStop = False
        Me.Piece_16.Tag = "16"
        Me.Piece_16.Visible = False
        '
        'Piece_15
        '
        Me.Piece_15.BackColor = System.Drawing.Color.Transparent
        Me.Piece_15.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_15.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_15.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_15.Image = Global.Chess2.My.Resources.Resources.pawn
        Me.Piece_15.Location = New System.Drawing.Point(358, 327)
        Me.Piece_15.Name = "Piece_15"
        Me.Piece_15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_15.Size = New System.Drawing.Size(52, 46)
        Me.Piece_15.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_15.TabIndex = 15
        Me.Piece_15.TabStop = False
        Me.Piece_15.Tag = "15"
        Me.Piece_15.Visible = False
        '
        'Piece_14
        '
        Me.Piece_14.BackColor = System.Drawing.Color.Transparent
        Me.Piece_14.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_14.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_14.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_14.Image = Global.Chess2.My.Resources.Resources.pawn
        Me.Piece_14.Location = New System.Drawing.Point(307, 327)
        Me.Piece_14.Name = "Piece_14"
        Me.Piece_14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_14.Size = New System.Drawing.Size(51, 46)
        Me.Piece_14.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_14.TabIndex = 14
        Me.Piece_14.TabStop = False
        Me.Piece_14.Tag = "14"
        Me.Piece_14.Visible = False
        '
        'Piece_13
        '
        Me.Piece_13.BackColor = System.Drawing.Color.Transparent
        Me.Piece_13.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_13.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_13.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_13.Image = Global.Chess2.My.Resources.Resources.pawn
        Me.Piece_13.Location = New System.Drawing.Point(256, 327)
        Me.Piece_13.Name = "Piece_13"
        Me.Piece_13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_13.Size = New System.Drawing.Size(51, 46)
        Me.Piece_13.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_13.TabIndex = 13
        Me.Piece_13.TabStop = False
        Me.Piece_13.Tag = "13"
        Me.Piece_13.Visible = False
        '
        'Piece_12
        '
        Me.Piece_12.BackColor = System.Drawing.Color.Transparent
        Me.Piece_12.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_12.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_12.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_12.Image = Global.Chess2.My.Resources.Resources.pawn
        Me.Piece_12.Location = New System.Drawing.Point(205, 326)
        Me.Piece_12.Name = "Piece_12"
        Me.Piece_12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_12.Size = New System.Drawing.Size(51, 47)
        Me.Piece_12.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_12.TabIndex = 12
        Me.Piece_12.TabStop = False
        Me.Piece_12.Tag = "12"
        Me.Piece_12.Visible = False
        '
        'Piece_11
        '
        Me.Piece_11.BackColor = System.Drawing.Color.Transparent
        Me.Piece_11.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_11.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_11.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_11.Image = Global.Chess2.My.Resources.Resources.pawn
        Me.Piece_11.Location = New System.Drawing.Point(154, 327)
        Me.Piece_11.Name = "Piece_11"
        Me.Piece_11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_11.Size = New System.Drawing.Size(51, 46)
        Me.Piece_11.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_11.TabIndex = 11
        Me.Piece_11.TabStop = False
        Me.Piece_11.Tag = "11"
        Me.Piece_11.Visible = False
        '
        'Piece_10
        '
        Me.Piece_10.BackColor = System.Drawing.Color.Transparent
        Me.Piece_10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_10.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_10.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_10.Image = Global.Chess2.My.Resources.Resources.pawn
        Me.Piece_10.Location = New System.Drawing.Point(102, 327)
        Me.Piece_10.Name = "Piece_10"
        Me.Piece_10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_10.Size = New System.Drawing.Size(52, 46)
        Me.Piece_10.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_10.TabIndex = 10
        Me.Piece_10.TabStop = False
        Me.Piece_10.Tag = "10"
        Me.Piece_10.Visible = False
        '
        'Piece_9
        '
        Me.Piece_9.BackColor = System.Drawing.Color.Transparent
        Me.Piece_9.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_9.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_9.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_9.Image = Global.Chess2.My.Resources.Resources.pawn
        Me.Piece_9.Location = New System.Drawing.Point(51, 327)
        Me.Piece_9.Name = "Piece_9"
        Me.Piece_9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_9.Size = New System.Drawing.Size(51, 46)
        Me.Piece_9.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_9.TabIndex = 9
        Me.Piece_9.TabStop = False
        Me.Piece_9.Tag = "9"
        Me.Piece_9.Visible = False
        '
        'Piece_8
        '
        Me.Piece_8.BackColor = System.Drawing.Color.Transparent
        Me.Piece_8.ContextMenuStrip = Me.ContextMenuStrip1
        Me.Piece_8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_8.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_8.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_8.Image = Global.Chess2.My.Resources.Resources.rook
        Me.Piece_8.Location = New System.Drawing.Point(410, 374)
        Me.Piece_8.Name = "Piece_8"
        Me.Piece_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_8.Size = New System.Drawing.Size(51, 45)
        Me.Piece_8.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_8.TabIndex = 8
        Me.Piece_8.TabStop = False
        Me.Piece_8.Tag = "8"
        Me.Piece_8.Visible = False
        '
        'Piece_1
        '
        Me.Piece_1.BackColor = System.Drawing.Color.Transparent
        Me.Piece_1.ContextMenuStrip = Me.ContextMenuStrip1
        Me.Piece_1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_1.Image = Global.Chess2.My.Resources.Resources.rook
        Me.Piece_1.Location = New System.Drawing.Point(51, 374)
        Me.Piece_1.Name = "Piece_1"
        Me.Piece_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_1.Size = New System.Drawing.Size(51, 45)
        Me.Piece_1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_1.TabIndex = 1
        Me.Piece_1.TabStop = False
        Me.Piece_1.Tag = "1"
        Me.Piece_1.Visible = False
        '
        'Piece_5
        '
        Me.Piece_5.BackColor = System.Drawing.Color.Transparent
        Me.Piece_5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_5.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_5.Image = Global.Chess2.My.Resources.Resources.king
        Me.Piece_5.Location = New System.Drawing.Point(256, 374)
        Me.Piece_5.Name = "Piece_5"
        Me.Piece_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_5.Size = New System.Drawing.Size(51, 45)
        Me.Piece_5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_5.TabIndex = 5
        Me.Piece_5.TabStop = False
        Me.Piece_5.Tag = "5"
        Me.Piece_5.Visible = False
        '
        'Piece_4
        '
        Me.Piece_4.BackColor = System.Drawing.Color.Transparent
        Me.Piece_4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_4.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_4.Image = Global.Chess2.My.Resources.Resources.queen
        Me.Piece_4.Location = New System.Drawing.Point(205, 374)
        Me.Piece_4.Name = "Piece_4"
        Me.Piece_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_4.Size = New System.Drawing.Size(51, 45)
        Me.Piece_4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_4.TabIndex = 4
        Me.Piece_4.TabStop = False
        Me.Piece_4.Tag = "4"
        Me.Piece_4.Visible = False
        '
        'Piece_6
        '
        Me.Piece_6.BackColor = System.Drawing.Color.Transparent
        Me.Piece_6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_6.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_6.Image = Global.Chess2.My.Resources.Resources.bishop
        Me.Piece_6.Location = New System.Drawing.Point(307, 374)
        Me.Piece_6.Name = "Piece_6"
        Me.Piece_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_6.Size = New System.Drawing.Size(51, 45)
        Me.Piece_6.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_6.TabIndex = 6
        Me.Piece_6.TabStop = False
        Me.Piece_6.Tag = "6"
        Me.Piece_6.Visible = False
        '
        'Piece_3
        '
        Me.Piece_3.BackColor = System.Drawing.Color.Transparent
        Me.Piece_3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_3.Image = Global.Chess2.My.Resources.Resources.bishop
        Me.Piece_3.Location = New System.Drawing.Point(154, 374)
        Me.Piece_3.Name = "Piece_3"
        Me.Piece_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_3.Size = New System.Drawing.Size(51, 45)
        Me.Piece_3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_3.TabIndex = 3
        Me.Piece_3.TabStop = False
        Me.Piece_3.Tag = "3"
        Me.Piece_3.Visible = False
        '
        'Piece_7
        '
        Me.Piece_7.BackColor = System.Drawing.Color.Transparent
        Me.Piece_7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_7.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_7.Image = Global.Chess2.My.Resources.Resources.knight
        Me.Piece_7.Location = New System.Drawing.Point(358, 374)
        Me.Piece_7.Name = "Piece_7"
        Me.Piece_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_7.Size = New System.Drawing.Size(52, 45)
        Me.Piece_7.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_7.TabIndex = 7
        Me.Piece_7.TabStop = False
        Me.Piece_7.Tag = "7"
        Me.Piece_7.Visible = False
        '
        'Piece_2
        '
        Me.Piece_2.BackColor = System.Drawing.Color.Transparent
        Me.Piece_2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Piece_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Piece_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Piece_2.Image = Global.Chess2.My.Resources.Resources.knight
        Me.Piece_2.Location = New System.Drawing.Point(102, 374)
        Me.Piece_2.Name = "Piece_2"
        Me.Piece_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Piece_2.Size = New System.Drawing.Size(52, 45)
        Me.Piece_2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.Piece_2.TabIndex = 2
        Me.Piece_2.TabStop = False
        Me.Piece_2.Tag = "2"
        Me.Piece_2.Visible = False
        '
        'OptionsToolStripMenuItem
        '
        Me.OptionsToolStripMenuItem.Name = "OptionsToolStripMenuItem"
        Me.OptionsToolStripMenuItem.Size = New System.Drawing.Size(270, 34)
        Me.OptionsToolStripMenuItem.Text = "&Options"
        '
        'FrmBoard
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(8, 19)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(318, 330)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.MnuStrip1)
        Me.Controls.Add(Me.Dragger)
        Me.Controls.Add(Me.Piece_25)
        Me.Controls.Add(Me.Piece_32)
        Me.Controls.Add(Me.Piece_31)
        Me.Controls.Add(Me.Piece_30)
        Me.Controls.Add(Me.Piece_29)
        Me.Controls.Add(Me.Piece_28)
        Me.Controls.Add(Me.Piece_27)
        Me.Controls.Add(Me.Piece_26)
        Me.Controls.Add(Me.Piece_24)
        Me.Controls.Add(Me.Piece_23)
        Me.Controls.Add(Me.Piece_22)
        Me.Controls.Add(Me.Piece_21)
        Me.Controls.Add(Me.Piece_20)
        Me.Controls.Add(Me.Piece_19)
        Me.Controls.Add(Me.Piece_18)
        Me.Controls.Add(Me.Piece_17)
        Me.Controls.Add(Me.Piece_16)
        Me.Controls.Add(Me.Piece_15)
        Me.Controls.Add(Me.Piece_14)
        Me.Controls.Add(Me.Piece_13)
        Me.Controls.Add(Me.Piece_12)
        Me.Controls.Add(Me.Piece_11)
        Me.Controls.Add(Me.Piece_10)
        Me.Controls.Add(Me.Piece_9)
        Me.Controls.Add(Me.Piece_8)
        Me.Controls.Add(Me.Piece_1)
        Me.Controls.Add(Me.Piece_5)
        Me.Controls.Add(Me.Piece_4)
        Me.Controls.Add(Me.Piece_6)
        Me.Controls.Add(Me.Piece_3)
        Me.Controls.Add(Me.Piece_7)
        Me.Controls.Add(Me.Piece_2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(20, 20)
        Me.MainMenuStrip = Me.MnuStrip1
        Me.Name = "FrmBoard"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Chess"
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.MnuStrip1.ResumeLayout(False)
        Me.MnuStrip1.PerformLayout()
        CType(Me.Dragger, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_25, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_32, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_31, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_30, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_29, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_28, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_27, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_26, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_24, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_23, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_22, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_21, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_20, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_19, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_18, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_17, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_16, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_15, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_14, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_13, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_12, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_11, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_10, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_9, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Piece_2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
#Region "Upgrade Support "
    Private Shared m_vb6FormDefInstance As FrmBoard
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As FrmBoard
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New FrmBoard()
                m_InitializingDefInstance = False
            End If
            DefInstance = m_vb6FormDefInstance
        End Get
        Set(ByVal value As FrmBoard)
            m_vb6FormDefInstance = value
        End Set
    End Property
#End Region
    'Pieces are controls on the board ranging from 0 to 31
    'aChessPiece is a corresponding array of chessPiece objects from 0 to 31
    Dim IsUDraggin As Boolean 'used to track dragging Pieces...
    Dim origX As Long 'the mouse point from which we dragged
    Dim origY As Long 'the mouse point from which we dragged

    Sub ComputerMove()
        'Find the best move and do it.
        Dim pbest As Byte 'best Piece (1..32)  where 1 is White (Red) Rook at bottom left and and 32 is Black (Blue) Rook at top right
        Dim mbest As Byte 'best move

        Do
            GetBestMove(pbest, mbest)
            Application.DoEvents()
            If pbest < 1 Or pbest > 32 Then 'there is no legal move
                If IsInCheck(Turn) Then
                    MsgBox("Check mate. You win.")
                Else 'No legal move and no check mate e.g. when there is just a king and any move would put him in check
                    MsgBox("Stalemate. We draw.")
                End If
                aGame.Playing = False
                Timer1.Enabled = False
            Else 'carry out the best move
                DoBestMove(pbest, mbest)
                Turn = -Turn 'next turn
            End If
            'If Not PlayerisHuman(0) And Not PlayerisHuman(2) Then 'When both opponents are computer's, currently pausing after each move to check action. 0 = top, 2 = bottom when it's bottoms turn turn = 1 and when it's top's turn = -1
            '    If MsgBox("pause", vbOKCancel) = vbCancel Then
            '        aGame.Playing = False
            '    End If
            'End If
        Loop While aGame.Playing And Not PlayerisHuman(1 + Turn) 'if then next turn is a computer loop
    End Sub

    Sub DrawBoard(ByVal formGraphics As System.Drawing.Graphics)
        'Draw an array of alternating colour squares: not the Pieces
        Dim xPos As Byte
        Dim yPos As Byte
        Dim myBrush As New System.Drawing.SolidBrush(System.Drawing.Color.Gray)

        For xPos = 1 To MaxX
            For yPos = 1 To MaxY
                If (xPos + yPos) Mod 2 = 0 Then
                    myBrush.Color = System.Drawing.Color.LightSlateGray 'White
                    formGraphics.FillRectangle(myBrush, xPos * PWIDTH, BBOTTOM - PHEIGHT * yPos, PWIDTH, PHEIGHT)  ', QBColor(10), B 'dotted
                Else
                    myBrush.Color = System.Drawing.Color.DarkGray  'Black 
                    formGraphics.FillRectangle(myBrush, xPos * PWIDTH, BBOTTOM - PHEIGHT * yPos, PWIDTH, PHEIGHT)  ', QBColor(10), B 'dotted
                End If
            Next yPos
        Next xPos
        myBrush.Dispose()
    End Sub

    Private Sub FrmBoard_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        aGame.Playing = True
        PlayerisHuman(1 + TOPGOINGDOWN) = False 'i.e. index = (0) 'i.e. Black/Blue at top is a computer
        PlayerisHuman(1 + BOTTOMGOINGUP) = True 'i.e. index = (2) 'i.e. White/Red at the bottom is a human (by default. See menu)
        MnuToolsTopHuman.Checked = PlayerisHuman(1 + TOPGOINGDOWN)
        MnuToolsBottomHuman.Checked = PlayerisHuman(1 + BOTTOMGOINGUP)
        Turn = BOTTOMGOINGUP 'white/red goes first
        StatusStrip1.Items("statFrom").Text = ""
        StatusStrip1.Items("statTo").Text = ""
        Cout("CLS")
        aBoard = New Board
        PHEIGHT = PWIDTH 'pixels
        PIECEWIDTH = PWIDTH * 15 '32 * 15 = 480 'twips: both Piece and square
        PIECEHEIGHT = PHEIGHT * 15 '
        BOARDBOTTOM = (MaxY + 1) * PIECEHEIGHT
        BBOTTOM = (MaxY + 1) * PHEIGHT
        Me.Width = (MaxX + 2) * PWIDTH
        Me.Height = (MaxY + 5) * PHEIGHT
        SetupPieces()
    End Sub

    Private Sub FrmBoard_Closed(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Closed
        End
    End Sub

    Private Sub HumanMove(ByRef Source As System.Windows.Forms.Control, ByRef X As Single, ByRef Y As Single)
        'What in VB6 was the source is now the sender; the control we're working with
        'Human move routine: called by mouse up at end of dragging
        'Called at the end of a dragging move when it's landed on a Piece
        'The source.tag identifies the Piece which has been moved
        Dim aMove As New Move 'create a new move object for storing the current move in 

        aMove.Clear()
        aMove.P1id = Source.Tag 'record the first Pieces move
        aMove.P1XOrig = aChessPiece(Source.Tag).XPos
        aMove.P1YOrig = aChessPiece(Source.Tag).YPos
        aMove.P1XDest = XPos(Source.Left + X) 'destX
        aMove.P1YDest = YPos(Source.Top + Y) 'destY
        aMove.P2id = aBoard.GetGBoardID(aMove.P1XDest, aMove.P1YDest)
        If aMove.P2id > 0 Then 'there's a Piece being taken
            aMove.P2XOrig = aMove.P1XDest
            aMove.P2YOrig = aMove.P1YDest
        End If
        If IsLegal1(aMove) Then
            Take(aMove.P2XOrig, aMove.P2YOrig) 'Take their Piece if it's there
            aChessPiece(aMove.P1id).Move(aMove.P1XDest, aMove.P1YDest) 'Move your Piece
            If Ispromoting(aMove.P1id, aMove.P1YDest) Then 'it realises it's reached the 8th rank
                Queening(aMove.P1id) 'change into a queen
                aMove.MCode = "q" 'in case we need to undo
            End If
            If IsInCheck(Turn) Then 'Check again for legality after the move & e.g. promoting
                MsgBox("You're in check. Either you moved into it or you didn't move out of it.")
                If aMove.MCode = "q" Then
                    UnQueen(Source.Tag)
                End If
                aChessPiece(Source.Tag).Move(aMove.P1XOrig, aMove.P1YOrig) 'move Piece back
                If aMove.P2id > 0 Then
                    aChessPiece(aMove.P2id).Move(aMove.P2XOrig, aMove.P2YOrig) 'move taken Piece back
                    ShowPiece(aMove.P2id) 'make visible again
                End If
            Else
                aGame.Add(aMove) 'add the move to the array
                Turn = -Turn
            End If
        End If
        StatusStrip1.Items("statFrom").Text = "From: " & aChessPiece(aMove.P1id).AlgNot & Chr(96 + aMove.P1XOrig) & aMove.P1YOrig
        StatusStrip1.Items("statTo").Text = "To: " & aChessPiece(aMove.P1id).AlgNot & Chr(96 + aMove.P1XDest) & aMove.P1YDest
        DoHistory(aChessPiece(aMove.P1id).AlgNot & Chr(96 + aMove.P1XDest) & aMove.P1YDest, aChessPiece(aMove.P1id).Direction)
        ShowPiece((aMove.P1id)) 'draw the Piece (if they moved the control but we didn't move the Piece position it will redraw in the original location
        Timer1.Enabled = True
    End Sub

    Private Sub ImgMousedown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Piece_1.MouseDown, Piece_2.MouseDown, Piece_3.MouseDown, Piece_4.MouseDown, Piece_5.MouseDown, Piece_6.MouseDown, Piece_7.MouseDown, Piece_8.MouseDown, Piece_9.MouseDown, Piece_10.MouseDown, Piece_11.MouseDown, Piece_12.MouseDown, Piece_13.MouseDown, Piece_14.MouseDown, Piece_15.MouseDown, Piece_16.MouseDown, Piece_17.MouseDown, Piece_18.MouseDown, Piece_19.MouseDown, Piece_20.MouseDown, Piece_21.MouseDown, Piece_22.MouseDown, Piece_23.MouseDown, Piece_24.MouseDown, Piece_25.MouseDown, Piece_26.MouseDown, Piece_27.MouseDown, Piece_28.MouseDown, Piece_29.MouseDown, Piece_30.MouseDown, Piece_31.MouseDown, Piece_32.MouseDown
        'The user commences dragging with the mousedown
        If Not IsUDraggin Then
            IsUDraggin = True
            Dragger = sender
            'Dragger.BackColor = System.Drawing.Color.Gray ' PieceArray(i).BackColor = System.Drawing.ColorTranslator.FromOle(QBColor(7))
            Dragger.Visible = True
            Dragger.BringToFront()
            origX = e.X 'used for displaying dragged Piece with same offset
            origY = e.Y 'e.x and e.y are the click position relative to the control
        End If
    End Sub

    Private Sub ImgMouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Piece_1.MouseUp, Piece_2.MouseUp, Piece_3.MouseUp, Piece_4.MouseUp, Piece_5.MouseUp, Piece_6.MouseUp, Piece_7.MouseUp, Piece_8.MouseUp, Piece_9.MouseUp, Piece_10.MouseUp, Piece_11.MouseUp, Piece_12.MouseUp, Piece_13.MouseUp, Piece_14.MouseUp, Piece_15.MouseUp, Piece_16.MouseUp, Piece_17.MouseUp, Piece_18.MouseUp, Piece_19.MouseUp, Piece_20.MouseUp, Piece_21.MouseUp, Piece_22.MouseUp, Piece_23.MouseUp, Piece_24.MouseUp, Piece_25.MouseUp, Piece_26.MouseUp, Piece_27.MouseUp, Piece_28.MouseUp, Piece_29.MouseUp, Piece_30.MouseUp, Piece_31.MouseUp, Piece_32.MouseUp
        'The user release the mouse button completing the dragging
        'I think e.x and e.y are relative to the control so the absolute position on the form is
        'sender.left + e.x and sender.top + e.y
        If IsUDraggin Then
            IsUDraggin = False
            Dragger.Visible = False
            HumanMove(sender, e.X, e.Y)
        End If
    End Sub

    Private Sub RP1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Piece_1.MouseMove, Piece_2.MouseMove, Piece_3.MouseMove, Piece_4.MouseMove, Piece_5.MouseMove, Piece_6.MouseMove, Piece_7.MouseMove, Piece_8.MouseMove, Piece_9.MouseMove, Piece_10.MouseMove, Piece_11.MouseMove, Piece_12.MouseMove, Piece_13.MouseMove, Piece_14.MouseMove, Piece_15.MouseMove, Piece_16.MouseMove, Piece_17.MouseMove, Piece_18.MouseMove, Piece_19.MouseMove, Piece_20.MouseMove, Piece_21.MouseMove, Piece_22.MouseMove, Piece_23.MouseMove, Piece_24.MouseMove, Piece_25.MouseMove, Piece_26.MouseMove, Piece_27.MouseMove, Piece_28.MouseMove, Piece_29.MouseMove, Piece_30.MouseMove, Piece_31.MouseMove, Piece_32.MouseMove
        'The user moves the mouse while dragging
        If IsUDraggin Then
            Dragger.Left = sender.left + e.X - origX 'baseX + e.X - origX
            Dragger.Top = sender.top + e.Y - origY 'baseY + e.Y - origY
        End If
    End Sub

    Public ReadOnly Property XPos(ByVal xmouse As Single) As Byte
        Get
            XPos = Int(xmouse / PWIDTH)
        End Get
    End Property

    Public ReadOnly Property YPos(ByVal ymouse As Single) As Byte
        Get
            YPos = MaxY + 1 - Int(ymouse / PHEIGHT)
        End Get
    End Property

    Private Sub FrmBoard_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        DrawBoard(e.Graphics)
    End Sub

    Sub ShowPiece(ByRef i As Short) 'I is the Piece
        'Move actual Piece (control) and display move in algebraic notation
        Dim Pieceexists As Boolean = True

        Try
            If aChessPiece(i) Is Nothing Then
                Pieceexists = False
            End If
        Catch ex As Exception
            'ignore errors: MsgBox("ShowPiece: " & ex.Message)
        End Try
        If Pieceexists Then
            If aChessPiece(i).OnBoard Then
                PieceArray(i).Visible = True
                PieceArray(i).Width = PWIDTH
                PieceArray(i).Height = PHEIGHT
                'On the form adjust the position of the Piece
                PieceArray(i).Left = TwipsToPixelsX(aChessPiece(i).XPos * PIECEWIDTH)
                PieceArray(i).Top = TwipsToPixelsY(((MaxY + 1) - aChessPiece(i).YPos) * PIECEHEIGHT)
            Else
                PieceArray(i).Visible = False
            End If
        End If
    End Sub

    Private Sub FrmBoard_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeEnd
        Dim i As Byte

        If Me.Width > Me.Height Then
            PWIDTH = Me.Height / (MaxX + 2) '
            PHEIGHT = Me.Height / (MaxY + 2) 'PWIDTH 'pixels
        Else
            PWIDTH = Me.Width / (MaxX + 2) '
            PHEIGHT = Me.Width / (MaxY + 2) 'PWIDTH 'pixels
        End If
        PIECEWIDTH = PWIDTH * 15 '32 * 15 = 480 'twips: both Piece and square
        PIECEHEIGHT = PHEIGHT * 15 '
        BOARDBOTTOM = (MaxY + 1) * PIECEHEIGHT
        BBOTTOM = (MaxY + 1) * PHEIGHT
        Me.Invalidate()
        For i = 1 To 32
            ShowPiece(i)
        Next
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        'Try to use this to allow redrawing board after user completes dragging (see Human move)
        Timer1.Enabled = False
        If aGame.Playing And Not PlayerisHuman(1 + Turn) Then
            ComputerMove()
        End If
    End Sub

    Private Sub ToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem1.Click
        'Castle
        Dim aMove As New Move 'create a new move object for storing the current move in 

        aMove.Clear()
        aMove.P1id = ContextMenuStrip1.SourceControl.Tag 'record the first Pieces move
        aMove.P1YOrig = aChessPiece(aMove.P1id).YPos 'top or bottom?
        aMove.P1XOrig = aChessPiece(aMove.P1id).XPos
        If aChessPiece(aMove.P1id).AlgNot = "R" And (aMove.P1YOrig = 1 Or aMove.P1YOrig = MaxY) Then 'it's a rook in position
            aMove.P2id = aBoard.GetGBoardID(5, aMove.P1YOrig) 'make sure king's there
            If aMove.P2id > 0 Then 'there is something there
                If aChessPiece(aMove.P2id).AlgNot = "K" Then 'it's a king!
                    If aMove.P1XOrig = 1 Then 'queenside and in legal position
                        If aBoard.GetGBoardID(2, aMove.P1YOrig) = 0 And aBoard.GetGBoardID(3, aMove.P1YOrig) = 0 And aBoard.GetGBoardID(4, aMove.P1YOrig) = 0 Then 'legal
                            DoCastle(aMove)
                            DoHistory("0-0-0", aChessPiece(aMove.P1id).Direction)
                            StatusStrip1.Items("statFrom").Text = "0-0-0"
                        End If 'no Pieces between
                    ElseIf aMove.P1XOrig = MaxX Then 'kingside and in legal position
                        If aBoard.GetGBoardID(7, aMove.P1YOrig) = 0 And aBoard.GetGBoardID(6, aMove.P1YOrig) = 0 Then 'legal
                            DoCastle(aMove)
                            DoHistory("0-0", aChessPiece(aMove.P1id).Direction)
                            StatusStrip1.Items("statFrom").Text = "0-0"
                        End If 'no Pieces between
                    End If
                End If 'it's a king
            End If 'it's a rook
        End If
        ShowPiece(aMove.P1id) 'draw the Piece (if they moved the control but we didn't move the Piece position it will redraw in the original location
        ShowPiece(aMove.P2id) 'draw the Piece (if they moved the control but we didn't move the Piece position it will redraw in the original location
        Timer1.Enabled = True
    End Sub

    Private Sub NewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuFileNew.Click
        'File New
        aGame.Clear()
        aBoard.Clear() ' = New Board
        SetupPieces()
        StatusStrip1.Items("statFrom").Text = ""
        StatusStrip1.Items("statTo").Text = ""
        Cout("CLS")
        aGame.Playing = True
        Turn = BOTTOMGOINGUP 'white/red goes first
        If PlayerisHuman(1 + Turn) Then
            ComputerMove()
        End If

    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuFileExit.Click
        'File Exit
        End
    End Sub

    Private Sub UndoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuEditUndo.Click
        'Edit Undo
        Dim aMove As Move

        Try
            aMove = aGame.Last
            If aMove.MCode = "q" Then
                UnQueen(aMove.P1id)
            End If
            aChessPiece(aMove.P1id).Move(aMove.P1XOrig, aMove.P1YOrig)
            ShowPiece(aMove.P1id)
            DoHistory("undo-> " & aChessPiece(aMove.P1id).AlgNot & Chr(96 + aMove.P1XDest) & aMove.P1YDest, aChessPiece(aMove.P1id).Direction)
            If aMove.P2id > 0 Then
                aChessPiece(aMove.P2id).Move(aMove.P2XOrig, aMove.P2YOrig)
                ShowPiece(aMove.P2id) 'make visible
            End If
            aGame.Remove() 'remove move from array
            Turn = -Turn 'end of turn
        Catch ex As Exception
            'MsgBox("MenuUndo: can't undo")
        End Try
    End Sub

    Private Sub HistoryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuToolsHistory.Click
        'Tools History
        frmCout.DefInstance.Show()
    End Sub

    Private Sub MoveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuToolsMove.Click
        'Tools Move
        If aGame.Playing Then
            ComputerMove()  'do a computer generated move
        Else
            MsgBox("Choose the New command from the File menu to start a new game.")
        End If
    End Sub

    Private Sub TopHumanToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuToolsTopHuman.Click
        'Tools Top Human
        PlayerisHuman(1 + TOPGOINGDOWN) = Not PlayerisHuman(1 + TOPGOINGDOWN)
        MnuToolsTopHuman.Checked = PlayerisHuman(1 + TOPGOINGDOWN)
        If aGame.Playing And Not PlayerisHuman(1 + Turn) Then
            ComputerMove()
        End If
    End Sub

    Private Sub BottomHumanToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MnuToolsBottomHuman.Click
        'Tools Bottom human
        PlayerisHuman(1 + BOTTOMGOINGUP) = Not PlayerisHuman(1 + BOTTOMGOINGUP)
        MnuToolsBottomHuman.Checked = PlayerisHuman(1 + BOTTOMGOINGUP)
        If aGame.Playing And Not PlayerisHuman(1 + Turn) Then
            ComputerMove()
        End If
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        dlgAboutBox1.ShowDialog()
    End Sub

    Private Function TwipsToPixelsX(twips As Integer) As Integer
        TwipsToPixelsX = CInt(twips / 15)
    End Function

    Private Function TwipsToPixelsY(twips As Integer) As Integer
        TwipsToPixelsY = CInt(twips / 15)
    End Function

    Private Sub MnuHelpInstructions_Click(sender As Object, e As EventArgs) Handles MnuHelpInstructions.Click
        Try
            Dim AppPath As String = System.AppDomain.CurrentDomain.BaseDirectory
            System.Diagnostics.Process.Start(AppPath + "ChessHelp.htm")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Protected Overloads Overrides Sub WndProc(ByRef m As Message)
        'Int wParam = (m.WParam.ToInt32() & 0xFFF0); if (wParam == 0xF030 || wParam == 0xF020 || wParam == 0xF120) { DoUpdate(); } worked for me. Also i called base.WndProc(ref m); before to be able to use updated window parameters 

        'Dim wParam As Integer = (m.WParam.ToInt32() & &HFFF0)

        MyBase.WndProc(m)
        If (m.Msg = &H112) Then '// WM_SYSCOMMAND

            '// Check your window state here
            'If (m.WParam = &HF030 Or m.WParam = &HF020 Or m.WParam = &HF120) Then '// Maximize Event - SC_MAXIMIZE from Winuser.h
            If m.WParam = New IntPtr(&HF030) Or m.WParam = New IntPtr(&HF020) Or m.WParam = New IntPtr(&HF120) Then
                Dim e As New EventArgs
                'MsgBox("Maximimised") '// THe window Is being maximized
                FrmBoard_Resize(vbNull, e)
            End If
        End If
    End Sub

    Private Sub OptionsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OptionsToolStripMenuItem.Click
        DlgOptions.ShowDialog()

    End Sub

End Class