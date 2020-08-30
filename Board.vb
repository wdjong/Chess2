Option Strict Off
Option Explicit On

Friend Class frmBoard
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
    Public WithEvents _piece_25 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_32 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_31 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_30 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_29 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_28 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_27 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_26 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_24 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_23 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_22 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_21 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_20 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_19 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_18 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_17 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_16 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_15 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_14 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_13 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_12 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_11 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_10 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_9 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_8 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_1 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_5 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_4 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_6 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_3 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_7 As System.Windows.Forms.PictureBox
    Public WithEvents _piece_2 As System.Windows.Forms.PictureBox
    Public WithEvents piece As System.Windows.Forms.PictureBox 'Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Public WithEvents Dragger As System.Windows.Forms.PictureBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents statFrom As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents statTo As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NewToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EditToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UndoToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HistoryToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MoveToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TopHumanToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents BottomHumanToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBoard))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.statFrom = New System.Windows.Forms.ToolStripStatusLabel()
        Me.statTo = New System.Windows.Forms.ToolStripStatusLabel()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EditToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.UndoToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HistoryToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MoveToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TopHumanToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.BottomHumanToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Dragger = New System.Windows.Forms.PictureBox()
        Me._piece_25 = New System.Windows.Forms.PictureBox()
        Me._piece_32 = New System.Windows.Forms.PictureBox()
        Me._piece_31 = New System.Windows.Forms.PictureBox()
        Me._piece_30 = New System.Windows.Forms.PictureBox()
        Me._piece_29 = New System.Windows.Forms.PictureBox()
        Me._piece_28 = New System.Windows.Forms.PictureBox()
        Me._piece_27 = New System.Windows.Forms.PictureBox()
        Me._piece_26 = New System.Windows.Forms.PictureBox()
        Me._piece_24 = New System.Windows.Forms.PictureBox()
        Me._piece_23 = New System.Windows.Forms.PictureBox()
        Me._piece_22 = New System.Windows.Forms.PictureBox()
        Me._piece_21 = New System.Windows.Forms.PictureBox()
        Me._piece_20 = New System.Windows.Forms.PictureBox()
        Me._piece_19 = New System.Windows.Forms.PictureBox()
        Me._piece_18 = New System.Windows.Forms.PictureBox()
        Me._piece_17 = New System.Windows.Forms.PictureBox()
        Me._piece_16 = New System.Windows.Forms.PictureBox()
        Me._piece_15 = New System.Windows.Forms.PictureBox()
        Me._piece_14 = New System.Windows.Forms.PictureBox()
        Me._piece_13 = New System.Windows.Forms.PictureBox()
        Me._piece_12 = New System.Windows.Forms.PictureBox()
        Me._piece_11 = New System.Windows.Forms.PictureBox()
        Me._piece_10 = New System.Windows.Forms.PictureBox()
        Me._piece_9 = New System.Windows.Forms.PictureBox()
        Me._piece_8 = New System.Windows.Forms.PictureBox()
        Me._piece_1 = New System.Windows.Forms.PictureBox()
        Me._piece_5 = New System.Windows.Forms.PictureBox()
        Me._piece_4 = New System.Windows.Forms.PictureBox()
        Me._piece_6 = New System.Windows.Forms.PictureBox()
        Me._piece_3 = New System.Windows.Forms.PictureBox()
        Me._piece_7 = New System.Windows.Forms.PictureBox()
        Me._piece_2 = New System.Windows.Forms.PictureBox()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.MenuStrip1.SuspendLayout()
        CType(Me.Dragger, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_25, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_32, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_31, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_30, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_29, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_28, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_27, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_26, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_24, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_23, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_22, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_21, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_20, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_19, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_18, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_17, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_16, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_15, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_14, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_13, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_12, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_11, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_10, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_9, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_8, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_6, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_7, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me._piece_2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripMenuItem1})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(107, 26)
        Me.ContextMenuStrip1.Tag = ""
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(106, 22)
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
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.statFrom, Me.statTo})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 308)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(318, 22)
        Me.StatusStrip1.TabIndex = 36
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'statFrom
        '
        Me.statFrom.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.statFrom.Name = "statFrom"
        Me.statFrom.Size = New System.Drawing.Size(35, 17)
        Me.statFrom.Text = "From"
        '
        'statTo
        '
        Me.statTo.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
        Me.statTo.Name = "statTo"
        Me.statTo.Size = New System.Drawing.Size(19, 17)
        Me.statTo.Text = "To"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.EditToolStripMenuItem, Me.ToolsToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(318, 24)
        Me.MenuStrip1.TabIndex = 37
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NewToolStripMenuItem, Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "&File"
        '
        'NewToolStripMenuItem
        '
        Me.NewToolStripMenuItem.Name = "NewToolStripMenuItem"
        Me.NewToolStripMenuItem.Size = New System.Drawing.Size(98, 22)
        Me.NewToolStripMenuItem.Text = "&New"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(98, 22)
        Me.ExitToolStripMenuItem.Text = "E&xit"
        '
        'EditToolStripMenuItem
        '
        Me.EditToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UndoToolStripMenuItem})
        Me.EditToolStripMenuItem.Name = "EditToolStripMenuItem"
        Me.EditToolStripMenuItem.Size = New System.Drawing.Size(39, 20)
        Me.EditToolStripMenuItem.Text = "&Edit"
        '
        'UndoToolStripMenuItem
        '
        Me.UndoToolStripMenuItem.Name = "UndoToolStripMenuItem"
        Me.UndoToolStripMenuItem.Size = New System.Drawing.Size(103, 22)
        Me.UndoToolStripMenuItem.Text = "&Undo"
        '
        'ToolsToolStripMenuItem
        '
        Me.ToolsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.HistoryToolStripMenuItem, Me.MoveToolStripMenuItem, Me.TopHumanToolStripMenuItem, Me.BottomHumanToolStripMenuItem})
        Me.ToolsToolStripMenuItem.Name = "ToolsToolStripMenuItem"
        Me.ToolsToolStripMenuItem.Size = New System.Drawing.Size(46, 20)
        Me.ToolsToolStripMenuItem.Text = "&Tools"
        '
        'HistoryToolStripMenuItem
        '
        Me.HistoryToolStripMenuItem.Name = "HistoryToolStripMenuItem"
        Me.HistoryToolStripMenuItem.Size = New System.Drawing.Size(157, 22)
        Me.HistoryToolStripMenuItem.Text = "&History"
        '
        'MoveToolStripMenuItem
        '
        Me.MoveToolStripMenuItem.Name = "MoveToolStripMenuItem"
        Me.MoveToolStripMenuItem.Size = New System.Drawing.Size(157, 22)
        Me.MoveToolStripMenuItem.Text = "&Move"
        '
        'TopHumanToolStripMenuItem
        '
        Me.TopHumanToolStripMenuItem.Name = "TopHumanToolStripMenuItem"
        Me.TopHumanToolStripMenuItem.Size = New System.Drawing.Size(157, 22)
        Me.TopHumanToolStripMenuItem.Text = "&Top Human"
        '
        'BottomHumanToolStripMenuItem
        '
        Me.BottomHumanToolStripMenuItem.Name = "BottomHumanToolStripMenuItem"
        Me.BottomHumanToolStripMenuItem.Size = New System.Drawing.Size(157, 22)
        Me.BottomHumanToolStripMenuItem.Text = "&Bottom Human"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AboutToolStripMenuItem})
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.HelpToolStripMenuItem.Text = "&Help"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(107, 22)
        Me.AboutToolStripMenuItem.Text = "&About"
        '
        'Dragger
        '
        Me.Dragger.BackColor = System.Drawing.Color.Transparent
        Me.Dragger.Cursor = System.Windows.Forms.Cursors.Default
        Me.Dragger.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Dragger.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Dragger.Image = Global.Chess2.My.Resources.Resources.pawn
        Me.Dragger.Location = New System.Drawing.Point(143, 147)
        Me.Dragger.Name = "Dragger"
        Me.Dragger.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Dragger.Size = New System.Drawing.Size(32, 31)
        Me.Dragger.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.Dragger.TabIndex = 35
        Me.Dragger.TabStop = False
        Me.Dragger.Tag = "20"
        Me.Dragger.Visible = False
        '
        '_piece_25
        '
        Me._piece_25.BackColor = System.Drawing.Color.Transparent
        Me._piece_25.ContextMenuStrip = Me.ContextMenuStrip1
        Me._piece_25.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_25.ErrorImage = CType(resources.GetObject("_piece_25.ErrorImage"), System.Drawing.Image)
        Me._piece_25.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_25.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_25.Image = Global.Chess2.My.Resources.Resources.rook1
        Me._piece_25.InitialImage = CType(resources.GetObject("_piece_25.InitialImage"), System.Drawing.Image)
        Me._piece_25.Location = New System.Drawing.Point(32, 24)
        Me._piece_25.Name = "_piece_25"
        Me._piece_25.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_25.Size = New System.Drawing.Size(32, 31)
        Me._piece_25.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_25.TabIndex = 25
        Me._piece_25.TabStop = False
        Me._piece_25.Tag = "25"
        '
        '_piece_32
        '
        Me._piece_32.BackColor = System.Drawing.Color.Transparent
        Me._piece_32.ContextMenuStrip = Me.ContextMenuStrip1
        Me._piece_32.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_32.ErrorImage = CType(resources.GetObject("_piece_32.ErrorImage"), System.Drawing.Image)
        Me._piece_32.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_32.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_32.Image = Global.Chess2.My.Resources.Resources.rook1
        Me._piece_32.InitialImage = CType(resources.GetObject("_piece_32.InitialImage"), System.Drawing.Image)
        Me._piece_32.Location = New System.Drawing.Point(256, 24)
        Me._piece_32.Name = "_piece_32"
        Me._piece_32.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_32.Size = New System.Drawing.Size(32, 31)
        Me._piece_32.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_32.TabIndex = 32
        Me._piece_32.TabStop = False
        Me._piece_32.Tag = "32"
        '
        '_piece_31
        '
        Me._piece_31.BackColor = System.Drawing.Color.Transparent
        Me._piece_31.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_31.ErrorImage = CType(resources.GetObject("_piece_31.ErrorImage"), System.Drawing.Image)
        Me._piece_31.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_31.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_31.Image = Global.Chess2.My.Resources.Resources.knight1
        Me._piece_31.InitialImage = CType(resources.GetObject("_piece_31.InitialImage"), System.Drawing.Image)
        Me._piece_31.Location = New System.Drawing.Point(224, 24)
        Me._piece_31.Name = "_piece_31"
        Me._piece_31.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_31.Size = New System.Drawing.Size(32, 31)
        Me._piece_31.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_31.TabIndex = 31
        Me._piece_31.TabStop = False
        Me._piece_31.Tag = "31"
        '
        '_piece_30
        '
        Me._piece_30.BackColor = System.Drawing.Color.Transparent
        Me._piece_30.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_30.ErrorImage = CType(resources.GetObject("_piece_30.ErrorImage"), System.Drawing.Image)
        Me._piece_30.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_30.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_30.Image = Global.Chess2.My.Resources.Resources.bishop1
        Me._piece_30.InitialImage = CType(resources.GetObject("_piece_30.InitialImage"), System.Drawing.Image)
        Me._piece_30.Location = New System.Drawing.Point(192, 24)
        Me._piece_30.Name = "_piece_30"
        Me._piece_30.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_30.Size = New System.Drawing.Size(32, 31)
        Me._piece_30.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_30.TabIndex = 30
        Me._piece_30.TabStop = False
        Me._piece_30.Tag = "30"
        '
        '_piece_29
        '
        Me._piece_29.BackColor = System.Drawing.Color.Transparent
        Me._piece_29.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_29.ErrorImage = CType(resources.GetObject("_piece_29.ErrorImage"), System.Drawing.Image)
        Me._piece_29.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_29.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_29.Image = Global.Chess2.My.Resources.Resources.king1
        Me._piece_29.InitialImage = CType(resources.GetObject("_piece_29.InitialImage"), System.Drawing.Image)
        Me._piece_29.Location = New System.Drawing.Point(160, 24)
        Me._piece_29.Name = "_piece_29"
        Me._piece_29.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_29.Size = New System.Drawing.Size(32, 31)
        Me._piece_29.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_29.TabIndex = 29
        Me._piece_29.TabStop = False
        Me._piece_29.Tag = "29"
        '
        '_piece_28
        '
        Me._piece_28.BackColor = System.Drawing.Color.Transparent
        Me._piece_28.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_28.ErrorImage = CType(resources.GetObject("_piece_28.ErrorImage"), System.Drawing.Image)
        Me._piece_28.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_28.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_28.Image = Global.Chess2.My.Resources.Resources.queen1
        Me._piece_28.InitialImage = CType(resources.GetObject("_piece_28.InitialImage"), System.Drawing.Image)
        Me._piece_28.Location = New System.Drawing.Point(128, 24)
        Me._piece_28.Name = "_piece_28"
        Me._piece_28.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_28.Size = New System.Drawing.Size(32, 31)
        Me._piece_28.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_28.TabIndex = 28
        Me._piece_28.TabStop = False
        Me._piece_28.Tag = "28"
        '
        '_piece_27
        '
        Me._piece_27.BackColor = System.Drawing.Color.Transparent
        Me._piece_27.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_27.ErrorImage = CType(resources.GetObject("_piece_27.ErrorImage"), System.Drawing.Image)
        Me._piece_27.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_27.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_27.Image = Global.Chess2.My.Resources.Resources.bishop1
        Me._piece_27.InitialImage = CType(resources.GetObject("_piece_27.InitialImage"), System.Drawing.Image)
        Me._piece_27.Location = New System.Drawing.Point(96, 24)
        Me._piece_27.Name = "_piece_27"
        Me._piece_27.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_27.Size = New System.Drawing.Size(32, 31)
        Me._piece_27.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_27.TabIndex = 27
        Me._piece_27.TabStop = False
        Me._piece_27.Tag = "27"
        '
        '_piece_26
        '
        Me._piece_26.BackColor = System.Drawing.Color.Transparent
        Me._piece_26.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_26.ErrorImage = CType(resources.GetObject("_piece_26.ErrorImage"), System.Drawing.Image)
        Me._piece_26.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_26.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_26.Image = Global.Chess2.My.Resources.Resources.knight1
        Me._piece_26.InitialImage = CType(resources.GetObject("_piece_26.InitialImage"), System.Drawing.Image)
        Me._piece_26.Location = New System.Drawing.Point(64, 24)
        Me._piece_26.Name = "_piece_26"
        Me._piece_26.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_26.Size = New System.Drawing.Size(32, 31)
        Me._piece_26.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_26.TabIndex = 26
        Me._piece_26.TabStop = False
        Me._piece_26.Tag = "26"
        '
        '_piece_24
        '
        Me._piece_24.BackColor = System.Drawing.Color.Transparent
        Me._piece_24.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_24.ErrorImage = CType(resources.GetObject("_piece_24.ErrorImage"), System.Drawing.Image)
        Me._piece_24.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_24.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_24.Image = Global.Chess2.My.Resources.Resources.pawn1
        Me._piece_24.InitialImage = CType(resources.GetObject("_piece_24.InitialImage"), System.Drawing.Image)
        Me._piece_24.Location = New System.Drawing.Point(256, 56)
        Me._piece_24.Name = "_piece_24"
        Me._piece_24.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_24.Size = New System.Drawing.Size(32, 31)
        Me._piece_24.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_24.TabIndex = 24
        Me._piece_24.TabStop = False
        Me._piece_24.Tag = "24"
        Me._piece_24.Visible = False
        '
        '_piece_23
        '
        Me._piece_23.BackColor = System.Drawing.Color.Transparent
        Me._piece_23.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_23.ErrorImage = CType(resources.GetObject("_piece_23.ErrorImage"), System.Drawing.Image)
        Me._piece_23.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_23.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_23.Image = Global.Chess2.My.Resources.Resources.pawn1
        Me._piece_23.InitialImage = CType(resources.GetObject("_piece_23.InitialImage"), System.Drawing.Image)
        Me._piece_23.Location = New System.Drawing.Point(224, 56)
        Me._piece_23.Name = "_piece_23"
        Me._piece_23.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_23.Size = New System.Drawing.Size(32, 31)
        Me._piece_23.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_23.TabIndex = 23
        Me._piece_23.TabStop = False
        Me._piece_23.Tag = "23"
        Me._piece_23.Visible = False
        '
        '_piece_22
        '
        Me._piece_22.BackColor = System.Drawing.Color.Transparent
        Me._piece_22.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_22.ErrorImage = CType(resources.GetObject("_piece_22.ErrorImage"), System.Drawing.Image)
        Me._piece_22.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_22.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_22.Image = Global.Chess2.My.Resources.Resources.pawn1
        Me._piece_22.InitialImage = CType(resources.GetObject("_piece_22.InitialImage"), System.Drawing.Image)
        Me._piece_22.Location = New System.Drawing.Point(192, 56)
        Me._piece_22.Name = "_piece_22"
        Me._piece_22.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_22.Size = New System.Drawing.Size(32, 31)
        Me._piece_22.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_22.TabIndex = 22
        Me._piece_22.TabStop = False
        Me._piece_22.Tag = "22"
        Me._piece_22.Visible = False
        '
        '_piece_21
        '
        Me._piece_21.BackColor = System.Drawing.Color.Transparent
        Me._piece_21.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_21.ErrorImage = CType(resources.GetObject("_piece_21.ErrorImage"), System.Drawing.Image)
        Me._piece_21.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_21.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_21.Image = Global.Chess2.My.Resources.Resources.pawn1
        Me._piece_21.InitialImage = CType(resources.GetObject("_piece_21.InitialImage"), System.Drawing.Image)
        Me._piece_21.Location = New System.Drawing.Point(160, 56)
        Me._piece_21.Name = "_piece_21"
        Me._piece_21.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_21.Size = New System.Drawing.Size(32, 31)
        Me._piece_21.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_21.TabIndex = 21
        Me._piece_21.TabStop = False
        Me._piece_21.Tag = "21"
        Me._piece_21.Visible = False
        '
        '_piece_20
        '
        Me._piece_20.BackColor = System.Drawing.Color.Transparent
        Me._piece_20.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_20.ErrorImage = CType(resources.GetObject("_piece_20.ErrorImage"), System.Drawing.Image)
        Me._piece_20.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_20.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_20.Image = Global.Chess2.My.Resources.Resources.pawn1
        Me._piece_20.InitialImage = CType(resources.GetObject("_piece_20.InitialImage"), System.Drawing.Image)
        Me._piece_20.Location = New System.Drawing.Point(128, 56)
        Me._piece_20.Name = "_piece_20"
        Me._piece_20.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_20.Size = New System.Drawing.Size(32, 31)
        Me._piece_20.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_20.TabIndex = 20
        Me._piece_20.TabStop = False
        Me._piece_20.Tag = "20"
        Me._piece_20.Visible = False
        '
        '_piece_19
        '
        Me._piece_19.BackColor = System.Drawing.Color.Transparent
        Me._piece_19.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_19.ErrorImage = CType(resources.GetObject("_piece_19.ErrorImage"), System.Drawing.Image)
        Me._piece_19.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_19.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_19.Image = Global.Chess2.My.Resources.Resources.pawn1
        Me._piece_19.InitialImage = CType(resources.GetObject("_piece_19.InitialImage"), System.Drawing.Image)
        Me._piece_19.Location = New System.Drawing.Point(96, 56)
        Me._piece_19.Name = "_piece_19"
        Me._piece_19.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_19.Size = New System.Drawing.Size(32, 31)
        Me._piece_19.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_19.TabIndex = 19
        Me._piece_19.TabStop = False
        Me._piece_19.Tag = "19"
        Me._piece_19.Visible = False
        '
        '_piece_18
        '
        Me._piece_18.BackColor = System.Drawing.Color.Transparent
        Me._piece_18.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_18.ErrorImage = CType(resources.GetObject("_piece_18.ErrorImage"), System.Drawing.Image)
        Me._piece_18.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_18.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_18.Image = Global.Chess2.My.Resources.Resources.pawn1
        Me._piece_18.InitialImage = CType(resources.GetObject("_piece_18.InitialImage"), System.Drawing.Image)
        Me._piece_18.Location = New System.Drawing.Point(64, 56)
        Me._piece_18.Name = "_piece_18"
        Me._piece_18.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_18.Size = New System.Drawing.Size(32, 31)
        Me._piece_18.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_18.TabIndex = 18
        Me._piece_18.TabStop = False
        Me._piece_18.Tag = "18"
        Me._piece_18.Visible = False
        '
        '_piece_17
        '
        Me._piece_17.BackColor = System.Drawing.Color.Transparent
        Me._piece_17.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_17.ErrorImage = CType(resources.GetObject("_piece_17.ErrorImage"), System.Drawing.Image)
        Me._piece_17.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_17.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_17.Image = Global.Chess2.My.Resources.Resources.pawn1
        Me._piece_17.InitialImage = CType(resources.GetObject("_piece_17.InitialImage"), System.Drawing.Image)
        Me._piece_17.Location = New System.Drawing.Point(32, 56)
        Me._piece_17.Name = "_piece_17"
        Me._piece_17.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_17.Size = New System.Drawing.Size(32, 31)
        Me._piece_17.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_17.TabIndex = 17
        Me._piece_17.TabStop = False
        Me._piece_17.Tag = "17"
        Me._piece_17.Visible = False
        '
        '_piece_16
        '
        Me._piece_16.BackColor = System.Drawing.Color.Transparent
        Me._piece_16.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_16.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_16.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_16.Image = Global.Chess2.My.Resources.Resources.pawn
        Me._piece_16.Location = New System.Drawing.Point(256, 224)
        Me._piece_16.Name = "_piece_16"
        Me._piece_16.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_16.Size = New System.Drawing.Size(32, 31)
        Me._piece_16.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_16.TabIndex = 16
        Me._piece_16.TabStop = False
        Me._piece_16.Tag = "16"
        Me._piece_16.Visible = False
        '
        '_piece_15
        '
        Me._piece_15.BackColor = System.Drawing.Color.Transparent
        Me._piece_15.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_15.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_15.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_15.Image = Global.Chess2.My.Resources.Resources.pawn
        Me._piece_15.Location = New System.Drawing.Point(224, 224)
        Me._piece_15.Name = "_piece_15"
        Me._piece_15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_15.Size = New System.Drawing.Size(32, 31)
        Me._piece_15.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_15.TabIndex = 15
        Me._piece_15.TabStop = False
        Me._piece_15.Tag = "15"
        Me._piece_15.Visible = False
        '
        '_piece_14
        '
        Me._piece_14.BackColor = System.Drawing.Color.Transparent
        Me._piece_14.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_14.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_14.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_14.Image = Global.Chess2.My.Resources.Resources.pawn
        Me._piece_14.Location = New System.Drawing.Point(192, 224)
        Me._piece_14.Name = "_piece_14"
        Me._piece_14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_14.Size = New System.Drawing.Size(32, 31)
        Me._piece_14.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_14.TabIndex = 14
        Me._piece_14.TabStop = False
        Me._piece_14.Tag = "14"
        Me._piece_14.Visible = False
        '
        '_piece_13
        '
        Me._piece_13.BackColor = System.Drawing.Color.Transparent
        Me._piece_13.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_13.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_13.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_13.Image = Global.Chess2.My.Resources.Resources.pawn
        Me._piece_13.Location = New System.Drawing.Point(160, 224)
        Me._piece_13.Name = "_piece_13"
        Me._piece_13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_13.Size = New System.Drawing.Size(32, 31)
        Me._piece_13.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_13.TabIndex = 13
        Me._piece_13.TabStop = False
        Me._piece_13.Tag = "13"
        Me._piece_13.Visible = False
        '
        '_piece_12
        '
        Me._piece_12.BackColor = System.Drawing.Color.Transparent
        Me._piece_12.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_12.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_12.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_12.Image = Global.Chess2.My.Resources.Resources.pawn
        Me._piece_12.Location = New System.Drawing.Point(128, 223)
        Me._piece_12.Name = "_piece_12"
        Me._piece_12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_12.Size = New System.Drawing.Size(32, 32)
        Me._piece_12.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_12.TabIndex = 12
        Me._piece_12.TabStop = False
        Me._piece_12.Tag = "12"
        Me._piece_12.Visible = False
        '
        '_piece_11
        '
        Me._piece_11.BackColor = System.Drawing.Color.Transparent
        Me._piece_11.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_11.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_11.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_11.Image = Global.Chess2.My.Resources.Resources.pawn
        Me._piece_11.Location = New System.Drawing.Point(96, 224)
        Me._piece_11.Name = "_piece_11"
        Me._piece_11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_11.Size = New System.Drawing.Size(32, 31)
        Me._piece_11.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_11.TabIndex = 11
        Me._piece_11.TabStop = False
        Me._piece_11.Tag = "11"
        Me._piece_11.Visible = False
        '
        '_piece_10
        '
        Me._piece_10.BackColor = System.Drawing.Color.Transparent
        Me._piece_10.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_10.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_10.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_10.Image = Global.Chess2.My.Resources.Resources.pawn
        Me._piece_10.Location = New System.Drawing.Point(64, 224)
        Me._piece_10.Name = "_piece_10"
        Me._piece_10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_10.Size = New System.Drawing.Size(32, 31)
        Me._piece_10.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_10.TabIndex = 10
        Me._piece_10.TabStop = False
        Me._piece_10.Tag = "10"
        Me._piece_10.Visible = False
        '
        '_piece_9
        '
        Me._piece_9.BackColor = System.Drawing.Color.Transparent
        Me._piece_9.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_9.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_9.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_9.Image = Global.Chess2.My.Resources.Resources.pawn
        Me._piece_9.Location = New System.Drawing.Point(32, 224)
        Me._piece_9.Name = "_piece_9"
        Me._piece_9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_9.Size = New System.Drawing.Size(32, 31)
        Me._piece_9.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_9.TabIndex = 9
        Me._piece_9.TabStop = False
        Me._piece_9.Tag = "9"
        Me._piece_9.Visible = False
        '
        '_piece_8
        '
        Me._piece_8.BackColor = System.Drawing.Color.Transparent
        Me._piece_8.ContextMenuStrip = Me.ContextMenuStrip1
        Me._piece_8.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_8.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_8.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_8.Image = Global.Chess2.My.Resources.Resources.rook
        Me._piece_8.Location = New System.Drawing.Point(256, 256)
        Me._piece_8.Name = "_piece_8"
        Me._piece_8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_8.Size = New System.Drawing.Size(32, 31)
        Me._piece_8.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_8.TabIndex = 8
        Me._piece_8.TabStop = False
        Me._piece_8.Tag = "8"
        Me._piece_8.Visible = False
        '
        '_piece_1
        '
        Me._piece_1.BackColor = System.Drawing.Color.Transparent
        Me._piece_1.ContextMenuStrip = Me.ContextMenuStrip1
        Me._piece_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_1.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_1.Image = Global.Chess2.My.Resources.Resources.rook
        Me._piece_1.Location = New System.Drawing.Point(32, 256)
        Me._piece_1.Name = "_piece_1"
        Me._piece_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_1.Size = New System.Drawing.Size(32, 31)
        Me._piece_1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_1.TabIndex = 1
        Me._piece_1.TabStop = False
        Me._piece_1.Tag = "1"
        Me._piece_1.Visible = False
        '
        '_piece_5
        '
        Me._piece_5.BackColor = System.Drawing.Color.Transparent
        Me._piece_5.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_5.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_5.Image = Global.Chess2.My.Resources.Resources.king
        Me._piece_5.Location = New System.Drawing.Point(160, 256)
        Me._piece_5.Name = "_piece_5"
        Me._piece_5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_5.Size = New System.Drawing.Size(32, 31)
        Me._piece_5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_5.TabIndex = 5
        Me._piece_5.TabStop = False
        Me._piece_5.Tag = "5"
        Me._piece_5.Visible = False
        '
        '_piece_4
        '
        Me._piece_4.BackColor = System.Drawing.Color.Transparent
        Me._piece_4.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_4.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_4.Image = Global.Chess2.My.Resources.Resources.queen
        Me._piece_4.Location = New System.Drawing.Point(128, 256)
        Me._piece_4.Name = "_piece_4"
        Me._piece_4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_4.Size = New System.Drawing.Size(32, 31)
        Me._piece_4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_4.TabIndex = 4
        Me._piece_4.TabStop = False
        Me._piece_4.Tag = "4"
        Me._piece_4.Visible = False
        '
        '_piece_6
        '
        Me._piece_6.BackColor = System.Drawing.Color.Transparent
        Me._piece_6.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_6.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_6.Image = Global.Chess2.My.Resources.Resources.bishop
        Me._piece_6.Location = New System.Drawing.Point(192, 256)
        Me._piece_6.Name = "_piece_6"
        Me._piece_6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_6.Size = New System.Drawing.Size(32, 31)
        Me._piece_6.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_6.TabIndex = 6
        Me._piece_6.TabStop = False
        Me._piece_6.Tag = "6"
        Me._piece_6.Visible = False
        '
        '_piece_3
        '
        Me._piece_3.BackColor = System.Drawing.Color.Transparent
        Me._piece_3.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_3.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_3.Image = Global.Chess2.My.Resources.Resources.bishop
        Me._piece_3.Location = New System.Drawing.Point(96, 256)
        Me._piece_3.Name = "_piece_3"
        Me._piece_3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_3.Size = New System.Drawing.Size(32, 31)
        Me._piece_3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_3.TabIndex = 3
        Me._piece_3.TabStop = False
        Me._piece_3.Tag = "3"
        Me._piece_3.Visible = False
        '
        '_piece_7
        '
        Me._piece_7.BackColor = System.Drawing.Color.Transparent
        Me._piece_7.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_7.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_7.Image = Global.Chess2.My.Resources.Resources.knight
        Me._piece_7.Location = New System.Drawing.Point(224, 256)
        Me._piece_7.Name = "_piece_7"
        Me._piece_7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_7.Size = New System.Drawing.Size(32, 31)
        Me._piece_7.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_7.TabIndex = 7
        Me._piece_7.TabStop = False
        Me._piece_7.Tag = "7"
        Me._piece_7.Visible = False
        '
        '_piece_2
        '
        Me._piece_2.BackColor = System.Drawing.Color.Transparent
        Me._piece_2.Cursor = System.Windows.Forms.Cursors.Default
        Me._piece_2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._piece_2.ForeColor = System.Drawing.SystemColors.WindowText
        Me._piece_2.Image = Global.Chess2.My.Resources.Resources.knight
        Me._piece_2.Location = New System.Drawing.Point(64, 256)
        Me._piece_2.Name = "_piece_2"
        Me._piece_2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._piece_2.Size = New System.Drawing.Size(32, 31)
        Me._piece_2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me._piece_2.TabIndex = 2
        Me._piece_2.TabStop = False
        Me._piece_2.Tag = "2"
        Me._piece_2.Visible = False
        '
        'frmBoard
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(318, 330)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Controls.Add(Me.Dragger)
        Me.Controls.Add(Me._piece_25)
        Me.Controls.Add(Me._piece_32)
        Me.Controls.Add(Me._piece_31)
        Me.Controls.Add(Me._piece_30)
        Me.Controls.Add(Me._piece_29)
        Me.Controls.Add(Me._piece_28)
        Me.Controls.Add(Me._piece_27)
        Me.Controls.Add(Me._piece_26)
        Me.Controls.Add(Me._piece_24)
        Me.Controls.Add(Me._piece_23)
        Me.Controls.Add(Me._piece_22)
        Me.Controls.Add(Me._piece_21)
        Me.Controls.Add(Me._piece_20)
        Me.Controls.Add(Me._piece_19)
        Me.Controls.Add(Me._piece_18)
        Me.Controls.Add(Me._piece_17)
        Me.Controls.Add(Me._piece_16)
        Me.Controls.Add(Me._piece_15)
        Me.Controls.Add(Me._piece_14)
        Me.Controls.Add(Me._piece_13)
        Me.Controls.Add(Me._piece_12)
        Me.Controls.Add(Me._piece_11)
        Me.Controls.Add(Me._piece_10)
        Me.Controls.Add(Me._piece_9)
        Me.Controls.Add(Me._piece_8)
        Me.Controls.Add(Me._piece_1)
        Me.Controls.Add(Me._piece_5)
        Me.Controls.Add(Me._piece_4)
        Me.Controls.Add(Me._piece_6)
        Me.Controls.Add(Me._piece_3)
        Me.Controls.Add(Me._piece_7)
        Me.Controls.Add(Me._piece_2)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Location = New System.Drawing.Point(20, 20)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "frmBoard"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Text = "Chess"
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.Dragger, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_25, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_32, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_31, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_30, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_29, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_28, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_27, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_26, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_24, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_23, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_22, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_21, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_20, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_19, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_18, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_17, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_16, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_15, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_14, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_13, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_12, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_11, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_10, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_9, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_8, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_6, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_7, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me._piece_2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
#Region "Upgrade Support "
    Private Shared m_vb6FormDefInstance As frmBoard
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As frmBoard
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New frmBoard()
                m_InitializingDefInstance = False
            End If
            DefInstance = m_vb6FormDefInstance
        End Get
        Set(ByVal value As frmBoard)
            m_vb6FormDefInstance = value
        End Set
    End Property
#End Region
    'Pieces are controls on the board ranging from 0 to 31
    'aChessPiece is a corresponding array of chesspiece objects from 0 to 31
    Dim IsUDraggin As Boolean 'used to track dragging pieces...
    Dim origX As Long 'the mouse point from which we dragged
    Dim origY As Long 'the mouse point from which we dragged

    Sub ComputerMove()
        'Find the best move and do it.
        Dim pbest As Byte 'best piece (1..32)  where 1 is White (Red) Rook at bottom left and and 32 is Black (Blue) Rook at top right
        Dim mbest As Byte 'best move

        Do
            getBestMove(pbest, mbest)
            If pbest < 1 Or pbest > 32 Then 'there is no legal move
                If isInCheck(Turn) Then
                    MsgBox("Check mate. You win.")
                Else 'No legal move and no check mate e.g. when there is just a king and any move would put him in check
                    MsgBox("Stalemate. We draw.")
                End If
                aGame.Playing = False
                Timer1.Enabled = False
            Else 'carry out the best move
                doBestMove(pbest, mbest)
                Turn = -Turn 'next turn
            End If
            If Not PlayerisHuman(0) And Not PlayerisHuman(2) Then 'When both opponents are computer's, currently pausing after each move to check action. 0 = top, 2 = bottom when it's bottoms turn turn = 1 and when it's top's turn = -1
                MsgBox("pause")
            End If
        Loop While aGame.Playing And Not PlayerisHuman(1 + Turn) 'if then next turn is a computer loop
    End Sub

    Sub DrawBoard(ByVal formGraphics As System.Drawing.Graphics)
        'Draw an array of alternating colour squares: not the pieces
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

    Private Sub frmBoard_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        aGame.Playing = True
        PlayerisHuman(1 + TOPGOINGDOWN) = False 'i.e. index = (0) 'i.e. Black/Blue at top is a computer
        PlayerisHuman(1 + BOTTOMGOINGUP) = True 'i.e. index = (2) 'i.e. White/Red at the bottom is a human (by default. See menu)
        TopHumanToolStripMenuItem.Checked = PlayerisHuman(1 + TOPGOINGDOWN)
        BottomHumanToolStripMenuItem.Checked = PlayerisHuman(1 + BOTTOMGOINGUP)
        Turn = BOTTOMGOINGUP 'white/red goes first
        StatusStrip1.Items("statFrom").Text = ""
        StatusStrip1.Items("statTo").Text = ""
        Cout("CLS")
        aBoard = New Board
        PHEIGHT = PWIDTH 'pixels
        PIECEWIDTH = PWIDTH * 15 '32 * 15 = 480 'twips: both piece and square
        PIECEHEIGHT = PHEIGHT * 15 '
        BOARDBOTTOM = (MaxY + 1) * PIECEHEIGHT
        BBOTTOM = (MaxY + 1) * PHEIGHT
        Me.Width = (MaxX + 2) * PWIDTH
        Me.Height = (MaxY + 5) * PHEIGHT
        SetupPieces()
    End Sub

    Private Sub frmBoard_Closed(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Closed
        End
    End Sub

    Private Sub HumanMove(ByRef Source As System.Windows.Forms.Control, ByRef X As Single, ByRef Y As Single)
        'What in VB6 was the source is now the sender; the control we're working with
        'Human move routine: called by mouse up at end of dragging
        'Called at the end of a dragging move when it's landed on a piece
        'The source.tag identifies the piece which has been moved
        Dim aMove As New Move 'create a new move object for storing the current move in 

        aMove.clear()
        aMove.p1id = Source.Tag 'record the first pieces move
        aMove.p1XOrig = aChessPiece(Source.Tag).xPos
        aMove.p1YOrig = aChessPiece(Source.Tag).yPos
        aMove.p1XDest = xpos(Source.Left + X) 'destX
        aMove.p1YDest = ypos(Source.Top + Y) 'destY
        aMove.p2id = aBoard.GetGBoardID(aMove.p1XDest, aMove.p1YDest)
        If aMove.p2id > 0 Then 'there's a piece being taken
            aMove.p2XOrig = aMove.p1XDest
            aMove.p2YOrig = aMove.p1YDest
        End If
        If isLegal1(aMove) Then
            Take(aMove.p2XOrig, aMove.p2YOrig) 'Take their piece if it's there
            aChessPiece(aMove.p1id).move(aMove.p1XDest, aMove.p1YDest) 'Move your piece
            If ispromoting(aMove.p1id, aMove.p1YDest) Then 'it realises it's reached the 8th rank
                Queening(aMove.p1id) 'change into a queen
                aMove.MCode = "q" 'in case we need to undo
            End If
            If isInCheck(Turn) Then 'Check again for legality after the move & e.g. promoting
                MsgBox("You're in check. Either you moved into it or you didn't move out of it.")
                If aMove.MCode = "q" Then
                    UnQueen(Source.Tag)
                End If
                aChessPiece(Source.Tag).move(aMove.p1XOrig, aMove.p1YOrig) 'move piece back
                If aMove.p2id > 0 Then
                    aChessPiece(aMove.p2id).move(aMove.p2XOrig, aMove.p2YOrig) 'move taken piece back
                    ShowPiece(aMove.p2id) 'make visible again
                End If
            Else
                aGame.add(aMove) 'add the move to the array
                Turn = -Turn
            End If
        End If
        StatusStrip1.Items("statFrom").Text = "From: " & aChessPiece(aMove.p1id).AlgNot & Chr(96 + aMove.p1XOrig) & aMove.p1YOrig
        StatusStrip1.Items("statTo").Text = "To: " & aChessPiece(aMove.p1id).AlgNot & Chr(96 + aMove.p1XDest) & aMove.p1YDest
        doHistory(aChessPiece(aMove.p1id).AlgNot & Chr(96 + aMove.p1XDest) & aMove.p1YDest, aChessPiece(aMove.p1id).direction)
        ShowPiece((aMove.p1id)) 'draw the piece (if they moved the control but we didn't move the piece position it will redraw in the original location
        Timer1.Enabled = True
    End Sub

    Private Sub imgMousedown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles _piece_1.MouseDown, _piece_2.MouseDown, _piece_3.MouseDown, _piece_4.MouseDown, _piece_5.MouseDown, _piece_6.MouseDown, _piece_7.MouseDown, _piece_8.MouseDown, _piece_9.MouseDown, _piece_10.MouseDown, _piece_11.MouseDown, _piece_12.MouseDown, _piece_13.MouseDown, _piece_14.MouseDown, _piece_15.MouseDown, _piece_16.MouseDown, _piece_17.MouseDown, _piece_18.MouseDown, _piece_19.MouseDown, _piece_20.MouseDown, _piece_21.MouseDown, _piece_22.MouseDown, _piece_23.MouseDown, _piece_24.MouseDown, _piece_25.MouseDown, _piece_26.MouseDown, _piece_27.MouseDown, _piece_28.MouseDown, _piece_29.MouseDown, _piece_30.MouseDown, _piece_31.MouseDown, _piece_32.MouseDown
        'The user commences dragging with the mousedown
        If Not IsUDraggin Then
            IsUDraggin = True
            Dragger = sender
            'Dragger.BackColor = System.Drawing.Color.Gray ' pieceArray(i).BackColor = System.Drawing.ColorTranslator.FromOle(QBColor(7))
            Dragger.Visible = True
            Dragger.BringToFront()
            origX = e.X 'used for displaying dragged piece with same offset
            origY = e.Y 'e.x and e.y are the click position relative to the control
        End If
    End Sub

    Private Sub imgMouseUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles _piece_1.MouseUp, _piece_2.MouseUp, _piece_3.MouseUp, _piece_4.MouseUp, _piece_5.MouseUp, _piece_6.MouseUp, _piece_7.MouseUp, _piece_8.MouseUp, _piece_9.MouseUp, _piece_10.MouseUp, _piece_11.MouseUp, _piece_12.MouseUp, _piece_13.MouseUp, _piece_14.MouseUp, _piece_15.MouseUp, _piece_16.MouseUp, _piece_17.MouseUp, _piece_18.MouseUp, _piece_19.MouseUp, _piece_20.MouseUp, _piece_21.MouseUp, _piece_22.MouseUp, _piece_23.MouseUp, _piece_24.MouseUp, _piece_25.MouseUp, _piece_26.MouseUp, _piece_27.MouseUp, _piece_28.MouseUp, _piece_29.MouseUp, _piece_30.MouseUp, _piece_31.MouseUp, _piece_32.MouseUp
        'The user release the mouse button completing the dragging
        'I think e.x and e.y are relative to the control so the absolute position on the form is
        'sender.left + e.x and sender.top + e.y
        If IsUDraggin Then
            IsUDraggin = False
            Dragger.Visible = False
            HumanMove(sender, e.X, e.Y)
        End If
    End Sub

    Private Sub RP1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles _piece_1.MouseMove, _piece_2.MouseMove, _piece_3.MouseMove, _piece_4.MouseMove, _piece_5.MouseMove, _piece_6.MouseMove, _piece_7.MouseMove, _piece_8.MouseMove, _piece_9.MouseMove, _piece_10.MouseMove, _piece_11.MouseMove, _piece_12.MouseMove, _piece_13.MouseMove, _piece_14.MouseMove, _piece_15.MouseMove, _piece_16.MouseMove, _piece_17.MouseMove, _piece_18.MouseMove, _piece_19.MouseMove, _piece_20.MouseMove, _piece_21.MouseMove, _piece_22.MouseMove, _piece_23.MouseMove, _piece_24.MouseMove, _piece_25.MouseMove, _piece_26.MouseMove, _piece_27.MouseMove, _piece_28.MouseMove, _piece_29.MouseMove, _piece_30.MouseMove, _piece_31.MouseMove, _piece_32.MouseMove
        'The user moves the mouse while dragging
        If IsUDraggin Then
            Dragger.Left = sender.left + e.X - origX 'baseX + e.X - origX
            Dragger.Top = sender.top + e.Y - origY 'baseY + e.Y - origY
        End If
    End Sub

    Public ReadOnly Property xpos(ByVal xmouse As Single) As Byte
        Get
            xpos = Int(xmouse / PWIDTH)
        End Get
    End Property

    Public ReadOnly Property ypos(ByVal ymouse As Single) As Byte
        Get
            ypos = MaxY + 1 - Int(ymouse / PHEIGHT)
        End Get
    End Property

    Private Sub frmBoard_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
        DrawBoard(e.Graphics)
    End Sub


    Sub ShowPiece(ByRef i As Short) 'I is the Piece
        'Move actual piece (control) and display move in algebraic notation
        Dim pieceexists As Boolean = True

        Try
            If aChessPiece(i) Is Nothing Then
                pieceexists = False
            End If
        Catch ex As Exception
            'ignore errors: MsgBox("Showpiece: " & ex.Message)
        End Try
        If pieceexists Then
            If aChessPiece(i).onboard Then
                pieceArray(i).Visible = True
                pieceArray(i).Width = PWIDTH
                pieceArray(i).Height = PHEIGHT
                'On the form adjust the position of the piece
                pieceArray(i).Left = TwipsToPixelsX(aChessPiece(i).xPos * PIECEWIDTH)
                pieceArray(i).Top = TwipsToPixelsY(((MaxY + 1) - aChessPiece(i).yPos) * PIECEHEIGHT)
            Else
                pieceArray(i).Visible = False
            End If
        End If
    End Sub

    Private Sub frmBoard_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeEnd
        Dim i As Byte

        PWIDTH = Me.Width / (MaxX + 2) '
        PHEIGHT = Me.Height / (MaxY + 5) 'PWIDTH 'pixels
        PIECEWIDTH = PWIDTH * 15 '32 * 15 = 480 'twips: both piece and square
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

        aMove.clear()
        aMove.p1id = ContextMenuStrip1.SourceControl.Tag 'record the first pieces move
        aMove.p1YOrig = aChessPiece(aMove.p1id).yPos 'top or bottom?
        aMove.p1XOrig = aChessPiece(aMove.p1id).xPos
        If aChessPiece(aMove.p1id).algnot = "R" And (aMove.p1YOrig = 1 Or aMove.p1YOrig = MaxY) Then 'it's a rook in position
            aMove.p2id = aBoard.GetGBoardID(5, aMove.p1YOrig) 'make sure king's there
            If aMove.p2id > 0 Then 'there is something there
                If aChessPiece(aMove.p2id).algnot = "K" Then 'it's a king!
                    If aMove.p1XOrig = 1 Then 'queenside and in legal position
                        If aBoard.GetGBoardID(2, aMove.p1YOrig) = 0 And aBoard.GetGBoardID(3, aMove.p1YOrig) = 0 And aBoard.GetGBoardID(4, aMove.p1YOrig) = 0 Then 'legal
                            doCastle(aMove)
                            doHistory("0-0-0", aChessPiece(aMove.p1id).direction)
                            StatusStrip1.Items("statFrom").Text = "0-0-0"
                        End If 'no pieces between
                    ElseIf aMove.p1XOrig = MaxX Then 'kingside and in legal position
                        If aBoard.GetGBoardID(7, aMove.p1YOrig) = 0 And aBoard.GetGBoardID(6, aMove.p1YOrig) = 0 Then 'legal
                            doCastle(aMove)
                            doHistory("0-0", aChessPiece(aMove.p1id).direction)
                            StatusStrip1.Items("statFrom").Text = "0-0"
                        End If 'no pieces between
                    End If
                End If 'it's a king
            End If 'it's a rook
        End If
        ShowPiece(aMove.p1id) 'draw the piece (if they moved the control but we didn't move the piece position it will redraw in the original location
        ShowPiece(aMove.p2id) 'draw the piece (if they moved the control but we didn't move the piece position it will redraw in the original location
        Timer1.Enabled = True
    End Sub

    Private Sub NewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripMenuItem.Click
        'File New
        aGame.Playing = True
        aBoard.Clear() ' = New Board
        Turn = BOTTOMGOINGUP 'white/red goes first
        SetupPieces()
        StatusStrip1.Items("statFrom").Text = ""
        StatusStrip1.Items("statTo").Text = ""
        Cout("CLS")
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        'File Exit
        End
    End Sub

    Private Sub UndoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UndoToolStripMenuItem.Click
        'Edit Undo
        Dim aMove As Move

        Try
            aMove = aGame.last
            If aMove.MCode = "q" Then
                UnQueen(aMove.p1id)
            End If
            aChessPiece(aMove.p1id).move(aMove.p1XOrig, aMove.p1YOrig)
            ShowPiece(aMove.p1id)
            doHistory("undo-> " & aChessPiece(aMove.p1id).AlgNot & Chr(96 + aMove.p1XDest) & aMove.p1YDest, aChessPiece(aMove.p1id).direction)
            If aMove.p2id > 0 Then
                aChessPiece(aMove.p2id).move(aMove.p2XOrig, aMove.p2YOrig)
                ShowPiece(aMove.p2id) 'make visible
            End If
            aGame.remove() 'remove move from array
            Turn = -Turn 'end of turn
        Catch ex As Exception
            'MsgBox("MenuUndo: can't undo")
        End Try
    End Sub

    Private Sub HistoryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HistoryToolStripMenuItem.Click
        'Tools History
        frmCout.DefInstance.Show()
    End Sub

    Private Sub MoveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MoveToolStripMenuItem.Click
        'Tools Move
        If aGame.Playing Then
            ComputerMove()  'do a computer generated move
        Else
            MsgBox("Choose the New command from the File menu to start a new game.")
        End If
    End Sub

    Private Sub TopHumanToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TopHumanToolStripMenuItem.Click
        'Tools Top Human
        PlayerisHuman(1 + TOPGOINGDOWN) = Not PlayerisHuman(1 + TOPGOINGDOWN)
        TopHumanToolStripMenuItem.Checked = PlayerisHuman(1 + TOPGOINGDOWN)
        If aGame.Playing And Not PlayerisHuman(1 + Turn) Then
            ComputerMove()
        End If
    End Sub

    Private Sub BottomHumanToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BottomHumanToolStripMenuItem.Click
        'Tools Bottom human
        PlayerisHuman(1 + BOTTOMGOINGUP) = Not PlayerisHuman(1 + BOTTOMGOINGUP)
        BottomHumanToolStripMenuItem.Checked = PlayerisHuman(1 + BOTTOMGOINGUP)
        If aGame.Playing And Not PlayerisHuman(1 + Turn) Then
            ComputerMove()
        End If
    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        'HelpAbout
        MsgBox("Not implemented") 'My.Forms.frmAbout.Show()
    End Sub

    Private Function TwipsToPixelsX(twips As Integer) As Integer
        TwipsToPixelsX = CInt(twips / 15)
    End Function
    Private Function TwipsToPixelsY(twips As Integer) As Integer
        TwipsToPixelsY = CInt(twips / 15)
    End Function
End Class