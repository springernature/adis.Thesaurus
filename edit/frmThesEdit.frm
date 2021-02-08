VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmThesEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Thesaurus Term"
   ClientHeight    =   6930
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   6630
   HelpContextID   =   10021
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   372
      Left            =   4080
      TabIndex        =   34
      Top             =   6480
      Width           =   1092
   End
   Begin VB.Timer timerScroll 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1080
      Top             =   6480
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   372
      Left            =   5400
      TabIndex        =   33
      Top             =   6480
      Width           =   1092
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   372
      Left            =   2760
      TabIndex        =   32
      Top             =   6480
      Width           =   1092
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   372
      Left            =   1440
      TabIndex        =   31
      Top             =   6480
      Width           =   1092
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Height          =   372
      Left            =   120
      TabIndex        =   30
      Top             =   6480
      Width           =   1092
   End
   Begin TabDlg.SSTab tabThesTerm 
      DragIcon        =   "frmThesEdit.frx":0000
      Height          =   6255
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   11033
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "&Term"
      TabPicture(0)   =   "frmThesEdit.frx":0152
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblStatic(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblStatic(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblStatic(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblStatic(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblStatic(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblStatic(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtTerm"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtUsedFor"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtPrintForm"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cboDesType"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtScopeNotes"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lstDateNotes"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "&Related"
      TabPicture(1)   =   "frmThesEdit.frx":016E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frameRelated"
      Tab(1).Control(1)=   "frameIndicator"
      Tab(1).Control(2)=   "tabGroup"
      Tab(1).Control(3)=   "lstGroup"
      Tab(1).Control(4)=   "cmdAddGroup"
      Tab(1).Control(5)=   "cmdDelGroup"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "&Hierarchy"
      TabPicture(2)   =   "frmThesEdit.frx":018A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblStatic(6)"
      Tab(2).Control(1)=   "lblStatic(7)"
      Tab(2).Control(2)=   "imgSplitter"
      Tab(2).Control(3)=   "treeHierarchy"
      Tab(2).Control(4)=   "cboHierarchy"
      Tab(2).Control(5)=   "chkInUseOnly"
      Tab(2).Control(6)=   "treeInverted"
      Tab(2).Control(7)=   "cboLocate"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Histor&y"
      TabPicture(3)   =   "frmThesEdit.frx":01A6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      Begin MSComctlLib.ListView lstDateNotes 
         Height          =   1095
         Left            =   1200
         TabIndex        =   11
         Top             =   4500
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   1931
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Note"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.ComboBox cboLocate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -73920
         TabIndex        =   26
         Top             =   780
         Width           =   2652
      End
      Begin MSComctlLib.TreeView treeInverted 
         DragIcon        =   "frmThesEdit.frx":01C2
         Height          =   915
         Left            =   -74760
         TabIndex        =   28
         Top             =   5160
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   1614
         _Version        =   393217
         Indentation     =   706
         LabelEdit       =   1
         Style           =   6
         HotTracking     =   -1  'True
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CheckBox chkInUseOnly 
         Caption         =   "In Use Only"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -71040
         TabIndex        =   24
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelGroup 
         Caption         =   "Delete"
         Height          =   325
         Left            =   -69780
         TabIndex        =   19
         Top             =   2760
         Width           =   852
      End
      Begin VB.CommandButton cmdAddGroup 
         Caption         =   "Add ..."
         Height          =   325
         Left            =   -69780
         TabIndex        =   18
         Top             =   2280
         Width           =   852
      End
      Begin VB.ListBox lstGroup 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1980
         Left            =   -74760
         Sorted          =   -1  'True
         TabIndex        =   17
         Top             =   2280
         Width           =   4755
      End
      Begin MSComctlLib.TabStrip tabGroup 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   16
         Top             =   1920
         Width           =   6075
         _ExtentX        =   10716
         _ExtentY        =   4471
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame frameIndicator 
         Caption         =   "Indicators"
         Height          =   1575
         Left            =   -74880
         TabIndex        =   20
         Top             =   4560
         Width           =   6075
         Begin VB.ListBox lstIndicator 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1140
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   21
            Top             =   240
            Width           =   4755
         End
      End
      Begin VB.Frame frameRelated 
         Caption         =   "Related Terms"
         Height          =   1452
         Left            =   -74880
         TabIndex        =   12
         Top             =   360
         Width           =   6075
         Begin VB.ListBox lstRelated 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1020
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   13
            Top             =   240
            Width           =   4755
         End
         Begin VB.CommandButton cmdDelRelated 
            Caption         =   "Delete"
            Height          =   325
            Left            =   5100
            TabIndex        =   15
            Top             =   720
            Width           =   852
         End
         Begin VB.CommandButton cmdAddRelated 
            Caption         =   "Add ..."
            Height          =   325
            Left            =   5100
            TabIndex        =   14
            Top             =   240
            Width           =   852
         End
      End
      Begin VB.TextBox txtScopeNotes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   3120
         Width           =   4995
      End
      Begin VB.ComboBox cboDesType 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1440
         Width           =   2292
      End
      Begin VB.TextBox txtPrintForm 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   1
         Top             =   480
         Width           =   3555
      End
      Begin VB.ComboBox cboHierarchy 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   -73920
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   360
         Width           =   2652
      End
      Begin MSComctlLib.TreeView treeHierarchy 
         DragIcon        =   "frmThesEdit.frx":0314
         Height          =   3855
         Left            =   -74760
         TabIndex        =   27
         Top             =   1200
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   6800
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   706
         LabelEdit       =   1
         Sorted          =   -1  'True
         Style           =   6
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
      End
      Begin VB.TextBox txtUsedFor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Top             =   1920
         Width           =   4995
      End
      Begin VB.TextBox txtTerm 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   3
         Top             =   960
         Width           =   3555
      End
      Begin VB.Label lblStatic 
         Caption         =   "Date Notes:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   10
         Top             =   4560
         Width           =   1095
      End
      Begin VB.Image imgSplitter 
         DragMode        =   1  'Automatic
         Height          =   132
         Left            =   -74760
         Top             =   4920
         Width           =   5412
      End
      Begin VB.Label lblStatic 
         Caption         =   "Find:"
         Height          =   255
         Index           =   7
         Left            =   -74820
         TabIndex        =   25
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblStatic 
         Caption         =   "Hierarchy:"
         Height          =   255
         Index           =   6
         Left            =   -74880
         TabIndex        =   22
         Top             =   420
         Width           =   855
      End
      Begin VB.Label lblStatic 
         Caption         =   "Scope Notes:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   3180
         Width           =   1095
      End
      Begin VB.Label lblStatic 
         Caption         =   "Type:"
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   852
      End
      Begin VB.Label lblStatic 
         Caption         =   "Print Form:"
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   1092
      End
      Begin VB.Label lblStatic 
         Caption         =   "Used For:"
         Height          =   252
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   1920
         Width           =   972
      End
      Begin VB.Label lblStatic 
         Caption         =   "Descriptor:"
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1092
      End
   End
End
Attribute VB_Name = "frmThesEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' **********************************************************************
'
' Module: $Id: frmThesEdit.frm,v 1.9 2000/06/20 03:45:26 AH2 Exp $
'
' Author: Andrew Howroyd
'
' Created: Jul 1999
'
' Description
'
' **********************************************************************
Option Explicit
Option Compare Text

' module name for exception reporting
Private Const m_module As String = "frmThesEdit"

' set by caller
Public m_context As HAppContext
Public m_thesObj As ThesObj
Public m_leadTerm As String
Public m_isNew As Boolean
Public m_ok As Boolean

' variables
Private m_termKv As New HKvList ' current term
Private m_orig As HKvList    ' original
Private m_isLocked As Boolean ' if thesaurus locked
Private m_refreshHierFlag As Boolean
Private m_curHier As HierObj
Private m_boldTerms As HStrVect
Private m_groupTabTag As String
Private m_groupSqlWhere As New HKvList

' icons for drag and drop
Private m_moveCursor As Object
Private m_copyCursor As Object

' drag and drop scrolling
Private m_inDrag As Boolean
Private m_scrollUp As Boolean

' declare windows functions
Private Declare Function SendMessage Lib "user32" Alias _
        "SendMessageA" (ByVal hwnd As Long, _
        ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
                  
' -------------------------------------------------------------------------
' fillDescriptorType
'    populates descriptor type combo
'    the list is dependent on hierarchies which term belongs to
'    Also, if new term, updates caption and editability of term name
'    (for a new term, term name only editable until placed in hierarchy)
' -------------------------------------------------------------------------
Private Sub fillDescriptorType()
  
  Dim hierarchies As HStrVect
  Set hierarchies = m_termKv.Object(g_sHierarchy)
  
  Dim AllDesTypes As HStrVect
  Set AllDesTypes = m_thesObj.AllDesTypes
  
  Dim curDesType As String
  curDesType = m_termKv(g_sDescriptorType)
  
  cboDesType.Clear
    
  Dim isLocked As Boolean
  Dim strVect As HStrVect
  Set strVect = m_termKv.Object(g_sGroupTerm)
  If strVect.Count > 0 Then isLocked = True
  Set strVect = m_termKv.Object(g_sMemberTerm)
  If strVect.Count > 0 Then isLocked = True
  Set strVect = m_termKv.Object(g_sIndicator)
  If strVect.Count > 0 Then isLocked = True
  
  Dim i As Long, k As Long
  For i = 1 To AllDesTypes.Count
    Dim Info As HKvList
    Set Info = m_thesObj.DesTypeInfo(AllDesTypes(i))
    
    If Info!Caption <> curDesType And Not isLocked Then
      Set strVect = Info.Object("Hierarchies")
      Dim bAdd As Boolean
      bAdd = True
      For k = 1 To hierarchies.Count
        If strVect Is Nothing Then
          bAdd = False
          Exit For
        ElseIf strVect.Find(hierarchies(k)) = 0 Then
          bAdd = False
          Exit For
        End If
      Next
      If bAdd Then
        cboDesType.AddItem AllDesTypes(i)
      End If
    End If
  Next
  
  ' enable / disable if any elements to select
  cboDesType.Enabled = cboDesType.ListCount > 0
  
  ' now add current descriptor type and select
  If curDesType <> "" Then
    cboDesType.AddItem curDesType
    cboDesType.ListIndex = cboDesType.NewIndex
  End If
  
  ' update other dialog components if new thesaurus term
  If m_isNew Then
    If hierarchies.Count() > 0 Then
      txtTerm.Enabled = False
      txtTerm.BackColor = Me.BackColor
      Caption = m_leadTerm
    Else
      txtTerm.Enabled = True
      txtTerm.BackColor = vbWindowBackground
      Caption = "New Thesaurus Term"
    End If
  End If
  
End Sub

' -------------------------------------------------------------------------
' fillInvertedLev
'   Helper to polulate inverted tree
' -------------------------------------------------------------------------
Private Sub fillInvertedLev(ByVal Node As Node)
  Dim parentVect As HStrVect
  Set parentVect = m_curHier.Parents(Node.Text)
  If Not parentVect Is Nothing Then
    Dim i As Long
    For i = 1 To parentVect.Count
      fillInvertedLev treeInverted.Nodes.Add(Node, tvwChild, , parentVect(i))
    Next
    Node.Sorted = True
  End If
End Sub

' -------------------------------------------------------------------------
' fillInverted
' -------------------------------------------------------------------------
Private Sub fillInverted()
  treeInverted.Nodes.Clear
  If Not treeHierarchy.SelectedItem Is Nothing Then
    Dim leadTerm As String
    leadTerm = treeHierarchy.SelectedItem.Text
    Dim Node As Node
    Set Node = treeInverted.Nodes.Add(, , , leadTerm)
    fillInvertedLev Node
    Node.Expanded = True
  End If
End Sub

' -------------------------------------------------------------------------
' checkNode
'   Called when node expanded or unexpanded
' -------------------------------------------------------------------------
Private Sub checkNode(ByVal Node As MSComctlLib.Node)
  Dim isExpanded As Boolean
  isExpanded = Node.Expanded
  While Node.Children > 0
    treeHierarchy.Nodes.Remove Node.Child.Index
  Wend
  If isExpanded Then
    Dim childTerms As HStrVect
    Set childTerms = m_curHier.Children(Node.Text)
    If Not childTerms Is Nothing Then
      Dim i As Long
      Dim NewNode As Node
      For i = 1 To childTerms.Count
        Set NewNode = treeHierarchy.Nodes.Add(Node, tvwChild, , childTerms(i))
        checkNode NewNode
        If Not m_boldTerms Is Nothing And Node.Bold Then
          If m_boldTerms.Find(NewNode.Text) > 0 Then NewNode.Bold = True
        End If
      Next
      Node.Sorted = True
    End If
  ElseIf m_curHier.NumChildren(Node.Text) > 0 Then
    treeHierarchy.Nodes.Add Node, tvwChild
  End If
End Sub

' -------------------------------------------------------------------------
' treeDelNode
'   Called after node deleted
'   Traverses tree recursively to update tree
'   NB: parentName, childName passed ByRef to avoid copying
' -------------------------------------------------------------------------
Private Sub treeDelNode(ByVal Node As MSComctlLib.Node, _
        parentName As String, childName As String)
  Dim isParent As Boolean
  isParent = Node.Text = parentName
  If Node.Expanded Then
    Set Node = Node.Child
    Do While Not Node Is Nothing
      If Not isParent Then
        treeDelNode Node, parentName, childName
      ElseIf Node.Text = childName Then
        treeHierarchy.Nodes.Remove Node.Index
        Exit Do
      End If
      Set Node = Node.Next
    Loop
  ElseIf isParent Then
    checkNode Node
  End If
End Sub

' -------------------------------------------------------------------------
' treeAddNode
'   Called after node added
'   Traverses tree recursively to update tree
'   NB: parentName, childName passed ByRef to avoid copying
' -------------------------------------------------------------------------
Private Sub treeAddNode(ByVal Node As MSComctlLib.Node, _
            parentName As String, childName As String)
  If Node.Text = parentName Then
    If Node.Expanded Then
      checkNode treeHierarchy.Nodes.Add(Node, tvwChild, , childName)
      Node.Sorted = True
    Else
      checkNode Node
    End If
  ElseIf Node.Expanded Then
    Set Node = Node.Child
    While Not Node Is Nothing
      treeAddNode Node, parentName, childName
      Set Node = Node.Next
    Wend
  End If
End Sub

' -------------------------------------------------------------------------
' treeExecCopy
'   Helper function for copying tree (called recursively)
'   called from treeCopyNode
' -------------------------------------------------------------------------
Private Sub treeExecCopy(ByVal Target As Node, ByVal SrcList As Node)
  While Not SrcList Is Nothing
    Dim NewNode As Node
    Set NewNode = treeHierarchy.Nodes.Add(Target, tvwChild)
    If SrcList.Expanded Then
      NewNode.Expanded = True
      treeExecCopy NewNode, SrcList.Child
    ElseIf SrcList.Children > 0 Then
      treeHierarchy.Nodes.Add Target, tvwChild
    End If
    NewNode.Text = SrcList.Text
    Set SrcList = SrcList.Next
  Wend
  Target.Sorted = True
End Sub

' -------------------------------------------------------------------------
' treeCopyNode
'   helper function for drag+drop move/copy (called recusively)
'   called from execCopyMove
'   NB: parentName passed ByRef to avoid copying
' -------------------------------------------------------------------------
Private Sub treeCopyNode(ByVal Node As Node, parentName As String, _
            ByVal srcNode As Node)
   While Not Node Is Nothing
     If Node.Text = parentName Then
       If Node.Expanded Then
         treeExecCopy treeHierarchy.Nodes.Add(Node, tvwChild, , srcNode.Text), _
                      srcNode.Child
         Node.Sorted = True
       ElseIf Node.Children = 0 Then
         treeHierarchy.Nodes.Add Node, tvwChild
       End If
     ElseIf Node.Expanded Then
       treeCopyNode Node.Child, parentName, srcNode
     End If
     Set Node = Node.Next
   Wend
End Sub

' -------------------------------------------------------------------------
' treeLocate
'   Moves tree to a specific position
' -------------------------------------------------------------------------
Private Sub treeLocate(ByVal Path As String)
  Dim pathVect As New HStrVect
  pathVect.Split Path, vbTab
  
  Dim Node As Node, Found As Node
  Set Node = treeHierarchy.Nodes(1)
  
  Dim i As Long
  For i = 1 To pathVect.Count
    Dim piece As String
    piece = UCase(pathVect(i))
    Do Until Node Is Nothing
      If UCase(Node.Text) = piece Then Exit Do
      Set Node = Node.Next
    Loop
    If Node Is Nothing Then Exit For
    If Not Node.Expanded Then
      If Node.Children > 0 Then Node.Expanded = True
    End If
    Set Found = Node
    Set Node = Node.Child
  Next
  If Not Found Is Nothing Then
    Set treeHierarchy.SelectedItem = Found
    Found.EnsureVisible
  End If
End Sub

' -------------------------------------------------------------------------
' treeNodePath
'   Moves tree to a specific position
' -------------------------------------------------------------------------
Private Function treeNodePath(ByVal Node As Node) As String
  Dim FullPath As String
  FullPath = Node.Text
  Set Node = Node.Parent
  Do Until Node Is Nothing
    FullPath = Node.Text & vbTab & FullPath
    Set Node = Node.Parent
  Loop
  treeNodePath = FullPath
End Function

' -------------------------------------------------------------------------
' treeRebold
'   helper function for rebolding tree
' -------------------------------------------------------------------------
Private Sub treeRebold(ByVal Node As Node)
  While Not Node Is Nothing
    Dim oldBold As Boolean
    oldBold = Node.Bold
    Node.Bold = m_boldTerms.Find(Node.Text) > 0
    If Node.Expanded Then
      If oldBold Or Node.Bold Then
        treeRebold Node.Child
      End If
    End If
    Set Node = Node.Next
  Wend
End Sub

' -------------------------------------------------------------------------
' refreshBold
'   Called after any edit to see if need to rebold hierarchy
' -------------------------------------------------------------------------
Private Sub refreshBold()
  Dim isSame As Boolean
  
  If m_leadTerm <> "" Then
    Dim boldTerms As HStrVect
    Set boldTerms = m_curHier.Ancestors(m_leadTerm)
    boldTerms.Add m_leadTerm
    If Not m_boldTerms Is Nothing Then
      isSame = m_boldTerms.IsEqual(boldTerms)
    End If
    
    If Not isSame Then
      Set m_boldTerms = boldTerms
      ' need to rebold tree
      treeRebold treeHierarchy.Nodes(1)
    End If
  End If

End Sub

' -------------------------------------------------------------------------
' execDelTerm
'   Top level function for deleting term from hierarchy
' -------------------------------------------------------------------------
Private Sub execDelTerm(ByVal Term As String, ByVal Parent As String)
  Dim hierVect As HStrVect
  Dim msg As String
  
  ' validate
  msg = m_curHier.ValidateMove(Term, Parent, "")
  If msg = "" And m_curHier.NumParents(Term) = 1 Then
    If Term <> m_leadTerm Then
      msg = "Cannot remove term from hierarchy"
    Else
      Set hierVect = m_termKv.Object(g_sHierarchy)
      If hierVect.Find(m_curHier.Root) = 0 Then
        msg = "Thought you could do that did ya!"
      End If
    End If
  End If
  
  ' report error if any
  If msg <> "" Then
    MsgBox msg, vbOKOnly, "Cannot remove term"
    Exit Sub
  End If
  
  ' confirm
  msg = "Are you sure you wish to remove '" & Term & _
        "'  from '" & Parent & "' ?"
  If MsgBox(msg, vbYesNo + vbDefaultButton2, "Confirmation") <> vbYes Then
    Exit Sub
  End If
  
  ' update hierarchy
  m_curHier.DoMove Term, Parent, ""

  ' update tree control
  treeDelNode treeHierarchy.Nodes(1), Parent, Term
  
  ' update bolding if required
  refreshBold
  
  ' if final delete
  If Not hierVect Is Nothing Then
    hierVect.Remove m_curHier.Root
    fillDescriptorType
  End If
End Sub

' -------------------------------------------------------------------------
' execAddTerm
'   Top level function for adding term to hierarchy
'   Returns True if item added
' -------------------------------------------------------------------------
Private Function execAddTerm(ByVal Term As String, _
                 ByVal Parent As String) As Boolean
  Dim hierVect As HStrVect
  Dim msg As String
  
  ' validate
  If Not m_curHier.IsTerm(Term) Then
    If Term <> m_leadTerm Then
      msg = "Cannot add new term to hierarchy"
    Else
      Set hierVect = m_termKv.Object(g_sHierarchy)
    End If
  End If
  
  If msg = "" Then msg = m_curHier.ValidateMove(Term, "", Parent)
  
  ' report error if any
  If msg <> "" Then
    MsgBox msg, vbOKOnly, "Cannot add term"
    Exit Function
  End If
  
  ' confirmation not required for add
  
  ' update hierarchy
  m_curHier.DoMove Term, "", Parent

  ' update tree control
  treeAddNode treeHierarchy.Nodes(1), Parent, Term
  
  ' update bolding if required
  refreshBold
 
  ' if first insert
  If Not hierVect Is Nothing Then
    If hierVect.Add(m_curHier.Root) Then
      fillDescriptorType
    End If
  End If
  
  ' return
  execAddTerm = True
End Function

' -------------------------------------------------------------------------
' execCopyMove
'   Top level function for drag and drop move or copy
' -------------------------------------------------------------------------
Private Sub execCopyMove(ByVal srcNode As Node, ByVal Parent As String, _
            ByVal isCopy As Boolean)
  ' validate
  Dim msg As String
  If isCopy Then
    msg = m_curHier.ValidateMove(srcNode.Text, "", Parent)
  Else
    msg = m_curHier.ValidateMove(srcNode.Text, srcNode.Parent.Text, Parent)
  End If
  
  ' report error if any
  If msg <> "" Then
    MsgBox msg, vbOKOnly, IIf(isCopy, "Cannot copy term", "Cannot move term")
    Exit Sub
  End If
  
  ' confirm
  If isCopy Then
    msg = "Copy '" & srcNode.Text & "' to '" & Parent & "'"
  Else
    msg = "Move '" & srcNode.Text & "' to '" & Parent & "'"
  End If
  If MsgBox(msg, vbYesNo + vbDefaultButton2, "Confirmation") <> vbYes Then
    Exit Sub
  End If
  
  ' update hierarchy
  If isCopy Then
    m_curHier.DoMove srcNode.Text, "", Parent
  Else
    m_curHier.DoMove srcNode.Text, srcNode.Parent.Text, Parent
  End If
  
  ' update tree control
  treeCopyNode treeHierarchy.Nodes(1), Parent, srcNode
  If Not isCopy Then
    treeDelNode treeHierarchy.Nodes(1), srcNode.Parent.Text, srcNode.Text
  End If
  
  ' update bolding if required
  refreshBold
 
End Sub

' -------------------------------------------------------------------------
' fillHierList
'    Populates hierarchy combo with available hierarchies
' -------------------------------------------------------------------------
Sub fillHierList()
  Dim hierList As HStrVect
  Dim i As Long
  Dim isFullList As Boolean
  
  If chkInUseOnly.Value Then
    Set hierList = m_termKv.Object(g_sHierarchy)
  Else
    Dim DesTypeInfo As HKvList
    Set DesTypeInfo = m_thesObj.DesTypeInfo(m_termKv(g_sDescriptorType))
    If Not DesTypeInfo Is Nothing Then
      Set hierList = DesTypeInfo.Object("Hierarchies")
    Else
      Set hierList = m_thesObj.HierManager.AllHierarchies
      isFullList = True
    End If
  End If
  
  cboHierarchy.Clear
  If Not hierList Is Nothing Then
    For i = 1 To hierList.Count
      cboHierarchy.AddItem hierList(i)
    Next
  
    If Not m_curHier Is Nothing Then
      HUtil.SelectCombo cboHierarchy, m_curHier.Root
    ElseIf Not isFullList And cboHierarchy.ListCount > 0 Then
      cboHierarchy.ListIndex = 0
    End If
  End If
  
  If cboHierarchy.ListIndex < 0 Then cboHierarchy_Click
    
End Sub

' -------------------------------------------------------------------------
' putGroup
' -------------------------------------------------------------------------
Sub putGroup(ByVal Tag As String)
  Dim termVect As HStrVect
  Set termVect = m_termKv.Object(Tag)
  m_groupTabTag = Tag
  
  lstGroup.Clear
  Dim i As Long
  For i = 1 To termVect.Count
    lstGroup.AddItem termVect(i)
  Next
End Sub

' -------------------------------------------------------------------------
' getGroup
' -------------------------------------------------------------------------
Sub getGroup()
  If m_groupTabTag <> "" Then
    Dim termVect As HStrVect
    Set termVect = m_termKv.Object(m_groupTabTag)
    termVect.Clear
    Dim i As Long
    For i = 0 To lstGroup.ListCount - 1
      termVect.Add lstGroup.List(i)
    Next
  End If
End Sub

' -------------------------------------------------------------------------
' prepareTabsForDesType
' -------------------------------------------------------------------------
Sub prepareTabsForDesType()
  Dim DesType As String
  DesType = m_termKv(g_sDescriptorType)
  Dim DesTypeInfo As HKvList
  Set DesTypeInfo = m_thesObj.DesTypeInfo(DesType)
  Dim i As Long
  
  Dim bEnable As Boolean
  Dim termVect As HStrVect
  ' class + member
  tabGroup.Tabs.Clear
  If Not DesTypeInfo Is Nothing Then
    If DesTypeInfo("MemberRelCode") <> "" Then
      tabGroup.Tabs.Add , g_sMemberTerm, DesTypeInfo("MemberCaption")
      m_groupSqlWhere(g_sMemberTerm) = DesTypeInfo("MemberSqlWhere")
    End If
    If DesTypeInfo("GroupRelCode") <> "" Then
      tabGroup.Tabs.Add , g_sGroupTerm, DesTypeInfo("GroupCaption")
      m_groupSqlWhere(g_sGroupTerm) = DesTypeInfo("GroupSqlWhere")
    End If
  End If
  m_groupTabTag = ""
  If tabGroup.Tabs.Count > 0 Then
    bEnable = True
    putGroup tabGroup.Tabs(1).key
  End If
  tabGroup.Enabled = bEnable
  lstGroup.Enabled = bEnable
  cmdAddGroup.Enabled = bEnable
  cmdDelGroup.Enabled = bEnable
  lstGroup.BackColor = IIf(bEnable, vbWindowBackground, BackColor)
  
  ' indicators
  Dim indicators As HStrVect
  If Not DesTypeInfo Is Nothing Then
    Set indicators = DesTypeInfo.Object("Indicators")
  End If
  bEnable = Not indicators Is Nothing
  lstIndicator.Clear
  If bEnable Then
    Set termVect = m_termKv.Object(g_sIndicator)
    For i = 1 To indicators.Count
      lstIndicator.AddItem indicators(i)
      If termVect.Find(indicators(i)) > 0 Then
        lstIndicator.Selected(lstIndicator.NewIndex) = True
      End If
    Next
  End If
  
  lstIndicator.BackColor = IIf(bEnable, vbWindowBackground, BackColor)
  frameIndicator.Enabled = bEnable
  
  ' update hierarchy combo if required
  If Not chkInUseOnly.Value Then
    m_refreshHierFlag = True
  End If
End Sub

' -------------------------------------------------------------------------
' putRecord
' -------------------------------------------------------------------------
Sub putRecord()
  Dim termVect As HStrVect
  Dim i As Long
  ' put term details
  txtPrintForm = m_termKv(g_sDisplayForm)
  txtTerm = m_termKv(g_sLeadTerm)
  
  ' use for terms
  Set termVect = m_termKv.Object(g_sUsedFor)
  If termVect.Count = 0 Then
    txtUsedFor = ""
  Else
    txtUsedFor = termVect.Join(vbCrLf) & vbCrLf
  End If
  
  ' related term
  lstRelated.Clear
  Set termVect = m_termKv.Object(g_sRelatedTerm)
  For i = 1 To termVect.Count
    lstRelated.AddItem termVect(i)
  Next
  
  ' notes
  Dim rowVect As HRowVect
  Set rowVect = m_termKv.Object(g_sScopeNote)
  Dim noteTexts As New HStrVect
  rowVect.ExtractStrVect noteTexts, g_sNoteText
  If rowVect.Count > 0 Then
    txtScopeNotes = noteTexts.Join(vbCrLf) & vbCrLf
  Else
    txtScopeNotes = ""
  End If
  
  Set rowVect = m_termKv.Object(g_sDateNote)
  lstDateNotes.ListItems.Clear
  For i = 1 To rowVect.Count
    Dim Item As ListItem
    Set Item = lstDateNotes.ListItems.Add
    Item.Text = HUtil.DateToDisplayStr(rowVect(i)(g_sNoteDate))
    Item.SubItems(1) = rowVect(i)(g_sNoteText)
  Next
  
End Sub

' -------------------------------------------------------------------------
' getRecord
' -------------------------------------------------------------------------
Function getRecord() As Boolean
  Dim termVect As HStrVect
  Dim i As Long
  
  m_termKv(g_sLeadTerm) = txtTerm.Text
  m_termKv(g_sDisplayForm) = txtPrintForm.Text
  If cboDesType <> m_termKv(g_sDescriptorType) Then
    m_termKv(g_sDescriptorType) = cboDesType.Text
    prepareTabsForDesType
  End If
  
  ' use for terms
  Set termVect = m_termKv.Object(g_sUsedFor)
  termVect.Split txtUsedFor, vbCrLf
  termVect.Remove ""
  
  ' related terms
  Set termVect = m_termKv.Object(g_sRelatedTerm)
  termVect.Clear
  For i = 0 To lstRelated.ListCount - 1
    termVect.Add lstRelated.List(i)
  Next
  
  ' indicators
  Set termVect = m_termKv.Object(g_sIndicator)
  termVect.Clear
  For i = 0 To lstIndicator.ListCount - 1
    If lstIndicator.Selected(i) Then
      termVect.Add lstIndicator.List(i)
    End If
  Next
  
  ' group / member terms
  If tabThesTerm.Tab = 1 Then getGroup
  
  ' notes
  Dim rowVect As HRowVect
  Set rowVect = m_termKv.Object(g_sScopeNote)
  rowVect.Clear
  Dim noteTexts As New HStrVect
  noteTexts.Split txtScopeNotes, vbCrLf
  For i = 1 To noteTexts.Count
    noteTexts(i) = HUtil.TrimSpaces(noteTexts(i))
  Next
  noteTexts.Remove ""
  rowVect.FillFromStrVect noteTexts, g_sNoteText
  For i = 1 To rowVect.Count
    rowVect(i)(g_sNoteId) = i
  Next
  
  getRecord = True
End Function

' -------------------------------------------------------------------------
' focusError
' -------------------------------------------------------------------------
Sub focusError(ByVal errInfo As HKvList)

  Dim field As String
  Dim ctl As Control
  Dim tabIdx As Integer
  
  tabIdx = 0
  Select Case errInfo("Field")
    Case g_sLeadTerm
      Set ctl = txtTerm
    Case g_sDisplayForm
      Set ctl = txtPrintForm
    Case g_sDescriptorType
      Set ctl = cboDesType
    Case g_sUsedFor
      Set ctl = txtUsedFor
    Case g_sScopeNote
      Set ctl = txtScopeNotes
    Case g_sRelatedTerm
      Set ctl = lstRelated
      tabIdx = 1
  End Select

  If Not ctl Is Nothing Then
    tabThesTerm.Tab = tabIdx
    If ctl.Enabled Then ctl.SetFocus
  End If

End Sub

' -------------------------------------------------------------------------
' saveChanges
'   Does database update
'   Returns True on success
'   Called from btnOk_click
' -------------------------------------------------------------------------
Private Function SaveChanges(ByVal checkOnly As Boolean) As Boolean
  Dim ok As Boolean
  Dim errInfo As New HKvList
  Dim key As New HKvList
  Dim pass As Integer
  
  
  ' transfer form to m_record
  If Not getRecord Then
    Exit Function
  End If
  
  m_thesObj.Normalise m_termKv
  MousePointer = vbHourglass
  ok = m_thesObj.ValidateObj(m_context, m_isNew, key, m_termKv, _
                              m_orig, errInfo)
  If Not ok Then
    ' a validation error occurred
    ' load form with normalised data
    putRecord
    ' set cursor on field
    focusError errInfo
    ' display error
    MsgBox errInfo("Field") & ": " & errInfo("ErrMsg"), _
           vbOKOnly, "Data entry error"
  ElseIf checkOnly Then
    ' when checking, always refresh because of Normalise
    putRecord
  Else
    ' commit changes
    ok = m_context.DB.BeginTransaction
    If ok Then
      Call m_thesObj.SaveObj(m_context, m_isNew, key, m_termKv, m_orig)
      ok = m_context.DB.EndTransaction(True)
    End If
    If ok Then
      m_thesObj.HierManager.OnCommit
    Else
      putRecord
    End If
  End If
  MousePointer = vbDefault

  SaveChanges = ok
End Function

' -------------------------------------------------------------------------
' cboHierarchy_Click
'   Called when selected hierarchy changed
' -------------------------------------------------------------------------
Private Sub cboHierarchy_Click()
  Dim Manager As HierMgr
  Set Manager = m_thesObj.HierManager
  
  ' if no change to hierarchy just return
  If Not m_curHier Is Nothing Then
    If m_curHier.Root = cboHierarchy.Text Then Exit Sub
  End If
  If cboHierarchy.ListIndex >= 0 Then
    Manager.OpenHier cboHierarchy, m_context
  End If
  If treeHierarchy.Nodes.Count > 0 Then
    treeHierarchy.Nodes(1).Expanded = False
    treeHierarchy.Nodes.Clear
  End If
  treeInverted.Nodes.Clear
  Set m_curHier = Manager.Hierarchy(cboHierarchy)
  Set m_boldTerms = Nothing
  If Not m_curHier Is Nothing Then
    Dim RootNode As Node
    Set RootNode = treeHierarchy.Nodes.Add(, , , m_curHier.Root)
    checkNode RootNode
    refreshBold
    RootNode.Expanded = True
  End If
End Sub

' -------------------------------------------------------------------------
' cboLocate_Click
' -------------------------------------------------------------------------
Private Sub cboLocate_Click()
  If cboLocate.Text <> "" And Not m_curHier Is Nothing Then
    Dim FullPath As String
    FullPath = m_curHier.MoveNext(cboLocate.Text, "")
    If FullPath <> "" Then
      treeLocate FullPath
    End If
  End If
End Sub

' -------------------------------------------------------------------------
' cboLocate_DropDown
' -------------------------------------------------------------------------
Private Sub cboLocate_DropDown()
  Dim pattStr As String
  pattStr = cboLocate.Text
  cboLocate.Clear
  If Not m_curHier Is Nothing Then
    Dim matchTerms As HStrVect
    Set matchTerms = m_curHier.MatchingTerms(pattStr & "*")
    matchTerms.Sort
    Dim i As Long
    For i = 1 To matchTerms.Count
      cboLocate.AddItem matchTerms(i)
    Next
  End If
  cboLocate = pattStr
End Sub

' -------------------------------------------------------------------------
' chkInUseOnly_Click
'   Called to change list of visible hierarchies
' -------------------------------------------------------------------------
Private Sub chkInUseOnly_Click()
  fillHierList
End Sub

' -------------------------------------------------------------------------
' cmdAddGroup
' -------------------------------------------------------------------------
Private Sub cmdAddGroup_Click()
  On Error GoTo ErrHandler
  Dim popup As Object
  Dim args As New HKvList, results As New HKvList

  ' set arguments
  args("Term") = ""
  args("Validate") = "Y"
  args("SqlWhere") = m_groupSqlWhere(m_groupTabTag)
  
  ' create and call popup object
  Set popup = CreateObject("AdisThesPopup.Popup")
  If popup.Invoke(Me, args, results, m_context) Then
    lstGroup.AddItem results("Term")
  End If

Exit Sub
ErrHandler:
  Screen.MousePointer = vbDefault
  m_context.ReportException m_module, "cmdAddGroup_Click"
End Sub

' -------------------------------------------------------------------------
' cmdAddRelated
' -------------------------------------------------------------------------
Private Sub cmdAddRelated_Click()
  On Error GoTo ErrHandler
  Dim popup As Object
  Dim args As New HKvList, results As New HKvList

  ' set arguments
  args("Term") = ""
  args("Validate") = "Y"
  
  ' create and call popup object
  Set popup = CreateObject("AdisThesPopup.Popup")
  If popup.Invoke(Me, args, results, m_context) Then
    lstRelated.AddItem results("Term")
  End If

Exit Sub
ErrHandler:
  Screen.MousePointer = vbDefault
  m_context.ReportException m_module, "cmdAddRelated_Click"
End Sub

' -------------------------------------------------------------------------
' cmdCancel_Click
' -------------------------------------------------------------------------
Private Sub cmdCancel_Click()
  Unload Me
End Sub

' -------------------------------------------------------------------------
' cmdDelete_Click
'    Deletes thesaurus term
' -------------------------------------------------------------------------
Private Sub cmdDelete_Click()
  
  ' should not get here if new term
  If m_isNew Then Exit Sub
  
  ' validate not in any hierarchy
  Dim hierVect As HStrVect
  Set hierVect = m_termKv.Object(g_sHierarchy)
  If hierVect.Count > 0 Then
    MsgBox "Term must first be removed from all hierarchies"
    Exit Sub
  End If
  
  ' confirm deletion
  If MsgBox("Are you sure you wish to delete this thesaurus term?", _
            vbYesNo, "Confirm deletion") <> vbYes Then
    Exit Sub
  End If
  
  ' do delete
  Dim key As New HKvList
  key(g_sLeadTerm) = m_leadTerm
  
  Dim ok As Boolean
  ok = m_context.DB.BeginTransaction
  If ok Then
    Call m_thesObj.DeleteObj(m_context, key)
    ok = m_context.DB.EndTransaction(True)
  End If
  If ok Then
    m_thesObj.HierManager.OnCommit
    Hide
    Unload Me
  End If
End Sub

' -------------------------------------------------------------------------
' cmdDelGroup
' -------------------------------------------------------------------------
Private Sub cmdDelGroup_Click()
  If lstGroup.ListIndex >= 0 Then
    lstGroup.RemoveItem lstGroup.ListIndex
  End If
End Sub

' -------------------------------------------------------------------------
' cmdDelRelated
' -------------------------------------------------------------------------
Private Sub cmdDelRelated_Click()
  If lstRelated.ListIndex >= 0 Then
    lstRelated.RemoveItem lstRelated.ListIndex
  End If
End Sub

' -------------------------------------------------------------------------
' cmdHelp_Click
' -------------------------------------------------------------------------
Private Sub cmdHelp_Click()
  SendKeys "{F1}"
End Sub

' -------------------------------------------------------------------------
' cmdOK_Click
' -------------------------------------------------------------------------
Private Sub cmdOK_Click()
  If SaveChanges(False) Then
    Hide
    Unload Me
  End If
End Sub

' -------------------------------------------------------------------------
' cmdSave_Click
' -------------------------------------------------------------------------
Private Sub cmdSave_Click()
  If SaveChanges(False) Then
    ' saved - ok
    ' firstly, after rename need to reload hierarchies
    If StrComp(m_leadTerm, m_termKv(g_sLeadTerm), vbBinaryCompare) <> 0 Then
      cboHierarchy.ListIndex = -1
      m_thesObj.HierManager.CloseAll
    End If
    Set m_orig = m_thesObj.CloneObj(m_termKv)
    m_leadTerm = m_termKv(g_sLeadTerm)
    m_isNew = False
    Caption = m_leadTerm
    txtTerm.Enabled = True
    txtTerm.BackColor = vbWindowBackground
    putRecord
    MsgBox "Thesaurus successfully updated"
  End If
End Sub

' -------------------------------------------------------------------------
' Form_Load
' -------------------------------------------------------------------------
Private Sub Form_Load()
  
  Set Icon = m_context.Icon
  Set m_moveCursor = treeHierarchy.DragIcon
  Set m_copyCursor = treeInverted.DragIcon

  Me.HelpContextID = Me.HelpContextID + tabThesTerm.Tab
  tabThesTerm.Tab = 0
  
  m_ok = True
  ' lock thesaurus
  ' don't do this anymore so that multiple users can use!!
'  m_isLocked = m_context.TryLock("THESAURUS", False)
'  m_ok = m_isLocked
  
  ' load record
  If m_ok Then
    If m_isNew Then
      m_ok = m_thesObj.NewObj(m_context, m_termKv)
    Else
      Dim key As New HKvList '
      key(g_sLeadTerm) = m_leadTerm
      m_ok = m_thesObj.LoadObj(m_context, key, m_termKv)
      If m_ok And m_termKv.Count() = 0 Then
        MsgBox "Unable to find thesaurus term '" & m_leadTerm & "'", _
               vbExclamation
      End If
    End If
  End If
  
  If m_ok Then
    Set m_orig = m_thesObj.CloneObj(m_termKv)
    m_leadTerm = m_termKv(g_sLeadTerm)
    
    ' set maximum widths of edit boxes
    txtTerm.MaxLength = MAX_TERM_LEN
    txtPrintForm.MaxLength = MAX_TERM_LEN
  
    ' set caption
    If Not m_isNew Then
      Caption = m_leadTerm
    End If
    cmdDelete.Enabled = Not m_isNew
  
    ' copy fields to form
    chkInUseOnly = IIf(m_isNew, 0, 1)
    putRecord
    fillDescriptorType
    prepareTabsForDesType
    m_refreshHierFlag = True
  End If
  
End Sub

' -------------------------------------------------------------------------
' Form_QueryUnload
' -------------------------------------------------------------------------
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If Me.Visible Then
    Const msg As String = _
      "Are you sure you wish to abandon editing thesaurus term?"
    If MsgBox(msg, vbYesNo + vbDefaultButton2 + vbQuestion) <> vbYes Then
      Cancel = True
    End If
  End If
End Sub

' -------------------------------------------------------------------------
' Form_Unload
' -------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
  If m_isLocked Then
    m_context.TryUnlock "THESAURUS"
  End If
End Sub

' -------------------------------------------------------------------------
' tabGroup_Click
' -------------------------------------------------------------------------
Private Sub tabGroup_Click()
   getGroup
   putGroup tabGroup.SelectedItem.key
End Sub

' -------------------------------------------------------------------------
' tabThesTerm_Click
' -------------------------------------------------------------------------
Private Sub tabThesTerm_Click(PreviousTab As Integer)
  Me.HelpContextID = Me.HelpContextID + tabThesTerm.Tab - PreviousTab
  ' if changing away from first tab, check descriptor type
  If PreviousTab = 0 Then
    If cboDesType <> m_termKv(g_sDescriptorType) Then
      m_termKv(g_sDescriptorType) = cboDesType.Text
      prepareTabsForDesType
    End If
    If m_isNew Then
      m_leadTerm = IIf(cboDesType.ListIndex >= 0, txtTerm.Text, "")
    End If
  End If
  If PreviousTab = 1 Then
    getGroup
  End If
  If tabThesTerm.Tab = 2 Then
    If m_refreshHierFlag Then fillHierList
    m_refreshHierFlag = False
  End If
  cmdDelete.Enabled = Not m_isNew And tabThesTerm.Tab = 0
  
End Sub

' -------------------------------------------------------------------------
' timerScroll_Timer
'    Called when scroll timer times out
' -------------------------------------------------------------------------
Private Sub timerScroll_Timer()
  SendMessage treeHierarchy.hwnd, 277&, IIf(m_scrollUp, 0, 1), vbNull
End Sub

' -------------------------------------------------------------------------
' treeHierachy_Collapse
' -------------------------------------------------------------------------
Private Sub treeHierarchy_Collapse(ByVal Node As MSComctlLib.Node)
   checkNode Node
End Sub

' -------------------------------------------------------------------------
' treeHierachy_Expand
' -------------------------------------------------------------------------
Private Sub treeHierarchy_Expand(ByVal Node As MSComctlLib.Node)
  checkNode Node
End Sub

' -------------------------------------------------------------------------
' treeHierachy_KeyDown
' -------------------------------------------------------------------------
Private Sub treeHierarchy_KeyDown(KeyCode As Integer, Shift As Integer)
  With treeHierarchy
    If KeyCode = vbKeyDelete Then
      ' delete key pressed: unlink current node
      If Not .SelectedItem Is Nothing Then
        If Not .SelectedItem.Parent Is Nothing Then
          execDelTerm .SelectedItem.Text, .SelectedItem.Parent.Text
        End If
      End If
    ElseIf KeyCode = vbKeyInsert And (Shift And vbShiftMask) <> 0 Then
      ' shift + insert pressed: insert current thesaurus term
      If Not .SelectedItem Is Nothing Then
        If m_leadTerm = "" Then
          Beep
        Else
          If execAddTerm(m_leadTerm, .SelectedItem.Text) Then
            .SelectedItem.Expanded = True
          End If
        End If
      End If
    ElseIf KeyCode = vbKeyF3 Then
      ' F3: find next item with same name
      If Not .SelectedItem Is Nothing Then
        Dim Path As String, NewPath As String
        Path = treeNodePath(.SelectedItem)
        NewPath = m_curHier.MoveNext(.SelectedItem.Text, Path)
        If Path <> NewPath Then
          treeLocate NewPath
        End If
      End If
    End If
  End With
End Sub

' -------------------------------------------------------------------------
' treeHierachy_MouseDown
'   Called when mouse button downed inside tree control
' -------------------------------------------------------------------------
Private Sub treeHierarchy_MouseDown(Button As Integer, _
            Shift As Integer, x As Single, y As Single)
  Dim dragNode As Node
  If Button = vbLeftButton Then
    Set dragNode = treeHierarchy.HitTest(x, y)
    If Not dragNode Is Nothing Then
      Set treeHierarchy.SelectedItem = dragNode
      Set treeHierarchy.DropHighlight = Nothing
      m_inDrag = True
      treeHierarchy.OLEDrag
    End If
  End If
End Sub

Private Sub treeHierarchy_NodeClick(ByVal Node As MSComctlLib.Node)
  fillInverted
End Sub

' -------------------------------------------------------------------------
' treeHierachy_OLECompleteDrag
' -------------------------------------------------------------------------
Private Sub treeHierarchy_OLECompleteDrag(Effect As Long)
  
  If m_inDrag Then
    treeHierarchy.MousePointer = vbDefault
    Set treeHierarchy.DropHighlight = Nothing
    timerScroll.Enabled = False
    m_inDrag = False
    fillInverted
  End If

End Sub

' -------------------------------------------------------------------------
' treeHierachy_OLEDragDrop
' -------------------------------------------------------------------------
Private Sub treeHierarchy_OLEDragDrop(Data As MSComctlLib.DataObject, _
            Effect As Long, Button As Integer, Shift As Integer, _
            x As Single, y As Single)
  treeHierarchy_OLEDragOver Data, Effect, Button, Shift, x, y, 0
  With treeHierarchy
    If m_inDrag And Not .DropHighlight Is Nothing Then
      Dim srcNode As Node
      Set srcNode = treeHierarchy.SelectedItem
        If Not srcNode Is .DropHighlight Then
          Set .SelectedItem = .DropHighlight
          execCopyMove srcNode, .DropHighlight.Text, _
                       (Effect And vbDropEffectCopy) <> 0
        End If
    End If
  End With
End Sub

' -------------------------------------------------------------------------
' treeHierachy_OLEDragOver
' -------------------------------------------------------------------------
Private Sub treeHierarchy_OLEDragOver(Data As MSComctlLib.DataObject, _
            Effect As Long, Button As Integer, Shift As Integer, _
            x As Single, y As Single, State As Integer)
  Dim hitNode As Node
  
  If m_inDrag Then
    If State = 1 Then
      ' cursor leaving window
      timerScroll.Enabled = False
      Exit Sub
    End If
    Effect = IIf(Shift And vbCtrlMask, vbDropEffectCopy, vbDropEffectMove)
    With treeHierarchy
      Set hitNode = .HitTest(x, y)
      If hitNode Is Nothing Then
        Effect = vbDropEffectNone
      ElseIf Not hitNode Is .SelectedItem Or Not .DropHighlight Is Nothing Then
        Set .DropHighlight = hitNode
      End If
      m_scrollUp = y + y <= .Height
      Dim t As Integer
      t = IIf(m_scrollUp, y, .Height - y)
      timerScroll.Enabled = t < 300
      timerScroll.Interval = IIf(t < 50, 50, t)
    End With
  End If
End Sub

' -------------------------------------------------------------------------
' treeHierachy_OLEGiveFeedback
' -------------------------------------------------------------------------
Private Sub treeHierarchy_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
  If Effect And (vbDropEffectMove + vbDropEffectCopy) Then
    Set treeHierarchy.MouseIcon = _
       IIf(Effect And vbDropEffectCopy, m_copyCursor, m_moveCursor)
    treeHierarchy.MousePointer = vbCustom
    DefaultCursors = False
  End If
End Sub

' -------------------------------------------------------------------------
' treeHierachy_OLEStartDrag
' -------------------------------------------------------------------------
Private Sub treeHierarchy_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
  Data.SetData treeHierarchy.SelectedItem.Text, vbCFText
  AllowedEffects = vbDropEffectCopy + vbDropEffectMove
End Sub

' -------------------------------------------------------------------------
' txtPrintForm_Change
' -------------------------------------------------------------------------
Private Sub txtPrintForm_Change()
  If txtTerm.Enabled And m_isNew Then
    If StrComp(Replace(m_termKv(g_sDisplayForm), " ", "-"), _
       txtTerm, vbTextCompare) = 0 Then
      txtTerm = HUtil.UCaseFirst(Replace(txtPrintForm, " ", "-"))
    End If
  End If
  m_termKv(g_sDisplayForm) = txtPrintForm.Text
End Sub
