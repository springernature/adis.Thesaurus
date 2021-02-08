VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmThesaurus 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Thesaurus Term Selection"
   ClientHeight    =   7740
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   6165
   HelpContextID   =   10050
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   6165
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdRemoveTerm 
      Caption         =   "« R&emove"
      Height          =   375
      Left            =   180
      TabIndex        =   8
      Top             =   4050
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddTerm 
      Caption         =   "&Add »"
      Height          =   375
      Left            =   180
      TabIndex        =   7
      Top             =   3570
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstLeadTerms 
      Height          =   1995
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   3519
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   9525
      EndProperty
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   372
      Left            =   3060
      TabIndex        =   16
      Top             =   7200
      Width           =   1212
   End
   Begin TabDlg.SSTab tabInfo 
      Height          =   2355
      Left            =   120
      TabIndex        =   10
      Top             =   4680
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4154
      _Version        =   393216
      Style           =   1
      Tab             =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Related Terms"
      TabPicture(0)   =   "frmThesaurus.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "treeRelated"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Notes"
      TabPicture(1)   =   "frmThesaurus.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtNotes"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Used For"
      TabPicture(2)   =   "frmThesaurus.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lstUsedFor"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.ListBox lstUsedFor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         Left            =   120
         TabIndex        =   13
         Top             =   420
         Width           =   5655
      End
      Begin MSComctlLib.TreeView treeRelated 
         Height          =   1815
         Left            =   -74880
         TabIndex        =   11
         Top             =   420
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   3201
         _Version        =   393217
         Indentation     =   529
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
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
      Begin VB.TextBox txtNotes 
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   -74880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   420
         Width           =   5655
      End
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   372
      Left            =   5220
      TabIndex        =   5
      Top             =   600
      Width           =   852
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
      Height          =   345
      Left            =   600
      TabIndex        =   4
      Top             =   600
      Width           =   4515
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
      ItemData        =   "frmThesaurus.frx":0054
      Left            =   4320
      List            =   "frmThesaurus.frx":0070
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   60
      Width           =   1755
   End
   Begin VB.CheckBox chkSynonyms 
      Caption         =   "Show S&ynonyms"
      Height          =   252
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1620
      TabIndex        =   15
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   180
      TabIndex        =   14
      Top             =   7200
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstSelectedTerms 
      Height          =   1515
      Left            =   1440
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   2672
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   8112
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Selected Terms:"
      Height          =   195
      Left            =   195
      TabIndex        =   17
      Top             =   3240
      Width           =   1155
   End
   Begin VB.Label lblLabel 
      Caption         =   "&Term:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   660
      Width           =   495
   End
   Begin VB.Label lblLabel 
      Caption         =   "&Descriptor Type:"
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmThesaurus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' **********************************************************************
'
' Module: $Id: frmThesaurus.frm,v 1.15 2004/03/30 04:06:32 GS2 Exp $
'
' Author: Andrew Howroyd
'
' Created: May 1999
'
' Description
'
' **********************************************************************
Option Explicit

' variables set by caller
Public m_context As HAppContext
Public m_thesObj As ThesBrowserObj
Public m_isModeless As Boolean
Public m_navigate As String
Public m_validate As Boolean
Public m_descriptorOnly As Boolean
Public m_endpointOnly As Boolean
Public m_sqlWhere As String
Public m_term As String
Public m_ok As Boolean
Public m_multipleAdd As Boolean

' internal stuff
Private m_isVerified As Boolean

' -------------------------------------------------------------------------
' fillDesTypeCombo
' -------------------------------------------------------------------------
Sub fillDesTypeCombo()
  If m_endpointOnly Then
    cboDesType.AddItem "endpoints"
  Else
    With cboDesType
      If Not m_descriptorOnly Then .AddItem "all"
      .AddItem "descriptor"
      .AddItem "drug"
      .AddItem "drug or drug class"
      .AddItem "drug class"
      .AddItem "disease"
      .AddItem "company"
      If Not m_descriptorOnly Then .AddItem "endpoints"
      .AddItem "media release type"
      .AddItem "non drug descriptor"
      If Not m_descriptorOnly Then .AddItem "non descriptor"
    End With
  End If
End Sub

' -------------------------------------------------------------------------
' formatTerm
'    Temporary function to format thesaurus terms
'    (deals with issue that thesaurus all upper-cased)
' -------------------------------------------------------------------------
Private Function formatTerm(ByVal term As String) As String
  formatTerm = HUtil.UCaseFirst(LCase(term))
End Function

' -------------------------------------------------------------------------
' addTree
'    Adds some terms to the tree
'    Called from loadTerm
' -------------------------------------------------------------------------
Private Sub addTree(ByVal Parent As Node, _
            ByVal strCaption As String, ByVal vect As HRowVect)
  Dim Folder As Node
  Dim text As String
  Dim i As Long
  
  If vect.Count > 0 Then
    Set Folder = treeRelated.Nodes.Add(Parent, tvwChild, , strCaption)
    For i = 1 To vect.Count
      text = vect(i)(1)
      treeRelated.Nodes.Add Folder, tvwChild, , text
    Next
    Folder.Expanded = True
  End If

End Sub

' -------------------------------------------------------------------------
' loadTerm
'    Loads information for current term
' -------------------------------------------------------------------------
Private Sub loadTerm()
  Dim ok As Boolean
  Dim key As New HKvList, obj As New HKvList
  Dim vect As HRowVect
  Dim i As Long
  
  Screen.MousePointer = vbHourglass
  key(g_sLeadTerm) = m_term
  
  ok = m_thesObj.LoadObj(m_context, key, obj)
  If ok And obj.Count > 0 Then
    ' load up related terms tree control
    Set treeRelated.SelectedItem = Nothing
    If treeRelated.Nodes.Count > 0 Then
      treeRelated.Nodes(1).Expanded = False
      treeRelated.Nodes.Clear
    End If
    Dim Parent As Node
    Set Parent = treeRelated.Nodes.Add(, , , m_term)
    addTree Parent, "Broader Terms", obj.Object("BT")
    addTree Parent, "Narrower Terms", obj.Object("NT")
    addTree Parent, "Related Terms", obj.Object("RT")
    Parent.Expanded = True
    
    ' now do notes
    Dim text As String
    Set vect = obj.Object("ScopeNote")
    If vect.Count > 0 Then
      text = "SCOPE NOTES" & vbCrLf
      For i = 1 To vect.Count
        text = text & vect(i)(g_sNoteText) & vbCrLf
      Next
    End If
    Set vect = obj.Object("DateNote")
    If vect.Count > 0 Then
      If text <> "" Then text = text & vbCrLf
      text = text & "DATE NOTES" & vbCrLf
      For i = 1 To vect.Count
        text = text & HUtil.ISODateToStr(vect(i)(g_sNoteDate)) & _
               ": " & vect(i)(g_sNoteText) & vbCrLf
      Next
    End If
    txtNotes = text

    ' used for terms
    Set vect = obj.Object("UF")
    lstUsedFor.Clear
    For i = 1 To vect.Count
      lstUsedFor.AddItem vect(i)(1)
    Next

    ' enable list control if more than 0 items
    txtTerm = m_term
    tabInfo.Tab = 0
    treeRelated.Enabled = Parent.Children > 0
    If treeRelated.Enabled Then
      treeRelated.SetFocus
      treeRelated.SelectedItem = Parent
    End If
  End If
  Screen.MousePointer = vbDefault
End Sub

' -------------------------------------------------------------------------
' cmdCancel_Click
' -------------------------------------------------------------------------
Private Sub cmdCancel_Click()
  If m_isModeless Then Unload Me Else Hide
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
  Dim errMsg As String
  Dim sDesType As String
  
  If m_validate Then
    If m_multipleAdd Then
      Dim i As Integer
      m_term = ""
      For i = 1 To lstSelectedTerms.ListItems.Count
        m_term = m_term & lstSelectedTerms.ListItems(i) & "#:#"
      Next
    ElseIf m_term = "" Then
      errMsg = "No term entered"
    Else
      Dim obj As New HKvList
      Dim sqlWhere As String
      If m_sqlWhere <> "" Then
        sqlWhere = m_sqlWhere
      ElseIf m_descriptorOnly Then
        sqlWhere = "DesType is not NULL"
      End If
      If Not m_thesObj.FindObj(m_context, m_term, sqlWhere, obj) Then
        Exit Sub ' error retrieving object
      End If
      If obj.Count = 0 Then
        errMsg = "Term '" & m_term & "' not found in thesaurus"
      Else
        m_term = obj(g_sLeadTerm)
        sDesType = obj(g_sDesType)
      End If
    End If
    
    If errMsg <> "" Then
      txtTerm.SetFocus
      MsgBox errMsg, vbOKOnly
      Exit Sub
    End If
  End If

  m_ok = True
  Hide
  
  If m_isModeless Then
    ' No navigation into Company type terms, ie no Edit.
    If m_navigate <> "" Then
      If sDesType <> "COY" Then
        m_context.Navigate m_navigate & "?Term=" & m_term
      Else
        MsgBox "Please edit a company term using company maintenance.", vbOKOnly
      End If
    End If
    Unload Me
  End If
  
End Sub

' -------------------------------------------------------------------------
' cmdSearch_Click
' -------------------------------------------------------------------------
Private Sub cmdSearch_Click()
  Dim args As New HKvList, result As New HKvList
  Dim i As Long
  
  Dim term As String
  term = txtTerm
  m_term = term
  
  ' don't do anything if nothing entered
  If term = "" Then Exit Sub
  
  ' fix search string to be sql search
  If HUtil.StrCSpan(term, 1, "*?%_[") <> Len(term) Then
    term = Replace(term, "*", "%")
    term = Replace(term, "?", "_")
  Else
    term = term & "%"
  End If
  
  ' clear list
  lstLeadTerms.ListItems.Clear
  
  Dim constraint As String
  constraint = m_sqlWhere
  If constraint = "" Then
    Select Case cboDesType.text
      Case "descriptor"
        constraint = "DesType Is Not Null"
      Case "drug"
        constraint = "DesType = 'DRUG'"
      Case "drug or drug class"
        constraint = "(DesType = 'DRUG' or DesType Like 'DC%')"
      Case "drug class"
        constraint = "DesType like 'DC%'"
      Case "disease"
        constraint = "DesType = 'DIS'"
      Case "company"
        constraint = "DesType = 'COY'"
      Case "media release type"
        constraint = "DesType = 'MRT'"
      Case "non drug descriptor"
        constraint = "DesType in ('GEN', 'DIS', 'MCT', 'DEA', 'COY', 'MRT')"
      Case "non descriptor"
        constraint = "DesType is Null"
      Case "endpoints"
        constraint = "leadterm in (select leadterm from thshierarchy " & _
                     "where hierarchy = '$Endpoints')"
      Case Else
        If m_descriptorOnly Then
          constraint = "DesType is Not Null"
        End If
    End Select
  End If
  
  ' search database
  MousePointer = vbHourglass
  args("Pattern") = term
  args("SqlWhere") = constraint
  If chkSynonyms Then args("IncludeSynonyms") = "Y"
  If m_thesObj.Search(m_context, args, result) Then
    Dim vect As HRowVect, r As HRow
    Set vect = result.Object("Terms")
    For i = 1 To vect.Count
      Set r = vect(i)
      term = r(g_sLeadTerm)
      If Not IsEmpty(r(g_sUseTerm)) Then
        term = formatTerm(term) & " -> " & r(g_sUseTerm)
      End If
      lstLeadTerms.ListItems.Add , , term
    Next
    If vect.Count > 0 Then
      lstLeadTerms.SelectedItem = lstLeadTerms.ListItems(1)
      lstLeadTerms_ItemClick lstLeadTerms.SelectedItem
      lstLeadTerms.SetFocus
    Else
      txtTerm.SetFocus
      MsgBox "No matching terms found", vbInformation
    End If
  End If
  MousePointer = vbDefault
End Sub

' -------------------------------------------------------------------------
' cmdAddTerm_Click
' -------------------------------------------------------------------------
Private Sub cmdAddTerm_Click()
  Dim i As Integer
  Dim bFound As Boolean
  Dim term As String
  Dim p As Integer
  
  If lstLeadTerms.SelectedItem Is Nothing Then Exit Sub
  
  term = lstLeadTerms.SelectedItem.text
  p = InStr(term, " -> ")
  If p > 0 Then term = Mid(term, p + 4)
  
  For i = 1 To lstSelectedTerms.ListItems.Count
    If lstSelectedTerms.ListItems.Item(i).text = term Then
      bFound = True
      Exit For
    End If
  Next
  If Not bFound Then
    lstSelectedTerms.ListItems.Add , , term
  Else
    MsgBox "The term '" & term & "' has already been added to the Selected Terms list.", vbExclamation
  End If
  
End Sub

' -------------------------------------------------------------------------
' cmdRemoveTerm_Click
' -------------------------------------------------------------------------
Private Sub cmdRemoveTerm_Click()
  If Not lstSelectedTerms.SelectedItem Is Nothing Then
    lstSelectedTerms.ListItems.Remove lstSelectedTerms.SelectedItem.Index
  End If
End Sub

' -------------------------------------------------------------------------
' Form_Activate
'   When form first entered, want to load info for term
'   Also set focus onto term text
' -------------------------------------------------------------------------
Private Sub Form_Activate()
  If m_ok Then
    m_ok = False
    txtTerm.SetFocus
    ' try to get window painted
    Enabled = False
    DoEvents
    Enabled = True
    
    If m_multipleAdd Then
      lstLeadTerms.Height = 1995 'half
      cmdAddTerm.Visible = True
      cmdRemoveTerm.Visible = True
      lstSelectedTerms.Visible = True
    Else
      lstLeadTerms.Height = 3555 'full
    End If
    
    cboDesType.Clear
    If m_sqlWhere <> "" Then
      cboDesType.Enabled = False
      cboDesType.BackColor = Me.BackColor
    Else
      fillDesTypeCombo
      cboDesType.ListIndex = 0
    End If
    
    If m_term <> "" Then
      If HUtil.StrCSpan(m_term, 1, "*?%_[") <> Len(m_term) Then
        cmdSearch_Click
      Else
        loadTerm
      End If
    End If
    txtTerm.SetFocus
  End If
End Sub

' -------------------------------------------------------------------------
' Form_Load
' -------------------------------------------------------------------------
Private Sub Form_Load()
  
  If Not m_isModeless Then
    Me.BorderStyle = vbFixedDialog
  End If
  
  Icon = m_context.Icon
  txtTerm = m_term
  m_ok = True
End Sub

' -------------------------------------------------------------------------
' lstLeadTerms_Dblclick
' -------------------------------------------------------------------------
Private Sub lstLeadTerms_DblClick()
  If Not lstLeadTerms.SelectedItem Is Nothing Then
    ' fetch term from database and populate form
    loadTerm
  End If
End Sub

' -------------------------------------------------------------------------
' lstLeadTerms_ItemClick
' -------------------------------------------------------------------------
Private Sub lstLeadTerms_ItemClick(ByVal Item As MSComctlLib.ListItem)
  m_term = Item.text
  Dim p As Integer
  p = InStr(m_term, " -> ")
  If p > 0 Then m_term = Mid(m_term, p + 4)

End Sub

' -------------------------------------------------------------------------
' lstLeadTerms_KeyDown
' -------------------------------------------------------------------------
Private Sub lstLeadTerms_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then lstLeadTerms_DblClick
End Sub

' -------------------------------------------------------------------------
' treeRelated_DblClick
' -------------------------------------------------------------------------
Private Sub treeRelated_DblClick()
  Dim selNode As Node
  Set selNode = treeRelated.SelectedItem
  If Not selNode Is Nothing Then
    If selNode.Children = 0 Then
      m_term = selNode.text
      loadTerm
    End If
  End If
End Sub

Private Sub treeRelated_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then treeRelated_DblClick
End Sub

'Private Sub treeRelated_NodeClick(ByVal Node As MSComctlLib.Node)
'  If Node.Children = 0 Then m_term = Node.text Else m_term = txtTerm
'End Sub

' -------------------------------------------------------------------------
' txtTerm_Change
' -------------------------------------------------------------------------
Private Sub txtTerm_Change()
  m_term = txtTerm
  cmdSearch.Enabled = m_term <> ""
End Sub

Private Sub txtTerm_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then cmdSearch_Click
End Sub

' -------------------------------------------------------------------------
' $Log: frmThesaurus.frm,v $
' Revision 1.15  2004/03/30 04:06:32  GS2
' Added multiple add mode.
'
' Revision 1.14  2001/12/10 04:46:48  jd4
' Thesaurus must not add or update Company Terms
'    Search on companies as a descriptor, but
'    cannot edit a company thesaurus term
'
' New level hierarchy Media-release-type
'    Allow search on media-release-type as a descriptor
'
' Revision 1.13  2000/08/15 02:00:37  AH2
' Changes for new Drug-class hierarchies
'
' Revision 1.12  2000/08/01 21:35:29  RB1
' Added support for endpoints hierarchy
'
' Revision 1.11  2000/07/06 07:11:16  GS2
' Changed button tab orders.
'
' Revision 1.10  1999/12/13 22:57:40  AH2
' Changed dialog style to include minimise button
'
' Revision 1.9  1999/11/26 01:25:48  AH2
' Added used for tab
' Help file integration
'
' Revision 1.8  1999/11/05 00:21:47  AH2
' Font size change
'
' Revision 1.7  1999/09/24 02:50:54  AH2
' Fixed static text size
'
' Revision 1.6  1999/09/12 21:59:53  AH2
' Added DescriptorOnly check and descriptor type support
'
' Revision 1.5  1999/09/05 23:59:46  AH2
' Fixed wildcarding
'
' Revision 1.4  1999/08/09 00:11:31  AH2
' Quick hack to get some validation going
'
' Revision 1.3  1999/08/01 22:52:34  AH2
' Added modeless Invoke to support launch of thesaurus editorial
'
' Revision 1.2  1999/07/11 21:46:44  AH2
' Added setting of icon
'
' Revision 1.1  1999/05/16 21:57:15  AH2
' Initial check-in
'

