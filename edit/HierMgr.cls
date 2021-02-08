VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HierMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' **********************************************************************
'
' Module: $Id: HierMgr.cls,v 1.9 2006/07/27 23:22:53 GS2 Exp $
'
' Author: Andrew Howroyd
'
' Created: Jul 1999
'
' Description
'
' **********************************************************************
Option Explicit

' names of hierarchies and lists
' initialised in Class_Initialise
Private m_allHierarchies As New HStrVect
Private m_numLists As Long

' KvList of open hierarchies
Private m_hier As New HKvList

' -------------------------------------------------------------------------
' Class_Initialize
'    Initialises class
' -------------------------------------------------------------------------
Private Sub Class_Initialize()

  With m_allHierarchies
    '.Add "Companies"
    .Add "Languages"
    .Add "Media-release-types"  'ONLY a list not a hierarchy!
    .Add "Major-clinical-trials"
    .Add "Routes-of-administration"
    .Add "Sports"
    m_numLists = .Count
    .Add "Body-systems"
    .Add "Drug-delivery-systems"
    .Add "$Non-disease-states"
    .Add "Geographical-locations"
    .Add "Special-interest-terms"
    .Add "Specialties"
    .Add "Therapeutic-procedures"
    .Add "Drug-classes-mechanism-of-action"
    .Add "Drug-classes-therapeutic-effect"
    .Add "Drug-classes-chemical"
    .Add "Drug-classes-biological"
'    .Add "Drug-classes"
'    .Add "Therapeutic-indications"
    .Add "$Diseases"
    .Add "$Endpoints"
' AH: Uncomment next line(s) to enable drugs hierarchy
'    .Add "Drugs"

' GS: added new BDL types 27/7/2006
    .Add "Diagnostics"
    .Add "Devices"
    .Add "Targets"
    .Add "Technology"
  End With

End Sub

' -------------------------------------------------------------------------
' AllHierarchies
'   returns list of all hierarchy names
' -------------------------------------------------------------------------
Public Property Get AllHierarchies() As HStrVect
  Set AllHierarchies = m_allHierarchies
End Property

' -------------------------------------------------------------------------
' Find
'   Finds a hierarchy by root term
' -------------------------------------------------------------------------
Public Property Get Hierarchy(ByVal RootTerm As String) As HierObj
  Set Hierarchy = m_hier.Object(RootTerm)
End Property

' -------------------------------------------------------------------------
' IsOpen
'   Tests if a hierarchy is open
' -------------------------------------------------------------------------
Public Property Get IsOpen(ByVal Hierarchy As String) As Boolean
  IsOpen = Not m_hier.Object(Hierarchy) Is Nothing
End Property

' -------------------------------------------------------------------------
' IsModified
'   Tests if a hierarchy has been modified
' -------------------------------------------------------------------------
Public Property Get IsModified(ByVal Hierarchy As String) As Boolean
  Dim HierObj As HierObj
  Set HierObj = m_hier.Object(Hierarchy)
  If Not HierObj Is Nothing Then
    IsModified = HierObj.IsModified
  End If
End Property

' -------------------------------------------------------------------------
' CloseHier
'   Closes hierarchy
' -------------------------------------------------------------------------
Public Sub CloseHier(ByVal Hierarchy As String)
  m_hier.Remove Hierarchy
End Sub

' -------------------------------------------------------------------------
' CloseAll
'   Closes all hierarchies
' -------------------------------------------------------------------------
Public Sub CloseAll()
  m_hier.Clear
End Sub

' -------------------------------------------------------------------------
' SaveChanges
'   Instructs all open hierarchies to save modifications
'   Return True on success
' -------------------------------------------------------------------------
Public Function SaveChanges(ByVal context As HAppContext) As Boolean
  SaveChanges = True
  Dim HierObj As HierObj
  Dim i As Long
  For i = 1 To m_allHierarchies.Count
    Set HierObj = m_hier.Object(m_allHierarchies(i))
    If Not HierObj Is Nothing Then
      If HierObj.IsModified Then
        If Not HierObj.SaveChanges(context) Then
          SaveChanges = False
          Exit For
        End If
      End If
    End If
  Next
End Function

' -------------------------------------------------------------------------
' OnCommit
'   Notifies hierarchies that transaction has been
'   successfully committed
' -------------------------------------------------------------------------
Public Sub OnCommit()
  Dim HierObj As HierObj
  Dim i As Long
  For i = 1 To m_allHierarchies.Count
    Set HierObj = m_hier.Object(m_allHierarchies(i))
    If Not HierObj Is Nothing Then
      If HierObj.IsModified Then HierObj.OnCommit
    End If
  Next
End Sub

' -------------------------------------------------------------------------
' OpenHier
'   Opens a hierarchy or list
'   Ignored if already open
'   Returns True on success, otherwise error reported
' -------------------------------------------------------------------------
Public Function OpenHier(ByVal Hierarchy As String, _
                 ByVal context As HAppContext) As Boolean
  ' just return success if already open
  If m_hier.Contains(Hierarchy) Then
    OpenHier = True
    Exit Function
  End If
  
  ' check that hieratchy name is valid
  Dim p As Long
  p = m_allHierarchies.Find(Hierarchy)
  If p = 0 Then
    MsgBox "Hierarchy name is not configured"
    Exit Function
  End If
  
  ' allocate new hierarchy object to manage this one
  Dim HierObj As New HierObj
  If HierObj.Load(context, m_allHierarchies(p), p <= m_numLists) Then
    m_hier.Add Hierarchy, HierObj
    OpenHier = True
  End If
  
End Function

' -------------------------------------------------------------------------
' $Log: HierMgr.cls,v $
' Revision 1.9  2006/07/27 23:22:53  GS2
' Added four new hierarchies for BDL.
'
' Revision 1.8  2002/01/10 22:29:29  gs2
' Add MRT to Hierarchy list.
'
' Revision 1.7  2001/12/10 04:44:08  jd4
' Thesaurus must not add or update Company Terms
'    removed company from available descriptor types
'    removed companies from available hierarchies
'
' New level hierarchy Media-release-type
'   added to available descriptor types
'   added to hierarchies
'
' Revision 1.6  2000/08/15 01:58:37  AH2
' Changes for new Drug-class hierarchies
'
' Revision 1.5  2000/06/20 03:37:35  AH2
' Added definitions for end-points hierarchy
'
' Revision 1.4  1999/11/26 01:27:47  AH2
' Replace 'Death' hierarchy with 'Non-disease-states'
' Fixed spelling of 'research-drug' term
'
' Revision 1.3  1999/09/12 21:58:53  AH2
' Source update: first releasable version
'
' Revision 1.2  1999/07/14 02:30:56  AH2
' Source update
'
' Revision 1.1  1999/07/11 22:11:37  AH2
' Initial check-in
'


