VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HierObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' **********************************************************************
'
' Module: $Id: HierObj.cls,v 1.4 1999/09/12 21:58:54 AH2 Exp $
'
' Author: Andrew Howroyd
'
' Created: Jul 1999
'
' Description
'     Object for managing changes to a single thesaurus hierarchy
'
' **********************************************************************
Option Explicit
Option Compare Text

Private m_root As String
Private m_isList As Boolean
Private m_childTerms As New HKvList
Private m_parentTerms As New HKvList
Private m_isModified As Boolean
Private m_inserts As HRowVect
Private m_deletes As HRowVect

' -------------------------------------------------------------------------
' IsValidTerm
'   tests if a string is a valid term name
' -------------------------------------------------------------------------
Public Function IsValidTerm(ByVal str As String) As Boolean
  If str <> "" And Len(str) <= MAX_TERM_LEN And InStr(str, " ") = 0 Then
    IsValidTerm = HUtil.IsAscii(str)
  End If
End Function

' -------------------------------------------------------------------------
' Root
'   Returns the root term of the hierarchy
' -------------------------------------------------------------------------
Public Property Get Root() As String
  Root = m_root
End Property

' -------------------------------------------------------------------------
' Children
'   Returns the immediate children of a node
'   The children are returned in no particular order
' -------------------------------------------------------------------------
Public Property Get Children(ByVal Parent As String) As HStrVect
  Set Children = m_childTerms.Object(Parent)
End Property

' -------------------------------------------------------------------------
' NumChildren
'   Returns number of children
' -------------------------------------------------------------------------
Public Property Get NumChildren(ByVal Term As String) As Long
  Dim vect As HStrVect
  Set vect = m_childTerms.Object(Term)
  If Not vect Is Nothing Then NumChildren = vect.Count
End Property

' -------------------------------------------------------------------------
' Parents
'   Returns the immediate parents of a node
'   The parents are returned in no particular order
' -------------------------------------------------------------------------
Public Property Get Parents(ByVal Term As String) As HStrVect
  Set Parents = m_parentTerms.Object(Term)
End Property

' -------------------------------------------------------------------------
' IsTerm
'   Tests if a term is a term
' -------------------------------------------------------------------------
Public Property Get IsTerm(ByVal Term As String) As Boolean
  IsTerm = Term = m_root Or Not m_parentTerms.Object(Term) Is Nothing
End Property

' -------------------------------------------------------------------------
' NumParents
' -------------------------------------------------------------------------
Public Property Get NumParents(ByVal Term As String) As Long
  Dim vect As HStrVect
  Set vect = m_parentTerms.Object(Term)
  If Not vect Is Nothing Then NumParents = vect.Count
End Property

' -------------------------------------------------------------------------
' Descendants
'   Returns all the descendants of a node
'   Duplicate names are pruned
'   An optional stop term can be specified: in this case the search
'   is not allowed to pass the stop term
' -------------------------------------------------------------------------
Public Property Get Descendants(ByVal Parent As String, _
                 Optional ByVal stopTerm As String) As HStrVect
  Dim all As New HStrVect
  Dim bucket As New HKvList
  Dim p As Long
  
  Do
    Dim items As HStrVect
    Set items = m_childTerms.Object(Parent)
    If Not items Is Nothing Then
      
      Dim i As Long
      For i = 1 To items.Count
        Dim Term As String
        Term = items(i)
        If Term <> stopTerm Then
          If Not bucket.Contains(Term) Then
            bucket.Add Term, True
            all.Add Term
          End If
        End If
      Next
    End If
    ' exit loop if checked all terms
    If p = all.Count Then Exit Do
    p = p + 1
    Parent = all(p)
  Loop
  
  ' return array
  Set Descendants = all
   
End Property

' -------------------------------------------------------------------------
' Ancestors
'   Returns all the ancestors of a node
'   Duplicate names are pruned
'   An optional stop term can be specified: in this case the search
'   is not allowed to pass the stop term
' -------------------------------------------------------------------------
Public Property Get Ancestors(ByVal Child As String, _
                 Optional ByVal stopTerm As String) As HStrVect
  Dim all As New HStrVect
  Dim bucket As New HKvList
  Dim p As Long
  
  Do
    Dim items As HStrVect
    Set items = m_parentTerms.Object(Child)
    If Not items Is Nothing Then
      
      Dim i As Long
      For i = 1 To items.Count
        Dim Term As String
        Term = items(i)
        If Term <> stopTerm Then
          If Not bucket.Contains(Term) Then
            bucket.Add Term, True
            all.Add Term
          End If
        End If
      Next
    End If
    ' exit loop if checked all terms
    If p = all.Count Then Exit Do
    p = p + 1
    Child = all(p)
  Loop
  
  ' return array
  Set Ancestors = all
   
End Property

' -------------------------------------------------------------------------
' IsChild
'   Tests if a term is a child of another
' -------------------------------------------------------------------------
Public Function IsChild(ByVal Term As String, ByVal Parent As String)
  Dim Parents As HStrVect
  Set Parents = m_parentTerms.Object(Term)
  If Not Parents Is Nothing Then
    IsChild = Parents.Find(Parent) > 0
  End If
End Function

' -------------------------------------------------------------------------
' MatchingTerms
'   Returns all terms that match a specified pattern
' -------------------------------------------------------------------------
Public Function MatchingTerms(ByVal pattern As String) As HStrVect
  Dim allTerms As HStrVect
  Set allTerms = Descendants(m_root)
  If Not m_isList Then allTerms.Add m_root
  Dim i As Long
  For i = 1 To allTerms.Count
    If Not UCase(allTerms(i)) Like UCase(pattern) Then
      allTerms(i) = ""
    End If
  Next
  allTerms.Remove ""
  Set MatchingTerms = allTerms
End Function

' -------------------------------------------------------------------------
' IsModified
'    Tests if modified flag set
' -------------------------------------------------------------------------
Public Property Get IsModified() As Boolean
  IsModified = m_isModified
End Property

' -------------------------------------------------------------------------
' DoMove
'   Called from RemoveTerm, AddTerm, MoveTerm to update
'   internal structures and record change
'   All required validation has been performed
' -------------------------------------------------------------------------
Public Sub DoMove(ByVal Term As String, ByVal y As String, ByVal z As String)
  Dim Children As HStrVect, Parents As HStrVect
  Dim r As HRow
  
  Set Parents = m_parentTerms.Object(Term)
  
  If z <> "" Then
    Set Children = m_childTerms.Object(z)
    If Children Is Nothing Then
      Set Children = New HStrVect
      m_childTerms.Add z, Children
    End If
    Children.Add Term
    If Parents Is Nothing Then
      Set Parents = New HStrVect
      m_parentTerms.Add Term, Parents
    End If
    Parents.Add z
    Set r = m_inserts.RowDef.CreateRow
    m_inserts.Add r
    r("LeadTerm") = Term
    If Not m_isList Then r("ParentTerm") = z
  End If
  
  If y <> "" Then
    Set Children = m_childTerms.Object(y)
    Children.Remove Term
    If Children.Count = 0 Then
      m_childTerms.Remove y
    End If
    Parents.Remove y
    If Parents.Count = 0 Then
      m_parentTerms.Remove Term
    End If
    Set r = m_deletes.RowDef.CreateRow
    m_deletes.Add r
    r("LeadTerm") = Term
    If Not m_isList Then r("ParentTerm") = y
  End If
  
  m_isModified = True

End Sub

' -------------------------------------------------------------------------
' ValidateMove
'     Validates a move operation
'     Returns empty string if ok to proceed, otherwise error message
'     This function must be called before DoMove
'     Supported operations:
'        Unlink Term:  z = "", z is Parent
'        Add Term:    y = "", z is new Parent
'        Move Term:  y is old Parent, z is new Parent
'     Rules: see comments in code below
' -------------------------------------------------------------------------
Public Function ValidateMove(ByVal Term As String, _
                ByVal y As String, ByVal z As String) As String
  Dim msg As String
  Dim termVect As HStrVect
  Dim i As Long
  
  ' check for really wrong stuff
  If y = z Or Term = "" Then msg = "Invalid operation"
    
  If msg = "" Then
    If y = "" Then
      ' adding a term, can be new term or existing
      If Not IsTerm(Term) And Not IsValidTerm(Term) Then
        msg = "Invalid name for new thesaurus term"
      End If
    Else
      ' Term must be child of y
      If Not IsChild(Term, y) Then
        msg = "Term is not child of indicated parent"
      End If
    End If
  End If
  
  If msg = "" Then
    If z = "" Then
      ' remove: either must be other parents or must be leaf
      If NumParents(Term) = 1 And NumChildren(Term) > 0 Then
        msg = "Child terms must first be removed"
      End If
    Else
      ' z is existing term
      ' Term is not z
      ' if list then z must be root term
      If Not IsTerm(z) Then
        msg = "Requested parent not in hierarchy"
      ElseIf z = Term Then
        msg = "Cannot add term to self"
      ElseIf m_isList And z <> m_root Then
        msg = "Can only add to root term of simple list"
      End If
      
      ' Term may not be descendant of z except through y
      If msg = "" Then
        Set termVect = Ancestors(Term, y)
        For i = 1 To termVect.Count
          If termVect(i) = z Then
            msg = "Term is already a descendant"
          End If
        Next
      End If
      
      ' z not descendant of Term
      ' no child of z may be descendant of Term
      If msg = "" Then
        Set termVect = Descendants(Term)
        For i = 1 To termVect.Count
          If termVect(i) = z Then
            msg = "Cyclic trees are not permitted"
          ElseIf IsChild(termVect(i), z) Then
            msg = "Relation with '" & termVect(i) & "' would not be minimal"
          End If
        Next
      End If
      
      ' no parent of term (except y) may be anscestor of z
      If msg = "" Then
        Set termVect = Ancestors(z, y)
        For i = 1 To termVect.Count
          If IsChild(Term, termVect(i)) Then
            msg = "Relation with '" & termVect(i) & "' would not be minimal"
          End If
        Next
      End If
    End If
  End If
  
  ' return
  ValidateMove = msg
  
End Function

' -------------------------------------------------------------------------
' Load
'   Loads hierarchy from database
'   Returns True on succcess, otherwise error reported
' -------------------------------------------------------------------------
Public Function Load(ByVal context As HAppContext, _
                ByVal Root As String, IsList As Boolean) As Boolean
  ' initialise
  m_isList = IsList
  m_root = ""
  m_childTerms.Clear
  m_parentTerms.Clear
  m_isModified = False
  
  Dim KeyDef As New HRowDef
  Dim RowDef As New HRowDef
  Dim rows As HRowVect, key As HRow
  Dim strVect As HStrVect
  Dim i As Long
  
  If IsList Then
    RowDef.Add "LeadTerm", vbString, MAX_TERM_LEN
    KeyDef.Add "ParentTerm", vbString, MAX_TERM_LEN
    Set rows = RowDef.CreateRowVect
    Set key = KeyDef.CreateRow
    key(1) = Root
    Load = context.DB.DBSelect("THSGroupTerm", rows, key, _
                               "RelCode='LIST'") >= 0
    If Load Then
      ' all data retrieved successfully
      m_root = Root
      Set strVect = New HStrVect
      rows.ExtractStrVect strVect
      strVect.Sort
      m_childTerms(Root) = strVect
      Dim rootVect As New HStrVect
      rootVect.Add Root
      For i = 1 To strVect.Count
        m_parentTerms(strVect(i)) = rootVect
      Next
    End If
  Else
    RowDef.Add "LeadTerm", vbString, MAX_TERM_LEN
    RowDef.Add "ParentTerm", vbString, MAX_TERM_LEN
    KeyDef.Add "Hierarchy", vbString, MAX_TERM_LEN
    Set rows = RowDef.CreateRowVect
    Set key = KeyDef.CreateRow
    key(1) = Root
    Load = context.DB.DBSelect("THSHierarchy", rows, key) >= 0
    If Load Then
      ' all data retrieved successfully
      ' loop over returned result set and build data structures
      Dim childTerm As String, parentTerm As String
      m_root = Root
      For i = 1 To rows.Count
        childTerm = rows(i)(1)
        parentTerm = rows(i)(2)
        Set strVect = m_parentTerms.Object(childTerm)
        If strVect Is Nothing Then
          Set strVect = New HStrVect
          m_parentTerms.Add childTerm, strVect
        End If
        strVect.Add parentTerm
        Set strVect = m_childTerms.Object(parentTerm)
        If strVect Is Nothing Then
          Set strVect = New HStrVect
          m_childTerms.Add parentTerm, strVect
        End If
        strVect.Add childTerm
      Next
    End If
  End If
  
  Set m_inserts = RowDef.CreateRowVect
  Set m_deletes = RowDef.CreateRowVect
  
End Function

' -------------------------------------------------------------------------
' SaveChanges
'   Saves all changes
'   Must be called inside transaction
' -------------------------------------------------------------------------
Public Function SaveChanges(ByVal context As HAppContext) As Boolean
  
  Dim KeyDef As New HRowDef, key As HRow, table As String
  
  ' set up key
  If m_isList Then
    KeyDef.Add "ParentTerm", vbString, MAX_TERM_LEN
    KeyDef.Add "RelCode", vbString, 4
    Set key = KeyDef.CreateRow
    key(1) = m_root
    key(2) = "LIST"
    table = "THSGroupTerm"
  Else
    KeyDef.Add "Hierarchy", vbString, MAX_TERM_LEN
    KeyDef.Add "RelCode", vbString, 4
    Set key = KeyDef.CreateRow
    key(1) = m_root
    key(2) = "HIER"
    table = "THSHierarchy"
  End If
  
  ' and perform update
  SaveChanges = context.DB.DBUpdateVect(table, key, m_inserts, m_deletes) >= 0
End Function

' -------------------------------------------------------------------------
' OnCommit
'   Called when transaction committed
'   Object cleans up change state
' -------------------------------------------------------------------------
Public Sub OnCommit()
  m_isModified = False
  m_deletes.Clear
  m_inserts.Clear
End Sub

' -------------------------------------------------------------------------
' MoveNext
' -------------------------------------------------------------------------
Public Function MoveNext(ByVal Term As String, ByVal Search As String) As String
  Dim bestMatch As String, FullPath As String
  Dim stack As New HObjVect
  Dim batch As HStrVect

  Set batch = Parents(Term)
  If Not batch Is Nothing Then Set batch = batch.Clone
  FullPath = Term
  Do Until batch Is Nothing And stack.Count = 0
    If batch Is Nothing Then
      If bestMatch = "" Then
        bestMatch = FullPath
      ElseIf Search < FullPath And FullPath < bestMatch Then
        bestMatch = FullPath
      ElseIf FullPath < bestMatch And bestMatch <= Search Then
        bestMatch = FullPath
      ElseIf bestMatch <= Search And Search < FullPath Then
        bestMatch = FullPath
      End If
      Set batch = New HStrVect
    ElseIf batch.Count > 0 Then
      Term = batch(batch.Count)
      batch.Count = batch.Count - 1
      stack.Add batch
      FullPath = Term & vbTab & FullPath
      Set batch = Parents(Term)
      If Not batch Is Nothing Then Set batch = batch.Clone
    ElseIf stack.Count > 0 Then
      Set batch = stack(stack.Count)
      stack.Count = stack.Count - 1
      FullPath = Mid(FullPath, InStr(FullPath, vbTab) + 1)
    Else
      Set batch = Nothing
    End If
  Loop
  
  MoveNext = bestMatch
End Function

' -------------------------------------------------------------------------
' $Log: HierObj.cls,v $
' Revision 1.4  1999/09/12 21:58:54  AH2
' Source update: first releasable version
'
' Revision 1.3  1999/07/23 05:35:28  AH2
' Added find/find next capability
'
' Revision 1.2  1999/07/14 02:30:56  AH2
' Source update
'
' Revision 1.1  1999/07/11 22:11:37  AH2
' Initial check-in
'

