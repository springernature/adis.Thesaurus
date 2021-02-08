VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ThesObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' **********************************************************************
'
' Module: $Id: ThesObj.cls,v 1.11 2006/09/18 04:11:45 GS2 Exp $
'
' Author: Andrew Howroyd
'
' Created: Jul 1999
'
' Description
'
' **********************************************************************
Option Explicit

' info fixed in Class_Initialise
Private m_allDesTypes As New HKvList
Private m_allDesTypeNames As New HStrVect
Private m_THSTermDef As New HRowDef
Private m_THSLeadTermDef As New HRowDef
Private m_THSNoteDef As New HRowDef

' hierarchy manager
Private m_hierManager As New HierMgr

' -------------------------------------------------------------------------
' NewDesType
'    Adds a new descriptor type to m_allDesTypes array
'    Helper function called from InitDesTypeInfo
' -------------------------------------------------------------------------
Private Function NewDesType(ByVal Caption As String) As HKvList
  Dim kv As New HKvList
  m_allDesTypes(Caption) = kv
  m_allDesTypeNames.Add Caption
  kv("Caption") = Caption
  Set NewDesType = kv
End Function

' -------------------------------------------------------------------------
' InitDesTypeInfo
'    Initialises descriptor type info (m_allDesTypes)
'    Also initialised m_allHierarchies, m_allList
'    Called from Class_Initialize
' -------------------------------------------------------------------------
Private Sub InitDescriptorTypeInfo()
  
  Dim kv As HKvList, strVect As HStrVect
  
  ' companies
  'Set kv = NewDesType("(companies)")
  'Set strVect = New HStrVect
  'kv!hierarchies = strVect
  'kv!required = "Companies"
  'strVect.Add kv!required
  
  ' languages
  Set kv = NewDesType("(languages)")
  Set strVect = New HStrVect
  kv!hierarchies = strVect
  kv!required = "Languages"
  strVect.Add kv!required
  
  ' geographical-locations
  Set kv = NewDesType("(geographical-locations)")
  Set strVect = New HStrVect
  kv!hierarchies = strVect
  kv!required = "Geographical-locations"
  strVect.Add kv!required
   
  ' non-descriptor
  Set kv = NewDesType("(non-descriptor)")
  Set strVect = New HStrVect
  kv!hierarchies = strVect
  strVect.Add "Specialties"
  strVect.Add "$Endpoints"
  
  ' media-release-types
  Set kv = NewDesType("media-release-types")
  kv!DesTypeCode = "MRT"
  Set strVect = New HStrVect
  kv!hierarchies = strVect
  kv!required = "Media-release-types"
  strVect.Add kv!required
  
  ' major-clinical-trials
  Set kv = NewDesType("major-clinical-trials")
  kv!DesTypeCode = "MCT"
  Set strVect = New HStrVect
  kv!hierarchies = strVect
  kv!required = "Major-clinical-trials"
  strVect.Add kv!required
   
  ' non disease states
  Set kv = NewDesType("non-disease-states")
  kv!DesTypeCode = "NDIS"
  Set strVect = New HStrVect
  kv!hierarchies = strVect
  kv!required = "$Non-disease-states"
  strVect.Add kv!required
  strVect.Add "$Endpoints"
  
  ' diseases
  Set kv = NewDesType("diseases")
  kv!DesTypeCode = "DIS"
  Set strVect = New HStrVect
  kv!hierarchies = strVect
  kv!required = "$Diseases"
  strVect.Add kv!required
  strVect.Add "Special-interest-terms"
  strVect.Add "Body-systems"
  strVect.Add "$Endpoints"
      
  ' drugs
  Set kv = NewDesType("drugs")
  kv!DesTypeCode = "DRUG"
  Set strVect = New HStrVect
  kv!indicators = strVect
  strVect.Add "Research-drug"
  strVect.Add "Fixed-combinations"
  strVect.Add "Company-specific-drugs"
  Set strVect = New HStrVect
  kv!hierarchies = strVect
  strVect.Add "Diagnostics"
  kv!GroupRelCode = "DRGB"
  kv!GroupCaption = "Drug classes"
  kv!GroupSqlWhere = "(DesType like 'DC%' Or DesType = 'TGT')"
  kv!IndRelCode = "DRGA"

  ' drug-classes (new)
  Set kv = NewDesType("drug-classes-mechanism-of-action")
  kv!DesTypeCode = "DC_M"
  Set strVect = New HStrVect
  kv!hierarchies = strVect
  kv!required = "Drug-classes-mechanism-of-action"
  strVect.Add kv!required
  strVect.Add "Body-systems"
  kv!MemberRelCode = "DRGB"
  kv!MemberCaption = "Members"
  kv!MemberSqlWhere = "DesType = 'DRUG'"

  Set kv = NewDesType("drug-classes-therapeutic-effect")
  kv!DesTypeCode = "DC_T"
  Set strVect = New HStrVect
  kv!hierarchies = strVect
  kv!required = "Drug-classes-therapeutic-effect"
  strVect.Add kv!required
  strVect.Add "Body-systems"
  kv!MemberRelCode = "DRGB"
  kv!MemberCaption = "Members"
  kv!MemberSqlWhere = "DesType = 'DRUG'"
  
  Set kv = NewDesType("drug-classes-biological")
  kv!DesTypeCode = "DC_B"
  Set strVect = New HStrVect
  kv!hierarchies = strVect
  kv!required = "Drug-classes-biological"
  strVect.Add kv!required
  strVect.Add "Body-systems"
  kv!MemberRelCode = "DRGB"
  kv!MemberCaption = "Members"
  kv!MemberSqlWhere = "DesType = 'DRUG'"

  Set kv = NewDesType("drug-classes-chemical")
  kv!DesTypeCode = "DC_C"
  Set strVect = New HStrVect
  kv!hierarchies = strVect
  kv!required = "Drug-classes-chemical"
  strVect.Add kv!required
  strVect.Add "Body-systems"
  kv!MemberRelCode = "DRGB"
  kv!MemberCaption = "Members"
  kv!MemberSqlWhere = "DesType = 'DRUG'"

  ' general-descriptor
  Set kv = NewDesType("general-descriptor")
  kv!DesTypeCode = "GEN"
  Set strVect = New HStrVect
  kv!hierarchies = strVect
  strVect.Add "Special-interest-terms"
  strVect.Add "Body-systems"
  strVect.Add "Routes-of-administration"
  strVect.Add "Drug-delivery-systems"
  strVect.Add "Therapeutic-procedures"
  strVect.Add "Sports"
  strVect.Add "Specialties"
  strVect.Add "$Endpoints"
     
  ' targets
  Set kv = NewDesType("targets")
  kv!DesTypeCode = "TGT"
  Set strVect = New HStrVect
  kv!hierarchies = strVect
  strVect.Add "Targets"
     
  ' devices
  Set kv = NewDesType("devices")
  kv!DesTypeCode = "DEV"
  Set strVect = New HStrVect
  kv!hierarchies = strVect
  strVect.Add "Devices"
     
  ' diagnostics
  Set kv = NewDesType("diagnostics")
  kv!DesTypeCode = "DIAG"
  Set strVect = New HStrVect
  kv!hierarchies = strVect
  strVect.Add "Diagnostics"
     
  ' technology
  Set kv = NewDesType("technology")
  kv!DesTypeCode = "TECH"
  Set strVect = New HStrVect
  kv!hierarchies = strVect
  strVect.Add "Technology"
     
End Sub

' -------------------------------------------------------------------------
' Class_Initialize
'    Initialises class
' -------------------------------------------------------------------------
Private Sub Class_Initialize()
  ' set up m_allDesTypes, m_allDesTypeNames
  InitDescriptorTypeInfo
  
  ' create row definitions
  m_THSLeadTermDef.Add g_sLeadTerm, vbString, 200
  
  m_THSTermDef.Add g_sLeadTerm, vbString, 200
  m_THSTermDef.Add g_sDisplayForm, vbString, 200
  m_THSTermDef.Add g_sIsApproved, vbString, 1
  m_THSTermDef.Add g_sDesType, vbString, 4
  m_THSTermDef.Add g_sUseTerm, vbString, 200

  m_THSNoteDef.Add g_sNoteId, vbLong, 0
  m_THSNoteDef.Add g_sNoteDate, vbDate, 0
  m_THSNoteDef.Add g_sNoteText, vbString, 1000
  
End Sub

' -------------------------------------------------------------------------
' AllDesTypes
'   Returns list of all descriptor types
' -------------------------------------------------------------------------
Public Property Get AllDesTypes() As HStrVect
  Set AllDesTypes = m_allDesTypeNames
End Property

' -------------------------------------------------------------------------
' DesTypeInfo
'   Returns descriptor type info for a descriptor type
' -------------------------------------------------------------------------
Public Property Get DesTypeInfo(ByVal DesType As String) As HKvList
  Set DesTypeInfo = m_allDesTypes.Object(DesType)
End Property

' -------------------------------------------------------------------------
' IsValidTerm
'   Tests if a term has a valid name
' -------------------------------------------------------------------------
Public Function IsValidTerm(ByVal str As String) As Boolean
  If str <> "" And Len(str) <= MAX_TERM_LEN And InStr(str, " ") = 0 Then
    IsValidTerm = HUtil.IsAscii(str)
  End If
End Function

' -------------------------------------------------------------------------
' HierManager
'   Returns associated hierarchy manager
' -------------------------------------------------------------------------
Public Property Get HierManager() As HierMgr
  Set HierManager = m_hierManager
End Property

' -------------------------------------------------------------------------
' NewObj
'   Initialises a new object
' -------------------------------------------------------------------------
Public Function NewObj(ByVal cxt As HAppContext, _
                ByVal obj As HKvList) As Boolean
  obj(g_sLeadTerm) = ""
  obj(g_sDisplayForm) = ""
  obj(g_sDescriptorType) = ""
  obj(g_sHierarchy) = New HStrVect
  obj(g_sUsedFor) = New HStrVect
  obj(g_sRelatedTerm) = New HStrVect
  obj(g_sGroupTerm) = New HStrVect
  obj(g_sMemberTerm) = New HStrVect
  obj(g_sIndicator) = New HStrVect
  obj(g_sScopeNote) = m_THSNoteDef.CreateRowVect
  obj(g_sDateNote) = m_THSNoteDef.CreateRowVect
  NewObj = True
End Function

' -------------------------------------------------------------------------
' LoadObj
'   Loads thesaurus object from database
'   Returns True on success. If not found, returned object empty
' -------------------------------------------------------------------------
Public Function LoadObj(ByVal cxt As HAppContext, _
                ByVal key As HKvList, _
                ByVal obj As HKvList) As Boolean

  Dim hdb As HDBBase
  Dim vect As HRowVect, r As HRow, strVect As HStrVect
  Dim sql As String
  Dim i As Long
  Dim ok  As Boolean
  Set hdb = cxt.DB
  
  Dim dbkey As HRow
  Set dbkey = m_THSLeadTermDef.CreateRow
  dbkey(1) = UCase(key(g_sLeadTerm))

  ' get THSTerm row
  Set vect = m_THSTermDef.CreateRowVect
  ok = hdb.DBSelect("THSTerm", vect, dbkey, "UseTerm is null") >= 0

  ' copy THSTerm row to obj - return if none
  If ok Then
    If vect.Count = 0 Then
      LoadObj = True
      Exit Function
    End If
    Set r = vect(1)
    obj(g_sLeadTerm) = r(g_sLeadTerm)
    obj(g_sDisplayForm) = r(g_sDisplayForm)
  End If

  ' get use terms
  If ok Then
    Set vect = m_THSLeadTermDef.CreateRowVect
    sql = "select LeadTerm from THSTerm where UseTerm=:LeadTerm" & _
          " order by LeadTerm"
    ok = hdb.SqlQuery(sql, vect, dbkey) >= 0
    Set strVect = New HStrVect
    vect.ExtractStrVect strVect
    obj(g_sUsedFor) = strVect
  End If
  
  ' get related terms
  If ok Then
    Set vect = m_THSLeadTermDef.CreateRowVect
    sql = "select RelatedTerm from THSRelatedTerm" & _
          " where LeadTerm=:LeadTerm order by RelatedTerm"
    ok = hdb.SqlQuery(sql, vect, dbkey) >= 0
    Set strVect = New HStrVect
    vect.ExtractStrVect strVect
    obj(g_sRelatedTerm) = strVect
  End If
  
  ' get hierarchies and lists
  Set vect = m_THSLeadTermDef.CreateRowVect
  sql = "select distinct Hierarchy from THSHierarchy" & _
        " where LeadTerm=:LeadTerm"
  If ok Then ok = hdb.SqlQuery(sql, vect, dbkey) >= 0
  sql = "select ParentTerm from THSGroupTerm" & _
        " where RelCode='LIST' and LeadTerm=:LeadTerm"
  If ok Then ok = hdb.SqlQuery(sql, vect, dbkey) >= 0
  Set strVect = New HStrVect
  vect.ExtractStrVect strVect
  obj(g_sHierarchy) = strVect
  
  ' determine descriptor type
  Dim DesTypeInfo As HKvList
  If ok Then
    For i = 1 To m_allDesTypeNames.Count
      Set DesTypeInfo = m_allDesTypes.Object(m_allDesTypeNames(i))
      If DesTypeInfo("DesTypeCode") = r(g_sDesType) Then
        If DesTypeInfo("Required") <> "" Then
          If strVect.Find(DesTypeInfo("Required")) > 0 Then
            obj(g_sDescriptorType) = DesTypeInfo("Caption")
            Exit For
          End If
        Else
            obj(g_sDescriptorType) = DesTypeInfo("Caption")
            Exit For
        End If
      End If
    Next
  End If
  
  ' THSGroup table
  obj(g_sGroupTerm) = New HStrVect
  obj(g_sMemberTerm) = New HStrVect
  obj(g_sIndicator) = New HStrVect
  If Not DesTypeInfo Is Nothing Then
    If ok And DesTypeInfo("GroupRelCode") <> "" Then
      Set vect = m_THSLeadTermDef.CreateRowVect
      sql = "select ParentTerm from THSGroupTerm where LeadTerm=:LeadTerm" & _
            " and RelCode='" & DesTypeInfo("GroupRelCode") & "'" & _
            " order by ParentTerm"
      ok = hdb.SqlQuery(sql, vect, dbkey) >= 0
      Set strVect = New HStrVect
      vect.ExtractStrVect strVect
      obj(g_sGroupTerm) = strVect
    End If
    If ok And DesTypeInfo("MemberRelCode") <> "" Then
      Set vect = m_THSLeadTermDef.CreateRowVect
      sql = "select LeadTerm from THSGroupTerm where ParentTerm=:LeadTerm" & _
            " and RelCode='" & DesTypeInfo("MemberRelCode") & "'" & _
            " order by LeadTerm"
      ok = hdb.SqlQuery(sql, vect, dbkey) >= 0
      Set strVect = New HStrVect
      vect.ExtractStrVect strVect
      obj(g_sMemberTerm) = strVect
    End If
    If ok And DesTypeInfo("IndRelCode") <> "" Then
      Set vect = m_THSLeadTermDef.CreateRowVect
      sql = "select ParentTerm from THSGroupTerm where LeadTerm=:LeadTerm" & _
            " and RelCode='" & DesTypeInfo("IndRelCode") & "'" & _
            " order by ParentTerm"
      ok = hdb.SqlQuery(sql, vect, dbkey) >= 0
      Set strVect = New HStrVect
      vect.ExtractStrVect strVect
      obj(g_sIndicator) = strVect
    End If
  End If
  
  ' Notes
  If ok Then
    Set vect = m_THSNoteDef.CreateRowVect
    ok = hdb.DBSelect("THSNote", vect, dbkey, _
                      "NoteType='SCOP'", "order by NoteId") >= 0
    
    obj(g_sScopeNote) = vect
  End If
  
  ' get date notes
  If ok Then
    Set vect = m_THSNoteDef.CreateRowVect
    ok = hdb.DBSelect("THSNote", vect, dbkey, _
                      "NoteType='DATE'", "order by NoteDate, NoteId") >= 0
    obj(g_sDateNote) = vect
  End If
  
  ' return
  LoadObj = ok
End Function

' -------------------------------------------------------------------------
' SaveObj
'   Saves an object to the database
'   Must be called inside transaction
'   Returns False if failed
' -------------------------------------------------------------------------
Public Function SaveObj(ByVal cxt As HAppContext, _
                ByVal isNew As Boolean, _
                ByVal keyObj As HKvList, _
                ByVal NewObj As HKvList, _
                ByVal origObj As HKvList) As Boolean
  Dim newVect As HRowVect, oldVect As HRowVect
  Dim RowDef As HRowDef, r As HRow, keyRow As HRow
  Dim isDesTypeModified As Boolean
  Dim leadTerm As String
  Dim sql As String

  Dim hdb As HDBBase
  Set hdb = cxt.DB
  
  ' look up descriptor type
  Dim DesTypeInfo As HKvList
  Set DesTypeInfo = m_allDesTypes.Object(NewObj(g_sDescriptorType))
  
  ' THSTerm
  Set keyRow = m_THSLeadTermDef.CreateRow
  Set r = m_THSTermDef.CreateRow
  r(g_sDisplayForm) = NewObj(g_sDisplayForm)
  r(g_sIsApproved) = "Y"
  r(g_sDesType) = DesTypeInfo!DesTypeCode
  
  If isNew Then
    keyRow(g_sLeadTerm) = NewObj(g_sLeadTerm)
    r(g_sLeadTerm) = keyRow(g_sLeadTerm)
    If hdb.DBInsert("THSTerm", True, r) = 0 Then
      hdb.FailTransaction
      MsgBox "Term already exists in thesaurus!"
    End If
  Else
    keyRow(g_sLeadTerm) = origObj(g_sLeadTerm)
    isDesTypeModified = origObj(g_sDescriptorType) <> NewObj(g_sDescriptorType)
    If isDesTypeModified Or _
       origObj(g_sDisplayForm) <> NewObj(g_sDisplayForm) Then
      r(g_sLeadTerm) = keyRow(g_sLeadTerm)
      hdb.DBUpdate "THSTerm", keyRow, r
    End If
  End If
  leadTerm = keyRow(g_sLeadTerm)
  
  ' if descriptor type modified need to do some cleaning
  If isDesTypeModified Then
    sql = "delete from THSGroupTerm" & _
          " where LeadTerm=:LeadTerm and RelCode<>'LIST'"
    hdb.SqlExec sql, False, keyRow
    sql = "delete from THSGroupTerm" & _
          " where ParentTerm=:LeadTerm and RelCode<>'LIST'"
    hdb.SqlExec sql, False, keyRow
    sql = "update THSTerm set DesType=" & _
          " (select DesType from THSTerm where LeadTerm=:LeadTerm)" & _
          " where UseTerm=:LeadTerm"
    hdb.SqlExec sql, False, keyRow
  End If
  
  ' THSRelated: related terms
  ' rows are inserted / deleted in pairs
  Set RowDef = New HRowDef
  RowDef.Add g_sRelatedTerm, vbString, MAX_TERM_LEN
  Set newVect = RowDef.CreateRowVect
  newVect.FillFromStrVect NewObj.Object(g_sRelatedTerm)
  If isNew Then
    hdb.DBInsert "THSRelatedTerm", False, keyRow, newVect
  Else
    Set oldVect = RowDef.CreateRowVect
    oldVect.FillFromStrVect origObj.Object(g_sRelatedTerm)
    hdb.DBUpdateVect "THSRelatedTerm", keyRow, newVect, oldVect
  End If
  Set keyRow = RowDef.CreateRow
  keyRow(g_sRelatedTerm) = leadTerm
  Set newVect = m_THSLeadTermDef.CreateRowVect
  newVect.FillFromStrVect NewObj.Object(g_sRelatedTerm)
  If isNew Then
    hdb.DBInsert "THSRelatedTerm", False, keyRow, newVect
  Else
    Set oldVect = m_THSLeadTermDef.CreateRowVect
    oldVect.FillFromStrVect origObj.Object(g_sRelatedTerm)
    hdb.DBUpdateVect "THSRelatedTerm", keyRow, newVect, oldVect
  End If
  
  ' used for terms
  Set RowDef = m_THSTermDef.SelectCols("UseTerm,DesType,IsApproved")
  Set keyRow = RowDef.CreateRow
  keyRow(g_sUseTerm) = leadTerm
  keyRow(g_sDesType) = r(g_sDesType)
  keyRow(g_sIsApproved) = "N"
  Set newVect = m_THSLeadTermDef.CreateRowVect
  newVect.FillFromStrVect NewObj.Object(g_sUsedFor)
  If isNew Then
    If hdb.DBInsert("THSTerm", True, keyRow, newVect) <> newVect.Count Then
      hdb.FailTransaction
      MsgBox "Use term is approved thesaurus term!"
    End If
  Else
    Set oldVect = m_THSLeadTermDef.CreateRowVect
    oldVect.FillFromStrVect origObj.Object(g_sUsedFor)
    hdb.DBUpdateVect "THSTerm", keyRow, newVect, oldVect
  End If

  ' THSGroup table
  Set RowDef = New HRowDef
  RowDef.Add g_sLeadTerm, vbString, MAX_TERM_LEN
  RowDef.Add g_sRelCode, vbString, 4
  Set keyRow = RowDef.CreateRow
  keyRow(g_sLeadTerm) = leadTerm
  Set RowDef = New HRowDef
  RowDef.Add g_sParentTerm, vbString, MAX_TERM_LEN
  keyRow(g_sRelCode) = DesTypeInfo("GroupRelCode")
  If keyRow(g_sRelCode) <> "" Then
    Set newVect = RowDef.CreateRowVect
    newVect.FillFromStrVect NewObj.Object(g_sGroupTerm)
    If isNew Or isDesTypeModified Then
      hdb.DBInsert "THSGroupTerm", False, keyRow, newVect
    Else
      Set oldVect = RowDef.CreateRowVect
      oldVect.FillFromStrVect origObj.Object(g_sGroupTerm)
      hdb.DBUpdateVect "THSGroupTerm", keyRow, newVect, oldVect
    End If
  End If
  keyRow(g_sRelCode) = DesTypeInfo("IndRelCode")
  If keyRow(g_sRelCode) <> "" Then
    Set newVect = RowDef.CreateRowVect
    newVect.FillFromStrVect NewObj.Object(g_sIndicator)
    If isNew Or isDesTypeModified Then
      hdb.DBInsert "THSGroupTerm", False, keyRow, newVect
    Else
      Set oldVect = RowDef.CreateRowVect
      oldVect.FillFromStrVect origObj.Object(g_sIndicator)
      hdb.DBUpdateVect "THSGroupTerm", keyRow, newVect, oldVect
    End If
  End If
  Set RowDef = New HRowDef
  RowDef.Add g_sParentTerm, vbString, MAX_TERM_LEN
  RowDef.Add g_sRelCode, vbString, 4
  Set keyRow = RowDef.CreateRow
  keyRow(g_sParentTerm) = leadTerm
  Set RowDef = New HRowDef
  RowDef.Add g_sLeadTerm, vbString, MAX_TERM_LEN
  keyRow(g_sRelCode) = DesTypeInfo("MemberRelCode")
  If keyRow(g_sRelCode) <> "" Then
    Set newVect = RowDef.CreateRowVect
    newVect.FillFromStrVect NewObj.Object(g_sMemberTerm)
    If isNew Or isDesTypeModified Then
      hdb.DBInsert "THSGroupTerm", False, keyRow, newVect
    Else
      Set oldVect = RowDef.CreateRowVect
      oldVect.FillFromStrVect origObj.Object(g_sMemberTerm)
      hdb.DBUpdateVect "THSGroupTerm", keyRow, newVect, oldVect
    End If
  End If
  
  ' THSNote table
  Set RowDef = New HRowDef
  RowDef.Add g_sLeadTerm, vbString, MAX_TERM_LEN
  RowDef.Add g_sNoteType, vbString, 4
  Set keyRow = RowDef.CreateRow
  keyRow(1) = leadTerm
  keyRow(2) = "SCOP"
  Set newVect = NewObj.Object(g_sScopeNote)
  If isNew Then
    hdb.DBInsert "THSNote", False, keyRow, newVect
  ElseIf Not newVect.IsEqual(origObj.Object(g_sScopeNote)) Then
    hdb.DBUpdateVect "THSNote", keyRow, newVect
  End If
  ' save changes to hierarchies
  If hdb.RowCount >= 0 Then m_hierManager.SaveChanges cxt
  
  ' need to rename term
  If leadTerm <> NewObj(g_sLeadTerm) Then
    Set RowDef = New HRowDef
    RowDef.Add "OrigTerm", vbString, MAX_TERM_LEN
    RowDef.Add "NewTerm", vbString, MAX_TERM_LEN
    Set keyRow = RowDef.CreateRow
    keyRow(1) = leadTerm
    keyRow(2) = NewObj(g_sLeadTerm)
    hdb.SqlExec "execute sp_THSRenameTerm :Origterm, :NewTerm", False, keyRow
  End If
  
  'return
  SaveObj = hdb.RowCount >= 0
  
End Function

' -------------------------------------------------------------------------
' DeleteObj
'   Deletes term
'   NB: caller is first repsponsible for ensuring term not in any hierarchies
'   Returns False if failed
' -------------------------------------------------------------------------
Public Function DeleteObj(ByVal cxt As HAppContext, _
                ByVal keyObj As HKvList) As Boolean
                
  Dim hdb As HDBBase
  Set hdb = cxt.DB
                
  Dim keyRow As HRow
  Set keyRow = m_THSLeadTermDef.CreateRow
  keyRow(g_sLeadTerm) = keyObj(g_sLeadTerm)
  
  ' called stored proc to validate and do bulk of work
  hdb.SqlExec "execute sp_THSDeleteTerm :LeadTerm", False, keyRow
                
  ' save changes to hierarchies
  If hdb.RowCount >= 0 Then m_hierManager.SaveChanges cxt
                
  'return
  DeleteObj = hdb.RowCount >= 0
                
End Function

' -------------------------------------------------------------------------
' CloneObj
'   Makes a copy of object for editing
' -------------------------------------------------------------------------
Public Function CloneObj(ByVal obj As HKvList) As HKvList
  Dim NewObj As New HKvList
  NewObj(g_sLeadTerm) = obj(g_sLeadTerm)
  NewObj(g_sDisplayForm) = obj(g_sDisplayForm)
  NewObj(g_sDescriptorType) = obj(g_sDescriptorType)
  NewObj(g_sHierarchy) = obj.Object(g_sHierarchy).Clone
  NewObj(g_sUsedFor) = obj.Object(g_sUsedFor).Clone
  NewObj(g_sRelatedTerm) = obj.Object(g_sRelatedTerm).Clone
  NewObj(g_sGroupTerm) = obj.Object(g_sGroupTerm).Clone
  NewObj(g_sMemberTerm) = obj.Object(g_sMemberTerm).Clone
  NewObj(g_sIndicator) = obj.Object(g_sIndicator).Clone
  NewObj(g_sScopeNote) = obj.Object(g_sScopeNote).Clone
  NewObj(g_sDateNote) = obj.Object(g_sDateNote).Clone
  Set CloneObj = NewObj
End Function

' -------------------------------------------------------------------------
' Normalise
'   Normalises fields in object
' -------------------------------------------------------------------------
Public Sub Normalise(ByVal obj As HKvList)
  obj(g_sDisplayForm) = HUtil.TrimSpaces(obj(g_sDisplayForm))
  Dim termVect As HStrVect
  
  Set termVect = obj.Object(g_sUsedFor)
  termVect.Sort
  termVect.Unique
  Set termVect = obj.Object(g_sRelatedTerm)
  termVect.Sort
  termVect.Unique
  Set termVect = obj.Object(g_sGroupTerm)
  termVect.Sort
  termVect.Unique
  Set termVect = obj.Object(g_sMemberTerm)
  termVect.Sort
  termVect.Unique
  Set termVect = obj.Object(g_sIndicator)
  termVect.Sort
  termVect.Unique
  
End Sub

' -------------------------------------------------------------------------
' ValidateObj
'   Validates an object
' -------------------------------------------------------------------------
Public Function ValidateObj(ByVal cxt As HAppContext, _
                ByVal isNew As Boolean, _
                ByVal keyObj As HKvList, _
                ByVal obj As HKvList, _
                ByVal origObj As HKvList, _
                ByVal errInfo As HKvList) As Boolean
  Dim msg As String, errItem As String
  Dim str As String, termVect As HStrVect
  Dim i As Long
  
  ' check term
  If msg = "" Then
    errItem = g_sLeadTerm
    str = obj(g_sLeadTerm)
    If str = "" Then
      msg = E_MISSING_VAL
    ElseIf Not IsValidTerm(str) Then
      msg = "Invalid thesaurus term name"
    End If
  End If
  
  ' check display form
  If msg = "" Then
    errItem = g_sDisplayForm
    str = obj(g_sDisplayForm)
    If str = "" Then
       msg = E_MISSING_VAL
    ElseIf Len(str) > MAX_TERM_LEN Then
       msg = E_STR_TOO_LONG
    ElseIf Not IsAscii(str) Then
       msg = E_STR_BAD_CHARS
    ElseIf HUtil.StrCSpan(str, 1, "<>""&|~\@^`#") <> Len(str) Then
       msg = E_STR_BAD_CHARS
    End If
  End If
        
  ' check descriptor type and validate type
  Dim Info As HKvList
  If msg = "" Then
    errItem = g_sDescriptorType
    str = obj(errItem)
    Set Info = DesTypeInfo(str)
    If str = "" Then
      msg = E_MISSING_VAL
    ElseIf Info Is Nothing Then
      msg = "Invalid descriptor type!"
    End If
  End If
  
  ' check used for terms
  If msg = "" Then
    errItem = g_sUsedFor
    Set termVect = obj.Object(g_sUsedFor)
    For i = 1 To termVect.Count
      str = termVect(i)
      If Len(str) > MAX_TERM_LEN Then
        msg = E_STR_TOO_LONG
      ElseIf Not IsValidTerm(str) Then
        msg = "Invalid thesaurus term name"
      End If
      If msg <> "" Then Exit For
    Next
  End If
  
  ' related terms - not to self
  If msg = "" Then
    errItem = g_sRelatedTerm
    Set termVect = obj.Object(g_sRelatedTerm)
    For i = 1 To termVect.Count
      str = termVect(i)
      If termVect(i) = obj(g_sLeadTerm) Then
        msg = "Term cannot be related to self"
      End If
      If msg <> "" Then Exit For
    Next
  End If
  
  ' check that have required hierarchy if any
  If msg = "" Then
    errItem = g_sHierarchy
    Set termVect = obj.Object(g_sHierarchy)
    Dim permitted As HStrVect
    Set permitted = Info.Object("Hierarchies")
    If permitted Is Nothing Then
      If termVect.Count <> 0 Then
        msg = "Term cannot belong to any hierarchies"
      End If
    Else
      For i = 1 To termVect.Count
        str = termVect(i)
        If permitted.Find(str) = 0 Then
          msg = "Term cannot belong to '" & str & "' hierarchy"
          Exit For
        End If
      Next
      ' added test for DRUG type since they dont require a hierarchy.
      If termVect.Count = 0 And obj(g_sDescriptorType) <> "drugs" Then
        msg = "Term must belong to at least one hierarchy"
      End If
      str = Info("Required")
      If str <> "" Then
        If termVect.Find(str) = 0 Then
          msg = "Term must belong to '" & str & "' hierarchy"
        End If
      End If
    End If
  End If
  
  ' check scope notes are ascii text
  Dim rowVect As HRowVect
  If msg = "" Then
    errItem = g_sScopeNote
    Set rowVect = obj.Object(g_sScopeNote)
    For i = 1 To rowVect.Count
      str = rowVect(i)(g_sNoteText)
      If Len(str) > rowVect.RowDef(g_sNoteText).MaxLength Then
        msg = E_STR_TOO_LONG
      ElseIf Not HUtil.IsAscii(str) Then
        msg = E_STR_BAD_CHARS
      End If
      If msg <> "" Then Exit For
    Next
  End If
  
  ' return
  ValidateObj = msg = ""
  If Not ValidateObj Then
    errInfo("Field") = errItem
    errInfo("ErrMsg") = msg
  End If
  
End Function

' -------------------------------------------------------------------------
' $Log: ThesObj.cls,v $
' Revision 1.11  2006/09/18 04:11:45  GS2
' Added Targets to Drug Class selection list.
'
' Revision 1.10  2006/09/14 04:03:20  GS2
' Added 4 new hierarchies and allowed Drugs type to be in Diagnostics hier.
'
' Revision 1.9  2001/12/10 04:44:08  jd4
' Thesaurus must not add or update Company Terms
'    removed company from available descriptor types
'    removed companies from available hierarchies
'
' New level hierarchy Media-release-type
'   added to available descriptor types
'   added to hierarchies
'
' Revision 1.8  2000/08/15 01:58:37  AH2
' Changes for new Drug-class hierarchies
'
' Revision 1.7  2000/06/20 03:37:35  AH2
' Added definitions for end-points hierarchy
'
' Revision 1.6  1999/12/13 22:56:56  AH2
' Added support for renaming/deletion of terms
'
' Revision 1.5  1999/11/26 01:27:47  AH2
' Replace 'Death' hierarchy with 'Non-disease-states'
' Fixed spelling of 'research-drug' term
'
' Revision 1.4  1999/11/12 02:55:19  AH2
' Fine tuned allowed hierarchies
'
' Revision 1.3  1999/09/12 21:58:54  AH2
' Source update: first releasable version
'
' Revision 1.2  1999/07/14 02:30:57  AH2
' Source update
'
' Revision 1.1  1999/07/11 22:11:38  AH2
' Initial check-in
'

