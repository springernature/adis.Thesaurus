VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ThesBrowserObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' **********************************************************************
'
' Module: $Id: ThesBrowserObj.cls,v 1.6 2004/03/22 02:05:50 sp2 Exp $
'
' Author: Andrew Howroyd
'
' Created: May 1999
'
' Description
'
' **********************************************************************
Option Explicit

Private m_THSTermDef As New HRowDef
Private m_THSLeadTermDef As New HRowDef
Private m_THSNoteDef As New HRowDef

' -------------------------------------------------------------------------
' Class_Initialize
' -------------------------------------------------------------------------
Private Sub Class_Initialize()

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
' Search
'    Carries out a search
' -------------------------------------------------------------------------
Public Function Search(ByVal cxt As HAppContext, _
                ByVal args As HKvList, _
                ByVal result As HKvList) As Boolean
  Dim vect As HRowVect
  Set vect = m_THSTermDef.CreateRowVect
  
  Dim sqlWhere As String
  sqlWhere = "LeadTerm Like '" & args("Pattern") & "'"
  If Not args.Contains("IncludeSynonyms") Then
    sqlWhere = sqlWhere & " and IsApproved='Y'"
  Else
    sqlWhere = sqlWhere & " and ((IsApproved='N' and useTerm is not NULL) or (isApproved='Y'))"
  End If
  If args("SqlWhere") <> "" Then
    sqlWhere = sqlWhere & " and (" & args("sqlWhere") & ")"
  End If
  
  If cxt.DB.DBSelect("THSTerm", vect, , sqlWhere, _
                     "order by LeadTerm") >= 0 Then
     result("Terms") = vect
     Search = True
  End If

End Function
                
' -------------------------------------------------------------------------
' FindObj
'    Validates the existance of a term
' -------------------------------------------------------------------------
Public Function FindObj(ByVal cxt As HAppContext, _
                        ByVal key As String, _
                        ByVal sqlWhere As String, _
                        ByVal obj) As Boolean
  Dim hdb As HDBBase
  Dim vect As HRowVect, r As HRow
  Dim ok  As Boolean
  Set hdb = cxt.DB

  Dim dbkey As HRow
  Set dbkey = m_THSLeadTermDef.CreateRow
  dbkey(1) = key

  ' get THSTerm row
  Set vect = m_THSTermDef.CreateRowVect
  ok = hdb.DBSelect("THSTerm", vect, dbkey, sqlWhere) >= 0

  If Not ok Then Exit Function
  
  If vect.Count > 0 Then
    Set r = vect(1)
    If r(g_sIsApproved) = "Y" Then
      obj(g_sLeadTerm) = r(g_sLeadTerm)
    Else
      obj(g_sLeadTerm) = r(g_sUseTerm)
    End If
    obj(g_sDesType) = r(g_sDesType)
  End If

  FindObj = True
End Function
                        

' -------------------------------------------------------------------------
' LoadObj
'    Loads a specific thesaurus object
'    If not found, returns True but result is empty
' -------------------------------------------------------------------------
Public Function LoadObj(ByVal cxt As HAppContext, _
                        ByVal key As HKvList, _
                        ByVal obj) As Boolean

  Dim hdb As HDBBase
  Dim vect As HRowVect
  Dim ok  As Boolean
  Set hdb = cxt.DB
  
  Dim dbkey As HRow
  Set dbkey = m_THSLeadTermDef.CreateRow
  dbkey(1) = UCase(key(g_sLeadTerm))
  
  ' get THSTerm row
  Set vect = m_THSTermDef.CreateRowVect
  ok = hdb.DBSelect("THSTerm", vect, dbkey) >= 0
  
  ' copy THSTerm row to obj - return if none
  If ok Then
    If vect.Count = 0 Then
      LoadObj = True
      Exit Function
    End If
    obj(g_sTerm) = vect(1)
  End If
  
  ' get related terms
  If ok Then
    Set vect = m_THSLeadTermDef.CreateRowVect
    ok = hdb.SqlQuery("select RelatedTerm from THSRelatedTerm" & _
                      " where LeadTerm=:LeadTerm order by RelatedTerm", _
                      vect, dbkey) >= 0
    obj("RT") = vect
  End If
  
  ' get broader terms
  If ok Then
    Set vect = m_THSLeadTermDef.CreateRowVect
    ok = hdb.SqlQuery("select ParentTerm from THSGroupTerm" & _
                      " where LeadTerm=:LeadTerm order by ParentTerm", _
                      vect, dbkey) >= 0
  End If
  If ok Then
    ok = hdb.SqlQuery("select distinct ParentTerm from THSHierarchy" & _
                      " where LeadTerm=:LeadTerm order by ParentTerm", _
                      vect, dbkey) >= 0
    obj("BT") = vect
  End If
  
  ' get narrower terms
  If ok Then
    Set vect = m_THSLeadTermDef.CreateRowVect
    ok = hdb.SqlQuery("select LeadTerm from THSGroupTerm" & _
                      " where ParentTerm=:LeadTerm order by LeadTerm", _
                      vect, dbkey) >= 0
  End If
  If ok Then
    ok = hdb.SqlQuery("select distinct LeadTerm from THSHierarchy" & _
                      " where ParentTerm=:LeadTerm order by LeadTerm", _
                      vect, dbkey) >= 0
    obj("NT") = vect
  End If
  
  ' get used for terms
  If ok Then
    Set vect = m_THSLeadTermDef.CreateRowVect
    ok = hdb.SqlQuery("select LeadTerm from THSTerm" & _
                      " where UseTerm=:LeadTerm order by LeadTerm", _
                      vect, dbkey) >= 0
    obj("UF") = vect
  End If
  
  ' get scope notes
  If ok Then
    Set vect = m_THSNoteDef.CreateRowVect
    ok = hdb.DBSelect("THSNote", vect, dbkey, _
                      "NoteType='SCOP'", "order by NoteId") >= 0
    
    obj("ScopeNote") = vect
  End If
  
  ' get date notes
  If ok Then
    Set vect = m_THSNoteDef.CreateRowVect
    ok = hdb.DBSelect("THSNote", vect, dbkey, _
                      "NoteType='DATE'", "order by NoteDate, NoteId") >= 0
    obj("DateNote") = vect
  End If
  LoadObj = ok
End Function

' -------------------------------------------------------------------------
' $Log: ThesBrowserObj.cls,v $
' Revision 1.6  2004/03/22 02:05:50  sp2
' Fix to thesaurus tool related to Ody 1337.  As disabled companies
' do not often have a use term, need to prevent tool showing any
' unapproved terms with no use-terms
'
' Revision 1.5  1999/11/26 01:25:48  AH2
' Added used for tab
' Help file integration
'
' Revision 1.4  1999/09/12 21:59:53  AH2
' Added DescriptorOnly check and descriptor type support
'
' Revision 1.3  1999/08/09 00:11:31  AH2
' Quick hack to get some validation going
'
' Revision 1.2  1999/08/01 22:51:34  AH2
' Made private
'
' Revision 1.1  1999/05/16 21:57:14  AH2
' Initial check-in
'

