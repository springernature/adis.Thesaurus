VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ThesEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' **********************************************************************
'
' Module: $Id: ThesEdit.cls,v 1.1 1999/07/11 22:11:38 AH2 Exp $
'
' Author: Andrew Howroyd
'
' Created: Jul 1999
'
' Description
'    Entry point module for thesaurus editor
'
' **********************************************************************
Option Explicit

Implements HAppBase

' -----------------------------------------------------------------------
' Invoke:
'    Application component entry function
' -----------------------------------------------------------------------
Private Sub HAppBase_Invoke(ByVal args As HUtil.HKvList, _
            ByVal cxt As HUtil.HAppContext)
  Dim frm As New frmThesEdit
  Dim ThesObj As New ThesObj
  
  Dim ok As Boolean
  ok = True
  
  Set frm.m_context = cxt
  Set frm.m_thesObj = ThesObj
  If args("Action") = "Edit" Then
    frm.m_isNew = False
    frm.m_leadTerm = args("Term")
  ElseIf args("Action") = "New" Then
    frm.m_isNew = True
    frm.m_leadTerm = ""
  Else
    Err.Raise -1, , "Invalid call to Invoke"
  End If
  
  If ok Then
    Load frm
    If frm.m_ok Then frm.Show Else Unload frm
  End If

End Sub

' -------------------------------------------------------------------------
' $Log: ThesEdit.cls,v $
' Revision 1.1  1999/07/11 22:11:38  AH2
' Initial check-in
'

