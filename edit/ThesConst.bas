Attribute VB_Name = "ThesConst"
' **********************************************************************
'
' Module: $Id: ThesConst.bas,v 1.3 1999/09/12 21:58:54 AH2 Exp $
'
' Author: Andrew Howroyd
'
' Created: Jul 1999
'
' Description
'
' **********************************************************************
Option Explicit

Public Const MAX_TERM_LEN As Integer = 200

' Error strings
Public Const E_MISSING_VAL As String = "Required field not entered"
Public Const E_INVALID_FIELD As String = "Invalid field value"
Public Const E_STR_TOO_LONG As String = "String too long"
Public Const E_STR_BAD_CHARS As String = "Invalid characters in string"

Public Const g_sLeadTerm As String = "LeadTerm"
Public Const g_sDisplayForm As String = "DisplayForm"
Public Const g_sIsApproved As String = "IsApproved"
Public Const g_sDesType As String = "DesType"
Public Const g_sUseTerm As String = "UseTerm"

Public Const g_sNoteType As String = "NoteType"
Public Const g_sNoteId As String = "NoteId"
Public Const g_sNoteDate As String = "NoteDate"
Public Const g_sNoteText As String = "NoteText"

Public Const g_sRelCode As String = "RelCode"
Public Const g_sParentTerm As String = "ParentTerm"

Public Const g_sDescriptorType As String = "DescriptorType"
Public Const g_sUsedFor As String = "UsedFor"
Public Const g_sHierarchy As String = "Hierarchy"
Public Const g_sRelatedTerm As String = "RelatedTerm"
Public Const g_sGroupTerm As String = "GroupTerm"
Public Const g_sMemberTerm As String = "MemberTerm"
Public Const g_sIndicator As String = "Indicator"
Public Const g_sScopeNote As String = "ScopeNote"
Public Const g_sDateNote As String = "DateNote"
' -------------------------------------------------------------------------
' $Log: ThesConst.bas,v $
' Revision 1.3  1999/09/12 21:58:54  AH2
' Source update: first releasable version
'
' Revision 1.2  1999/07/14 02:30:56  AH2
' Source update
'
' Revision 1.1  1999/07/11 22:11:37  AH2
' Initial check-in
'

