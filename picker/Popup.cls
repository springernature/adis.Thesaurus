VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Popup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' **********************************************************************
'
' Module: $Id: Popup.cls,v 1.7 2004/03/30 04:06:28 GS2 Exp $
'
' Author: Andrew Howroyd
'
' Created: May 1999
'
' Description
'     Popup for locating valid thesaurus terms
'     Also invoked modelessly through HAppBase to start editorial
'
' **********************************************************************
Option Explicit
Implements HAppBase

' -------------------------------------------------------------------------
' Invoke
'   Thesaurus popup component entry point
'   Input parameters:
'      Term: initial search string
'   Output parameters
'      Term: selected term
' -------------------------------------------------------------------------
Function Invoke(ByVal owner As Object, _
                ByVal args As HKvList, _
                ByVal results As HKvList, _
                ByVal cxt As HAppContext) As Boolean

  Dim frm As New frmThesaurus
  Set frm.m_context = cxt
  Set frm.m_thesObj = New ThesBrowserObj
  frm.m_term = args("Term")
  frm.m_validate = args("Validate") = "Y"
  frm.m_sqlWhere = args("SqlWhere")
  frm.m_descriptorOnly = args("DescriptorOnly") = "Y"
  frm.m_endpointOnly = args("EndpointsOnly") = "Y"
  frm.m_multipleAdd = args("MultipleAdd") = "Y"
  
  Load frm
  If frm.m_ok Then
    If owner Is Nothing Then
      frm.Show 1
    Else
      frm.Show 1, owner
    End If
    If frm.m_ok Then
      Invoke = True
      results("Term") = frm.m_term
    End If
  End If
  Unload frm
End Function

' -------------------------------------------------------------------------
' Invoke (HAppBase)
'   modeless entry point
' -------------------------------------------------------------------------
Private Sub HAppBase_Invoke(ByVal args As HUtil.HKvList, _
            ByVal cxt As HUtil.HAppContext)
  Dim ok As Boolean
  ok = True
  
  Dim f As New frmThesaurus
  Set f.m_context = cxt
  Set f.m_thesObj = New ThesBrowserObj
  f.m_term = args("Term")
  f.m_ok = True
  f.m_isModeless = True
  f.m_navigate = args("Navigate")
  
  Load f
  If f.m_ok Then f.Show Else Unload f

End Sub

' -------------------------------------------------------------------------
' Validate
'   Thesaurus validation entry point
'   Input parameters:
'      key: initial search string
' -------------------------------------------------------------------------
Function Validate(ByVal cxt As HAppContext, _
                        ByVal key As String, _
                        ByVal sqlWhere As String, _
                        ByVal obj) As Boolean
  
  Dim m_thesObj As New ThesBrowserObj
  
  Validate = m_thesObj.FindObj(cxt, key, sqlWhere, obj)
  
End Function


' -------------------------------------------------------------------------
' $Log: Popup.cls,v $
' Revision 1.7  2004/03/30 04:06:28  GS2
' Added multiple add mode.
'
' Revision 1.6  2001/10/29 20:55:31  jm5
' CTI264 added public function Validate which validates an entry in the thesaurus
' Modified Invoke so that it can be called from a web page without a owner window
'
' Revision 1.5  2000/08/01 21:35:29  RB1
' Added support for endpoints hierarchy
'
' Revision 1.4  1999/09/12 21:59:53  AH2
' Added DescriptorOnly check and descriptor type support
'
' Revision 1.3  1999/08/09 00:11:31  AH2
' Quick hack to get some validation going
'
' Revision 1.2  1999/08/01 22:52:34  AH2
' Added modeless Invoke to support launch of thesaurus editorial
'
' Revision 1.1  1999/05/16 21:57:14  AH2
' Initial check-in
'

