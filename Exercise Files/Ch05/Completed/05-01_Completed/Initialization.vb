﻿Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices

Public Class Initialization
    Implements IExtensionApplication

    <CommandMethod("LIDBProps")> Public Sub cmdDataBaseProps()
        Dim db As Database = Application.DocumentManager.MdiActiveDocument.Database
        db.Ltscale = 48

    End Sub

#Region "Commands"
    <CommandMethod("MyFirstCommand")> Public Sub cmdMyFirst()
        Dim ed As Editor
        Dim doc As Document
        doc = Application.DocumentManager.MdiActiveDocument
        ed = doc.Editor

        If IsSavedFile() = True Then
            ed.WriteMessage(vbLf & "File is an existing file")
        Else
            ed.WriteMessage(vbLf & "File is a new drawing")
        End If

    End Sub

    <CommandMethod("LIGetVersion")> Public Sub cmdAcadVersion()
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor

        Select Case Application.DocumentManager.MdiActiveDocument.Database.LastSavedAsVersion
            Case Autodesk.AutoCAD.DatabaseServices.DwgVersion.AC1032
                ed.WriteMessage(vbLf & "AutoCAD 2018")
            Case Autodesk.AutoCAD.DatabaseServices.DwgVersion.AC1027
                ed.WriteMessage(vbLf & "AutoCAD 2013")
            Case Autodesk.AutoCAD.DatabaseServices.DwgVersion.AC1024
                ed.WriteMessage(vbLf & "AutoCAD 2010")
            Case Autodesk.AutoCAD.DatabaseServices.DwgVersion.AC1021
                ed.WriteMessage(vbLf & "AutoCAD 2007")
            Case Autodesk.AutoCAD.DatabaseServices.DwgVersion.AC1800, Autodesk.AutoCAD.DatabaseServices.DwgVersion.AC1800a
                ed.WriteMessage(vbLf & "AutoCAD 2004")
            Case Else
                ed.WriteMessage(vbLf & "Uknown or too old")
        End Select

    End Sub

    <CommandMethod("LIGetLoginLastName")> Public Sub cmdLogin()
        Dim logName As String = Application.GetSystemVariable("LoginName")
        Dim lastNameLetter As String = logName.Substring(1, 1).ToUpper

        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        ed.WriteMessage(vbLf & lastNameLetter & logName.Substring(2))

    End Sub

    <CommandMethod("LILoops")> Public Sub cmdLoops()
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor

        Dim arry(5) As Integer
        arry(0) = 10
        arry(1) = 9
        arry(2) = 8
        arry(3) = 7
        arry(4) = 6
        arry(5) = 5

        For Each iObj As Integer In arry
            ed.WriteMessage(vbLf & iObj)
        Next

        For i As Integer = arry.Length - 1 To 0 Step -1
            ed.WriteMessage(vbLf & arry(i))
        Next

        Dim iList As New List(Of Integer)
        iList.Add(1)
        iList.Add(2)
        iList.Add(3)
        iList.Add(4)
        iList.Add(5)

        For Each iObj As Integer In iList
            ed.WriteMessage(vbLf & iObj.ToString)
        Next
    End Sub

    <CommandMethod("LIDocCount")> Public Sub cmdDocCount()
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        ed.WriteMessage(vbLf & "There are " & Application.DocumentManager.Count & " drawings opened.")
    End Sub

    <CommandMethod("LIDocProp")> Public Sub cmdDocProp()
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        acDoc.Editor.WriteMessage(vbLf & acDoc.Name)
    End Sub

    <CommandMethod("LINewDoc")> Public Sub cmdNewDoc()
        Dim docCol As DocumentCollection = Application.DocumentManager
        Dim newDoc As Document = docCol.Add("C:\Users\jmodglin\AppData\Local\Autodesk\AutoCAD 2018\R22.0\enu\Template\acad.dwt")

        docCol.MdiActiveDocument = newDoc
    End Sub
#End Region

#Region "Support Functions"
    Private Function IsSavedFile() As Boolean
        If Application.GetSystemVariable("DWGTITLED") <> 0 Then
            Return True
        Else
            Return False
        End If
    End Function

#End Region

#Region "Extension Routines"
    Public Sub Initialize() Implements IExtensionApplication.Initialize

    End Sub

    Public Sub Terminate() Implements IExtensionApplication.Terminate

    End Sub

#End Region

End Class
