Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports Zoople

Namespace mshtml


    <ComImport>
    <Guid("3050F220-98B5-11CF-BB82-00AA00BDCE0B")>
    <InterfaceType(ComInterfaceType.InterfaceIsIDispatch)>
    Interface IHTMLTxtRange

        Function duplicate() As IHTMLTxtRange
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1010)>
        Function inRange(
    <[In]>
    <MarshalAs(UnmanagedType.[Interface])> ByVal range As IHTMLTxtRange) As Boolean
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1011)>
        Function isEqual(
    <[In]>
    <MarshalAs(UnmanagedType.[Interface])> ByVal range As IHTMLTxtRange) As Boolean
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1012)>
        Sub scrollIntoView(
    <[In]> ByVal Optional fStart As Boolean = True)
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1013)>
        Sub collapse(
    <[In]> ByVal Optional Start As Boolean = True)
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1014)>
        Function expand(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal Unit As String) As Boolean
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1015)>
        Function move(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal Unit As String,
    <[In]> ByVal Optional Count As Integer = 1) As Integer
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1016)>
        Function moveStart(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal Unit As String,
    <[In]> ByVal Optional Count As Integer = 1) As Integer
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1017)>
        Function moveEnd(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal Unit As String,
    <[In]> ByVal Optional Count As Integer = 1) As Integer

        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1024)>
        Sub [select]()

        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1026)>
        Sub pasteHTML(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal html As String)
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1006)>
        Function parentElement() As Object ' IHTMLElement
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1001)>
        Sub moveToElementText(
    <[In]>
    <MarshalAs(UnmanagedType.[Interface])> ByVal element As Object)
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1025)>
        Sub setEndPoint(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal how As String,
    <[In]>
    <MarshalAs(UnmanagedType.[Interface])> ByVal SourceRange As IHTMLTxtRange)
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1018)>
        Function compareEndPoints(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal how As String,
    <[In]>
    <MarshalAs(UnmanagedType.[Interface])> ByVal SourceRange As IHTMLTxtRange) As Integer
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1019)>
        Function findText(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal _String As String,
    <[In]> ByVal Optional Count As Integer = 1073741823,
    <[In]> ByVal Optional Flags As Integer = 0) As Boolean
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1020)>
        Sub moveToPoint(
    <[In]> ByVal x As Integer,
    <[In]> ByVal y As Integer)

        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1009)>
        Function moveToBookmark(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal Bookmark As String) As Boolean
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1027)>
        Function queryCommandSupported(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal cmdID As String) As Boolean
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1028)>
        Function queryCommandEnabled(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal cmdID As String) As Boolean
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1029)>
        Function queryCommandState(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal cmdID As String) As Boolean
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1030)>
        Function queryCommandIndeterm(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal cmdID As String) As Boolean
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1031)>
        Function queryCommandText(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal cmdID As String) As String
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1032)>
        Function queryCommandValue(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal cmdID As String) As Object
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1033)>
        Function execCommand(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal cmdID As String,
    <[In]> ByVal Optional showUI As Boolean = False,
    <[Optional]>
    <[In]>
    <MarshalAs(UnmanagedType.Struct)> ByVal Optional value As Object = Nothing) As Boolean
        <MethodImpl(MethodImplOptions.InternalCall)>
        <DispId(1034)>
        Function execCommandShowHelp(
    <[In]>
    <MarshalAs(UnmanagedType.BStr)> ByVal cmdID As String) As Boolean
    End Interface
End Namespace


Public Class Form1
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.HtmlEditControl1.DocumentHTML = "<p>Dear <span class=""field"">Enter Name</span></p>
                <p>Please be advised that your <span class=""field"">enter item</span> will be delivered shortly.</p>
                    <br /><p>Mr Delivery</p>"

        Me.HtmlEditControl1.CSSText = "body {font-family: arial} span.field {background-color: #F29AA4}"

    End Sub

    Private Enum SelectionType
        DoesNotContainField = 0
        ContainsFieldAndOtherElements = 1
        OnlyContainsField = 2
    End Enum

    Private Function AllowKeyAction(InteractionType As String, ctlKey As Boolean, keyCode As Integer) As Boolean

        If ctlKey And (keyCode = Keys.A) Then
            Return True
        End If

        If (ctlKey And keyCode = Keys.C) Or (ctlKey And keyCode = Keys.Z) Then
            Return True
        End If

        If keyCode = Keys.Up Or keyCode = Keys.Down Or keyCode = Keys.Left Or keyCode = Keys.Right Then
            Return True
        End If

        If IsCaretInField() Then

            If keyCode = (Keys.Delete) Then
                If CurrentFieldLength() > 1 And CompareStartOrEnd("EndToEnd") < 0 Then
                    Return True
                Else
                    Return False
                End If
            End If

            If keyCode = (Keys.Back) Then
                If CurrentFieldLength() > 1 And CompareStartOrEnd("StartToStart") > 0 Then
                    Return True
                Else
                    Return False
                End If
            End If

            Return True

        End If

        Return False

    End Function

    Private Function IsCaretInField() As SelectionType

        Dim x As Integer

        Dim COMDoc = Me.HtmlEditControl1.Document.DomDocument

        If (COMDoc.selection.Type = "Control") Then
            ' this is for singularly selected images and table element selections etc
            Return False
            Exit Function
        End If

        Dim Range As mshtml.IHTMLTxtRange

        Range = COMDoc.selection.createRange()

        If Range.parentElement.getAttribute("class") IsNot DBNull.Value AndAlso Range.parentElement.getAttribute("class") = "field" Then
            Return SelectionType.OnlyContainsField
        End If

        Do While (x < Range.parentElement.all.length)

            If Range.parentElement.all(x).getAttribute("class") IsNot DBNull.Value AndAlso Range.parentElement.all(x).getAttribute("class") = "field" Then

                Dim testRange As mshtml.IHTMLTxtRange
                testRange = COMDoc.selection.createRange()
                testRange.moveToElementText(Range.parentElement.all(x))

                Dim startPoints As Long, endPoints As Long

                endPoints = Range.compareEndPoints("EndToStart", testRange)
                startPoints = Range.compareEndPoints("StartToEnd", testRange)

                If ((startPoints >= 0) And (endPoints <= 0)) Then
                    Return SelectionType.ContainsFieldAndOtherElements
                End If

                If Not ((startPoints = 1 And endPoints = 1) Or (startPoints = -1 And endPoints = -1)) Then
                    Return SelectionType.DoesNotContainField
                End If

            End If

            x += 1

        Loop

        Return SelectionType.DoesNotContainField

    End Function

    Private Function CompareStartOrEnd(StartToStartOrEndToEnd As String) As Long

        Dim COMDoc = Me.HtmlEditControl1.Document.DomDocument

        If (COMDoc.selection.Type = "Control") Then
            ' this is for singularly selected images and table element selections etc
            Return 0
        End If

        Dim Range As mshtml.IHTMLTxtRange

        Range = COMDoc.selection.createRange()

        If Range.parentElement.getAttribute("class") IsNot DBNull.Value AndAlso Range.parentElement.getAttribute("class") = "field" Then

            Dim testRange As mshtml.IHTMLTxtRange
            testRange = COMDoc.selection.createRange()
            testRange.moveToElementText(Range.parentElement)

            Return Range.compareEndPoints(StartToStartOrEndToEnd, testRange)

        End If

        Return 0

    End Function

    Private Function CurrentFieldLength() As Integer

        Dim Range As mshtml.IHTMLTxtRange
        Dim COMDoc = Me.HtmlEditControl1.Document.DomDocument

        Range = COMDoc.selection.createRange()

        Return Range.parentElement.innerText.length

    End Function

    Private Sub HtmlEditControl1_CancellableUserInteraction(sender As Object, e As CancellableUserInteractionEventsArgs) Handles HtmlEditControl1.CancellableUserInteraction

        If e.InteractionType = "onkeyup" Or e.InteractionType = "onkeydown" Then
            e.Cancel = Not AllowKeyAction(e.InteractionType, e.Keys.Control, e.Keys.Keycode)
        End If

        If e.InteractionType = "onmouseup" Then
            SelectFieldOnMouseUp()
        End If

    End Sub

    Private Sub htmlEditControl1_CommandsToolbarButtonClicked(ByVal sender As Object, ByVal e As CommandsToolbarButtonClickedEventArgs) Handles HtmlEditControl1.CommandsToolbarButtonClicked

        If e.ButtonIdentifier = "tsbCut" AndAlso IsCaretInField() <> SelectionType.OnlyContainsField Then
            MessageBox.Show("Selection contains a read only area and cannot be 'Cut'")
            e.Cancelled = True
        End If

    End Sub

    Private Sub SelectFieldOnMouseUp()

        If HtmlEditControl1.CurrentWindowsFormsElement.GetAttribute("className") = "field" Then

            Dim oRange As mshtml.IHTMLTxtRange
            oRange = HtmlEditControl1.Document.DomDocument.selection.createRange()

            oRange.moveToElementText(HtmlEditControl1.CurrentWindowsFormsElement.DomElement)
            oRange.select()

        End If

    End Sub

End Class
