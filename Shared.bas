Attribute VB_Name = "modShared"
Public MainFrm As Form
Public Locations() As String
Public PageNames() As String
Public PageChild() As Boolean
Public UnLoaded As Boolean

Public Sub RetForm(FormClose As Form, OpenFrm As Boolean)

  If OpenFrm = False Then
    MainFrm.Visible = True
    Unload FormClose
  Else
    If FormClose.Name = "frmAbout" Then
        FormClose.Show 1
    Else
        FormClose.Visible = True
    End If
    MainFrm.Hide
  End If
End Sub

Public Sub RefreshFromFile()

  'Load Locations and Page Names from File
  ChDir App.Path
  Dim FileString As String
  
  Dim n As Integer
  n = 0
  FileString = "Initiation Files\pages.nod"
  Open FileString For Input As #1
  
  Do While Not EOF(1)
    n = n + 1
    ReDim Preserve PageNames(n)
    ReDim Preserve PageChild(n)
    Line Input #1, PageNames(n)
    
    If Left(PageNames(n), 1) = vbTab Then
      PageNames(n) = Mid(PageNames(n), 2, Len(PageNames(n)) - 1)
      PageChild(n) = True
    Else
      PageChild(n) = False
    End If
  Loop
  Close #1
  
  ReDim Locations(UBound(PageNames))
  n = 0
  FileString = "Initiation Files\pages.loc"
  Open FileString For Input As #1

  Do While Not EOF(1)
    If n < UBound(PageNames) + 1 Then Line Input #1, Locations(n)
    If n > 300 Then Exit Do
    n = n + 1
  Loop
  Close #1
  
End Sub

Public Sub UpdateLocation(Location As String, Index As Integer)
  'Update Single Location
  Locations(Index) = Location
End Sub

Public Sub SubmitDoc(Location As String, Copies As Integer)
  'Submit Documents to Print Queue
  frmPrint.lstQueue.AddItem (Copies & ":" & Location)
End Sub

