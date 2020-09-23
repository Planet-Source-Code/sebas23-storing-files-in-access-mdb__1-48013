VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   3165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3120
   ScaleHeight     =   3165
   ScaleWidth      =   3120
   Begin VB.PictureBox Picture1 
      Height          =   525
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   1185
      TabIndex        =   0
      Top             =   0
      Width           =   1245
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim cn As DBFileBlob
Dim rs As Recordset
Dim mstream As ADODB.Stream
'Default Property Values:
Const m_def_BackColor = 0
Const m_def_ForeColor = 0
Const m_def_Enabled = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
'Property Variables:
Dim m_BackColor As Long
Dim m_ForeColor As Long
Dim m_Enabled As Boolean
Dim m_Font As Font
Dim m_BackStyle As Integer
Dim m_BorderStyle As Integer
'Event Declarations:
Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick()
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."



Private Sub UserControl_Resize()
Picture1.Height = Height
Picture1.Width = Width
End Sub


Public Property Get BackColor() As Long
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As Long)
    m_BackColor = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get ForeColor() As Long
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As Long)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    PropertyChanged "Font"
End Property

Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = m_BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
End Property

Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property


Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = Picture1.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set Picture1.Picture = New_Picture
    PropertyChanged "Picture"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_Enabled = m_def_Enabled
    Set m_Font = Ambient.Font
    m_BackStyle = m_def_BackStyle
    m_BorderStyle = m_def_BorderStyle
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub

Private Sub UserControl_Terminate()
On Local Error Resume Next
cn.Conn.Close
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
End Sub


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function GetFileFromDB(DatabaseName As String, Query As String, FieldName As String, FileName As String) As Variant
Dim sFilename As String
Set cn = New DBFileBlob
Set rs = New Recordset
On Error GoTo Command1_Error
sFilename = FileName
cn.DDataBase = DatabaseName
cn.Connect
rs.Open Query, cn.Conn, adOpenKeyset, adLockOptimistic
If rs.RecordCount = 0 Then Picture1.Picture = Nothing: Exit Function
Set mstream = New ADODB.Stream
mstream.Type = adTypeBinary
mstream.Open
mstream.Write rs.Fields(FieldName).Value
mstream.SaveToFile sFilename, adSaveCreateOverWrite
Picture1.Picture = LoadPicture(sFilename)
rs.Close
Command1_Error:
 If err.Number <> 0 Then MsgBox Str(err) & " - " & Error, vbExclamation
End Function

Public Function PutFileToDB(DatabaseName As String, Query As String, FieldName As String, FileNameToPut As String) As Variant
Set cn = New DBFileBlob
Set rs = New ADODB.Recordset
cn.DDataBase = DatabaseName
cn.Connect
rs.Open Query, cn.Conn, adOpenKeyset, adLockOptimistic
On Error GoTo Command2_Error
Set mstream = New ADODB.Stream
mstream.Type = adTypeBinary
mstream.Open
    If Len(FileNameToPut) <> 0 Then
        mstream.LoadFromFile FileNameToPut
        Picture1.Picture = LoadPicture(FileNameToPut)
        rs.Fields(FieldName).Value = mstream.Read
        rs.Update
    End If
rs.Close
Exit Function
Command2_Error:
 If err.Number <> 0 Then MsgBox Str(err) & " - " & Error, vbExclamation

End Function

