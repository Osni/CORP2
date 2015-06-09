
Public Class ClsFilterCols
    Inherits PropertyCollection

    Private _Name As String = String.Empty
    Private _Label As String = String.Empty
    Private _Title As String = String.Empty
    Private _Size As Integer = 10
    Private _TypeDB As Integer = TypeDB.STRING_T
    Private _Visible As Boolean = True
    Private _Style As TStyle = TStyle.FilterField
    Private _TextValue As String = String.Empty
    Private _PageURLDestino As String = String.Empty
    Private _PageURLColVar As String = String.Empty
    Private prt As PropertyCollection    
    Public Enum TypeDB As Integer
        STRING_T = 1
        DATE_T = 2
        DATE_TIME_T = 3
        NUMERIC_T = 4
        MONEY_T = 5
    End Enum
    Public Enum TStyle
        FilterField        
        ImageButton
    End Enum
    Public Property Style() As TStyle
        Get
            Return _Style
        End Get
        Set(ByVal value As TStyle)
            _Style = value
        End Set
    End Property
    Public Property PageURLDestino() As String
        Get
            Return _PageURLDestino
        End Get
        Set(ByVal value As String)
            _PageURLDestino = value
        End Set
    End Property
    Public Property PageURLColVar() As String
        Get
            Return _PageURLColVar
        End Get
        Set(ByVal value As String)
            _PageURLColVar = value
        End Set
    End Property

    Public Property TextValue() As String
        Get
            Return _TextValue
        End Get
        Set(ByVal value As String)
            _TextValue = value
        End Set
    End Property
    Public ReadOnly Property FilterColsReadOnly() As PropertyCollection
        Get
            Return prt
        End Get
    End Property
    Public Property TypeCOL() As TypeDB
        Get
            Return _TypeDB
        End Get
        Set(ByVal value As TypeDB)
            _TypeDB = value
        End Set
    End Property
    Public Property Name() As String
        Get
            Return _Name
        End Get
        Set(ByVal value As String)
            _Name = value
        End Set
    End Property
    Public Property Size() As String
        Get
            Return _Size
        End Get
        Set(ByVal value As String)
            _Size = value
        End Set
    End Property
    Public Property Label() As String
        Get
            Return _Label
        End Get
        Set(ByVal value As String)
            _Label = value
        End Set
    End Property

    Public Property Title() As String
        Get
            Return _Title
        End Get
        Set(ByVal value As String)
            _Title = value
        End Set
    End Property

    Public Property Visible() As Boolean
        Get
            Return _Visible
        End Get
        Set(ByVal value As Boolean)
            _Visible = value
        End Set
    End Property

    Public Function GetFilterCols() As PropertyCollection
        Try
            prt = New PropertyCollection
            With prt
                .Add("Label", _Label)
                .Add("Name", _Name)
                .Add("Title", _Title)
                .Add("Visible", _Visible)
                .Add("Size", _Size)
                .Add("Type", Int(_TypeDB))
                .Add("TextValue", _TextValue)
                .Add("PageURLDestino", _PageURLDestino)
                .Add("PageURLColVar", _PageURLColVar)
                .Add("Style", _Style)
            End With
            Return prt
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
            Return Nothing
        End Try
    End Function
    Public Sub New(Optional ByVal sName As String = "", _
            Optional ByVal sLabel As String = "", _
            Optional ByVal sTitle As String = "", _
            Optional ByVal bVisible As Boolean = True, _
            Optional ByVal iSize As Integer = 10, _
            Optional ByVal eTypeDB As TypeDB = TypeDB.STRING_T, _
            Optional ByVal sTextValue As String = "", _
            Optional ByVal sPageURLDestino As String = "", _
            Optional ByVal sPageURLColVar As String = "", _
            Optional ByVal Style As TStyle = TStyle.FilterField)
        Try
            If sName = "" Then
                Exit Sub
            End If
            prt = New PropertyCollection
            With prt
                .Add("Name", sName)
                .Add("Label", sLabel)
                .Add("Title", sTitle)
                .Add("Visible", bVisible)
                .Add("Size", iSize)
                .Add("Type", Int(eTypeDB))
                .Add("TextValue", sTextValue)
                .Add("PageURLDestino", sPageURLDestino)
                .Add("PageURLColVar", sPageURLColVar)
                .Add("Style", Style)
            End With
        Catch ex As Exception
            Throw New System.Exception(ex.Message)
        End Try
    End Sub

End Class

