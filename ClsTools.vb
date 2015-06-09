Imports MSXML2
Imports MSScriptControl
Imports System
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.HtmlControls
Imports System.Web
Imports System.Text
Imports System.Text.RegularExpressions
Imports Microsoft.VisualBasic
Imports System.Security.Cryptography
Imports System.IO

Public Class ClsTools
    Inherits WebControl

    Public Enum TMsgStyleIcon
        MSG_ERROR
        MSG_WARNING
        MSG_INFORMATION
    End Enum

    '==========================================================================
    ' Limpar campos na tela
    '==========================================================================
    Public Sub LimpaCampos(ByRef Form As Object)
        For Each ctrl As Object In Form.Controls
            Select Case ctrl.GetType.ToString()
                Case "System.Web.UI.WebControls.TextBox"
                    CType(ctrl, TextBox).Text = String.Empty

                Case "System.Web.UI.WebControls.DropDownList"
                    CType(ctrl, DropDownList).SelectedIndex = -1

                Case "System.Web.UI.WebControls.RadioButton"
                    CType(ctrl, RadioButton).Checked = False

                Case "System.Web.UI.WebControls.CheckBox"
                    CType(ctrl, CheckBox).Checked = False
            End Select
        Next
    End Sub

    Public Sub SetReadOnly(ByRef Form As Object, Optional ByVal blnStatus As Boolean = True)
        For Each ctrl As Object In Form.Controls
            Select Case ctrl.GetType.ToString()
                Case "System.Web.UI.WebControls.TextBox"
                    With CType(ctrl, TextBox)
                        .ReadOnly = blnStatus
                        If blnStatus Then
                            .BackColor = System.Drawing.Color.FromArgb(0, 240, 240, 240)
                        Else
                            .BackColor = System.Drawing.Color.FromArgb(0, 255, 255, 255)
                        End If
                    End With
                Case "System.Web.UI.WebControls.RadioButton", "System.Web.UI.WebControls.CheckBox", _
                     "System.Web.UI.WebControls.DropDownList"
                    If ctrl.GetType.ToString() = "System.Web.UI.WebControls.DropDownList" Then
                        If blnStatus Then
                            ctrl.BackColor = System.Drawing.Color.FromArgb(0, 240, 240, 240)
                        Else
                            ctrl.BackColor = System.Drawing.Color.FromArgb(0, 255, 255, 255)
                        End If
                    End If
                    ctrl.Enabled = Not blnStatus
            End Select
        Next
    End Sub

    Public Sub ShowMessage(ByVal strTextoMsg As String, ByVal strTituloMsg As String, _
            ByVal sender As Object, _
            Optional ByVal msgStyle As TMsgStyleIcon = TMsgStyleIcon.MSG_ERROR, _
            Optional ByVal strScriptTag As String = "", _
            Optional ByVal strScript As String = "")


        Dim TabCentro As New Table
        Dim TabMens As New Table
        Dim Row As TableRow
        Dim Cell As TableCell


        Dim strMyScript As New StringBuilder
        Dim imgIcone As New Image
        Dim lblTitulo As New Label
        Dim lblTexto As New Label
        Dim btnOk As New HtmlInputButton

        With TabMens
            .ID = "TabMens"
            .Style.Add("width", "350px")
            .CellPadding = 0
            .CellSpacing = 0
        End With
        '-------------------------------
        Row = New TableRow
        '-------------------------------
        Select Case msgStyle
            Case TMsgStyleIcon.MSG_ERROR
                imgIcone.ImageUrl = "imagens/msg/error.gif"

            Case TMsgStyleIcon.MSG_INFORMATION
                imgIcone.ImageUrl = "imagens/msg/information.gif"

            Case TMsgStyleIcon.MSG_WARNING
                imgIcone.ImageUrl = "imagens/msg/warning.gif"

        End Select
        '-------------------------------
        'Registrando script
        With strMyScript
            .AppendLine()
            '.AppendLine("function ShowMsg() {")
            '.AppendLine("	var x = document.getElementById('TabMens');")
            '.AppendLine(" 	var ifra = document.getElementById('MsgIframe').style;")
            '.AppendLine(" 	var vTabCentro = document.getElementById('TabCentro');")
            '.AppendLine("	vTabCentro.style.height = document.body.offsetHeight + 'px';")
            '.AppendLine("	vTabCentro.style.width = document.body.offsetWidth + 'px';")
            '.AppendLine("	ifra.top=x.offsetTop + 'px';")
            '.AppendLine("	ifra.left=x.offsetLeft + 'px';")
            '.AppendLine("	ifra.width=x.offsetWidth + 'px';")
            '.AppendLine("	ifra.height=x.offsetHeight + 'px';")
            '.AppendLine("	ifra.position='absolute';")
            '.AppendLine("	ifra.visibility='inherit';")
            '.AppendLine("	ifra.zIndex=2;")
            '.AppendLine("	x.style.zIndex=3;")
            '.AppendLine("}")
            '.AppendLine()
            '.AppendLine("window.onload = ShowMsg;")
            '.AppendLine("window.onresize = ShowMsg; ")

            .AppendLine("function ShowMsg() {")
            .AppendLine("	var x = document.getElementById('TabMens');")
            .AppendLine(" 	var ifra = document.getElementById('MsgIframe').style;")
            .AppendLine(" 	var shd = document.getElementById('shd').style;")
            .AppendLine(" 	var vTabCentro = document.getElementById('TabCentro');")
            .AppendLine("	vTabCentro.style.height = document.body.offsetHeight + 'px';")
            .AppendLine("	vTabCentro.style.width = document.body.offsetWidth + 'px';")
            .AppendLine()
            .AppendLine("   with (shd) {")
            .AppendLine("        top=x.offsetTop + 10 + 'px';")
            .AppendLine("        left=x.offsetLeft + 10  + 'px';")
            .AppendLine("        width=x.offsetWidth + 'px';")
            .AppendLine("        height=x.offsetHeight + 'px';")
            .AppendLine("        visibility='inherit';")
            .AppendLine("        zIndex=0;")
            .AppendLine("    }")
            .AppendLine()
            .AppendLine("    with (ifra) {")
            '.AppendLine("    	top=x.offsetTop + 'px';")
            '.AppendLine("    	left=x.offsetLeft + 'px';")
            '.AppendLine("    	width=x.offsetWidth + 12 +'px';")
            '.AppendLine("    	height=x.offsetHeight + 12 + 'px';")


            .AppendLine("    	top = '0px';")
            .AppendLine("    	left = '0px';")
            .AppendLine("    	width = document.body.offsetWidth + 'px';")
            .AppendLine("    	height = document.body.offsetHeight + 'px';")

            .AppendLine("    	position='absolute';")
            .AppendLine("    	visibility='inherit';")
            .AppendLine("    	zIndex=2;")
            .AppendLine("    	border =0;")
            .AppendLine("	}")
            .AppendLine("	x.style.zIndex=9;")
            .AppendLine("}")
            .AppendLine()
            .AppendLine("window.onload = ShowMsg;")
            .AppendLine("window.onresize = ShowMsg; ")
            .AppendLine()
            .AppendLine("function CloseMsg() {")
            .AppendLine("	document.getElementById('TabCentro').style.display='none';")
            .AppendLine("	document.getElementById('MsgIframe').style.display='none';")
            .AppendLine("	document.getElementById('shd').style.display='none';")
            .AppendLine("}")
        End With
        '-------------------------------
        If Not sender.ClientScript.IsClientScriptBlockRegistered("msgscript") Then _
            sender.ClientScript.RegisterClientScriptBlock(Me.GetType, "msgscript", strMyScript.ToString, True)
        '-------------------------------
        'Titulo da caixa
        Cell = New TableHeaderCell
        With Cell
            .ColumnSpan = 2
            .VerticalAlign = VerticalAlign.Middle
            .HorizontalAlign = HorizontalAlign.Center
            .Style.Add("height", "31px")
            If strTituloMsg.Trim = "" Then strTituloMsg = "Mensagem"
            With lblTitulo
                .Text = strTituloMsg.Replace(vbCrLf, "<br />")
            End With
            .Controls.Add(lblTitulo)
        End With
        '-------------------------------
        Row.Cells.Add(Cell)
        TabMens.Rows.Add(Row)
        '-------------------------------
        'Ícone
        Row = New TableRow
        Cell = New TableCell
        '-------------------------------
        With Cell
            .ID = "CellIMG"
            .VerticalAlign = VerticalAlign.Top
            .HorizontalAlign = HorizontalAlign.Center
            .Style.Add("height", "70px")
            .Controls.Add(imgIcone)
        End With
        '-------------------------------
        Row.Cells.Add(Cell)
        '-------------------------------          
        'Corpo da mensagem
        Cell = New TableCell
        With Cell
            .ID = "CellText"
            .VerticalAlign = VerticalAlign.Middle
            .HorizontalAlign = HorizontalAlign.Center
            With .Style
                .Add("width", "290px")
                .Add("text-align", "justify")
                .Add("vertical-align", "middle")
            End With
            If strTextoMsg.Length > 100 Then
                lblTexto.Text = "<div style='height:100px;overflow:auto'>" & strTextoMsg.Replace(vbCrLf, "<br />") & "</div>"
            Else
                lblTexto.Text = strTextoMsg
            End If
            .Controls.Add(lblTexto)
        End With
        '-------------------------------
        Row.Cells.Add(Cell)
        '-------------------------------
        TabMens.Rows.Add(Row)
        '-------------------------------            
        'Inserindo linha do botão inferior
        With btnOk
            .ID = "btnOK"
            .Value = "Ok"
            If strScript.Trim = "" Then
                .Attributes.Add("onclick", "javascript:CloseMsg();")
            Else
                If strScriptTag.Trim = "" Then
                    .Attributes.Add("onclick", "javascript:CloseMsg();" & strScript)
                Else
                    .Attributes.Add("onclick", "javascript:CloseMsg();" & strScriptTag)
                    sender.ClientScript.RegisterClientScriptBlock(sender.GetType, strScriptTag, strScript, True)
                End If
            End If
            .Attributes.Add("title", " Clique para fechar janela ")
        End With
        '-------------------------------            
        Row = New TableRow
        Cell = New TableCell
        With Cell
            .ColumnSpan = 2
            .VerticalAlign = VerticalAlign.Middle
            .HorizontalAlign = HorizontalAlign.Center
            .Style.Add("height", "50px")
            .Controls.Add(btnOk)
        End With
        '-------------------------------
        Row.Cells.Add(Cell)
        TabMens.Rows.Add(Row)
        '-------------------------------
        'Inserindo na tabela de centro
        With TabCentro
            .ID = "TabCentro"
            .Attributes.Add("onload", "javascript:ShowMsg()")
            With .Style
                .Add("top", "1px")
                .Add("left", "1px")
                .Add("z-index", "3")
                .Add("position", "absolute")
            End With
            .BorderWidth = Unit.Pixel(0)
            .CellPadding = 0
            .CellSpacing = 0
        End With
        '-------------------------------
        Row = New TableRow
        Cell = New TableCell
        With Cell
            .HorizontalAlign = HorizontalAlign.Center
            .VerticalAlign = VerticalAlign.Middle
            .Height = Unit.Pixel("420")
            .Controls.Add(TabMens)
        End With
        Row.Cells.Add(Cell)
        TabCentro.Rows.Add(Row)
        '-------------------------------
        sender.Controls.Add(TabCentro)
        sender.Controls.Add(New LiteralControl("<DIV ID=""shd""></DIV>"))
        sender.Controls.Add(New LiteralControl("<IFRAME style='visibility:hidden;z-index:2;FILTER: Alpha(Opacity=0);' src='about:blank' id='MsgIframe'></IFRAME>"))
    End Sub

    'Public Function RemTags2(ByVal strTexto As String) As String
    '    Dim strTexto2 As String = String.Empty
    '
    '
    '   Dim Server As HttpServerUtility = HttpContext.Current.Server
    '
    '       For lngCount As Long = 0 To strTexto.Length - 1
    '          If strTexto.Substring(lngCount, 1) = "<" Then
    '             While strTexto.Substring(lngCount, 1) <> ">"
    '                lngCount += 1
    '           End While
    '      Else
    '               strTexto2 = strTexto2 & strTexto.Substring(lngCount, 1)
    '          End If
    '     Next
    '    If Server IsNot Nothing Then
    '       Return Server.HtmlDecode(strTexto2)
    '   Else
    '        Return strTexto2
    '   End If
    'End Function


    Function RemTags(ByVal strHTML As String) As String
        'Remove tags HTML

        Dim strOutput As String
        Dim objRegExp As Regex = New Regex("<(.|\n)+?>")

        strOutput = objRegExp.Replace(strHTML, "")

        Return strOutput
    End Function


    ''' <summary>
    ''' Checa comando arbitrário.
    ''' </summary>
    ''' <param name="strExpressao">Expressão a ser testada.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function CheckCommandSQL(ByVal strExpressao As String) As String
        Dim vCmdArbitrario As Object
        Dim vArrExpr As Object
        Try
            vArrExpr = Split(strExpressao, " ")
            vCmdArbitrario = Split("select#insert#update#delete#drop#--#'", "#")
            For i As Integer = 0 To UBound(vCmdArbitrario)
                For l As Integer = 0 To UBound(vArrExpr)
                    If LCase(vCmdArbitrario(i)) = LCase(vArrExpr(l)) Then Throw New Exception("Comando Arbitrário")
                Next
            Next
            Return String.Empty
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

End Class
Public Structure CorpCripto

    Private aCorpCrip As String
    Private Shared des As New TripleDESCryptoServiceProvider()
    Private Shared k() As Byte = Encoding.Unicode.GetBytes("etujwxrr")
    Private Shared v() As Byte = Encoding.Unicode.GetBytes("26zzgg4t")

    '####################################################
    Public Shared Function EncryptString(ByVal encryptValue As String) As String
        Dim valBytes As Byte() = Encoding.Unicode.GetBytes(encryptValue)
        Dim transform As System.Security.Cryptography.ICryptoTransform = des.CreateEncryptor(k, v)
        Dim ms As New IO.MemoryStream()
        Dim cs As New System.Security.Cryptography.CryptoStream(ms, transform, CryptoStreamMode.Write)
        cs.Write(valBytes, 0, valBytes.Length)
        cs.FlushFinalBlock()
        Dim returnBytes As Byte() = ms.ToArray()
        cs.Close()
        Return ToHex(Convert.ToBase64String(returnBytes))
    End Function 'EncryptString
    '####################################################
    Public Shared Function DecryptString(ByVal encryptedValue As String) As String
        Dim valBytes As Byte() = Convert.FromBase64String(HexToString(encryptedValue))

        Dim transform As System.Security.Cryptography.ICryptoTransform = des.CreateDecryptor(k, v)

        Dim ms As New MemoryStream()
        Dim cs As New System.Security.Cryptography.CryptoStream(ms, transform, CryptoStreamMode.Write)
        cs.Write(valBytes, 0, valBytes.Length)
        cs.FlushFinalBlock()
        Dim returnBytes As Byte() = ms.ToArray()
        cs.Close()

        Return Encoding.Unicode.GetString(returnBytes)
    End Function
    '####################################################
    Public Shared Function ToHex(ByVal byteArray As String) As String
        Dim outString As String = ""
        Dim returnBytes() As Byte = ASCIIEncoding.ASCII.GetBytes(byteArray)
        Dim b As [Byte]
        For Each b In returnBytes
            outString &= Hex(b)
        Next b
        Return outString
    End Function 'ByteToHex



    '####################################################
    Public Shared Function HexToString(ByVal hexString As String) As String
        Dim returnBytes(hexString.Length / 2) As Byte
        Dim outString As String = ""
        Dim i As Integer
        For i = 1 To hexString.Length Step 2
            outString &= Chr(Convert.ToByte(Mid(hexString, i, 2), 16))
        Next i
        Return outString
    End Function 'HexToByte

End Structure