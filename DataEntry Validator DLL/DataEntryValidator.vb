Imports System.Windows.Forms

Namespace DynamicComponents
    <ComClass(DataEntryValidator.ClassId, DataEntryValidator.InterfaceId, DataEntryValidator.EventsId)> _
    Public Class DataEntryValidator

#Region "COM GUIDs"
        ' These  GUIDs provide the COM identity for this class 
        ' and its COM interfaces. If you change them, existing 
        ' clients will no longer be able to access the class.
        Public Const ClassId As String = "f260e473-8aa6-44bf-9600-59cda99dbb17"
        Public Const InterfaceId As String = "966e6695-163d-4ff4-ac0d-81e2b6c3aa4b"
        Public Const EventsId As String = "49c251aa-422e-4690-a95e-0b6d97476b73"
#End Region

        ' A creatable COM class must have a Public Sub New() 
        ' with no parameters, otherwise, the class will not be 
        ' registered in the COM registry and cannot be created 
        ' via CreateObject.
        Public Sub New()
            MyBase.New()
        End Sub


        Private Col_ControlName As New Collection()
        Private Col_ControlIndex As New Collection()
        Private Col_FieldsType(30, 1) As String
        Private Col_FieldsTypePos As Byte = 0
        Private m_SpecialChars As String
        Private MyForm As New System.Windows.Forms.Form()
        Private m_DecimalPlaces As Byte
        'Public Shared Internal_Registration_Key As String = "0000-000-000-000-0000"

        Private Function ZeroPad(ByVal str_String As String, ByVal int_Count As Byte) As String
            If str_String <> "" And int_Count <> 0 Then
                Return (New String("0", int_Count - Len(Trim(str_String))) & Trim(str_String))
            ElseIf int_Count = 0 Then
                Return str_String
            Else : Return str_String
            End If
        End Function

        Public Sub InitForm(ByRef dm_Form As System.Windows.Forms.Form)

            Dim TxtCtrl As New Control()
            Dim X As Byte


            MyForm = dm_Form

            X = 0
            For Each TxtCtrl In dm_Form.Controls
                If TypeName(TxtCtrl) = "TextBox" Then
                    Col_ControlName.Add(TxtCtrl)
                    Col_ControlIndex.Add(X)
                End If
                X += 1
            Next TxtCtrl

        End Sub

        Private Sub SpecialChars(ByVal str_Chars As String)
            m_SpecialChars = str_Chars
        End Sub

        Private Sub SpecialCharsFields(ByVal ParamArray str_SpecialFields() As String)
            Dim Num As Integer
            Dim cField As String

            For Each cField In str_SpecialFields
                For Num = 1 To Col_ControlName.Count()
                    If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                        Dim txt_control As TextBox = CType(MyForm.Controls(Col_ControlIndex(Num)), TextBox)
                        AddHandler txt_control.KeyPress, AddressOf SpecialCharsFields_KeyPress
                        Col_FieldsType(Col_FieldsTypePos, 0) = cField
                        Col_FieldsType(Col_FieldsTypePos, 1) = "Special Characters (" + m_SpecialChars + ")"
                        Col_FieldsTypePos += 1
                    End If
                Next Num
            Next cField

        End Sub

        Private Sub SpecialCharsFields_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
            On Error Resume Next
            Dim KeyAscii As Short = Asc(e.KeyChar)
            If (m_SpecialChars <> "" And (InStr(1, m_SpecialChars, Chr(KeyAscii)) > 0)) Or (KeyAscii = 13) Or (KeyAscii = 27) Then
            Else
                'sender.text = Mid(sender.text, 1, Len(sender.text) - 1)
                System.Windows.Forms.SendKeys.Send("{BS}")
            End If
        End Sub

        Public Sub AlphaNumericFields(ByVal ParamArray str_NumericFields() As String)
            Dim Num As Integer
            Dim cField As String

            For Each cField In str_NumericFields
                For Num = 1 To Col_ControlName.Count()
                    If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                        Dim txt_control As TextBox = CType(MyForm.Controls(Col_ControlIndex(Num)), TextBox)
                        AddHandler txt_control.KeyPress, AddressOf AlphaNumericFields_KeyPress
                        Col_FieldsType(Col_FieldsTypePos, 0) = cField
                        Col_FieldsType(Col_FieldsTypePos, 1) = "Alphabetic & Numeric"
                        Col_FieldsTypePos += 1
                    End If
                Next Num
            Next cField

        End Sub
        Private Sub AlphaNumericFields_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
            Dim KeyAscii As Short = Asc(e.KeyChar)
            If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 32) Or (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8) Or (KeyAscii = 13) Or (KeyAscii = 27) Then
            Else
                System.Windows.Forms.SendKeys.Send("{BS}")
            End If
        End Sub

        Public Sub AlphabeticFields(ByVal ParamArray str_NumericFields() As String)
            Dim Num As Integer
            Dim cField As String

            For Each cField In str_NumericFields
                For Num = 1 To Col_ControlName.Count()
                    If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                        Dim txt_control As TextBox = CType(MyForm.Controls(Col_ControlIndex(Num)), TextBox)
                        AddHandler txt_control.KeyPress, AddressOf AlphabeticFields_KeyPress
                        Col_FieldsType(Col_FieldsTypePos, 0) = cField
                        Col_FieldsType(Col_FieldsTypePos, 1) = "Alphabetic"
                        Col_FieldsTypePos += 1
                    End If
                Next Num
            Next cField

        End Sub
        Private Sub AlphabeticFields_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
            Dim KeyAscii As Short = Asc(e.KeyChar)
            If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 32) Or (KeyAscii = 8) Or (KeyAscii = 13) Or (KeyAscii = 27) Then
            Else
                System.Windows.Forms.SendKeys.Send("{BS}")
            End If
        End Sub

        Public Sub DecimalFields(ByVal ParamArray str_NumericFields() As String)
            Dim Num As Integer
            Dim cField As String

            For Each cField In str_NumericFields
                For Num = 1 To Col_ControlName.Count()
                    If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                        Dim txt_control As TextBox = CType(MyForm.Controls(Col_ControlIndex(Num)), TextBox)
                        AddHandler txt_control.Leave, AddressOf DecimalFields_Leave
                        AddHandler txt_control.KeyPress, AddressOf DecimalFields_KeyPress
                        Col_FieldsType(Col_FieldsTypePos, 0) = cField
                        Col_FieldsType(Col_FieldsTypePos, 1) = "Integer & Decimal" + IIf(m_DecimalPlaces <> 0, " with " + m_DecimalPlaces.ToString + " Decimal Digits", "")
                        Col_FieldsTypePos += 1
                    End If
                Next Num
            Next cField

        End Sub
        Private Sub DecimalFields_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
            Dim KeyAscii As Short = Asc(e.KeyChar)
            If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 45) Or (KeyAscii = 46) Or (KeyAscii = 8) Or (KeyAscii = 13) Or (KeyAscii = 27) Then
            Else
                System.Windows.Forms.SendKeys.Send("{BS}")
            End If
        End Sub

        Private Sub DecimalFields_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            Static Flag As Byte = 0
            Dim TextBoxColor As System.Drawing.Color
            On Error Resume Next
            Flag += 1
            If Flag = 1 Then
                TextBoxColor = sender.ForeColor
            End If
            Dim j As TextBox

            If sender.Text <> "" And Not IsNumeric(sender.Text) Then
                sender.ForeColor = System.Drawing.Color.Red
            Else
                sender.ForeColor = TextBoxColor
                Dim Num As Decimal = CDec(sender.Text)
                If m_DecimalPlaces <> 0 Then
                    Num = Format(Num, "############." + New String("0", m_DecimalPlaces))
                End If
                sender.Text = Num
            End If
        End Sub

        Public Sub DecimalPlaces(ByVal n_DecimalPlaces As Byte)
            m_DecimalPlaces = n_DecimalPlaces
        End Sub

        Public Sub NumericFields(ByVal ParamArray str_NumericFields() As String)
            Dim Num As Integer
            Dim cField As String

            For Each cField In str_NumericFields
                For Num = 1 To Col_ControlName.Count()
                    If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                        Dim txt_control As TextBox = CType(MyForm.Controls(Col_ControlIndex(Num)), TextBox)
                        AddHandler txt_control.Leave, AddressOf NumericFields_Leave
                        AddHandler txt_control.KeyPress, AddressOf NumericFields_KeyPress
                        Col_FieldsType(Col_FieldsTypePos, 0) = cField
                        Col_FieldsType(Col_FieldsTypePos, 1) = "Numeric"
                        Col_FieldsTypePos += 1
                    End If
                Next Num
            Next cField

        End Sub
        Private Sub NumericFields_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
            Dim KeyAscii As Short = Asc(e.KeyChar)
            If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8) Or (KeyAscii = 13) Or (KeyAscii = 27) Then
            Else
                System.Windows.Forms.SendKeys.Send("{BS}")
            End If
        End Sub

        Private Sub NumericFields_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            Static Flag As Byte = 0
            Dim TextBoxColor As System.Drawing.Color
            On Error Resume Next
            Flag += 1
            If Flag = 1 Then
                TextBoxColor = sender.ForeColor
            End If

            If sender.Text <> "" And Not IsNumeric(sender.Text) Then
                sender.ForeColor = System.Drawing.Color.Red
            Else
                sender.ForeColor = TextBoxColor
            End If
        End Sub

        Public Sub DateFields(ByVal ParamArray str_DateFields() As String)
            Dim Num As Integer
            Dim cField As String

            For Each cField In str_DateFields
                For Num = 1 To Col_ControlName.Count()
                    If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                        Dim txt_control As TextBox = CType(MyForm.Controls(Col_ControlIndex(Num)), TextBox)
                        AddHandler txt_control.Leave, AddressOf DateFields_Leave
                        AddHandler txt_control.KeyPress, AddressOf DateFields_KeyPress
                        Col_FieldsType(Col_FieldsTypePos, 0) = cField
                        Col_FieldsType(Col_FieldsTypePos, 1) = "Date"
                        Col_FieldsTypePos += 1
                    End If
                Next Num
            Next cField

        End Sub

        Private Sub DateFields_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs)
            Dim KeyAscii As Short = Asc(e.KeyChar)
            If (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 45) Or (KeyAscii = 47) Or (KeyAscii = 92) Or (KeyAscii = 8) Or (KeyAscii = 13) Then
            Else
                System.Windows.Forms.SendKeys.Send("{BS}")
            End If
        End Sub

        Private Sub DateFields_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            Static Flag As Byte = 0
            Dim TextBoxColor As System.Drawing.Color
            On Error Resume Next
            Flag += 1
            If Flag = 1 Then
                TextBoxColor = sender.ForeColor
            End If

            If sender.Text <> "" And Not IsDate(sender.Text) Then
                sender.ForeColor = System.Drawing.Color.Red
            Else
                sender.ForeColor = TextBoxColor
            End If
        End Sub


        Public Sub UpperCaseFields(ByVal ParamArray str_UpperCaseFields() As String)
            Dim Num As Integer
            Dim cField As String

            For Each cField In str_UpperCaseFields
                For Num = 1 To Col_ControlName.Count()
                    If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                        Dim txt_control As TextBox = CType(MyForm.Controls(Col_ControlIndex(Num)), TextBox)
                        AddHandler txt_control.Leave, AddressOf UpperCaseFields_Leave
                    End If
                Next Num
            Next cField

        End Sub

        Private Sub UpperCaseFields_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            sender.text = UCase(sender.text)
        End Sub


        Public Sub LowerCaseFields(ByVal ParamArray str_LowerCaseFields() As String)
            Dim Num As Integer
            Dim cField As String

            For Each cField In str_LowerCaseFields
                For Num = 1 To Col_ControlName.Count()
                    If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                        Dim txt_control As TextBox = CType(MyForm.Controls(Col_ControlIndex(Num)), TextBox)
                        AddHandler txt_control.Leave, AddressOf LowerCaseFields_Leave
                    End If
                Next Num
            Next cField

        End Sub

        Private Sub LowerCaseFields_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            sender.text = LCase(sender.text)
        End Sub

        Public Sub FirstCharOnlyFields(ByVal ParamArray str_FirstCharOnlyFields() As String)
            Dim Num As Integer
            Dim cField As String

            For Each cField In str_FirstCharOnlyFields
                For Num = 1 To Col_ControlName.Count()
                    If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                        Dim txt_control As TextBox = CType(MyForm.Controls(Col_ControlIndex(Num)), TextBox)
                        AddHandler txt_control.Leave, AddressOf FirstCharOnlyFields_Leave
                    End If
                Next Num
            Next cField

        End Sub

        Private Sub FirstCharOnlyFields_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            sender.Text = UCase(Left(sender.Text, 1)) + Right(sender.Text, Len(sender.Text) - 1)
        End Sub


        Public Sub FirstCharOfWordsFields(ByVal ParamArray str_FirstCharOfWordsFields() As String)
            Dim Num As Integer
            Dim cField As String

            For Each cField In str_FirstCharOfWordsFields
                For Num = 1 To Col_ControlName.Count()
                    If UCase(Col_ControlName(Num).Name) = UCase(cField) Then
                        Dim txt_control As TextBox = CType(MyForm.Controls(Col_ControlIndex(Num)), TextBox)
                        AddHandler txt_control.Leave, AddressOf FirstCharOfWordsFields_Leave
                    End If
                Next Num
            Next cField

        End Sub

        Private Sub FirstCharOfWordsFields_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim MyChar As String
            Dim PrevChar As String = ""
            Dim newValue As String = ""
            Dim X As Integer

            sender.Text = UCase(Left(sender.Text, 1)) + Right(sender.Text, Len(sender.Text) - 1)

            For X = 1 To Len(sender.Text)
                MyChar = Mid(sender.Text, X, 1)
                If PrevChar = " " Then
                    newValue = newValue + UCase(MyChar)
                Else
                    newValue = newValue + MyChar
                End If
                PrevChar = MyChar
            Next X
            sender.Text = newValue
        End Sub

    End Class
End Namespace



