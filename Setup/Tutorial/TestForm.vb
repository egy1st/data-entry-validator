Public Class TestForm
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Public WithEvents ExitButton As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents FirstButton As System.Windows.Forms.Button
    Friend WithEvents PreviousButton As System.Windows.Forms.Button
    Friend WithEvents NextButton As System.Windows.Forms.Button
    Friend WithEvents LastButton As System.Windows.Forms.Button
    Public WithEvents CustomerName As System.Windows.Forms.TextBox
    Public WithEvents CustomerID As System.Windows.Forms.TextBox
    Public WithEvents Address As System.Windows.Forms.TextBox
    Public WithEvents MaxDebit As System.Windows.Forms.TextBox
    Public WithEvents LastDeal As System.Windows.Forms.TextBox
    Public WithEvents Phone As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(TestForm))
        Me.CustomerName = New System.Windows.Forms.TextBox
        Me.ExitButton = New System.Windows.Forms.Button
        Me.CustomerID = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Address = New System.Windows.Forms.TextBox
        Me.Phone = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.MaxDebit = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.LastDeal = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label10 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label13 = New System.Windows.Forms.Label
        Me.FirstButton = New System.Windows.Forms.Button
        Me.PreviousButton = New System.Windows.Forms.Button
        Me.NextButton = New System.Windows.Forms.Button
        Me.LastButton = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'CustomerName
        '
        Me.CustomerName.AcceptsReturn = True
        Me.CustomerName.BackColor = System.Drawing.SystemColors.Window
        Me.CustomerName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.CustomerName.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomerName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.CustomerName.Location = New System.Drawing.Point(103, 42)
        Me.CustomerName.MaxLength = 50
        Me.CustomerName.Name = "CustomerName"
        Me.CustomerName.Size = New System.Drawing.Size(154, 23)
        Me.CustomerName.TabIndex = 1
        '
        'ExitButton
        '
        Me.ExitButton.BackColor = System.Drawing.SystemColors.Control
        Me.ExitButton.Cursor = System.Windows.Forms.Cursors.Default
        Me.ExitButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.ExitButton.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ExitButton.Image = CType(resources.GetObject("ExitButton.Image"), System.Drawing.Image)
        Me.ExitButton.Location = New System.Drawing.Point(355, 249)
        Me.ExitButton.Name = "ExitButton"
        Me.ExitButton.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.ExitButton.Size = New System.Drawing.Size(72, 48)
        Me.ExitButton.TabIndex = 14
        Me.ExitButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ExitButton.UseVisualStyleBackColor = False
        '
        'CustomerID
        '
        Me.CustomerID.AcceptsReturn = True
        Me.CustomerID.BackColor = System.Drawing.SystemColors.Window
        Me.CustomerID.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.CustomerID.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomerID.ForeColor = System.Drawing.SystemColors.WindowText
        Me.CustomerID.Location = New System.Drawing.Point(104, 9)
        Me.CustomerID.MaxLength = 4
        Me.CustomerID.Name = "CustomerID"
        Me.CustomerID.Size = New System.Drawing.Size(73, 23)
        Me.CustomerID.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(90, 24)
        Me.Label1.TabIndex = 60
        Me.Label1.Text = "Cust ID"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label2.Location = New System.Drawing.Point(8, 41)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(90, 24)
        Me.Label2.TabIndex = 61
        Me.Label2.Text = "Cust Name"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label3.Location = New System.Drawing.Point(8, 73)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 24)
        Me.Label3.TabIndex = 62
        Me.Label3.Text = "Address"
        '
        'Address
        '
        Me.Address.AcceptsReturn = True
        Me.Address.BackColor = System.Drawing.SystemColors.Window
        Me.Address.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Address.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Address.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Address.Location = New System.Drawing.Point(103, 75)
        Me.Address.MaxLength = 50
        Me.Address.Name = "Address"
        Me.Address.Size = New System.Drawing.Size(154, 23)
        Me.Address.TabIndex = 2
        '
        'Phone
        '
        Me.Phone.AcceptsReturn = True
        Me.Phone.BackColor = System.Drawing.SystemColors.Window
        Me.Phone.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Phone.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Phone.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Phone.Location = New System.Drawing.Point(102, 112)
        Me.Phone.MaxLength = 12
        Me.Phone.Name = "Phone"
        Me.Phone.Size = New System.Drawing.Size(124, 23)
        Me.Phone.TabIndex = 3
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label4.Location = New System.Drawing.Point(6, 111)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(90, 24)
        Me.Label4.TabIndex = 64
        Me.Label4.Text = "Phone"
        '
        'MaxDebit
        '
        Me.MaxDebit.AcceptsReturn = True
        Me.MaxDebit.BackColor = System.Drawing.SystemColors.Window
        Me.MaxDebit.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.MaxDebit.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MaxDebit.ForeColor = System.Drawing.SystemColors.WindowText
        Me.MaxDebit.Location = New System.Drawing.Point(103, 146)
        Me.MaxDebit.MaxLength = 12
        Me.MaxDebit.Name = "MaxDebit"
        Me.MaxDebit.Size = New System.Drawing.Size(123, 23)
        Me.MaxDebit.TabIndex = 4
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label5.Location = New System.Drawing.Point(8, 146)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(88, 24)
        Me.Label5.TabIndex = 66
        Me.Label5.Text = "Max Debit"
        '
        'LastDeal
        '
        Me.LastDeal.AcceptsReturn = True
        Me.LastDeal.BackColor = System.Drawing.SystemColors.Window
        Me.LastDeal.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.LastDeal.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LastDeal.ForeColor = System.Drawing.SystemColors.WindowText
        Me.LastDeal.Location = New System.Drawing.Point(103, 177)
        Me.LastDeal.MaxLength = 10
        Me.LastDeal.Name = "LastDeal"
        Me.LastDeal.Size = New System.Drawing.Size(123, 23)
        Me.LastDeal.TabIndex = 5
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ActiveCaption
        Me.Label6.Location = New System.Drawing.Point(8, 180)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(88, 24)
        Me.Label6.TabIndex = 68
        Me.Label6.Text = "Last Deal"
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label8.ForeColor = System.Drawing.Color.ForestGreen
        Me.Label8.Location = New System.Drawing.Point(241, 7)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(163, 24)
        Me.Label8.TabIndex = 72
        Me.Label8.Text = "Numeric characters"
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label9.ForeColor = System.Drawing.Color.ForestGreen
        Me.Label9.Location = New System.Drawing.Point(271, 40)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(168, 24)
        Me.Label9.TabIndex = 73
        Me.Label9.Text = "Alphabetic characters"
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label10.ForeColor = System.Drawing.Color.ForestGreen
        Me.Label10.Location = New System.Drawing.Point(271, 69)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(155, 44)
        Me.Label10.TabIndex = 74
        Me.Label10.Text = "Alphabetic or numeic characters"
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label11.ForeColor = System.Drawing.Color.ForestGreen
        Me.Label11.Location = New System.Drawing.Point(241, 147)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(187, 24)
        Me.Label11.TabIndex = 75
        Me.Label11.Text = "2 decimal digit number"
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label12.ForeColor = System.Drawing.Color.ForestGreen
        Me.Label12.Location = New System.Drawing.Point(241, 113)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(198, 24)
        Me.Label12.TabIndex = 76
        Me.Label12.Text = "Numeic characters and  ( - )"
        '
        'Label13
        '
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(178, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.ForestGreen
        Me.Label13.Location = New System.Drawing.Point(241, 177)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(169, 24)
        Me.Label13.TabIndex = 77
        Me.Label13.Text = "Valid date"
        '
        'FirstButton
        '
        Me.FirstButton.Location = New System.Drawing.Point(24, 249)
        Me.FirstButton.Name = "FirstButton"
        Me.FirstButton.Size = New System.Drawing.Size(72, 48)
        Me.FirstButton.TabIndex = 7
        Me.FirstButton.Text = "First"
        '
        'PreviousButton
        '
        Me.PreviousButton.Location = New System.Drawing.Point(102, 249)
        Me.PreviousButton.Name = "PreviousButton"
        Me.PreviousButton.Size = New System.Drawing.Size(72, 48)
        Me.PreviousButton.TabIndex = 8
        Me.PreviousButton.Text = "Previous"
        '
        'NextButton
        '
        Me.NextButton.Location = New System.Drawing.Point(180, 249)
        Me.NextButton.Name = "NextButton"
        Me.NextButton.Size = New System.Drawing.Size(72, 48)
        Me.NextButton.TabIndex = 9
        Me.NextButton.Text = "Next"
        '
        'LastButton
        '
        Me.LastButton.Location = New System.Drawing.Point(258, 249)
        Me.LastButton.Name = "LastButton"
        Me.LastButton.Size = New System.Drawing.Size(72, 48)
        Me.LastButton.TabIndex = 10
        Me.LastButton.Text = "last"
        '
        'TestForm
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(441, 304)
        Me.Controls.Add(Me.LastButton)
        Me.Controls.Add(Me.NextButton)
        Me.Controls.Add(Me.PreviousButton)
        Me.Controls.Add(Me.FirstButton)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.LastDeal)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.MaxDebit)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Phone)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Address)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CustomerName)
        Me.Controls.Add(Me.ExitButton)
        Me.Controls.Add(Me.CustomerID)
        Me.ForeColor = System.Drawing.Color.Blue
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "TestForm"
        Me.Tag = "Orders Form"
        Me.Text = "Data Entry Validator Tutorial"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Dim DV As New DynamicComponents.DataEntryValidator()
    Dim CN As New ADODB.Connection()
    Dim oCust As New ADODB.Recordset()
    Dim oAccess As New Access.Application()
    Dim DAO_DBEngine As New DAO.DBEngine()

    Private Sub TestForm_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        'establish DSN
        oAccess.DBEngine.RegisterDatabase("DCDM_Nwind", "Microsoft Access Driver (*.mdb)", True, "DBQ=" & VB6.GetPath & "\Nwind.mdb")
        CN.Open("DSN=DCDM_NWind")
        oCust.Open("Customers", CN, ADODB.CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
        PopulateData()
        DV.InitForm(Me)                              'must be your first assignment , an error occurs if not
        DV.NumericFields("CustomerId")               'CustomerId must be numeric characters(0123456789)
        DV.AlphabeticFields("CustomerName")          'CustomerName must be alphabetic characters (abcdefghijklmnopqrstuvwzyzABCDEFGHIJKLMNOPQRSTUVWXYZ) 
        DV.FirstCharOfWordsFields("CustomerName")    'First charecter of every word will be in upper case
        DV.AlphaNumericFields("Address")             'Address must be numeric or alphabetic characters (0123456789abcdefghijklmnopqrstuvwzyzABCDEFGHIJKLMNOPQRSTUVWXYZ)  
        DV.FirstCharOnlyFields("Address")            'First charecter of first word only will be in uooer case 
        DV.AlphaNumericFields("phone")               'Phone must be numeric characters(0123456789)
        DV.DecimalFields("MaxDebit")                 'MaxDebit must be decimal characters(0123456789 & .) 
        DV.DecimalPlaces(2)                          'MaxDebit will be formatted with 2 decimal digits
        DV.DateFields("LastDeal")                    'LastDeal must be accepted date(0123456789-\/)
    End Sub


    Private Sub ExitButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitButton.Click
        oAccess.Quit()
        Me.Close()

    End Sub

    Private Sub PopulateData()
        Me.CustomerID.Text = oCust.Fields("CustomerID").Value
        Me.CustomerName.Text = oCust.Fields("CustomerName").Value
        Me.Address.Text = oCust.Fields("Address").Value
        Me.Phone.Text = oCust.Fields("Phone").Value
        Me.MaxDebit.Text = oCust.Fields("MaxDebit").Value
        Me.LastDeal.Text = oCust.Fields("LastDeal").Value

    End Sub

    Private Sub FirstButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FirstButton.Click
        oCust.MoveFirst()
        PopulateData()
    End Sub

    Private Sub PreviousButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PreviousButton.Click
        On Error Resume Next

        oCust.MovePrevious()
        PopulateData()
    End Sub

    Private Sub NextButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NextButton.Click
        On Error Resume Next

        oCust.MoveNext()
        PopulateData()
    End Sub

    Private Sub LastButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LastButton.Click
        oCust.MoveLast()
        PopulateData()
    End Sub
End Class
