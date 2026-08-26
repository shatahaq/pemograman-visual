<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        lblNama = New Label()
        txtNama = New TextBox()
        lblNIM = New Label()
        txtNIM = New TextBox()
        txtKOM = New TextBox()
        lblKOM = New Label()
        btnTampilkan = New Button()
        txtHapus = New Button()
        btnKeluar = New Button()
        SuspendLayout()
        ' 
        ' lblNama
        ' 
        lblNama.AutoSize = True
        lblNama.Location = New Point(149, 15)
        lblNama.Name = "lblNama"
        lblNama.Size = New Size(60, 20)
        lblNama.TabIndex = 0
        lblNama.Text = "Nama  :"
        ' 
        ' txtNama
        ' 
        txtNama.Location = New Point(221, 12)
        txtNama.Name = "txtNama"
        txtNama.Size = New Size(134, 27)
        txtNama.TabIndex = 1
        ' 
        ' lblNIM
        ' 
        lblNIM.AutoSize = True
        lblNIM.Location = New Point(149, 57)
        lblNIM.Name = "lblNIM"
        lblNIM.Size = New Size(56, 20)
        lblNIM.TabIndex = 2
        lblNIM.Text = "NIM    :"
        ' 
        ' txtNIM
        ' 
        txtNIM.Location = New Point(221, 57)
        txtNIM.Name = "txtNIM"
        txtNIM.Size = New Size(134, 27)
        txtNIM.TabIndex = 3
        ' 
        ' txtKOM
        ' 
        txtKOM.Location = New Point(221, 105)
        txtKOM.Name = "txtKOM"
        txtKOM.Size = New Size(134, 27)
        txtKOM.TabIndex = 4
        ' 
        ' lblKOM
        ' 
        lblKOM.AutoSize = True
        lblKOM.Location = New Point(149, 105)
        lblKOM.Name = "lblKOM"
        lblKOM.Size = New Size(56, 20)
        lblKOM.TabIndex = 5
        lblKOM.Text = "KOM   :"
        ' 
        ' btnTampilkan
        ' 
        btnTampilkan.Location = New Point(149, 162)
        btnTampilkan.Name = "btnTampilkan"
        btnTampilkan.Size = New Size(94, 29)
        btnTampilkan.TabIndex = 6
        btnTampilkan.Text = "Tampilkan"
        btnTampilkan.UseVisualStyleBackColor = True
        ' 
        ' txtHapus
        ' 
        txtHapus.Location = New Point(275, 162)
        txtHapus.Name = "txtHapus"
        txtHapus.Size = New Size(94, 29)
        txtHapus.TabIndex = 7
        txtHapus.Text = "Hapus"
        txtHapus.UseVisualStyleBackColor = True
        ' 
        ' btnKeluar
        ' 
        btnKeluar.Location = New Point(198, 218)
        btnKeluar.Name = "btnKeluar"
        btnKeluar.Size = New Size(94, 29)
        btnKeluar.TabIndex = 8
        btnKeluar.Text = "Keluar"
        btnKeluar.UseVisualStyleBackColor = True
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(8F, 20F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(549, 328)
        Controls.Add(btnKeluar)
        Controls.Add(txtHapus)
        Controls.Add(btnTampilkan)
        Controls.Add(lblKOM)
        Controls.Add(txtKOM)
        Controls.Add(txtNIM)
        Controls.Add(lblNIM)
        Controls.Add(txtNama)
        Controls.Add(lblNama)
        Name = "Form1"
        Text = "Profile Mahasiswa"
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents lblNama As Label
    Friend WithEvents txtNama As TextBox
    Friend WithEvents lblNIM As Label
    Friend WithEvents txtNIM As TextBox
    Friend WithEvents txtKOM As TextBox
    Friend WithEvents lblKOM As Label
    Friend WithEvents btnTampilkan As Button
    Friend WithEvents txtHapus As Button
    Friend WithEvents btnKeluar As Button

End Class
