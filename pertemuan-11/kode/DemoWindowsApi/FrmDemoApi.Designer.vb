<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmDemoApi
    Inherits System.Windows.Forms.Form

    Private components As System.ComponentModel.IContainer

    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then components.Dispose()
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        components = New ComponentModel.Container()
        pnlTimer = New Panel()
        lblTahap = New Label()
        lblSisa = New Label()
        lblKeterangan = New Label()
        btnMulai = New Button()
        btnUlangi = New Button()
        chkTeratas = New CheckBox()
        btnRingkas = New Button()
        pnlPengaturan = New Panel()
        lblDurasi = New Label()
        lblMenit = New Label()
        nudMenit = New NumericUpDown()
        lblDetik = New Label()
        nudDetik = New NumericUpDown()
        chkBunyi = New CheckBox()
        lblShortcut = New Label()
        lblStatus = New Label()
        tmrHitung = New Timer(components)
        pnlTimer.SuspendLayout()
        pnlPengaturan.SuspendLayout()
        CType(nudMenit, ComponentModel.ISupportInitialize).BeginInit()
        CType(nudDetik, ComponentModel.ISupportInitialize).BeginInit()
        SuspendLayout()
        ' 
        ' pnlTimer
        ' 
        pnlTimer.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        pnlTimer.BackColor = Color.FromArgb(CByte(23), CByte(50), CByte(77))
        pnlTimer.Controls.Add(lblTahap)
        pnlTimer.Controls.Add(lblSisa)
        pnlTimer.Controls.Add(lblKeterangan)
        pnlTimer.ForeColor = Color.White
        pnlTimer.Location = New Point(24, 13)
        pnlTimer.Name = "pnlTimer"
        pnlTimer.Size = New Size(512, 202)
        pnlTimer.TabIndex = 2
        ' 
        ' lblTahap
        ' 
        lblTahap.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        lblTahap.Location = New Point(16, 17)
        lblTahap.Name = "lblTahap"
        lblTahap.Size = New Size(480, 26)
        lblTahap.TabIndex = 0
        lblTahap.Text = "Timer Presentasi"
        lblTahap.TextAlign = ContentAlignment.MiddleCenter
        ' 
        ' lblSisa
        ' 
        lblSisa.AccessibleName = "Sisa waktu presentasi"
        lblSisa.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        lblSisa.Font = New Font("Segoe UI", 58F, FontStyle.Bold)
        lblSisa.Location = New Point(12, 44)
        lblSisa.Name = "lblSisa"
        lblSisa.Size = New Size(488, 113)
        lblSisa.TabIndex = 1
        lblSisa.Text = "05:00"
        lblSisa.TextAlign = ContentAlignment.MiddleCenter
        ' 
        ' lblKeterangan
        ' 
        lblKeterangan.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        lblKeterangan.Location = New Point(16, 164)
        lblKeterangan.Name = "lblKeterangan"
        lblKeterangan.Size = New Size(480, 24)
        lblKeterangan.TabIndex = 2
        lblKeterangan.Text = "Atur durasi, lalu mulai ketika Anda siap."
        lblKeterangan.TextAlign = ContentAlignment.MiddleCenter
        ' 
        ' btnMulai
        ' 
        btnMulai.BackColor = Color.FromArgb(CByte(0), CByte(82), CByte(164))
        btnMulai.Cursor = Cursors.Hand
        btnMulai.FlatAppearance.BorderSize = 0
        btnMulai.FlatAppearance.MouseDownBackColor = Color.FromArgb(CByte(0), CByte(42), CByte(95))
        btnMulai.FlatAppearance.MouseOverBackColor = Color.FromArgb(CByte(0), CByte(62), CByte(130))
        btnMulai.FlatStyle = FlatStyle.Flat
        btnMulai.Font = New Font("Segoe UI", 11F, FontStyle.Bold)
        btnMulai.ForeColor = Color.White
        btnMulai.Location = New Point(24, 233)
        btnMulai.Name = "btnMulai"
        btnMulai.Size = New Size(324, 44)
        btnMulai.TabIndex = 0
        btnMulai.Text = "&Mulai presentasi"
        btnMulai.UseVisualStyleBackColor = False
        ' 
        ' btnUlangi
        ' 
        btnUlangi.Cursor = Cursors.Hand
        btnUlangi.Location = New Point(364, 233)
        btnUlangi.Name = "btnUlangi"
        btnUlangi.Size = New Size(172, 44)
        btnUlangi.TabIndex = 1
        btnUlangi.Text = "&Ulangi"
        ' 
        ' chkTeratas
        ' 
        chkTeratas.Cursor = Cursors.Hand
        chkTeratas.Location = New Point(24, 295)
        chkTeratas.Name = "chkTeratas"
        chkTeratas.Size = New Size(324, 30)
        chkTeratas.TabIndex = 2
        chkTeratas.Text = "&Tetap di atas aplikasi lain"
        ' 
        ' btnRingkas
        ' 
        btnRingkas.Cursor = Cursors.Hand
        btnRingkas.Location = New Point(364, 293)
        btnRingkas.Name = "btnRingkas"
        btnRingkas.Size = New Size(172, 34)
        btnRingkas.TabIndex = 3
        btnRingkas.Text = "Mode &ringkas"
        ' 
        ' pnlPengaturan
        ' 
        pnlPengaturan.Controls.Add(lblDurasi)
        pnlPengaturan.Controls.Add(lblMenit)
        pnlPengaturan.Controls.Add(nudMenit)
        pnlPengaturan.Controls.Add(lblDetik)
        pnlPengaturan.Controls.Add(nudDetik)
        pnlPengaturan.Controls.Add(chkBunyi)
        pnlPengaturan.Controls.Add(lblShortcut)
        pnlPengaturan.Location = New Point(24, 340)
        pnlPengaturan.Name = "pnlPengaturan"
        pnlPengaturan.Size = New Size(512, 126)
        pnlPengaturan.TabIndex = 4
        ' 
        ' lblDurasi
        ' 
        lblDurasi.Font = New Font("Segoe UI", 11F, FontStyle.Bold)
        lblDurasi.Location = New Point(0, 0)
        lblDurasi.Name = "lblDurasi"
        lblDurasi.Size = New Size(512, 25)
        lblDurasi.TabIndex = 0
        lblDurasi.Text = "Durasi presentasi"
        ' 
        ' lblMenit
        ' 
        lblMenit.Location = New Point(0, 35)
        lblMenit.Name = "lblMenit"
        lblMenit.Size = New Size(50, 26)
        lblMenit.TabIndex = 0
        lblMenit.Text = "M&enit"
        ' 
        ' nudMenit
        ' 
        nudMenit.AccessibleName = "Durasi menit"
        nudMenit.Location = New Point(52, 33)
        nudMenit.Maximum = New Decimal(New Integer() {180, 0, 0, 0})
        nudMenit.Name = "nudMenit"
        nudMenit.Size = New Size(74, 25)
        nudMenit.TabIndex = 1
        nudMenit.Value = New Decimal(New Integer() {5, 0, 0, 0})
        ' 
        ' lblDetik
        ' 
        lblDetik.Location = New Point(146, 35)
        lblDetik.Name = "lblDetik"
        lblDetik.Size = New Size(50, 26)
        lblDetik.TabIndex = 2
        lblDetik.Text = "&Detik"
        ' 
        ' nudDetik
        ' 
        nudDetik.AccessibleName = "Durasi detik"
        nudDetik.Location = New Point(198, 33)
        nudDetik.Maximum = New Decimal(New Integer() {59, 0, 0, 0})
        nudDetik.Name = "nudDetik"
        nudDetik.Size = New Size(74, 25)
        nudDetik.TabIndex = 3
        ' 
        ' chkBunyi
        ' 
        chkBunyi.Checked = True
        chkBunyi.CheckState = CheckState.Checked
        chkBunyi.Cursor = Cursors.Hand
        chkBunyi.Location = New Point(310, 33)
        chkBunyi.Name = "chkBunyi"
        chkBunyi.Size = New Size(202, 30)
        chkBunyi.TabIndex = 4
        chkBunyi.Text = "&Bunyi saat selesai"
        ' 
        ' lblShortcut
        ' 
        lblShortcut.Font = New Font("Segoe UI", 9F)
        lblShortcut.Location = New Point(0, 77)
        lblShortcut.Name = "lblShortcut"
        lblShortcut.Size = New Size(512, 48)
        lblShortcut.TabIndex = 5
        lblShortcut.Text = "Shortcut disiapkan saat aplikasi dibuka."
        ' 
        ' lblStatus
        ' 
        lblStatus.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        lblStatus.Font = New Font("Segoe UI", 9F)
        lblStatus.Location = New Point(24, 469)
        lblStatus.Name = "lblStatus"
        lblStatus.Size = New Size(512, 42)
        lblStatus.TabIndex = 5
        lblStatus.Text = "Timer siap."
        ' 
        ' tmrHitung
        ' 
        ' 
        ' FrmDemoApi
        ' 
        AutoScaleDimensions = New SizeF(7F, 17F)
        AutoScaleMode = AutoScaleMode.Font
        BackColor = Color.FromArgb(CByte(246), CByte(248), CByte(250))
        ClientSize = New Size(560, 642)
        Controls.Add(pnlTimer)
        Controls.Add(btnMulai)
        Controls.Add(btnUlangi)
        Controls.Add(chkTeratas)
        Controls.Add(btnRingkas)
        Controls.Add(pnlPengaturan)
        Controls.Add(lblStatus)
        Font = New Font("Segoe UI", 10F)
        ForeColor = Color.FromArgb(CByte(23), CByte(50), CByte(77))
        MaximizeBox = False
        Name = "FrmDemoApi"
        StartPosition = FormStartPosition.CenterScreen
        Text = "Asisten Presentasi | Pertemuan 11"
        pnlTimer.ResumeLayout(False)
        pnlPengaturan.ResumeLayout(False)
        CType(nudMenit, ComponentModel.ISupportInitialize).EndInit()
        CType(nudDetik, ComponentModel.ISupportInitialize).EndInit()
        ResumeLayout(False)
    End Sub
    Friend WithEvents pnlTimer As System.Windows.Forms.Panel
    Friend WithEvents lblTahap As System.Windows.Forms.Label
    Friend WithEvents lblSisa As System.Windows.Forms.Label
    Friend WithEvents lblKeterangan As System.Windows.Forms.Label
    Friend WithEvents btnMulai As System.Windows.Forms.Button
    Friend WithEvents btnUlangi As System.Windows.Forms.Button
    Friend WithEvents chkTeratas As System.Windows.Forms.CheckBox
    Friend WithEvents btnRingkas As System.Windows.Forms.Button
    Friend WithEvents pnlPengaturan As System.Windows.Forms.Panel
    Friend WithEvents lblDurasi As System.Windows.Forms.Label
    Friend WithEvents lblMenit As System.Windows.Forms.Label
    Friend WithEvents nudMenit As System.Windows.Forms.NumericUpDown
    Friend WithEvents lblDetik As System.Windows.Forms.Label
    Friend WithEvents nudDetik As System.Windows.Forms.NumericUpDown
    Friend WithEvents chkBunyi As System.Windows.Forms.CheckBox
    Friend WithEvents lblShortcut As System.Windows.Forms.Label
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents tmrHitung As System.Windows.Forms.Timer
End Class
