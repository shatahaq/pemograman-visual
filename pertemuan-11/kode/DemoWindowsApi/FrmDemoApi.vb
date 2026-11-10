Option Strict On
Imports System.Runtime.InteropServices

Partial Class FrmDemoApi
    Private durasi As ULong
    Private terpakai As ULong
    Private mulaiTick As ULong
    Private berjalan As Boolean
    Private selesai As Boolean
    Private ringkas As Boolean
    Private mengubahPin As Boolean
    Private handleShortcut As IntPtr
    Private shortcutMulai As Boolean
    Private shortcutUlangi As Boolean
    Private ukuranLengkap As Size

    Private Sub FrmDemoApi_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ukuranLengkap = Me.ClientSize
        UlangiTimer()
    End Sub

    Private Sub FrmDemoApi_HandleCreated(sender As Object, e As EventArgs) Handles MyBase.HandleCreated
        handleShortcut = Me.Handle
        Dim modifier As UInteger = MOD_CONTROL Or MOD_ALT Or MOD_NOREPEAT
        shortcutMulai = RegisterHotKey(handleShortcut, ID_MULAI, modifier, VK_F9)
        shortcutUlangi = RegisterHotKey(handleShortcut, ID_ULANGI, modifier, VK_F10)
        Dim mulai As String = If(shortcutMulai, "Ctrl+Alt+F9: mulai/jeda", "Shortcut mulai tidak tersedia")
        Dim ulang As String = If(shortcutUlangi, "Ctrl+Alt+F10: ulangi", "Shortcut ulangi tidak tersedia")
        lblShortcut.Text = mulai & vbCrLf & ulang & ". Tombol tetap dapat digunakan."
        If chkTeratas.Checked Then chkTeratas_CheckedChanged(Me, EventArgs.Empty)
    End Sub

    Private Sub FrmDemoApi_HandleDestroyed(sender As Object, e As EventArgs) Handles MyBase.HandleDestroyed
        If shortcutMulai Then UnregisterHotKey(handleShortcut, ID_MULAI)
        If shortcutUlangi Then UnregisterHotKey(handleShortcut, ID_ULANGI)
        shortcutMulai = False
        shortcutUlangi = False
        handleShortcut = IntPtr.Zero
    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message)
        If m.Msg = WM_HOTKEY Then
            If m.WParam.ToInt32() = ID_MULAI AndAlso shortcutMulai Then MulaiAtauJeda()
            If m.WParam.ToInt32() = ID_ULANGI AndAlso shortcutUlangi Then UlangiTimer()
        End If
        MyBase.WndProc(m)
    End Sub

    Private Sub btnMulai_Click(sender As Object, e As EventArgs) Handles btnMulai.Click
        MulaiAtauJeda()
    End Sub

    Private Sub btnUlangi_Click(sender As Object, e As EventArgs) Handles btnUlangi.Click
        UlangiTimer()
    End Sub

    Private Sub MulaiAtauJeda()
        If berjalan Then
            PerbaruiHitungan()
            If selesai Then Return
            terpakai += GetTickCount64() - mulaiTick
            berjalan = False
            tmrHitung.Stop()
            btnMulai.Text = "&Lanjutkan"
            lblTahap.Text = "Dijeda"
            lblKeterangan.Text = "Waktu berhenti sampai Anda melanjutkan."
            lblStatus.Text = "Timer dijeda. Durasi dapat diubah setelah memilih Ulangi."
            Return
        End If
        If selesai Then UlangiTimer()
        If durasi = 0UL Then
            lblStatus.Text = "Isi durasi lebih dari 0 detik sebelum mulai."
            nudDetik.Focus()
            Return
        End If
        mulaiTick = GetTickCount64()
        berjalan = True
        nudMenit.Enabled = False
        nudDetik.Enabled = False
        btnMulai.Text = "&Jeda"
        lblTahap.Text = "Presentasi berlangsung"
        lblKeterangan.Text = "Sisa waktu presentasi Anda"
        lblStatus.Text = "Timer berjalan. Anda dapat berpindah ke aplikasi slide."
        tmrHitung.Start()
    End Sub

    Private Sub UlangiTimer()
        tmrHitung.Stop()
        berjalan = False
        selesai = False
        terpakai = 0UL
        durasi = CULng(nudMenit.Value * 60D + nudDetik.Value) * 1000UL
        nudMenit.Enabled = True
        nudDetik.Enabled = True
        btnMulai.Text = "&Mulai presentasi"
        lblTahap.Text = "Timer presentasi"
        lblSisa.Text = FormatSisa(durasi)
        lblKeterangan.Text = "Atur durasi, lalu mulai ketika Anda siap."
        pnlTimer.BackColor = Color.FromArgb(23, 50, 77)
        lblStatus.Text = "Timer siap. Aktifkan Tetap di atas agar terlihat saat berganti aplikasi."
    End Sub

    Private Sub Durasi_ValueChanged(sender As Object, e As EventArgs) Handles nudMenit.ValueChanged, nudDetik.ValueChanged
        If Not berjalan AndAlso nudMenit.Enabled Then UlangiTimer()
    End Sub

    Private Sub tmrHitung_Tick(sender As Object, e As EventArgs) Handles tmrHitung.Tick
        PerbaruiHitungan()
    End Sub

    Private Sub PerbaruiHitungan()
        If Not berjalan Then Return
        Dim elapsed As ULong = terpakai + GetTickCount64() - mulaiTick
        Dim sisa As ULong = HitungSisa(durasi, elapsed)
        lblSisa.Text = FormatSisa(sisa)
        If sisa = 0UL Then WaktuHabis()
    End Sub

    Private Sub WaktuHabis()
        tmrHitung.Stop()
        berjalan = False
        selesai = True
        lblTahap.Text = "Waktu habis"
        lblKeterangan.Text = "Akhiri presentasi dan masuk ke tanya jawab."
        pnlTimer.BackColor = Color.FromArgb(133, 52, 21)
        btnMulai.Text = "&Mulai lagi"
        lblStatus.Text = "Durasi selesai. Pilih Ulangi untuk mengatur sesi berikutnya."
        If chkBunyi.Checked Then
            If Not MessageBeep(MB_ICONINFORMATION) Then
                lblStatus.Text = "Waktu habis. Bunyi tidak dapat diputar; penanda visual tetap aktif."
            End If
        End If
    End Sub

    Private Sub chkTeratas_CheckedChanged(sender As Object, e As EventArgs) Handles chkTeratas.CheckedChanged
        If mengubahPin Then Return
        Dim posisi As IntPtr = If(chkTeratas.Checked, HWND_TOPMOST, HWND_NOTOPMOST)
        Dim flags As UInteger = SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE
        If SetWindowPos(Me.Handle, posisi, 0, 0, 0, 0, flags) Then
            lblStatus.Text = If(chkTeratas.Checked, "Timer tetap di atas jendela biasa. Gunakan Mode ringkas saat presentasi.", "Jendela kembali mengikuti urutan aplikasi biasa.")
        Else
            Dim kode As Integer = Marshal.GetLastPInvokeError()
            mengubahPin = True
            chkTeratas.Checked = Not chkTeratas.Checked
            mengubahPin = False
            lblStatus.Text = $"Posisi jendela tidak dapat diubah. Kode Windows: {kode}."
        End If
    End Sub

    Private Sub btnRingkas_Click(sender As Object, e As EventArgs) Handles btnRingkas.Click
        If Not ringkas Then ukuranLengkap = Me.ClientSize
        ringkas = Not ringkas
        pnlPengaturan.Visible = Not ringkas
        btnRingkas.Text = If(ringkas, "&Pengaturan", "Mode &ringkas")
        Dim bingkai As New Size(Me.Width - Me.ClientSize.Width, Me.Height - Me.ClientSize.Height)
        Dim tinggi As Integer = If(ringkas, pnlPengaturan.Top + lblStatus.Height + 24, ukuranLengkap.Height)
        Me.MinimumSize = New Size(ukuranLengkap.Width + bingkai.Width, tinggi + bingkai.Height)
        Me.ClientSize = New Size(ukuranLengkap.Width, tinggi)
    End Sub
End Class
