Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        FunkcjaDoTegoObwodu()

    End Sub
    Private Function OporPolaczeniaSzeregowego(PierwszyOdbiornik As Double, DrugiOdbiornik As Double) As Double
        Return PierwszyOdbiornik + DrugiOdbiornik
    End Function
    Private Function PradWGalezi(OporGalezi As Double, NapiecieNaGalezi As Double) As Double
        Return NapiecieNaGalezi / OporGalezi
    End Function
    Private Function NapiecieNaGalezi(OporGalezi As Double, PradWGalezi As Double) As Double
        Return OporGalezi * PradWGalezi
    End Function
    Private Function OporPolaczeniaRownoleglego(ByVal R1 As Double, ByVal R2 As Double, Optional ByVal R3 As Double = -1, Optional ByVal R4 As Double = -1, Optional ByVal R5 As Double = -1, Optional ByVal R6 As Double = -1, Optional ByVal R7 As Double = -1) As Double
        Dim opory As New List(Of Double) From {R1, R2, R3, R4, R5, R6, R7}
        Dim Rz As Double
        Rz = 0
        For Each r As Double In opory
            If (r <> -1) Then
                Rz = Rz + 1 / r
            End If
        Next
        Return 1 / Rz
    End Function
    Private Function OporPolaczeniaRownoleglego(ByVal opory As List(Of Double)) As Double
        Dim Rz As Double
        Rz = 0
        For Each r As Double In opory
            Rz = Rz + 1 / r
        Next
        Return 1 / Rz
    End Function
    Private Function Zmiana(ByVal poczatekZakresuZmiennosci As Double, ByVal koniecZakresuZmiennosi As Double, ByVal Krok As Double) As List(Of Double)
        Dim iloscKrokow As Integer
        iloscKrokow = Math.Floor((koniecZakresuZmiennosi - poczatekZakresuZmiennosci) / Krok)
        Dim x As Double
        Dim wynik As New List(Of Double)
        For i As Integer = 0 To iloscKrokow
            x = poczatekZakresuZmiennosci + i * Krok


            Dim V, f, R2, C, L, R1 As Double
            R1 = 100
            Dim opory As New List(Of Double)
            pobierzDowolnaWartoscTypuDouble(TextBox1, V)
            If Not pobierzWartoscTypuDoubleRoznaOdZera(TextBox2, f) Then

            End If
            opory.Add(x)
            If CheckBox2.Checked Then
                If Not pobierzDowolnaWartoscTypuDouble(TextBox4, L) Then

                End If
                opory.Add(2 * Math.PI * f * L)
            End If
            If CheckBox3.Checked Then
                If Not pobierzDowolnaWartoscTypuDouble(TextBox5, C) Then

                End If
                opory.Add(1 / (2 * Math.PI * f * C))
            End If

            Dim Rz As Double = OporPolaczeniaSzeregowego(100, OporPolaczeniaRownoleglego(opory))



            wynik.Add(NapiecieNaGalezi(100, PradWGalezi(Rz, V))) 'w nawiasie podać wyrażenie zwracające wynik dla danej wartości x
        Next
        Return wynik
    End Function
    Private Sub kopiowanieKoelkcjiDoListBox(ByRef kolekcjaPierwotna As List(Of Double), ByVal kolokcjaWtorna As ListBox.ObjectCollection)
        kolokcjaWtorna.Clear()
        For Each x As Double In kolekcjaPierwotna
            kolokcjaWtorna.Add(x)
        Next
    End Sub
    Private Sub prostyWykres(ByVal wartosci As List(Of Double), ByRef poczatekDziedziny As Double, ByRef koniecDziedziny As Double)
        Dim Top, Left, Right, Bottom, x1, y1, x2, y2 As Integer
        Top = 50
        Bottom = 500
        Left = 250
        Right = 700
        Dim skala, krokow, krok As Double
        skala = (Bottom - Top) / wartosci.Max
        krokow = Math.Ceiling((Right - Left) / wartosci.Count)
        krok = (Right - Left) / wartosci.Count
        Dim graph As Graphics
        graph = Me.CreateGraphics
        graph.Clear(BackColor)
        Dim p As New Pen(Color.Black, 2)
        graph.DrawLine(p, Left, Top, Left, Bottom)
        graph.DrawLine(p, Left, Top, Right, Top)
        graph.DrawLine(p, Left, Bottom, Right, Bottom)
        graph.DrawLine(p, Right, Top, Right, Bottom)
        p.Color = Color.Red
        For x As Integer = 0 To wartosci.Count - 2
            x1 = Left + x * krok
            y1 = Math.Round(Bottom - wartosci(x) * skala)
            x2 = Left + (x + 1) * krok
            y2 = Math.Round(Bottom - wartosci(x + 1) * skala)
            graph.DrawLine(p, x1, y1, x2, y2)
        Next
    End Sub
    Private Sub przecietnyWykres(ByVal wartosci As List(Of Double))
        Dim Top, Left, Right, Bottom, x1, y1, x2, y2, ox, rozmiarXStrzalki, rozmiarYStrzalki, oxr As Integer
        Top = 50
        Bottom = 500
        Left = 250
        Right = 1000
        rozmiarXStrzalki = (Right - Left) / 25
        rozmiarYStrzalki = (Bottom - Top) / 25
        Dim skala, krokow, krok As Double
        skala = (Bottom - Top) / (wartosci.Max - wartosci.Min)
        krokow = Math.Ceiling((Right - Left) / wartosci.Count)
        krok = (Right - Left) / wartosci.Count
        ox = wartosci.Min * skala
        If (ox > 0) Then
            oxr = 0
        Else
            oxr = ox
        End If

        Dim graph As Graphics
        graph = Me.CreateGraphics
        graph.Clear(BackColor)
        Dim p As New Pen(Color.Black, 2)
        graph.DrawLine(p, Left, Top, Left, Bottom)
        graph.DrawLine(p, Left, Top, Right, Top)
        graph.DrawLine(p, Left, Bottom, Right, Bottom)
        graph.DrawLine(p, Right, Top, Right, Bottom)

        graph.DrawLine(p, Left, Bottom + oxr, Right, Bottom + oxr)
        graph.DrawLine(p, Right, Bottom + oxr, Right - rozmiarXStrzalki, Bottom + oxr - rozmiarYStrzalki)
        graph.DrawLine(p, Right, Bottom + oxr, Right - rozmiarXStrzalki, Bottom + oxr + rozmiarYStrzalki)

        p.Color = Color.Red
        For x As Integer = 0 To wartosci.Count - 2
            x1 = Left + x * krok
            y1 = Math.Round(Bottom + ox - wartosci(x) * skala)
            x2 = Left + (x + 1) * krok
            y2 = Math.Round(Bottom + ox - wartosci(x + 1) * skala)
            graph.DrawLine(p, x1, y1, x2, y2)
        Next

    End Sub
    Private Function pobierzDowolnaWartoscTypuDouble(ByVal wart As TextBox, ByRef liczba As Double) As Boolean
        Dim wartosc As Double
        Try
            wartosc = Double.Parse(wart.Text)
        Catch
            MessageBox.Show("Coś poszło nie tak, być może wpisałeś literę zamiast cyfr. Popraw to!")
            Return False
        End Try
        liczba = wartosc
        Return True
    End Function
    Private Function pobierzWartoscTypuDoubleRoznaOdZera(ByVal wart As TextBox, ByRef liczba As Double) As Boolean
        Dim wartosc As Double
        Try
            wartosc = Double.Parse(wart.Text)
        Catch
            MessageBox.Show("Coś poszło nie tak, być może wpisałeś literę zamiast cyfr. Popraw to!")
            Return False
        End Try
        If (wartosc = 0) Then
            MessageBox.Show("Podana wartosc to 0, co jest niedopuszczalne. Popraw to!")
            Return False
        Else
            liczba = wartosc
            Return True
        End If
        liczba = wartosc
        Return True
    End Function
    Private Sub wpiszDoListy(ByVal dane As List(Of Double), ByRef list As ListBox)
        For Each x As Double In dane
            list.Items.Add("Napięcie to " & x)
        Next
    End Sub
    Private Sub FunkcjaDoTegoObwodu()
        ListBox1.Items.Clear()
        Dim V, f, R2, C, L, R1 As Double
        R1 = 100
        Dim opory As New List(Of Double)
        If Not pobierzDowolnaWartoscTypuDouble(TextBox1, V) Then
            Return
        End If
        If Not pobierzWartoscTypuDoubleRoznaOdZera(TextBox2, f) Then
            Return
        End If
        If CheckBox1.Checked Then
            If Not pobierzDowolnaWartoscTypuDouble(TextBox3, R2) Then
                Return
            End If
            opory.Add(R2)
        End If
        If CheckBox2.Checked Then
            If Not pobierzDowolnaWartoscTypuDouble(TextBox4, L) Then
                Return
            End If
            opory.Add(2 * Math.PI * f * L)
        End If
        If CheckBox3.Checked Then
            If Not pobierzDowolnaWartoscTypuDouble(TextBox5, C) Then
                Return
            End If
            opory.Add(1 / (2 * Math.PI * f * C))
        End If
        If Not (CheckBox1.Checked Or CheckBox2.Checked Or CheckBox3.Checked) Then
            MessageBox.Show("Obwód nie jest zamknięty, prąd nie płynie")
            Return
        End If
        Dim Rz As Double = OporPolaczeniaSzeregowego(100, OporPolaczeniaRownoleglego(opory))
        ListBox1.Items.Add("Napięcie na oporze R1 wynosi" & NapiecieNaGalezi(100, PradWGalezi(Rz, V)))
        przecietnyWykres(Zmiana(0, R2, 0.1))
        wpiszDoListy(Zmiana(0, R2, 0.1), ListBox1)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class
