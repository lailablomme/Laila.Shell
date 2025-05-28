Imports System.Globalization
Imports System.Text
Imports System.Windows.Data
Imports System.Windows.Media

Namespace Converters
    Public Class LightnessColorConverter
        Implements IValueConverter

        Public Function Convert(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.Convert
            ' Get HLS values of color passed in
            Dim rgbColorIn = value
            Dim hlsColor = RgbToHls(rgbColorIn)

            'Console.WriteLine("*** RGB IN ***")
            'Console.WriteLine("R={0}", rgbColorIn.R)
            'Console.WriteLine("G={0}", rgbColorIn.G)
            'Console.WriteLine("B={0}", rgbColorIn.B)
            'Console.WriteLine("A={0}", rgbColorIn.A)
            'Console.WriteLine()

            'Console.WriteLine("*** HLS IN ***")
            'Console.WriteLine("H={0}", hlsColor.H.ToString("#"))
            'Console.WriteLine("L={0}", hlsColor.L.ToString("P0"))
            'Console.WriteLine("S={0}", hlsColor.S.ToString("P0"))
            'Console.WriteLine("A={0}", hlsColor.A.ToString("P0"))
            'Console.WriteLine()

            ' Adjust color by factor passed in
            Dim brightnessAdjustment As Double = Double.Parse(String.Format(CultureInfo.InvariantCulture, "{0}", parameter), CultureInfo.InvariantCulture)
            hlsColor.L *= brightnessAdjustment
            If hlsColor.L > 1 Then hlsColor.L = 1
            If hlsColor.L < 0 Then hlsColor.L = 0

            'Console.WriteLine("*** HLS OUT ***")
            'Console.WriteLine("H={0}", hlsColor.H.ToString("#"))
            'Console.WriteLine("L={0}", hlsColor.L.ToString("P0"))
            'Console.WriteLine("S={0}", hlsColor.S.ToString("P0"))
            'Console.WriteLine("A={0}", hlsColor.A.ToString("P0"))
            'Console.WriteLine()

            ' Return result
            Dim rgbColorOut = HlsToRgb(hlsColor)
            Return rgbColorOut
        End Function

        Public Function ConvertBack(value As Object, targetType As Type, parameter As Object, culture As CultureInfo) As Object Implements IValueConverter.ConvertBack
            Throw New NotImplementedException()
        End Function

        Private Structure ColorHls
            Public H As Double
            Public L As Double
            Public S As Double
            Public A As Double
        End Structure

        Private Function HexToRgb(hexRgbColor As String) As Color
            Dim rgbColor = New Color()

            Dim s = hexRgbColor.Substring(1, 2)
            Dim i = Int32.Parse(s, NumberStyles.HexNumber)
            rgbColor.A = System.Convert.ToByte(i)

            s = hexRgbColor.Substring(3, 2)
            i = Int32.Parse(s, NumberStyles.HexNumber)
            rgbColor.R = System.Convert.ToByte(i)

            s = hexRgbColor.Substring(5, 2)
            i = Int32.Parse(s, NumberStyles.HexNumber)
            rgbColor.G = System.Convert.ToByte(i)

            s = hexRgbColor.Substring(7, 2)
            i = Int32.Parse(s, NumberStyles.HexNumber)
            rgbColor.B = System.Convert.ToByte(i)

            Return rgbColor
        End Function

        Private Function RgbToHex(rgbColor As Color) As String
            Dim sb As StringBuilder = New StringBuilder("#")
            sb.Append(rgbColor.A.ToString("X"))
            sb.Append(rgbColor.R.ToString("X"))
            sb.Append(rgbColor.G.ToString("X"))
            sb.Append(rgbColor.B.ToString("X"))
            Return sb.ToString()
        End Function

        Private Function RgbToHls(rgbColor As Color) As ColorHls
            ' Initialize result
            Dim hlsColor = New ColorHls()

            ' Convert RGB values to percentages
            Dim r As Double = rgbColor.R / 255
            Dim g = rgbColor.G / 255
            Dim b = rgbColor.B / 255
            Dim a = rgbColor.A / 255

            ' Find min And max RGB values
            Dim min = Math.Min(r, Math.Min(g, b))
            Dim max = Math.Max(r, Math.Max(g, b))
            Dim delta = max - min

            ' If max And min are equal, that means we are dealing with 
            ' a shade of gray. So we set H And S to zero, And L to either
            ' max Or min (it doesn't matter which), and  then we exit. */

            ' Special case Gray
            If (max = min) Then
                hlsColor.H = 0
                hlsColor.S = 0
                hlsColor.L = max
                hlsColor.A = a
                Return hlsColor
            End If

            ' If we get to this point, we know we don't have a shade of gray. 

            ' Set L
            hlsColor.L = (min + max) / 2

            ' Set S
            If (hlsColor.L < 0.5) Then
                hlsColor.S = delta / (max + min)
            Else
                hlsColor.S = delta / (2.0 - max - min)
            End If

            ' Set H
            If (r = max) Then hlsColor.H = (g - b) / delta
            If (g = max) Then hlsColor.H = 2.0 + (b - r) / delta
            If (b = max) Then hlsColor.H = 4.0 + (r - g) / delta
            hlsColor.H *= 60
            If (hlsColor.H < 0) Then hlsColor.H += 360

            ' Set A
            hlsColor.A = a

            ' Set return value
            Return hlsColor
        End Function

        Private Function HlsToRgb(hlsColor As ColorHls) As Color
            ' Initialize result
            Dim rgbColor = New Color()

            ' If S = 0, that means we are dealing with a shade 
            ' of gray. So, we set R, G, And B to L And exit. */

            ' Special case Gray
            If (hlsColor.S = 0) Then
                rgbColor.R = (hlsColor.L * 255)
                rgbColor.G = (hlsColor.L * 255)
                rgbColor.B = (hlsColor.L * 255)
                rgbColor.A = (hlsColor.A * 255)
                Return rgbColor
            End If

            Dim t1 As Double
            If (hlsColor.L < 0.5) Then
                t1 = hlsColor.L * (1.0 + hlsColor.S)
            Else
                t1 = hlsColor.L + hlsColor.S - (hlsColor.L * hlsColor.S)
            End If

            Dim t2 = 2.0 * hlsColor.L - t1

            ' Convert H from degrees to a percentage
            Dim h = hlsColor.H / 360

            ' Set colors as percentage values 
            Dim tR = h + (1.0 / 3.0)
            Dim r = SetColor(t1, t2, tR)

            Dim tG = h
            Dim g = SetColor(t1, t2, tG)

            Dim tB = h - (1.0 / 3.0)
            Dim b = SetColor(t1, t2, tB)

            ' Assign colors to Color object
            rgbColor.R = (r * 255)
            rgbColor.G = (g * 255)
            rgbColor.B = (b * 255)
            rgbColor.A = (hlsColor.A * 255)

            ' Set return value
            Return rgbColor
        End Function

        Private Function SetColor(t1 As Double, t2 As Double, t3 As Double) As Double
            If (t3 < 0) Then t3 += 1.0
            If (t3 > 1) Then t3 -= 1.0

            Dim color As Double
            If (6.0 * t3 < 1) Then
                color = t2 + (t1 - t2) * 6.0 * t3
            ElseIf (2.0 * t3 < 1) Then
                color = t1
            ElseIf (3.0 * t3 < 2) Then
                color = t2 + (t1 - t2) * ((2.0 / 3.0) - t3) * 6.0
            Else
                color = t2
            End If

            ' Set return value
            Return color
        End Function
    End Class
End Namespace