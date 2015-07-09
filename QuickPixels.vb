Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Runtime.InteropServices

''' <summary>
''' LockBits method on BitMap class, used to process data faster.  Made by Caleb Woolley, 2015.
''' </summary>
''' <remarks></remarks>
Public Class QuickPixels
    ' Bitmap reference and data.
    Private bmp As Bitmap
    Private bitmapData As BitmapData
    Private bounds As Rectangle

    ' Bitmap byte data, public.
    Public imgBytes() As Byte
    Public stride As ULong
    Public Const PixelDataSize As Integer = 1
    Public totalSize As Integer

    Public Sub New(ByRef bm As Bitmap)
        bmp = bm
        bounds = New Rectangle(0, 0, bmp.Width, bmp.Height)
        bitmapData = bmp.LockBits(bounds, Imaging.ImageLockMode.ReadWrite, Imaging.PixelFormat.Format24bppRgb)
        stride = BitmapData.Stride
        totalSize = Math.Abs(BitmapData.Stride) * BitmapData.Height
        ReDim imgBytes(totalSize)
        Marshal.Copy(BitmapData.Scan0, imgBytes, 0, totalSize)
    End Sub

    Public Sub LockBitmap()
        bitmapData = bmp.LockBits(bounds, Imaging.ImageLockMode.ReadWrite, Imaging.PixelFormat.Format24bppRgb)
        Marshal.Copy(BitmapData.Scan0, imgBytes, 0, totalSize)
    End Sub

    Public Sub UnlockBitmap()
        Marshal.Copy(imgBytes, 0, BitmapData.Scan0, totalSize)
        bmp.UnlockBits(bitmapData)
    End Sub

    ''' <summary>
    ''' Gets the color of the pixel at the specified X and Y coordinates.
    ''' </summary>
    ''' <param name="x">X location on specified bitmap.</param>
    ''' <param name="y">Y location on specified bitmap.</param>
    ''' <returns>Color of pixel locaton.</returns>
    Public Function GetPixel(ByRef x As Integer, ByRef y As Integer) As Color
        Dim l As Integer = (y * stride) + x * 3
        Dim valB As Integer = imgBytes(l)
        Dim valG As Integer = imgBytes(l + 1)
        Dim valR As Integer = imgBytes(l + 2)
        Return Color.FromArgb(valR, valG, valB)
    End Function

End Class
