Imports stdole
Imports System.Text
Imports System.Runtime.InteropServices
Module MFreeImage
    Private Const ERROR_SUCCESS As Integer = 0
    Public Const FIF_LOAD_NOPIXELS = &H8000              ' load the image header only (not supported by all plugins)

    Public Const BMP_DEFAULT As Integer = 0
    Public Const BMP_SAVE_RLE As Integer = 1
    Public Const CUT_DEFAULT As Integer = 0
    Public Const DDS_DEFAULT As Integer = 0
    Public Const EXR_DEFAULT As Integer = 0                 ' save data as half with piz-based wavelet compression
    Public Const EXR_FLOAT As Integer = &H1                 ' save data as float instead of as half (not recommended)
    Public Const EXR_NONE As Integer = &H2                  ' save with no compression
    Public Const EXR_ZIP As Integer = &H4                   ' save with zlib compression, in blocks of 16 scan lines
    Public Const EXR_PIZ As Integer = &H8                   ' save with piz-based wavelet compression
    Public Const EXR_PXR24 As Integer = &H10                ' save with lossy 24-bit float compression
    Public Const EXR_B44 As Integer = &H20                  ' save with lossy 44% float compression - goes to 22% when combined with EXR_LC
    Public Const EXR_LC As Integer = &H40                   ' save images with one luminance and two chroma channels, rather than as RGB (lossy compression)
    Public Const FAXG3_DEFAULT As Integer = 0
    Public Const GIF_DEFAULT As Integer = 0
    Public Const GIF_LOAD256 As Integer = 1                 ' Load the image as a 256 color image with ununsed palette entries, if it's 16 or 2 color
    Public Const GIF_PLAYBACK As Integer = 2                ''Play' the GIF to generate each frame (as 32bpp) instead of returning raw frame data when loading
    Public Const HDR_DEFAULT As Integer = 0
    Public Const ICO_DEFAULT As Integer = 0
    Public Const ICO_MAKEALPHA As Integer = 1               ' convert to 32bpp and create an alpha channel from the AND-mask when loading
    Public Const IFF_DEFAULT As Integer = 0
    Public Const J2K_DEFAULT As Integer = 0                ' save with a 16:1 rate
    Public Const JP2_DEFAULT As Integer = 0                 ' save with a 16:1 rate
    Public Const JPEG_DEFAULT As Integer = 0                ' loading (see JPEG_FAST); saving (see JPEG_QUALITYGOOD|JPEG_SUBSAMPLING_420)
    Public Const JPEG_FAST As Integer = &H1                 ' load the file as fast as possible, sacrificing some quality
    Public Const JPEG_ACCURATE As Integer = &H2             ' load the file with the best quality, sacrificing some speed
    Public Const JPEG_CMYK As Integer = &H4                 ' load separated CMYK "as is" (use 'OR' to combine with other flags)
    Public Const JPEG_EXIFROTATE As Integer = &H8           ' load and rotate according to Exif 'Orientation' tag if available
    Public Const JPEG_GREYSCALE As Integer = &H10           ' load and convert to a 8-bit greyscale image
    Public Const JPEG_QUALITYSUPERB As Integer = &H80       ' save with superb quality (100:1)
    Public Const JPEG_QUALITYGOOD As Integer = &H100        ' save with good quality (75:1)
    Public Const JPEG_QUALITYNORMAL As Integer = &H200      ' save with normal quality (50:1)
    Public Const JPEG_QUALITYAVERAGE As Integer = &H400     ' save with average quality (25:1)
    Public Const JPEG_QUALITYBAD As Integer = &H800         ' save with bad quality (10:1)
    Public Const JPEG_PROGRESSIVE As Integer = &H2000       ' save as a progressive-JPEG (use 'OR' to combine with other save flags)
    Public Const JPEG_SUBSAMPLING_411 As Integer = &H1000   ' save with high 4x1 chroma subsampling (4:1:1)
    Public Const JPEG_SUBSAMPLING_420 As Integer = &H4000   ' save with medium 2x2 medium chroma subsampling (4:2:0) - default value
    Public Const JPEG_SUBSAMPLING_422 As Integer = &H8000   ' save with low 2x1 chroma subsampling (4:2:2)
    Public Const JPEG_SUBSAMPLING_444 As Integer = &H10000  ' save with no chroma subsampling (4:4:4)
    Public Const JPEG_OPTIMIZE As Integer = &H20000         ' on saving, compute optimal Huffman coding tables (can reduce a few percent of file size)
    Public Const JPEG_BASELINE As Integer = &H40000         ' save basic JPEG, without metadata or any markers
    Public Const KOALA_DEFAULT As Integer = 0
    Public Const LBM_DEFAULT As Integer = 0
    Public Const MNG_DEFAULT As Integer = 0
    Public Const PCD_DEFAULT As Integer = 0
    Public Const PCD_BASE As Integer = 1                    ' load the bitmap sized 768 x 512
    Public Const PCD_BASEDIV4 As Integer = 2                ' load the bitmap sized 384 x 256
    Public Const PCD_BASEDIV16 As Integer = 3               ' load the bitmap sized 192 x 128
    Public Const PCX_DEFAULT As Integer = 0
    Public Const PFM_DEFAULT As Integer = 0
    Public Const PICT_DEFAULT As Integer = 0
    Public Const PNG_DEFAULT As Integer = 0
    Public Const PNG_IGNOREGAMMA As Integer = 1             ' avoid gamma correction
    Public Const PNG_Z_BEST_SPEED As Integer = &H1          ' save using ZLib level 1 compression flag (default value is 6)
    Public Const PNG_Z_DEFAULT_COMPRESSION As Integer = &H6 ' save using ZLib level 6 compression flag (default recommended value)
    Public Const PNG_Z_BEST_COMPRESSION As Integer = &H9    ' save using ZLib level 9 compression flag (default value is 6)
    Public Const PNG_Z_NO_COMPRESSION As Integer = &H100    ' save without ZLib compression
    Public Const PNG_INTERLACED As Integer = &H200          ' save using Adam7 interlacing (use | to combine with other save flags)
    Public Const PNM_DEFAULT As Integer = 0
    Public Const PNM_SAVE_RAW As Integer = 0                ' if set, the writer saves in RAW format (i.e. P4, P5 or P6)
    Public Const PNM_SAVE_ASCII As Integer = 1              ' if set, the writer saves in ASCII format (i.e. P1, P2 or P3)
    Public Const PSD_DEFAULT As Integer = 0
    Public Const PSD_CMYK As Integer = 1                    ' reads tags for separated CMYK (default is conversion to RGB)
    Public Const PSD_LAB As Integer = 2                     ' reads tags for CIELab (default is conversion to RGB)
    Public Const RAS_DEFAULT As Integer = 0
    Public Const RAW_DEFAULT As Integer = 0                 ' load the file as linear RGB 48-bit
    Public Const RAW_PREVIEW As Integer = 1                 ' try to load the embedded JPEG preview with included Exif Data or default to RGB 24-bit
    Public Const RAW_DISPLAY As Integer = 2                 ' load the file as RGB 24-bit
    Public Const RAW_HALFSIZE As Integer = 4                ' load the file as half-size color image
    Public Const RAW_UNPROCESSED As Integer = 8             ' load the file as FIT_UINT16 raw Bayer image
    Public Const SGI_DEFAULT As Integer = 0
    Public Const TARGA_DEFAULT As Integer = 0
    Public Const TARGA_LOAD_RGB888 As Integer = 1           ' if set, the loader converts RGB555 and ARGB8888 -> RGB888
    Public Const TARGA_SAVE_RLE As Integer = 2              ' if set, the writer saves with RLE compression
    Public Const TIFF_DEFAULT As Integer = 0
    Public Const TIFF_CMYK As Integer = &H1                 ' reads/stores tags for separated CMYK (use 'OR' to combine with compression flags)
    Public Const TIFF_PACKBITS As Integer = &H100           ' save using PACKBITS compression
    Public Const TIFF_DEFLATE As Integer = &H200            ' save using DEFLATE compression (a.k.a. ZLIB compression)
    Public Const TIFF_ADOBE_DEFLATE As Integer = &H400      ' save using ADOBE DEFLATE compression
    Public Const TIFF_NONE As Integer = &H800               ' save without any compression
    Public Const TIFF_CCITTFAX3 As Integer = &H1000         ' save using CCITT Group 3 fax encoding
    Public Const TIFF_CCITTFAX4 As Integer = &H2000         ' save using CCITT Group 4 fax encoding
    Public Const TIFF_LZW As Integer = &H4000               ' save using LZW compression
    Public Const TIFF_JPEG As Integer = &H8000              ' save using JPEG compression
    Public Const TIFF_LOGLUV As Integer = &H10000           ' save using LogLuv compression
    Public Const WBMP_DEFAULT As Integer = 0
    Public Const XBM_DEFAULT As Integer = 0
    Public Const XPM_DEFAULT As Integer = 0
    Public Const WEBP_DEFAULT As Integer = 0                ' save with good quality (75:1)
    Public Const WEBP_LOSSLESS As Integer = &H100           ' save in lossless mode
    Public Const JXR_DEFAULT As Integer = 0                 ' save with quality 80 and no chroma subsampling (4:4:4)
    Public Const JXR_LOSSLESS As Integer = &H64             ' save in lossless mode
    Public Const JXR_PROGRESSIVE As Integer = &H2000        ' save as a progressive-JXR (use Or to combine with other save flags)

    Public Enum FREE_IMAGE_LOAD_OPTIONS
        FILO_LOAD_NOPIXELS = FIF_LOAD_NOPIXELS         ' load the image header only (not supported by all plugins)
        FILO_LOAD_DEFAULT = 0
        FILO_GIF_DEFAULT = GIF_DEFAULT
        FILO_GIF_LOAD256 = GIF_LOAD256                 ' load the image as a 256 color image with ununsed palette entries, if it's 16 or 2 color
        FILO_GIF_PLAYBACK = GIF_PLAYBACK               ' 'play' the GIF to generate each frame (as 32bpp) instead of returning raw frame data when loading
        FILO_ICO_DEFAULT = ICO_DEFAULT
        FILO_ICO_MAKEALPHA = ICO_MAKEALPHA             ' convert to 32bpp and create an alpha channel from the AND-mask when loading
        FILO_JPEG_DEFAULT = JPEG_DEFAULT               ' for loading this is a synonym for FILO_JPEG_FAST
        FILO_JPEG_FAST = JPEG_FAST                     ' load the file as fast as possible, sacrificing some quality
        FILO_JPEG_ACCURATE = JPEG_ACCURATE             ' load the file with the best quality, sacrificing some speed
        FILO_JPEG_CMYK = JPEG_CMYK                     ' load separated CMYK "as is" (use 'OR' to combine with other load flags)
        FILO_JPEG_EXIFROTATE = JPEG_EXIFROTATE         ' load and rotate according to Exif 'Orientation' tag if available
        FILO_JPEG_GREYSCALE = JPEG_GREYSCALE           ' load and convert to a 8-bit greyscale image
        FILO_PCD_DEFAULT = PCD_DEFAULT
        FILO_PCD_BASE = PCD_BASE                       ' load the bitmap sized 768 x 512
        FILO_PCD_BASEDIV4 = PCD_BASEDIV4               ' load the bitmap sized 384 x 256
        FILO_PCD_BASEDIV16 = PCD_BASEDIV16             ' load the bitmap sized 192 x 128
        FILO_PNG_DEFAULT = PNG_DEFAULT
        FILO_PNG_IGNOREGAMMA = PNG_IGNOREGAMMA         ' avoid gamma correction
        FILO_PSD_CMYK = PSD_CMYK                       ' reads tags for separated CMYK (default is conversion to RGB)
        FILO_PSD_LAB = PSD_LAB                         ' reads tags for CIELab (default is conversion to RGB)
        FILO_RAW_DEFAULT = RAW_DEFAULT                 ' load the file as linear RGB 48-bit
        FILO_RAW_PREVIEW = RAW_PREVIEW                 ' try to load the embedded JPEG preview with included Exif Data or default to RGB 24-bit
        FILO_RAW_DISPLAY = RAW_DISPLAY                 ' load the file as RGB 24-bit
        FILO_RAW_HALFSIZE = RAW_HALFSIZE               ' load the file as half-size color image
        FILO_RAW_UNPROCESSED = RAW_UNPROCESSED         ' load the file as FIT_UINT16 raw Bayer image
        FILO_TARGA_DEFAULT = TARGA_LOAD_RGB888
        FILO_TARGA_LOAD_RGB888 = TARGA_LOAD_RGB888     ' if set, the loader converts RGB555 and ARGB8888 -> RGB888
        FISO_TIFF_DEFAULT = TIFF_DEFAULT
        FISO_TIFF_CMYK = TIFF_CMYK                     ' reads tags for separated CMYK
    End Enum
    Public Enum FREE_IMAGE_FILTER
        FILTER_BOX = 0            ' Box, pulse, Fourier window, 1st order (constant) b-spline
        FILTER_BICUBIC = 1        ' Mitchell & Netravali's two-param cubic filter
        FILTER_BILINEAR = 2       ' Bilinear filter
        FILTER_BSPLINE = 3        ' 4th order (cubic) b-spline
        FILTER_CATMULLROM = 4     ' Catmull-Rom spline, Overhauser spline
        FILTER_LANCZOS3 = 5       ' Lanczos3 filter
    End Enum
    Public Enum FREE_IMAGE_FORMAT
        FIF_UNKNOWN = -1
        FIF_BMP = 0
        FIF_ICO = 1
        FIF_JPEG = 2
        FIF_JNG = 3
        FIF_KOALA = 4
        FIF_LBM = 5
        FIF_IFF = FIF_LBM
        FIF_MNG = 6
        FIF_PBM = 7
        FIF_PBMRAW = 8
        FIF_PCD = 9
        FIF_PCX = 10
        FIF_PGM = 11
        FIF_PGMRAW = 12
        FIF_PNG = 13
        FIF_PPM = 14
        FIF_PPMRAW = 15
        FIF_RAS = 16
        FIF_TARGA = 17
        FIF_TIFF = 18
        FIF_WBMP = 19
        FIF_PSD = 20
        FIF_CUT = 21
        FIF_XBM = 22
        FIF_XPM = 23
        FIF_DDS = 24
        FIF_GIF = 25
        FIF_HDR = 26
        FIF_FAXG3 = 27
        FIF_SGI = 28
        FIF_EXR = 29
        FIF_J2K = 30
        FIF_JP2 = 31
        FIF_PFM = 32
        FIF_PICT = 33
        FIF_RAW = 34
        FIF_WEBP = 35
        FIF_JXR = 36
    End Enum

    Private Structure PictDesc
        Dim cbSizeofStruct As Integer
        Dim picType As Integer
        Dim hImage As Integer
        Dim xExt As Integer
        Dim yExt As Integer
    End Structure

    Private Structure GUID
        Dim Data1 As Integer
        Dim Data2 As Integer
        Dim Data3 As Integer
        Dim Data4() As Byte
    End Structure

    Private Declare Function FreeImage_GetVersionInt Lib "FreeImage.dll" Alias "_FreeImage_GetVersion@0" () As Integer
    Public Declare Function FreeImage_GetFileType Lib "FreeImage.dll" Alias "_FreeImage_GetFileType@8" (
           ByVal FileName As String,
  Optional ByVal Size As Integer = 0) As FREE_IMAGE_FORMAT

    Public Declare Function FreeImage_Load Lib "FreeImage.dll" Alias "_FreeImage_Load@12" (
           ByVal Format As FREE_IMAGE_FORMAT,
           ByVal FileName As String,
  Optional ByVal Flags As FREE_IMAGE_LOAD_OPTIONS = 0) As Integer

    Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (
    ByRef lpPictDesc As PictDesc,
    ByRef riid As GUID,
    ByVal fOwn As Integer,
    ByRef lplpvObj As IPicture) As Integer



    Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As Integer) As Integer

    Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (
    ByRef Destination As Decimal,
    ByRef Source As Decimal,
    ByVal Length As Integer)

    Private Declare Function FreeImage_FIFSupportsReadingInt Lib "FreeImage.dll" Alias "_FreeImage_FIFSupportsReading@4" (
           ByVal Format As FREE_IMAGE_FORMAT) As Integer
    Public Declare Function FreeImage_GetWidth Lib "FreeImage.dll" Alias "_FreeImage_GetWidth@4" (
           ByVal BITMAP As Integer) As Integer
    Public Declare Function FreeImage_GetHeight Lib "FreeImage.dll" Alias "_FreeImage_GetHeight@4" (
           ByVal BITMAP As Integer) As Integer
    Public Declare Function FreeImage_Rescale Lib "FreeImage.dll" Alias "_FreeImage_Rescale@16" (
           ByVal BITMAP As Integer,
           ByVal Width As Integer,
           ByVal Height As Integer,
           ByVal Filter As FREE_IMAGE_FILTER) As Integer
    Public Declare Sub FreeImage_Unload Lib "FreeImage.dll" Alias "_FreeImage_Unload@4" (
           ByVal BITMAP As Integer)

    Private Declare Function FreeImage_GetFormatFromFIFInt Lib "FreeImage.dll" Alias "_FreeImage_GetFormatFromFIF@4" (
           ByVal Format As FREE_IMAGE_FORMAT) As Integer
    Public Declare Function FreeImage_HasPixelsInt Lib "FreeImage.dll" Alias "_FreeImage_HasPixels@4" (
           ByVal BITMAP As Integer) As Integer

    Private Declare Function CreateDIBitmap Lib "gdi32.dll" (
    ByVal hDC As Integer,
    ByVal lpInfoHeader As Integer,
    ByVal dwUsage As Integer,
    ByVal lpInitBits As Integer,
    ByVal lpInitInfo As Integer,
    ByVal wUsage As Integer) As Integer

    Public Declare Function FreeImage_GetInfoHeader Lib "FreeImage.dll" Alias "_FreeImage_GetInfoHeader@4" (
           ByVal BITMAP As Integer) As Integer

    Private Const CBM_INIT As Integer = &H4

    Public Declare Function FreeImage_GetBits Lib "FreeImage.dll" Alias "_FreeImage_GetBits@4" (
           ByVal BITMAP As Integer) As Integer

    Public Declare Function FreeImage_GetInfo Lib "FreeImage.dll" Alias "_FreeImage_GetInfo@4" (
           ByVal BITMAP As Integer) As Integer

    Private Const DIB_RGB_COLORS As Integer = 0

    Private Declare Function ReleaseDC Lib "user32.dll" (
    ByVal hwnd As Integer,
    ByVal hDC As Integer) As Integer

    Public Function FreeImage_IsAvailable(Optional ByRef Version As String = "") As Boolean
        On Error Resume Next
        Version = FreeImage_GetVersion()
        FreeImage_IsAvailable = (Err.Number = ERROR_SUCCESS)
        On Error GoTo 0
    End Function

    Public Function LoadPictureEx(Optional ByRef FileName As String = Nothing,
                              Optional ByRef Options As FREE_IMAGE_LOAD_OPTIONS = Nothing,
                              Optional ByRef Width As Object = Nothing,
                              Optional ByRef Height As Object = Nothing,
                              Optional ByRef InPercent As Boolean = Nothing,
                              Optional ByRef Filter As FREE_IMAGE_FILTER = Nothing,
                              Optional ByRef Format As FREE_IMAGE_FORMAT = Nothing) As IPicture

        Dim hDIB As Integer

        ' This function is an extended version of the VB method 'LoadPicture'. As
        ' the VB version it takes a filename parameter to load the image and throws
        ' the same errors in most cases.

        ' This function now is only a thin wrapper for the FreeImage_LoadEx() wrapper
        ' function (as compared to releases of this wrapper prior to version 1.8). So,
        ' have a look at this function's discussion of the parameters.

        ' However, we do mask out the FILO_LOAD_NOPIXELS load option, since this
        ' function shall create a VB Picture object, which does not support
        ' FreeImage's header-only loading option.


        If (Not IsNothing(FileName)) Then
            hDIB = FreeImage_LoadEx(FileName, (Options And (Not FREE_IMAGE_LOAD_OPTIONS.FILO_LOAD_NOPIXELS)), Width, Height, InPercent, Filter, Format)
            LoadPictureEx = FreeImage_GetOlePicture(hDIB, , True)
        End If

    End Function
    Public Function FreeImage_GetVersion() As String

        ' This function returns the version of the FreeImage 3 library
        ' as VB String.

        FreeImage_GetVersion = pGetStringFromPointerA(FreeImage_GetVersionInt)

    End Function
    Public Function FreeImage_LoadEx(ByVal FileName As String,
                        Optional ByVal Options As FREE_IMAGE_LOAD_OPTIONS = Nothing,
                        Optional ByVal Width As Object = Nothing,
                        Optional ByVal Height As Object = Nothing,
                        Optional ByVal InPercent As Boolean = False,
                        Optional ByVal Filter As FREE_IMAGE_FILTER = Nothing,
                        Optional ByRef Format As FREE_IMAGE_FORMAT = Nothing) As Integer

        Const vbInvalidPictureError As Integer = 481

        ' The function provides all image formats, the FreeImage library can read. The
        ' image format is determined from the image file to load, the optional parameter
        ' 'Format' is an OUT parameter that will contain the image format that has
        ' been loaded.

        ' The parameters 'Width', 'Height', 'InPercent' and 'Filter' make it possible
        ' to "load" the image in a resized version. 'Width', 'Height' specify the desired
        ' width and height, 'Filter' determines, what image filter should be used
        ' on the resizing process.

        ' The parameters 'Width', 'Height', 'InPercent' and 'Filter' map directly to the
        ' according parameters of the 'FreeImage_RescaleEx' function. So, read the
        ' documentation of the 'FreeImage_RescaleEx' for a complete understanding of the
        ' usage of these parameters.


        Format = FreeImage_GetFileType(FileName)
        If (Format <> FREE_IMAGE_FORMAT.FIF_UNKNOWN) Then
            If (FreeImage_FIFSupportsReading(Format)) Then
                FreeImage_LoadEx = FreeImage_Load(Format, FileName, Options)
                If (FreeImage_LoadEx) Then

                    If ((Not IsNothing(Width)) Or
                (Not IsNothing(Height))) Then
                        FreeImage_LoadEx = FreeImage_RescaleEx(FreeImage_LoadEx, Width, Height,
                     InPercent, True, Filter)
                    End If
                Else
                    Call Err.Raise(vbInvalidPictureError)
                End If
            Else
                'Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf &
                Call Err.Raise(5, "MFreeImage", vbCrLf & vbCrLf &
                        "FreeImage Library plugin '" & FreeImage_GetFormatFromFIF(Format) & "' " &
                        "does not support reading.")
            End If
        Else
            'Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf &
            Call Err.Raise(5, "MFreeImage", vbCrLf & vbCrLf &
                     "The file specified has an unknown image format.")
        End If

    End Function
    Public Function FreeImage_GetOlePicture(ByVal BITMAP As Integer,
                               Optional ByVal hDC As Integer = 0,
                               Optional ByVal UnloadSource As Boolean = False) As IPicture

        Dim hBitmap As Integer
        Dim tPicDesc As PictDesc
        Dim tGuid As GUID
        Dim cPictureDisp As IPictureDisp

        ' This function creates a VB Picture object (OlePicture) from a FreeImage DIB.
        ' The original image need not remain valid nor loaded after the VB Picture
        ' object has been created.

        ' The optional parameter 'hDC' determines the device context (DC) used for
        ' transforming the device independent bitmap (DIB) to a device dependent
        ' bitmap (DDB). This device context's color depth is responsible for this
        ' transformation. This parameter may be null or omitted. In that case, the
        ' windows desktop's device context will be used, what will be the desired
        ' way in almost any cases.

        ' The optional 'UnloadSource' parameter is for unloading the original image
        ' after the OlePicture has been created, so you can easily "switch" from a
        ' FreeImage DIB to a VB Picture object. There is no need to unload the DIB
        ' at the caller's site if this argument is True.


        If (BITMAP) Then

            If (Not FreeImage_HasPixels(BITMAP)) Then
                'Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf &
                Call Err.Raise(5, "MFreeImage", vbCrLf & vbCrLf &
                        "Unable to create a picture from a 'header-only' bitmap.")
            End If

            hBitmap = FreeImage_GetBitmapForDevice(BITMAP, hDC, UnloadSource)
            If (hBitmap) Then
                ' fill tPictDesc structure with necessary parts
                With tPicDesc
                    .cbSizeofStruct = Len(tPicDesc)
                    ' the vbPicTypeBitmap constant is not available in VBA environemnts
                    .picType = 1  'vbPicTypeBitmap
                    .hImage = hBitmap
                End With

                ' fill in IDispatch Interface ID
                With tGuid
                    ReDim tGuid.Data4(8)
                    .Data1 = &H20400
                    .Data4(0) = &HC0
                    .Data4(7) = &H46
                End With

                ' create a picture object
                cPictureDisp = Nothing
                Call OleCreatePictureIndirect(tPicDesc, tGuid, True, cPictureDisp)

                FreeImage_GetOlePicture = cPictureDisp
            End If
        End If

    End Function


    Private Function pGetStringFromPointerA(ByRef ptr As Integer) As String

        Dim abBuffer() As Byte
        Dim lLength As Integer


        'Dim ute32 As UTF32Encoding = New UTF32Encoding
        ' This function creates and returns a VB BSTR variable from
        ' a C/C++ style string pointer by making a redundant deep
        ' copy of the string's characters.

        If (ptr) Then
            ' get the length of the ANSI string pointed to by ptr
            lLength = lstrlen(ptr)
            If (lLength) Then
                ' copy characters to a byte array
                ReDim abBuffer(lLength - 1)
                'Call CopyMemory(abBuffer(0), ptr, lLength)
                'abBuffer(0) = ptr.ToString
                ' convert from byte array to unicode BSTR
                'pGetStringFromPointerA = StrConv(abBuffer, vbUnicode)

                abBuffer = Encoding.Default.GetBytes(ptr.ToString)
                'pGetStringFromPointerA = Text.Encoding.Default.GetString(abBuffer, 0, abBuffer.Length)
                pGetStringFromPointerA = Text.Encoding.Default.GetString(abBuffer)
                'pGetStringFromPointerA = abBuffer.ToString

                '--------------
                'abBuffer = Encoding.Default.GetBytes(ptr.ToString)

            End If
        End If

    End Function
    Public Function FreeImage_FIFSupportsReading(ByVal Format As FREE_IMAGE_FORMAT) As Boolean

        ' Thin wrapper function returning a real VB Boolean value

        FreeImage_FIFSupportsReading = (FreeImage_FIFSupportsReadingInt(Format) = 1)

    End Function
    Public Function FreeImage_RescaleEx(ByVal BITMAP As Integer,
                           Optional ByVal Width As Object = Nothing,
                           Optional ByVal Height As Object = Nothing,
                           Optional ByVal IsPercentValue As Boolean = False,
                           Optional ByVal UnloadSource As Boolean = False,
                           Optional ByVal Filter As FREE_IMAGE_FILTER = FREE_IMAGE_FILTER.FILTER_BICUBIC,
                           Optional ByVal ForceCloneCreation As Boolean = False) As Integer

        Dim lNewWidth As Integer
        Dim lNewHeight As Integer
        Dim hDIBNew As Integer

        ' This function is a easy-to-use wrapper for rescaling an image with the
        ' FreeImage library. It returns a pointer to a new rescaled DIB provided
        ' by FreeImage.

        ' The parameters 'Width', 'Height' and 'IsPercentValue' control
        ' the size of the new image. Here, the function tries to fake something like
        ' overloading known from Java. It depends on the parameter's data type passed
        ' through the Variant, how the provided values for width and height are
        ' actually interpreted. The following rules apply:

        ' In general, non integer values are either interpreted as percent values or
        ' factors, the original image size will be multiplied with. The 'IsPercentValue'
        ' parameter controls whether the values are percent values or factors. Integer
        ' values are always considered to be the direct new image size, not depending on
        ' the original image size. In that case, the 'IsPercentValue' parameter has no
        ' effect. If one of the parameters is omitted, the image will not be resized in
        ' that direction (either in width or height) and keeps it's original size. It is
        ' possible to omit both, but that makes actually no sense.

        ' The following table shows some of possible data type and value combinations
        ' that might by used with that function: (assume an original image sized 100x100 px)

        ' Parameter         |  Values |  Values |  Values |  Values |     Values |
        ' ----------------------------------------------------------------------
        ' Width             |    75.0 |    0.85 |     200 |     120 |      400.0 |
        ' Height            |   120.0 |     1.3 |     230 |       - |      400.0 |
        ' IsPercentValue    |    True |   False |    d.c. |    d.c. |      False | <- wrong option?
        ' ----------------------------------------------------------------------
        ' Result Size       |  75x120 |  85x130 | 200x230 | 120x100 |40000x40000 |
        ' Remarks           | percent |  factor |  direct |         |maybe not   |
        '                                                           |what you    |
        '                                                           |wanted,     |
        '                                                           |right?      |

        ' The optional 'UnloadSource' parameter is for unloading the original image, so
        ' you can "change" an image with this function rather than getting a new DIB
        ' pointer. There is no more need for a second DIB variable at the caller's site.

        ' As of version 2.0 of the FreeImage VB wrapper, this function and all it's derived
        ' functions like FreeImage_RescaleByPixel() or FreeImage_RescaleByPercent(), do NOT
        ' return a clone of the image, if the new size desired is the same as the source
        ' image's size. That behaviour can be forced by setting the new parameter
        ' 'ForceCloneCreation' to True. Then, an image is also rescaled (and so
        ' effectively cloned), if the new width and height is exactly the same as the source
        ' image's width and height.

        ' Since this diversity may be confusing to VB developers, this function is also
        ' callable through three different functions called 'FreeImage_RescaleByPixel',
        ' 'FreeImage_RescaleByPercent' and 'FreeImage_RescaleByFactor'.

        If (BITMAP) Then

            If (Not FreeImage_HasPixels(BITMAP)) Then
                'Call Err.Raise(5, "MFreeImage", Error$(5) & vbCrLf & vbCrLf &
                Call Err.Raise(5, "MFreeImage", vbCrLf & vbCrLf &
                        "Unable to rescale a 'header-only' bitmap.")
            End If

            If (Not IsNothing(Width)) Then
                Select Case VarType(Width)

                    Case vbDouble, vbSingle, vbDecimal, vbCurrency
                        lNewWidth = FreeImage_GetWidth(BITMAP) * Width
                        If (IsPercentValue) Then
                            lNewWidth = lNewWidth / 100
                        End If

                    Case Else
                        lNewWidth = Width

                End Select
            End If

            If (Not IsNothing(Height)) Then
                Select Case VarType(Height)

                    Case vbDouble, vbSingle, vbDecimal
                        lNewHeight = FreeImage_GetHeight(BITMAP) * Height
                        If (IsPercentValue) Then
                            lNewHeight = lNewHeight / 100
                        End If

                    Case Else
                        lNewHeight = Height

                End Select
            End If

            If ((lNewWidth > 0) And (lNewHeight > 0)) Then
                If (ForceCloneCreation) Then
                    hDIBNew = FreeImage_Rescale(BITMAP, lNewWidth, lNewHeight, Filter)

                ElseIf ((lNewWidth <> FreeImage_GetWidth(BITMAP)) Or
                 (lNewHeight <> FreeImage_GetHeight(BITMAP))) Then
                    hDIBNew = FreeImage_Rescale(BITMAP, lNewWidth, lNewHeight, Filter)

                End If

            ElseIf (lNewWidth > 0) Then
                If ((lNewWidth <> FreeImage_GetWidth(BITMAP)) Or
             (ForceCloneCreation)) Then
                    lNewHeight = lNewWidth / (FreeImage_GetWidth(BITMAP) / FreeImage_GetHeight(BITMAP))
                    hDIBNew = FreeImage_Rescale(BITMAP, lNewWidth, lNewHeight, Filter)
                End If

            ElseIf (lNewHeight > 0) Then
                If ((lNewHeight <> FreeImage_GetHeight(BITMAP)) Or
             (ForceCloneCreation)) Then
                    lNewWidth = lNewHeight * (FreeImage_GetWidth(BITMAP) / FreeImage_GetHeight(BITMAP))
                    hDIBNew = FreeImage_Rescale(BITMAP, lNewWidth, lNewHeight, Filter)
                End If

            End If

            If (hDIBNew) Then
                FreeImage_RescaleEx = hDIBNew
                If (UnloadSource) Then
                    Call FreeImage_Unload(BITMAP)
                End If
            Else
                FreeImage_RescaleEx = BITMAP
            End If
        End If

    End Function
    Public Function FreeImage_GetFormatFromFIF(ByVal Format As FREE_IMAGE_FORMAT) As String

        ' This function returns the result of the 'FreeImage_GetFormatFromFIF' function
        ' as VB String.

        ' The parameter 'Format' works according to the FreeImage 3 API documentation.

        FreeImage_GetFormatFromFIF = pGetStringFromPointerA(FreeImage_GetFormatFromFIFInt(Format))

    End Function
    Public Function FreeImage_HasPixels(ByVal BITMAP As Integer) As Boolean

        ' Thin wrapper function returning a real VB Boolean value

        FreeImage_HasPixels = (FreeImage_HasPixelsInt(BITMAP) = 1)

    End Function
    Public Function FreeImage_GetBitmapForDevice(ByVal BITMAP As Integer,
                                    Optional ByVal hDC As Integer = 0,
                                    Optional ByVal UnloadSource As Boolean = False) As Integer

        Dim bReleaseDC As Boolean

        ' This function returns an HBITMAP created by the CreateDIBitmap() function which
        ' in turn has always the same color depth as the reference DC, which may be provided
        ' through the 'hDC' parameter. The desktop DC will be used, if no reference DC is
        ' specified.

        If (BITMAP) Then

            If (Not FreeImage_HasPixels(BITMAP)) Then
                Call Err.Raise(5, "MFreeImage", vbCrLf & vbCrLf &
                        "Unable to create a bitmap from a 'header-only' bitmap.")
            End If

            If (hDC = 0) Then
                hDC = GetDC(0)
                bReleaseDC = True
            End If
            If (hDC) Then
                FreeImage_GetBitmapForDevice =
               CreateDIBitmap(hDC, FreeImage_GetInfoHeader(BITMAP), CBM_INIT,
                     FreeImage_GetBits(BITMAP), FreeImage_GetInfo(BITMAP),
                           DIB_RGB_COLORS)
                If (UnloadSource) Then
                    Call FreeImage_Unload(BITMAP)
                End If
                If (bReleaseDC) Then
                    Call ReleaseDC(0, hDC)
                End If
            End If
        End If

    End Function

End Module
