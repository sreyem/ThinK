

Imports System.ComponentModel
Imports System.Drawing.Design
Imports essentials
Imports ThinK.toolkit

Public Class toolkit


#Region "     Hypres"

#End Region

    <TypeConverter(GetType(enumConverter(Of eHypresClass)))>
    Public Enum eHypresClass

        <Description("Coarse")>
        coarse

        <Description("Fine")>
        fine

        <Description("Medium")>
        medium

        <Description("Medium - Fine")>
        mediumFine

        <Description("Organic")>
        organic

        <Description(enumConverter(Of Type).not_defined)>
        not_def = -1

    End Enum

    ''' <summary>
    ''' get Hypres-Class out of clay and sand content
    ''' </summary>
    ''' <param name="sand">
    ''' Sand content 0 - 100%, Integer.MaxValue if to be calculated
    ''' </param>
    ''' <param name="clay">
    ''' Clay content 0 - 100%, Integer.MaxValue if to be calculated
    ''' </param>
    ''' <param name="silt">
    ''' Silt content 0 - 100%, Integer.MaxValue if to be calculated
    ''' </param>
    ''' <returns>
    ''' not_def, if something went wrong
    ''' </returns>
    Public Shared Function getHypresClass(sand As Integer, clay As Integer, silt As Integer) As eHypresClass

        Try

            If sand = Integer.MaxValue Then
                sand = 100 - (clay + silt)
            ElseIf clay = Integer.MaxValue Then
                clay = 100 - (sand + silt)
            ElseIf silt = Integer.MaxValue Then
                silt = 100 - (sand + clay)
            Else
                If clay + silt + sand <> 100 Then
                    Return eHypresClass.not_def
                End If
            End If

            If sand < 0 OrElse silt < 0 OrElse clay < 0 Then
                Return eHypresClass.not_def
            End If

        Catch ex As Exception
            Return eHypresClass.not_def
        End Try


        'Rules

        If (clay >= 60) Then
            Return eHypresClass.organic
        End If

        If (clay < 60 AndAlso clay >= 35) Then
            Return eHypresClass.fine
        End If

        If (clay < 18 AndAlso sand >= 65) Then
            Return eHypresClass.coarse
        End If

        If (sand <= 15 AndAlso clay < 35) Then
            Return eHypresClass.mediumFine
        End If

        If ((clay >= 18 AndAlso clay < 35 AndAlso sand > 15) OrElse
            (clay < 18 AndAlso sand >= 15 AndAlso sand <= 65)) Then
            Return eHypresClass.medium
        End If

        Return eHypresClass.not_def

    End Function


    Public Enum eUSDAclass

        clay = 1
        sandyClay
        siltyClay
        loam
        clayLoam
        siltyClayLoam
        siltLoam
        silt
        sandyClayLoam
        sandyLoam
        sand
        loamySand

        not_def = -1

    End Enum


    Public Shared Function getUSDAClass(xSand As Single, yClay As Single) As eUSDAclass

        Dim SandTexC(eUSDAclass.loamySand, 10) As Single
        Dim ClayTexC(eUSDAclass.loamySand, 10) As Single
        Dim Sand(10) As Single
        Dim Clay(10) As Single
        Dim nPoly(eUSDAclass.loamySand) As Integer
        Dim texStr(eUSDAclass.loamySand) As String


#Region "    Definitions"

        nPoly(1) = 6
        nPoly(2) = 4
        nPoly(3) = 4
        nPoly(4) = 6
        nPoly(5) = 5
        nPoly(6) = 5
        nPoly(7) = 7
        nPoly(8) = 5
        nPoly(9) = 6
        nPoly(10) = 8
        nPoly(11) = 4
        nPoly(12) = 5

        SandTexC(1, 1) = 0
        SandTexC(1, 2) = 0.45
        SandTexC(1, 3) = 0.45
        SandTexC(1, 4) = 0.2
        SandTexC(1, 5) = 0
        SandTexC(1, 6) = 0

        ClayTexC(1, 1) = 1
        ClayTexC(1, 2) = 0.55
        ClayTexC(1, 3) = 0.4
        ClayTexC(1, 4) = 0.4
        ClayTexC(1, 5) = 0.6
        ClayTexC(1, 6) = 1

        SandTexC(2, 1) = 0.45
        SandTexC(2, 2) = 0.65
        SandTexC(2, 3) = 0.45
        SandTexC(2, 4) = 0.45

        ClayTexC(2, 1) = 0.55
        ClayTexC(2, 2) = 0.35
        ClayTexC(2, 3) = 0.35
        ClayTexC(2, 4) = 0.55

        SandTexC(3, 1) = 0
        SandTexC(3, 2) = 0.2
        SandTexC(3, 3) = 0
        SandTexC(3, 4) = 0

        ClayTexC(3, 1) = 0.6
        ClayTexC(3, 2) = 0.4
        ClayTexC(3, 3) = 0.4
        ClayTexC(3, 4) = 0.6

        SandTexC(4, 1) = 0.45
        SandTexC(4, 2) = 0.52
        SandTexC(4, 3) = 0.52
        SandTexC(4, 4) = 0.43
        SandTexC(4, 5) = 0.23
        SandTexC(4, 6) = 0.45

        ClayTexC(4, 1) = 0.27
        ClayTexC(4, 2) = 0.2
        ClayTexC(4, 3) = 0.07
        ClayTexC(4, 4) = 0.07
        ClayTexC(4, 5) = 0.27
        ClayTexC(4, 6) = 0.27

        SandTexC(5, 1) = 0.45
        SandTexC(5, 2) = 0.45
        SandTexC(5, 3) = 0.2
        SandTexC(5, 4) = 0.2
        SandTexC(5, 5) = 0.45

        ClayTexC(5, 1) = 0.4
        ClayTexC(5, 2) = 0.27
        ClayTexC(5, 3) = 0.27
        ClayTexC(5, 4) = 0.4
        ClayTexC(5, 5) = 0.4

        SandTexC(6, 1) = 0.2
        SandTexC(6, 2) = 0.2
        SandTexC(6, 3) = 0
        SandTexC(6, 4) = 0
        SandTexC(6, 5) = 0.2

        ClayTexC(6, 1) = 0.4
        ClayTexC(6, 2) = 0.27
        ClayTexC(6, 3) = 0.27
        ClayTexC(6, 4) = 0.4
        ClayTexC(6, 5) = 0.4

        SandTexC(7, 1) = 0.23
        SandTexC(7, 2) = 0.5
        SandTexC(7, 3) = 0.2
        SandTexC(7, 4) = 0.08
        SandTexC(7, 5) = 0
        SandTexC(7, 6) = 0
        SandTexC(7, 7) = 0.23

        ClayTexC(7, 1) = 0.27
        ClayTexC(7, 2) = 0
        ClayTexC(7, 3) = 0
        ClayTexC(7, 4) = 0.12
        ClayTexC(7, 5) = 0.12
        ClayTexC(7, 6) = 0.27
        ClayTexC(7, 7) = 0.27

        SandTexC(8, 1) = 0.08
        SandTexC(8, 2) = 0.2
        SandTexC(8, 3) = 0
        SandTexC(8, 4) = 0
        SandTexC(8, 5) = 0.08

        ClayTexC(8, 1) = 0.12
        ClayTexC(8, 2) = 0
        ClayTexC(8, 3) = 0
        ClayTexC(8, 4) = 0.12
        ClayTexC(8, 5) = 0.12

        SandTexC(9, 1) = 0.65
        SandTexC(9, 2) = 0.8
        SandTexC(9, 3) = 0.52
        SandTexC(9, 4) = 0.45
        SandTexC(9, 5) = 0.45
        SandTexC(9, 6) = 0.65

        ClayTexC(9, 1) = 0.35
        ClayTexC(9, 2) = 0.2
        ClayTexC(9, 3) = 0.2
        ClayTexC(9, 4) = 0.27
        ClayTexC(9, 5) = 0.35
        ClayTexC(9, 6) = 0.35

        SandTexC(10, 1) = 0.8
        SandTexC(10, 2) = 0.85
        SandTexC(10, 3) = 0.7
        SandTexC(10, 4) = 0.5
        SandTexC(10, 5) = 0.43
        SandTexC(10, 6) = 0.52
        SandTexC(10, 7) = 0.52
        SandTexC(10, 8) = 0.8

        ClayTexC(10, 1) = 0.2
        ClayTexC(10, 2) = 0.15
        ClayTexC(10, 3) = 0
        ClayTexC(10, 4) = 0
        ClayTexC(10, 5) = 0.07
        ClayTexC(10, 6) = 0.07
        ClayTexC(10, 7) = 0.2
        ClayTexC(10, 8) = 0.2

        SandTexC(11, 1) = 0.9
        SandTexC(11, 2) = 1
        SandTexC(11, 3) = 0.87
        SandTexC(11, 4) = 0.9

        ClayTexC(11, 1) = 0.1
        ClayTexC(11, 2) = 0
        ClayTexC(11, 3) = 0
        ClayTexC(11, 4) = 0.1

        SandTexC(12, 1) = 0.85
        SandTexC(12, 2) = 0.9
        SandTexC(12, 3) = 0.87
        SandTexC(12, 4) = 0.7
        SandTexC(12, 5) = 0.85

        ClayTexC(12, 1) = 0.15
        ClayTexC(12, 2) = 0.1
        ClayTexC(12, 3) = 0
        ClayTexC(12, 4) = 0
        ClayTexC(12, 5) = 0.15

#End Region

        For USDAClass As eUSDAclass = eUSDAclass.clay To eUSDAclass.loamySand
            For j = 1 To nPoly(USDAClass)
                Sand(j) = SandTexC(USDAClass, j)
                Clay(j) = ClayTexC(USDAClass, j)
            Next j
            If inRegion(Sand, Clay, nPoly(USDAClass), xSand, yClay) = 1 Then
                Return USDAClass
            End If
        Next

        Return eUSDAclass.not_def

    End Function

    Private Shared Function inRegion(xPoly() As Single, yPoly() As Single, nPoly As Integer, xp As Single, yp As Single) As Integer


#Region "    Definitions"

        Const eps As Single = 0.000001

        Dim M As Single()
        Dim y As Single()
        Dim upper As Integer()
        Dim lower As Integer()
        Dim xMax As Single
        Dim xMin As Single
        Dim yMax As Single
        Dim yMin As Single
        Dim i As Integer
        Dim k As Integer
        Dim l As Integer
        Dim nc As Integer
        Dim dup As Single
        Dim dlo As Single
        Dim dupk As Integer
        Dim dlok As Integer

        ReDim M(nPoly)
        ReDim y(nPoly)
        ReDim upper(nPoly)
        ReDim lower(nPoly)

        Dim out As Integer

#End Region

        xp /= 100
        yp /= 100


        xMax = xPoly(1)
        xMin = xPoly(1)
        yMax = yPoly(1)
        yMin = yPoly(1)

        out = 0

        k = 1

        For i = 1 To nPoly
            If (xPoly(i) > xMax) Then xMax = xPoly(i)
            If (xPoly(i) < xMin) Then xMin = xPoly(i)
            If (yPoly(i) > yMax) Then yMax = yPoly(i)
            If (yPoly(i) < yMin) Then yMin = yPoly(i)
        Next i

        If (xp > xMax OrElse xp < xMin OrElse yp > yMax OrElse yp < yMin) Then Return out
        out = 2

        For i = 1 To nPoly

            If (Math.Abs(xPoly(i) - xp) < eps And Math.Abs(yPoly(i) - yp) < eps) Then

                out = 1
                k = k + 1

                Exit For

            End If

            If (i > 1) Then

                If ((xPoly(i - 1) <= xp AndAlso xPoly(i) >= xp) OrElse (xPoly(i - 1) >= xp AndAlso xPoly(i) <= xp)) Then
                    If (xPoly(i - 1) <= xp AndAlso xPoly(i) >= xp) Then upper(k) = 1
                    If (xPoly(i - 1) >= xp AndAlso xPoly(i) <= xp) Then lower(k) = 1
                    If (xPoly(i - 1) - xPoly(i) <> 0) Then
                        M(k) = (yPoly(i - 1) - yPoly(i)) / (xPoly(i - 1) - xPoly(i))
                        y(k) = M(k) * (xp - xPoly(i - 1)) + yPoly(i - 1)
                        If (Math.Abs(y(k) - yp) < eps) Then
                            out = 1
                            Return out
                        End If
                        k = k + 1
                    Else
                        If ((yPoly(i - 1) <= yp AndAlso yPoly(i) >= yp) OrElse (yPoly(i - 1) >= yp AndAlso yPoly(i) <= yp)) Then
                            out = 1
                            Return out
                        End If
                    End If
                End If
            End If

        Next i

        If (k > 2) Then
            dup = 1.0E+21
            dlo = 1.0E+21
            dupk = -1
            dlok = -1
            For l = 1 To k
                If (y(l) >= yp) Then
                    If (y(l) - yp < dup) Then
                        dup = y(l) - yp
                        dupk = l
                    End If
                End If
                If (y(l) <= yp) Then
                    If (yp - y(l) < dlo) Then
                        dlo = yp - y(l)
                        dlok = l
                    End If
                End If
            Next l

            If (dupk > 0 AndAlso dlok > 0) Then ' upper and lower y-value on polygon present
                If (upper(dupk) = 1 AndAlso yp <= y(dupk) AndAlso lower(dlok) = 1 AndAlso yp >= y(dlok)) Then
                    out = 1
                End If
            End If
        End If

        Return out

    End Function

    Public Enum eSoilTexture
        sand
        silt
        clay
    End Enum


End Class




<TypeConverter(GetType(propGridConverter))>
<Serializable>
Public Class soilTexture

    Public Sub New()
        init()
    End Sub

    Private Sub init()

        _sand = 10
        _clay = 60
        _silt = 30

        toCalculate = _toCalculate

    End Sub

    Private Sub soilTexture_update() Handles Me.update

        hypres =
            getHypresClass(
            sand:=sand,
            silt:=silt,
            clay:=clay)

        USDA =
            getUSDAClass(
            xSand:=sand,
            yClay:=clay)

    End Sub

    Private _hypres As eHypresClass
    Private _USDA As eUSDAclass

    Private _sand As Integer
    Private _silt As Integer
    Private _clay As Integer

    Private _toCalculate As eSoilTexture = eSoilTexture.clay

    Private _om As Double = Double.NaN
    Private _oc As Double = Double.NaN

    Public Const catTexture As String = "01  Texture"

#Region "    Texture"

    ''' <summary>
    ''' Sand   0.05  - 2.00 mm
    ''' Coarsest mineral particles
    ''' feel gritty when rubbed between your fingers
    ''' </summary>
    ''' <returns></returns>
    <DisplayName(
        "Sand")>
    <Description(
    "0.05  - 2.00 mm, coarsest mineral particles" & vbCrLf &
    "feel gritty when rubbed between your fingers")>
    <Category(catTexture)>
    <Editor(GetType(nudValueEditorInteger), GetType(UITypeEditor))>
    <RefreshProperties(RefreshProperties.All)>
    <[ReadOnly](False)>
    Public Property sand As Integer
        Get
            Return _sand
        End Get
        Set

            Select Case toCalculate

                Case eSoilTexture.clay

                    If Value + _silt > 100 Then _silt = 100 - Value

                    _clay = 100 - (Value + _silt)
                    _sand = Value

                Case eSoilTexture.silt

                    If Value + _clay > 100 Then _clay = 100 - Value

                    _silt = 100 - (Value + _clay)
                    _sand = Value

                Case eSoilTexture.sand

            End Select

            RaiseEvent update()

        End Set
    End Property

    ''' <summary>
    ''' Silt  0.002 - 0.05 mm
    ''' feel similar to flour when dry
    ''' </summary>
    ''' <returns></returns>
    <DisplayName(
        "Silt")>
    <Description(
    "0.002 - 0.05 mm" & vbCrLf &
    "Feel similar to flour when dry")>
    <Category(catTexture)>
    <Editor(GetType(nudValueEditorInteger), GetType(UITypeEditor))>
    <RefreshProperties(RefreshProperties.All)>
    <[ReadOnly](False)>
    Public Property silt As Integer
        Get
            Return _silt
        End Get
        Set

            Select Case toCalculate

                Case eSoilTexture.sand

                    If Value + _clay > 100 Then _clay = 100 - Value

                    _sand = 100 - (Value + _clay)
                    _silt = Value

                Case eSoilTexture.clay

                    If Value + _sand > 100 Then _sand = 100 - Value

                    _clay = 100 - (Value + _sand)
                    _silt = Value

                Case eSoilTexture.silt


            End Select

            RaiseEvent update()

        End Set
    End Property

    ''' <summary>
    ''' Clay  less than 0.002 mm
    ''' feel sticky in your fingers when wet and clump to the point
    ''' that you can't see an individual particle without a microscope
    ''' </summary>
    ''' <returns></returns>
    <DisplayName(
        "Clay")>
    <Description(
    "<0.002 mm, feel sticky in your fingers when wet and clump to the point " & vbCrLf &
    "that you can't see an individual particle without a microscope")>
    <Category(catTexture)>
    <Editor(GetType(nudValueEditorInteger), GetType(UITypeEditor))>
    <RefreshProperties(RefreshProperties.All)>
    <[ReadOnly](True)>
    Public Property clay As Integer
        Get
            Return _clay
        End Get
        Set

            Select Case toCalculate

                Case eSoilTexture.sand

                    If Value + _silt > 100 Then _silt = 100 - Value

                    _sand = 100 - (Value + _silt)
                    _clay = Value

                Case eSoilTexture.silt

                    If Value + _sand > 100 Then _sand = 100 - Value

                    _silt = 100 - (Value + _sand)
                    _clay = Value

                Case eSoilTexture.clay

            End Select

            RaiseEvent update()

        End Set
    End Property

    ''' <summary>
    ''' 'Unknown' part to the texture
    ''' to be calculated by 100 - (texture a + b)
    ''' </summary>
    ''' <returns></returns>
    <Description(
    "'Unknown' part to the texture" & vbCrLf &
    "to be calculated by 100 - (texture a + b)")>
    <DisplayName(
    " ... to calculate")>
    <Category(catTexture)>
    Public Property toCalculate As eSoilTexture
        Get
            Return _toCalculate
        End Get
        Set

            _toCalculate = Value

            Select Case _toCalculate

                Case eSoilTexture.clay

                    setReadOnlyProperty(propertyName:="sand", [readOnly]:=False)
                    setReadOnlyProperty(propertyName:="silt", [readOnly]:=False)
                    setReadOnlyProperty(propertyName:="clay", [readOnly]:=True)

                Case eSoilTexture.silt


                    setReadOnlyProperty(propertyName:="sand", [readOnly]:=False)
                    setReadOnlyProperty(propertyName:="silt", [readOnly]:=True)
                    setReadOnlyProperty(propertyName:="clay", [readOnly]:=False)

                Case eSoilTexture.sand


                    setReadOnlyProperty(propertyName:="sand", [readOnly]:=True)
                    setReadOnlyProperty(propertyName:="silt", [readOnly]:=False)
                    setReadOnlyProperty(propertyName:="clay", [readOnly]:=False)

            End Select

            RaiseEvent update()

        End Set
    End Property


    ''' <summary>
    ''' Switch for the Attribute 'read only'
    ''' </summary>
    ''' <param name="propertyName"></param>
    ''' Name of the property to change
    ''' <param name="readOnly"></param>
    ''' read only or not , true or false
    ''' <remarks></remarks>
    Private Sub setReadOnlyProperty(ByVal propertyName As String,
                                    ByVal [readOnly] As Boolean)

        Try

            Dim descriptor As PropertyDescriptor =
                           TypeDescriptor.GetProperties([GetType]())(propertyName)
            Dim attribute As ReadOnlyAttribute =
                DirectCast(descriptor.Attributes(GetType(ReadOnlyAttribute)), ReadOnlyAttribute)



            Dim fieldToChange As Reflection.FieldInfo =
                        attribute.[GetType]().GetField("isReadOnly",
                                                    System.Reflection.BindingFlags.NonPublic Or
                                                    System.Reflection.BindingFlags.Instance)

            fieldToChange.SetValue(attribute, [readOnly])

        Catch ex As Exception
            Console.WriteLine(
                "Error setting '[readOnly]' property for " &
                propertyName & " to " &
                [readOnly].ToString & vbCrLf &
                ex.Message)
        End Try

    End Sub

#End Region

    Public Const catClassifications As String = "02 Classifications"

#Region "    Classifications"

    <RefreshProperties(RefreshProperties.All)>
    <[ReadOnly](True)>
    <Category(catClassifications)>
    Public Property hypres As eHypresClass
        Get
            Return _hypres
        End Get
        Set
            _hypres = Value
        End Set
    End Property

    <RefreshProperties(RefreshProperties.All)>
    <[ReadOnly](True)>
    <Category(catClassifications)>
    Public Property USDA As eUSDAclass
        Get
            Return _USDA
        End Get
        Set
            _USDA = Value
        End Set
    End Property

#End Region


    Public Const catOmOc As String = "03 Organic matter/carbon"

#Region "    Organic matter/carbon"

    Private Const factorOmOc As Double = 1.724
    Private Const digits As Integer = 2

    <Category(catOmOc)>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(dblConv))>
    Public Property om As Double
        Get
            Return _om
        End Get
        Set

            _om = Value
            _oc = Math.Round(_om / factorOmOc, digits:=digits)

        End Set
    End Property


    <Category(catOmOc)>
    <RefreshProperties(RefreshProperties.All)>
    <TypeConverter(GetType(dblConv))>
    Public Property oc As Double
        Get
            Return _oc
        End Get
        Set

            _oc = Value
            _om = Math.Round(_oc * factorOmOc, digits:=digits)

        End Set
    End Property

#End Region


    Public Event update()


End Class

