Imports Microsoft.Office.Interop.Excel

Friend Module Service
    ''' <summary>
    ''' Weight Ibaraki Pca BF.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub WtIbarakiPcaBF(xlApp As Application)
        ' Fare
        Dim truck2Ton = HdrYNQ(vbTab & vbTab & "運賃 (2トン車): ")
        Fare(xlApp, truck2Ton)
        ' Lower end
        PrefWarn(vbTab & vbTab & "下端 (D13)")
        LwrEnd(xlApp, truck2Ton)
        ' Edge
        Edge(xlApp, HdrYNQ(vbTab & vbTab & "端部 (D10): "))
        ' Deep foundation
        DeepFnd(xlApp, HdrYNQ(vbTab & vbTab & "深基礎 (D16): "), truck2Ton)
        ' Corner joint
        PrefWarn(vbTab & vbTab & "コーナー")
        JtCor(xlApp)
        ' Dirt floor scissors
        PubDVal(xlApp, "BA129", HdrDInp(vbTab & vbTab & "土間用さし: "))
        ' End slab for deep foundation
        EndSlabForDeepFnd(xlApp, HdrYNQ(vbTab & vbTab & "深基礎用端部スラブ (D10): "))
        ' U type
        UType(xlApp, HdrYNQ(vbTab & vbTab & "Ｕ型 (D16): "))
        ' Haunch
        Haunch(xlApp, HdrYNQ(vbTab & vbTab & "ハンチ (H250): "))
        ' Electric water heater
        ElecWtrHtr(xlApp, HdrDInp(vbTab & vbTab & "電気温水器: "))
        ' Sleeve
        Sleeve(xlApp, HdrDInp(vbTab & vbTab & "スリーブ: "))
        ' Spare material
        SprMatl(xlApp, HdrYNQ(vbTab & vbTab & "予備材 (D16): "), truck2Ton)
        ' Slab straight
        SlabStr(xlApp, HdrYNQ(vbTab & vbTab & "スラブ直 (D13): "), truck2Ton)
        'Slab L type
        SlabLType(xlApp, HdrYNQ(vbTab & vbTab & "スラブＬ型 (D13): "), truck2Ton)
        ' Slab hook type
        SlabHookType(xlApp, HdrYNQ(vbTab & vbTab & "スラブフック型 (D13): "), truck2Ton)
        ' Slab bending L type
        SlabReinfLType(xlApp, HdrYNQ(vbTab & vbTab & "スラブ補強Ｌ型 (D10): "), truck2Ton)
        ' Parts
        PrefWarn(vbTab & vbTab & "副資材リスト")
        Parts(xlApp)
    End Sub
End Module
