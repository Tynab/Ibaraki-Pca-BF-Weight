Imports Microsoft.Office.Interop.Excel

Friend Module Service
    ''' <summary>
    ''' Weight Ibaraki Pca BF.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub WtIbarakiPcaBF(xlApp As Application)
        ' Fare
        Fare(xlApp, HdrYNQ(vbTab & vbTab & "運賃 (2トン車): "))
        ' Slab hook type
        SlabHookType(xlApp, HdrYNQ(vbTab & vbTab & "スラブフック型 (D13): "))
        'Slab L type
        SlabLType(xlApp, HdrYNQ(vbTab & vbTab & "スラブＬ型 (D13): "))
        ' Slab straight
        SlabStr(xlApp, HdrYNQ(vbTab & vbTab & "スラブ直 (D13): "))
        ' Slab reinforcement hook type
        SlabReinfHookType(xlApp, HdrYNQ(vbTab & vbTab & "スラブ補強フック型 (D10): "))
        ' Slab reinforcement straight
        SlabReinfStr(xlApp, HdrYNQ(vbTab & vbTab & "スラブ補強直 (D10): "))
        ' Lower end d13
        PrefWarn(vbTab & vbTab & "下端 (D13)")
        LwrEndD13(xlApp)
        ' Lower end d16
        LwrEndD16(xlApp, HdrYNQ(vbTab & vbTab & "下端 (D16): "))
        ' Edge
        Edge(xlApp, HdrYNQ(vbTab & vbTab & "端部 (D10): "))
        ' Sleeve
        Sleeve(xlApp, HdrDInp(vbTab & vbTab & "スリーブ: "))
        ' Corner joint
        PrefWarn(vbTab & vbTab & "コーナー")
        JtCor(xlApp)
        ' Dirt floor scissors
        PubDVal(xlApp, "BA138", HdrDInp(vbTab & vbTab & "土間用さし: "))
        ' U type
        PubDModVal(xlApp, "126", "（Ｕノ字型）", "900×80×900", 3.1, HdrDInp(vbTab & vbTab & "Ｕ型 (D16): "))
        ' Haunch
        Haunch(xlApp, HdrYNQ(vbTab & vbTab & "ハンチ (H250): "))
        ' End slab for deep foundation
        PubDModVal(xlApp, "129", "650×250　　フック付", 0.6, HdrYNQ(vbTab & vbTab & "深基礎用端部スラブ (D10): "))
        ' Electric water heater
        ElecWtrHtr(xlApp, HdrDInp(vbTab & vbTab & "電気温水器: "))
        ' Parts
        PrefWarn(vbTab & vbTab & "副資材リスト")
        Parts(xlApp)
    End Sub
End Module
