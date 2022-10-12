Imports Microsoft.Office.Interop.Excel

Friend Module Util
    ''' <summary>
    ''' 運賃 (2トン車).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Fare(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            DctVal(xlApp, "BA151", choosen)
        End If
    End Sub

    ''' <summary>
    ''' 下端 (D13).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="truck2Ton">2 ton truck.</param>
    Friend Sub LwrEnd(xlApp As Application, truck2Ton As Double)
        If Not truck2Ton = 1 Then
            PubDVal(xlApp, "BA92", DtlDInp(vbTab & "5500: "))
            PubDVal(xlApp, "BA93", DtlDInp(vbTab & "5000: "))
        End If
        PubDVal(xlApp, "BA94", DtlDInp(vbTab & "4500: "))
        PubDVal(xlApp, "BA95", DtlDInp(vbTab & "4000: "))
        PubDVal(xlApp, "BA96", DtlDInp(vbTab & "3500: "))
        PubDVal(xlApp, "BA97", DtlDInp(vbTab & "3000: "))
        PubDVal(xlApp, "BA98", DtlDInp(vbTab & "2500: "))
        PubDVal(xlApp, "BA99", DtlDInp(vbTab & "2000: "))
        PubDVal(xlApp, "BA100", DtlDInp(vbTab & "1500: "))
        PubDVal(xlApp, "BA101", DtlDInp(vbTab & "1000: "))
    End Sub

    ''' <summary>
    ''' 端部 (D10).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="truck2Ton">2 ton truck.</param>
    Friend Sub Edge(xlApp As Application, truck2Ton As Double)
        If Not truck2Ton = 1 Then
            PubDVal(xlApp, "BA79", DtlDInp(vbTab & "5500: "))
            PubDVal(xlApp, "BA80", DtlDInp(vbTab & "5000: "))
        End If
        PubDVal(xlApp, "BA81", DtlDInp(vbTab & "4500: "))
        PubDVal(xlApp, "BA82", DtlDInp(vbTab & "4000: "))
        PubDVal(xlApp, "BA83", DtlDInp(vbTab & "3500: "))
        PubDVal(xlApp, "BA84", DtlDInp(vbTab & "3000: "))
        PubDVal(xlApp, "BA85", DtlDInp(vbTab & "2500: "))
        PubDVal(xlApp, "BA86", DtlDInp(vbTab & "2000: "))
        PubDVal(xlApp, "BA87", DtlDInp(vbTab & "1500: "))
        PubDVal(xlApp, "BA88", DtlDInp(vbTab & "1000: "))
    End Sub

    ''' <summary>
    ''' 深基礎 (D16).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    ''' <param name="truck2Ton">2 ton truck.</param>
    Friend Sub DeepFnd(xlApp As Application, choosen As Double, truck2Ton As Double)
        If choosen = 1 Then
            If Not truck2Ton = 1 Then
                PubDVal(xlApp, "BA102", DtlDInp(vbTab & "5500: "))
                PubDVal(xlApp, "BA103", DtlDInp(vbTab & "5000: "))
            End If
            PubDVal(xlApp, "BA104", DtlDInp(vbTab & "4500: "))
            PubDVal(xlApp, "BA105", DtlDInp(vbTab & "4000: "))
            PubDVal(xlApp, "BA106", DtlDInp(vbTab & "3500: "))
            PubDVal(xlApp, "BA107", DtlDInp(vbTab & "3000: "))
            PubDVal(xlApp, "BA108", DtlDInp(vbTab & "2500: "))
            PubDVal(xlApp, "BA109", DtlDInp(vbTab & "2000: "))
            PubDVal(xlApp, "BA110", DtlDInp(vbTab & "1500: "))
            PubDVal(xlApp, "BA111", DtlDInp(vbTab & "1000: "))
        End If
    End Sub

    ''' <summary>
    ''' コーナー.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub JtCor(xlApp As Application)
        PubDVal(xlApp, "BA114", DtlDInp(vbTab & "D16: "))
        PubDVal(xlApp, "BA113", DtlDInp(vbTab & "D13: "))
        PubDVal(xlApp, "BA112", DtlDInp(vbTab & "D10: "))
    End Sub

    ''' <summary>
    ''' 深基礎用端部スラブ (D10).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub EndSlabForDeepFnd(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDModVal(xlApp, "116", "600×450　　フック付", 0.7, DtlDInpDesc(vbTab & "600×450 ", "[0.7]" & vbTab))
            PubDVal(xlApp, "BA118", DtlDInp(vbTab & "600×350" & vbTab & vbTab & ": "))
            PubDModVal(xlApp, "117", "600×250　　フック付", 0.6, DtlDInpDesc(vbTab & "600×250 ", "[0.6]" & vbTab))
        End If
    End Sub

    ''' <summary>
    ''' Ｕ型 (D16).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub UType(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDModVal(xlApp, "119", "D16", "（コノ字型）", "900×280×900", 3.4, DtlDInpDesc(vbTab & "900×280×900 ", "[3.4]" & vbTab))
            PubDModVal(xlApp, "121", "900×180×900", 3.2, DtlDInpDesc(vbTab & "900×180×900 ", "[3.2]" & vbTab))
            PubDModVal(xlApp, "120", "D16", "（Ｕノ字型）", "900×80×900", 3.1, DtlDInpDesc(vbTab & "900× 80×900 ", "[3.1]" & vbTab))
        End If
    End Sub

    ''' <summary>
    ''' ハンチ (H250).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    Friend Sub Haunch(xlApp As Application, choosen As Double)
        If choosen = 1 Then
            PubDModVal(xlApp, "123", "900×曲（H250）×900", 3.5, DtlDInpDesc(vbTab & "D16 ", "[3.5]" & vbTab))
            PubDModVal(xlApp, "122", "D13", "750×曲（H250）×750", 2, DtlDInpDesc(vbTab & "D13 ", "[2.0]" & vbTab))
        End If
    End Sub

    ''' <summary>
    ''' 電気温水器.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="value">Value.</param>
    Friend Sub ElecWtrHtr(xlApp As Application, value As Double)
        If value > 0 Then
            DctVal(xlApp, "BA31", value)
        Else
            ClrVal(xlApp, "BA31")
        End If
    End Sub

    ''' <summary>
    ''' スリーブ.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="value">Value.</param>
    Friend Sub Sleeve(xlApp As Application, value As Double)
        If value > 0 Then
            DctVal(xlApp, "BA30", value)
            DctVal(xlApp, "BA115", value)
        End If
    End Sub

    ''' <summary>
    ''' 予備材 (D16).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    ''' <param name="truck2Ton">2 ton truck.</param>
    Friend Sub SprMatl(xlApp As Application, choosen As Double, truck2Ton As Double)
        If Not truck2Ton = 1 Then
            DctVal(xlApp, "BA90", 2)
            DctVal(xlApp, "BA91", 2)
            If choosen = 1 Then
                DctVal(xlApp, "BA89", 2)
            End If
        Else
            PubDModVal(xlApp, "90", "4500", 4.5, 3)
            PubDModVal(xlApp, "91", "4500", 2.6, 3)
            If choosen = 1 Then
                PubDModVal(xlApp, "89", "4500", 7.1, 3)
            End If
        End If
    End Sub

    ''' <summary>
    ''' スラブ直 (D13).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    ''' <param name="truck2Ton">2 ton truck.</param>
    Friend Sub SlabStr(xlApp As Application, choosen As Double, truck2Ton As Double)
        If choosen = 1 Then
            If Not truck2Ton = 1 Then
                PubDVal(xlApp, "BA57", DtlDInp(vbTab & "5500" & vbTab & vbTab & ":"))
                PubDVal(xlApp, "BA58", DtlDInp(vbTab & "5000" & vbTab & vbTab & ":"))
            End If
            PubDVal(xlApp, "BA59", DtlDInp(vbTab & "4500" & vbTab & vbTab & ":"))
            PubDVal(xlApp, "BA60", DtlDInp(vbTab & "4000" & vbTab & vbTab & ":"))
            PubDVal(xlApp, "BA61", DtlDInp(vbTab & "3500" & vbTab & vbTab & ":"))
            PubDVal(xlApp, "BA62", DtlDInp(vbTab & "3000" & vbTab & vbTab & ":"))
            PubDVal(xlApp, "BA63", DtlDInp(vbTab & "2500" & vbTab & vbTab & ":"))
            PubDVal(xlApp, "BA64", DtlDInp(vbTab & "2000" & vbTab & vbTab & ":"))
            PubDVal(xlApp, "BA65", DtlDInp(vbTab & "1500" & vbTab & vbTab & ":"))
            PubDModVal(xlApp, "66", "1300", 1.4, DtlDInpDesc(vbTab & "1300 ", "[1.4]" & vbTab))
            PubDVal(xlApp, "BA67", DtlDInp(vbTab & "1000" & vbTab & vbTab & ":"))
            PubDVal(xlApp, "BA68", DtlDInp(vbTab & " 900" & vbTab & vbTab & ":"))
        End If
    End Sub

    ''' <summary>
    ''' スラブＬ型 (D13).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    ''' <param name="truck2Ton">2 ton truck.</param>
    Friend Sub SlabLType(xlApp As Application, choosen As Double, truck2Ton As Double)
        If choosen = 1 Then
            If Not truck2Ton = 1 Then
                PubDVal(xlApp, "BA36", DtlDInp(vbTab & "180×5320: "))
                PubDVal(xlApp, "BA37", DtlDInp(vbTab & "180×4820: "))
            End If
            PubDVal(xlApp, "BA38", DtlDInp(vbTab & "180×4320: "))
            PubDVal(xlApp, "BA39", DtlDInp(vbTab & "180×3820: "))
            PubDVal(xlApp, "BA40", DtlDInp(vbTab & "180×3320: "))
            PubDVal(xlApp, "BA41", DtlDInp(vbTab & "180×2820: "))
            PubDVal(xlApp, "BA42", DtlDInp(vbTab & "180×2320: "))
            PubDVal(xlApp, "BA43", DtlDInp(vbTab & "180×1820: "))
            PubDVal(xlApp, "BA44", DtlDInp(vbTab & "180×1320: "))
            PubDVal(xlApp, "BA45", DtlDInp(vbTab & "180× 820: "))
        End If
    End Sub

    ''' <summary>
    ''' スラブフック型 (D13).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    ''' <param name="truck2Ton">2 ton truck.</param>
    Friend Sub SlabHookType(xlApp As Application, choosen As Double, truck2Ton As Double)
        If choosen = 1 Then
            If Not truck2Ton = 1 Then
                PubDVal(xlApp, "BA46", DtlDInp(vbTab & "95×5405" & vbTab & vbTab & ":"))
                PubDVal(xlApp, "BA47", DtlDInp(vbTab & "95×4905" & vbTab & vbTab & ":"))
            End If
            PubDVal(xlApp, "BA48", DtlDInp(vbTab & "95×4405" & vbTab & vbTab & ":"))
            PubDVal(xlApp, "BA49", DtlDInp(vbTab & "95×3905" & vbTab & vbTab & ":"))
            PubDVal(xlApp, "BA50", DtlDInp(vbTab & "95×3405" & vbTab & vbTab & ":"))
            PubDVal(xlApp, "BA51", DtlDInp(vbTab & "95×2905" & vbTab & vbTab & ":"))
            PubDVal(xlApp, "BA52", DtlDInp(vbTab & "95×2405" & vbTab & vbTab & ":"))
            PubDVal(xlApp, "BA53", DtlDInp(vbTab & "95×1905" & vbTab & vbTab & ":"))
            PubDVal(xlApp, "BA54", DtlDInp(vbTab & "95×1405" & vbTab & vbTab & ":"))
            PubDVal(xlApp, "BA55", DtlDInp(vbTab & "95× 905" & vbTab & vbTab & ":"))
            PubDModVal(xlApp, "56", "490(両フック)", 1.2, DtlDInpDesc(vbTab & "95×490×95 ", "[1.2]" & vbTab))
        End If
    End Sub

    ''' <summary>
    ''' スラブ補強Ｌ型 (D10).
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    ''' <param name="choosen">Selection.</param>
    ''' <param name="truck2Ton">2 ton truck.</param>
    Friend Sub SlabReinfLType(xlApp As Application, choosen As Double, truck2Ton As Double)
        If choosen = 1 Then
            If Not truck2Ton = 1 Then
                PubDVal(xlApp, "BA69", DtlDInp(vbTab & "180×5320: "))
                PubDVal(xlApp, "BA70", DtlDInp(vbTab & "180×4820: "))
            End If
            PubDVal(xlApp, "BA71", DtlDInp(vbTab & "180×4320: "))
            PubDVal(xlApp, "BA72", DtlDInp(vbTab & "180×3820: "))
            PubDVal(xlApp, "BA73", DtlDInp(vbTab & "180×3320: "))
            PubDVal(xlApp, "BA74", DtlDInp(vbTab & "180×2820: "))
            PubDVal(xlApp, "BA75", DtlDInp(vbTab & "180×2320: "))
            PubDVal(xlApp, "BA76", DtlDInp(vbTab & "180×1820: "))
            PubDVal(xlApp, "BA77", DtlDInp(vbTab & "180×1320: "))
            PubDVal(xlApp, "BA78", DtlDInp(vbTab & "180× 820: "))
        End If
    End Sub

    ''' <summary>
    ''' 副資材リスト.
    ''' </summary>
    ''' <param name="xlApp">Excel Application.</param>
    Friend Sub Parts(xlApp As Application)
        Dim name = $"{DtlSInp(vbTab & "邸名" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & ": ")}様邸"
        DctVal(xlApp, "BJ13", name)
        CType(xlApp.ActiveSheet, Worksheet).Name = name
        PubSVal(xlApp, vbTab & "邸名コード" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & ": ", "AD6")
        PubSVal(xlApp, vbTab & "納品日" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & ": ", "BO3")
        PubSVal(xlApp, vbTab & "住所" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & ": ", "BJ14")
        Dim curingShRingTree = DtlDInpDesc(vbTab & "養生シート輪木 (ｾｯﾄ)", vbTab & vbTab & vbTab & vbTab & "[3.6×5.4]" & vbTab & vbTab)
        If curingShRingTree > 0 Then
            DctVal(xlApp, "BA144", curingShRingTree)
        Else
            DctVal(xlApp, "BA144", 1)
            ClrVal(xlApp, "BF144")
            ClrVal(xlApp, "CB144")
        End If
        PubDVal(xlApp, "BA141", DtlDInpDesc(vbTab & "給水用スリーブホルダー・D10用 (箱)", vbTab & vbTab & "[50ﾊﾟｲ]" & vbTab & vbTab & vbTab))
        PubDVal(xlApp, "BA142", DtlDInpDesc(vbTab & "排水用スリーブホルダー・D10用 (箱)", vbTab & vbTab & "[50ﾊﾟｲ 75ﾊﾟｲ]" & vbTab & vbTab))
        PubDVal(xlApp, "BA133", DtlDInpDesc(vbTab & "カットスクリュー・Ⅱ (袋)", vbTab & vbTab & vbTab & "[M12ﾖｳ 50ｺｲﾘ]" & vbTab & vbTab))
        PubDVal(xlApp, "BA147", DtlDInpDesc(vbTab & "スペーサーブロック (個)", vbTab & vbTab & vbTab & vbTab & "[H60・70・80]" & vbTab & vbTab))
        PubDVal(xlApp, "BA134", DtlDInpDesc(vbTab & "PCa基礎梁 敷き鉄板 (枚)", vbTab & vbTab & vbTab & vbTab & "[1.6t×200×300]" & vbTab & vbTab))
        PubDVal(xlApp, "BA135", DtlDInpDesc(vbTab & "PCa基礎梁 敷き鉄板 (枚)", vbTab & vbTab & vbTab & vbTab & "[3.2t×200×300]" & vbTab & vbTab))
        PubDVal(xlApp, "BA136", DtlDInpDesc(vbTab & "PCa基礎梁 土台用アンカー (本)", vbTab & vbTab & vbTab & "[M12×147]" & vbTab & vbTab))
        PubDVal(xlApp, "BA138", DtlDInpDesc(vbTab & "PCa基礎梁 ホールダウンカンカーボルト (本)", vbTab & "[M12×170]" & vbTab & vbTab))
        PubDVal(xlApp, "BA143", DtlDInpDesc(vbTab & "PCa基礎梁 BF用柱脚両ネジボルト (本)", vbTab & vbTab & "[M16×70]" & vbTab & vbTab))
        PubDVal(xlApp, "BA137", DtlDInpDesc(vbTab & "PCa基礎梁 ホールダウンカンカーボルト (本)", vbTab & "[M12×387]" & vbTab & vbTab))
        PubDVal(xlApp, "BA139", DtlDInpDesc(vbTab & "マグネット差筋アンカーD13 (ｾｯﾄ)", vbTab & vbTab & vbTab & "[直]" & vbTab & vbTab & vbTab))
        PubDVal(xlApp, "BA140", DtlDInpDesc(vbTab & "マグネット差筋アンカーD13 (ｾｯﾄ)", vbTab & vbTab & vbTab & "[曲]" & vbTab & vbTab & vbTab))
        PubDVal(xlApp, "BA146", DtlDInp(vbTab & "カットスクリューⅡ・専用ピット (個)" & vbTab & vbTab & vbTab & vbTab & vbTab & ": "))
        PubDVal(xlApp, "BA149", DtlDInpDesc(vbTab & "グリッパーM12アンカー用D16 (箱)", vbTab & vbTab & vbTab & "[TG1216D 50ｺ/ﾊｺ]" & vbTab))
        PubDVal(xlApp, "BA145", DtlDInpDesc(vbTab & "結束線メッキ450 (ｹｰｽ)", vbTab & vbTab & vbTab & vbTab & "[20kg]" & vbTab & vbTab & vbTab))
        PubDVal(xlApp, "BA150", DtlDInpDesc(vbTab & "Ｕボルト (ｾｯﾄ)", vbTab & vbTab & vbTab & vbTab & vbTab & "[M8]" & vbTab & vbTab & vbTab))
        ' Extend
        PubDVal(xlApp, "BA148", DtlDInpDesc(vbTab & "樹脂スペーサー (個)", vbTab & vbTab & vbTab & vbTab & "[70×80 50ｺ/ﾊｺ]" & vbTab & vbTab))
    End Sub
End Module
