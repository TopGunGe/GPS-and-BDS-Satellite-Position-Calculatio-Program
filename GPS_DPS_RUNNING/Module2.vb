Option Explicit On
Imports System.Math
Imports GPS_DPS_RUNNING.Form2
Public Module Module2
    Public Sub Data_operater(Data, key_L)

        Dim n#, n0#, tn#, GM#, a#     ''1.mean angular speed
        Dim toe#, tk#, toc#, t#, a0#, a1#, a2#         ''2.normal time
        Dim Mk#, M0#                        ''3.mean anomaly
        Dim Ek#, e#, tEk#                   ''4.eccentric anomaly  E(initional) = M
        Dim Vk#                             ''5.true anomaly
        Dim fi0#, W#                       ''6.argument of ascending node
        Dim rk#                             ''7.orbital radius
        Dim Su#, Cuc#, Cus#
        Dim Sr#, Crc#, Crs#
        Dim Si#, Cic#, Cis#                 ''8.perturbed correction
        Dim uk#
        Dim ik#, i0#, IDOT#               ''9.corrected argument
        Dim xk#, yk#                        ''10.expressed in right angle coordinate
        Dim tvt#, OL#, OU#, we#             ''11.longitude of ascending node
        Dim XXk#, YYk#, ZZk#                ''12.expressed in spheric coordinate system

        Dim i%

        Dim year%, month%, day%, hour%, minute%
        Dim second#
        Dim IDOE#
        Dim sqrtA#
        Dim Ol0#
        Dim L2Data#, PS#, L2P#
        Dim ST_PRC#, ST_HEL#, TGD#, IDOC#
        Dim Send_Time#, Countin_h#, ps_1#, ps_2#

        ''*******************************取参数********************************''
        Dim prn
        prn = Val(Microsoft.VisualBasic.Left(Data(key_L + 1), 2))

        year = Val(Mid(Data(key_L + 1), 4, 2))
        month = Val(Mid(Data(key_L + 1), 7, 2))
        day = Val(Mid(Data(key_L + 1), 10, 2))
        hour = Val(Mid(Data(key_L + 1), 13, 2))
        minute = Val(Mid(Data(key_L + 1), 16, 2))
        second = Val(Mid(Data(key_L + 1), 18, 5))

        a0 = Val(Mid(Data(key_L + 1), 24, 19))
        a1 = Val(Mid(Data(key_L + 1), 43, 19))
        a2 = Val(Mid(Data(key_L + 1), 62, 19))

        ''************************************************************************''
        ''第2行

        IDOE = Val(Mid(Data(key_L + 2), 5, 19))
        Crs = Val(Mid(Data(key_L + 2), 24, 19))
        tn = Val(Mid(Data(key_L + 2), 43, 19))
        M0 = Val(Mid(Data(key_L + 2), 62, 19))
        ''************************************************************************''
        ''第3行

        Cuc = Val(Mid(Data(key_L + 3), 5, 19))
        e = Val(Mid(Data(key_L + 3), 24, 19))
        Cus = Val(Mid(Data(key_L + 3), 43, 19))
        sqrtA = Val(Mid(Data(key_L + 3), 62, 19))
        ''************************************************************************''
        ''第4行

        toe = Val(Mid(Data(key_L + 4), 5, 19))
        Cic = Val(Mid(Data(key_L + 4), 24, 19))
        Ol0 = Val(Mid(Data(key_L + 4), 43, 19))
        Cis = Val(Mid(Data(key_L + 4), 62, 19))

        ''************************************************************************''
        ''第5行
        i0 = Val(Mid(Data(key_L + 5), 5, 19))
        Crc = Val(Mid(Data(key_L + 5), 24, 19))
        W = Val(Mid(Data(key_L + 5), 43, 19))
        OU = Val(Mid(Data(key_L + 5), 62, 19))
        ''************************************************************************''
        ''第6行

        IDOT = Val(Mid(Data(key_L + 6), 5, 19))
        L2Data = Val(Mid(Data(key_L + 6), 24, 19))
        PS = Val(Mid(Data(key_L + 6), 43, 19))
        L2P = Val(Mid(Data(key_L + 6), 62, 19))
        ''************************************************************************''
        ''第7行

        ST_PRC = Val(Mid(Data(key_L + 7), 5, 19))
        ST_HEL = Val(Mid(Data(key_L + 7), 24, 19))
        TGD = Val(Mid(Data(key_L + 7), 43, 19))
        IDOC = Val(Mid(Data(key_L + 7), 62, 19))
        ''************************************************************************''
        ''第8行

        Send_Time = Val(Mid(Data(key_L + 8), 5, 19))
        Countin_h = Val(Mid(Data(key_L + 8), 24, 19))
        ps_1 = Val(Mid(Data(key_L + 8), 43, 19))
        ps_2 = Val(Mid(Data(key_L + 8), 62, 19))
        ''*************************输出各参数至frame1******************************************************''
        'Form2.ComboBox1.Items.Add(prn)
        'Combo1.listlndex = 0
        ''*************************输出各参数至text2*************************************''

        Form2.RichTextBox2.Text = prn
        'Form2.RichTextBox2.Text + "PRN =" + prn + vbCrLf
        'Form2.RichTextBox2.Text = Form2.RichTextBox2.Text & "----" & vbCrLf
        ' Form2.RichTextBox2.Text = Form2.RichTextBox2.Text & "y,m,d=" & " " & year & " " & month & " " & day & vbCrLf
        'Form2.RichTextBox2.Text = Form2.RichTextBox2.Text & "h,min,sec=" & hour & " " & minute & " " & second & " " & vbCrLf
        ' Form2.RichTextBox2.Text = Form2.RichTextBox2.Text & "a0=" & a0 & vbCrLf
        ' Form2.RichTextBox2.Text = Form2.RichTextBox2.Text & "a1=" & a1 & vbCrLf
        ' Form2.RichTextBox2.Text = Form2.RichTextBox2.Text & "a2=" & a2 & vbCrLf
        ''*************************************************************************''
        '   Form2.RichTextBox2.Text = Form2.RichTextBox2.Text & "IDOE=" & IDOE & vbCrLf
        '   Form2.RichTextBox2.Text = Form2.RichTextBox2.Text & "Crs=" & Crs & vbCrLf
        '   Form2.RichTextBox2.Text = Form2.RichTextBox2.Text & "tn=" & tn & vbCrLf
        '  Form2.RichTextBox2.Text = Form2.RichTextBox2.Text & "M0=" & M0 & vbCrLf
        ''*************************************************************************''
        '  Form2.RichTextBox2.Text = Form2.RichTextBox2.Text & "Cuc=" & Cuc & vbCrLf
        '  Form2.RichTextBox2.Text = Form2.RichTextBox2.Text & "e=" & e & vbCrLf
        '   Form2.RichTextBox2.Text = Form2.RichTextBox2.Text & "Cus=" & Cus & vbCrLf
        '   Form2.RichTextBox2.Text = Form2.RichTextBox2.Text & "SqrtA=" & sqrtA & vbCrLf
        ''*************************************************************************''
        '  Form2.RichTextBox2.Text = Form2.RichTextBox2.Text & "toe=" & toe & vbCrLf
        ' Form2.RichTextBox2.Text = Form2.RichTextBox2.Text & "Cic=" & Cic & vbCrLf
        '  Form2.RichTextBox2.Text = Form2.RichTextBox2.Text & "Ol0=" & Ol0 & vbCrLf
        ' Form2.RichTextBox2.Text = Form2.RichTextBox2.Text & "Cis=" & Cis & vbCrLf
        ''*************************************************************************''
        ' Form2.RichTextBox2.Text = Form2.RichTextBox2.Text & "i0=" & i0 & vbCrLf
        ' Form2.RichTextBox2.Text = Form2.RichTextBox2.Text & "Crc=" & Crc & vbCrLf
        ' Form2.RichTextBox2.Text = Form2.RichTextBox2.Text & "w=" & W & vbCrLf
        '  Form2.RichTextBox2.Text = Form2.RichTextBox2.Text & "OU=" & OU & vbCrLf
        ''*************************************************************************''
        ' Form2.RichTextBox2.Text = Form2.RichTextBox2.Text & "IDOT=" & IDOT & vbCrLf
        ' Form2.RichTextBox2.Text = Form2.RichTextBox2.Text & "L2Data=" & L2Data & vbCrLf
        ' Form2.RichTextBox2.Text = Form2.RichTextBox2.Text & "ps=" & PS & vbCrLf
        'Form2.RichTextBox2.Text = Form2.RichTextBox2.Text & "L2P=" & L2P & vbCrLf
        'Form2.RichTextBox2.Text = Form2.RichTextBox2.Text & "******************************************" & vbCrLf
        ''********************************************************************************************************''


        ''******************************计算**************************************''
        ' Form2.RichTextBox3.Text = Form2.RichTextBox3.Text & "PRN=" & prn & vbCrLf
        ' Form2.RichTextBox3.Text = Form2.RichTextBox3.Text & "------" & vbCrLf
        GM = 398600500000000.0#
        a = sqrtA ^ 2
        n0 = (GM / (a ^ 3)) ^ (1 / 2)
        n = n0 + tn
        ' Form2.RichTextBox3.Text = Form2.RichTextBox3.Text & "a=" & a & vbCrLf
        ' Form2.RichTextBox3.Text = Form2.RichTextBox3.Text & "n0=" & n0 & vbCrLf
        ' Form2.RichTextBox3.Text = Form2.RichTextBox3.Text & "n=" & n & vbCrLf
        ''*************************************************************************''
        toc = (((CalaWeek_(year, month, day) - 0) * 24) + hour) * 60 * 60
        'Form2.RichTextBox3.Text = Form2.RichTextBox3.Text & "toc=" & toc & vbCrLf
        t = toc - a0
        tk = t - toe
        If tk > 302400 Then
            tk -= 604800
        ElseIf tk < -302400 Then
            tk += 604800
        End If
        ' Form2.RichTextBox3.Text = Form2.RichTextBox3.Text & "tk=" & tk & vbCrLf
        ''*************************************************************************''
        Mk = M0 + (n * tk)
        'Form2.RichTextBox3.Text = Form2.RichTextBox3.Text & "Mk=" & Mk & vbCrLf
        ''*************************************************************************''
        'Dim msg%
        i = 0
        Ek = Mk
        Do
            '' tEk = calcEk(Mk)
            Ek = Mk + (e * Sin(Ek))
            tEk = Mk + e * Sin(Ek)
            i += 1
        Loop Until i > 10
        'Form2.RichTextBox3.Text = Form2.RichTextBox3.Text & "Ek = " & Ek & vbCrLf
        ''***********************************************************************''
        Vk = Atan(Sqrt(1 - e ^ 2) * Sin(Ek) / (Cos(Ek) - e))
        ' Form2.RichTextBox3.Text = Form2.RichTextBox3.Text & "Vk=" & Vk & vbCrLf
        ''***********************************************************************''
        fi0 = Vk + W
        ' Form2.RichTextBox3.Text = Form2.RichTextBox3.Text & "fi0=" & fi0 & vbCrLf
        ''***********************************************************************''
        rk = a * (1 - (e * Cos(Ek)))
        ' Form2.RichTextBox3.Text = Form2.RichTextBox3.Text & "rk=" & rk & vbCrLf
        ''***********************************************************************''
        Su = (Cuc * Cos(2 * fi0)) + (Cus * Sin(2 * fi0))
        Sr = (Crc * Cos(2 * fi0)) + (Crs * Sin(2 * fi0))
        Si = (Cic * Cos(2 * fi0)) + (Cis * Sin(2 * fi0))
        '  Form2.RichTextBox3.Text = Form2.RichTextBox3.Text & "Su=" & Su & vbCrLf
        ' Form2.RichTextBox3.Text = Form2.RichTextBox3.Text & "Sr=" & Sr & vbCrLf
        ' Form2.RichTextBox3.Text = Form2.RichTextBox3.Text & "Si=" & Si & vbCrLf
        ''***********************************************************************''
        uk = fi0 + Su
        rk += Sr
        ik = i0 + Si + (IDOT * tk)
        ' Form2.RichTextBox3.Text = Form2.RichTextBox3.Text & "uk=" & uk & vbCrLf
        ' Form2.RichTextBox3.Text = Form2.RichTextBox3.Text & "rk=" & rk & vbCrLf
        ' Form2.RichTextBox3.Text = Form2.RichTextBox3.Text & "ik=" & ik & vbCrLf
        ''***********************************************************************''
        xk = rk * Cos(uk)
        yk = rk * Sin(uk)
        ' Form2.RichTextBox3.Text = Form2.RichTextBox3.Text & "xk=" & xk & vbCrLf
        ' Form2.RichTextBox3.Text = Form2.RichTextBox3.Text & "yk=" & yk & vbCrLf
        ''***********************************************************************''
        we = 0.0000729211567
        tvt = OL + ((OU - we) * tk) - (we * toe)
        ' Form2.RichTextBox3.Text = Form2.RichTextBox3.Text & "tvt=" & tvt & vbCrLf
        ''***********************************************************************''
        XXk = (xk * Cos(tvt)) - yk * Cos(ik) * Sin(tvt)
        YYk = (xk * Sin(tvt)) + yk * Cos(ik) * Cos(tvt)
        ZZk = yk * Sin(ik)
        '   Form2.RichTextBox3.Text = Form2.RichTextBox3.Text & "Xk=" & XXk & vbCrLf
        ' Form2.RichTextBox3.Text = Form2.RichTextBox3.Text & "Yk=" & YYk & vbCrLf
        ' Form2.RichTextBox3.Text = Form2.RichTextBox3.Text & "ZK=" & ZZk & vbCrLf
        ' Form2.RichTextBox3.Text = Form2.RichTextBox3.Text & "********************************************" & vbCrLf
        ''*******************************************************************************************************''
    End Sub
    ' Public Function calcEk(Ek As Double)
    '
    '     calcEk = Mk + E * Sin(Ek)
    ' End Function
End Module
