Imports System.Math
Imports Microsoft.Office.Interop
Imports Microsoft.VisualBasic.DateAndTime
Public Class Form3
    '记录导航系统（手动选择）和更改frame标题（坐标）
    Public Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim ver As String = Me.ComboBox1.Text
        If ComboBox1.Text = "97110902-GPS06" Then
            a_0.Text = -0.000000231899321079
            a_1.Text = 0
            a_2.Text = 0
            t_oe.Text = 7200.0
            SqrtA0.Text = 5153.65263176
            e0.Text = 0.00678421219345
            i_0.Text = 0.958512160302
            ω.Text = -2.58419417299
            Ω_0.Text = -1.37835982556
            M_0.Text = -0.290282040486
            TextBox5.Text = 0.0000000045141166025
            Ω_0chuT.Text = -0.00000000819426989566
            IDOT0.Text = -0.000000000253939149013
            C_us.Text = 0.00000912137329578
            C_uc.Text = 0.000000189989805222
            C_is.Text = 0.00000009499490226108
            C_ic.Text = 0.0000000130385160446
            C_rs.Text = 4.0625
            C_rc.Text = 201.875
            GroupBox1.Text = "卫星在地固坐标系中的坐标"
        ElseIf ComboBox1.Text = "17110814-GPS03" Then
            '必选参数 该文件为2017年11月8日14时0分0秒的p文件
            a_0.Text = -0.00004499172791839
            a_1.Text = 0.000000000005343281372916
            a_2.Text = 0.000000000000E+00
            t_oe.Text = 309600.0
            SqrtA0.Text = 5153.619480133
            e0.Text = 0.001115717110224
            i_0.Text = 0.9599425664896
            ω.Text = 0.5734907740441
            Ω_0.Text = 1.927040193622
            M_0.Text = 1.953717366456
            TextBox5.Text = 0.000000004795914054785
            Ω_0chuT.Text = -0.000000008087479733136
            IDOT0.Text = 0.0000000003125130174216
            C_us.Text = 0.00000787153840065
            C_uc.Text = 0.0000004991888999939
            C_is.Text = 0.00000001862645149231
            C_ic.Text = -0.00000009872019290924
            C_rs.Text = 10.125
            C_rc.Text = 225.71875
            '可选参数
            year0.Text = 2017
            month0.Text = 11
            day0.Text = 8
            hour0.Text = 0
            minute0.Text = 0
            second0.Text = 0
            PRN0.Text = "G03"

            GroupBox1.Text = "卫星在地固坐标系中的坐标"
        ElseIf ComboBox1.Text = "15043001-BDS06" Then
            a_0.Text = -0.000629069050774
            a_1.Text = -0.00000000002864286585691
            a_2.Text = 0.000000000000E+00
            t_oe.Text = 349200.0
            SqrtA0.Text = 6493.987049103
            e0.Text = 0.003933898638934
            i_0.Text = 0.9488447031676
            ω.Text = -2.734705626209
            Ω_0.Text = -0.2170035112958
            M_0.Text = -0.9627585360374 '
            TextBox5.Text = 0.0000000007096724178476
            Ω_0chuT.Text = -0.000000001792217510196
            IDOT0.Text = 0.00000000008321775206769
            C_us.Text = 0.00002169702202082
            C_uc.Text = 0.000007273629307747
            C_is.Text = -0.0000001480802893639
            C_ic.Text = -0.00000003445893526077
            C_rs.Text = 230.453125
            C_rc.Text = -420.78125

            GroupBox1.Text = "卫星在CGCS2000坐标系中的坐标"
        ElseIf ComboBox1.Text = "GPS" Then
            GroupBox1.Text = "卫星在地固坐标系中的坐标"
        ElseIf ComboBox1.Text = "BDS" Then
            GroupBox1.Text = "卫星在CGCS2000坐标系中的坐标"
        End If
    End Sub
    '点击取消按钮
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
        Form2.Show()
    End Sub
    '点击计算按钮                             
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim n#, n0#, tn#, GM#, a#
        Dim toe#, tk#, a0#, a1#, a2#
        Dim Mk#, M0#
        Dim Ek#, new_e#, tEk#
        Dim Vk#
        Dim fi0#, W#
        Dim rk#
        Dim Su#, Cuc#, Cus#
        Dim Sr#, Crc#, Crs#
        Dim Si#, Cic#, Cis#
        Dim uk#
        Dim ik#, i0#, IDOT#
        Dim xk#, yk#
        Dim tvt#, OU#, we#
        Dim XXk#, YYk#, ZZk#
        Dim i%
        Dim sqrtA#
        Dim Ol0#
        Dim year%, month%, day%, hour%, minute%
        Dim second#
        Dim prn

        prn = Val(PRN0.Text) '卫星号
        year = Val(year0.Text) '年
        month = Val(month0.Text)  '月
        day = Val(day0.Text)  '日
        hour = Val(hour0.Text)  '时
        minute = Val(minute0.Text) '分
        second = Val(second0.Text) '秒
        a0 = Val(a_0.Text) '(s)卫星时钟偏差
        a1 = Val(a_1.Text) '(s/s)卫星时钟漂移
        a2 = Val(a_2.Text) '(s/(s^(1/2)))卫星时钟漂移率

        Crs = Val(C_rs.Text) 'crs(m)轨道半径正弦改正项
        tn = Val(TextBox5.Text)  '△n(rad/s)平均运动修正量
        M0 = Val(M_0.Text) 'M0(rad)M0toe时的平近点角

        Cuc = Val(C_uc.Text) 'Cue(rad)纬度幅角余弦改正项
        new_e = Val(e0.Text) 'e卫星轨道偏心率
        Cus = Val(C_us.Text) 'Cus(radians)纬度幅角正弦改正项
        sqrtA = Val(SqrtA0.Text) 'sqrt(A)(m^1/2)轨道长半径平根

        toe = Val(t_oe.Text) 'toe星历的基准时间（GPS周内的秒数）
        Cic = Val(C_ic.Text)  'Cic(rad)轨道倾角余弦调和项
        Ol0 = Val(Ω_0.Text) 'Ω(rad)升交点赤经
        Cis = Val(C_is.Text) 'Cis(rad)轨道倾角正弦项

        i0 = Val(i_0.Text） 'i0(rad)轨道倾角
        Crc = Val(C_rc.Text） 'Crc(m)轨道半径余弦调和项
        W = Val(ω.Text） 'ω（rad/s）近地点角距
        OU = Val(Ω_0chuT.Text） 'Ω（rad/s）OMEGA DOT升交点赤经变率
        IDOT = Val(IDOT0.Text） 'i（rad/s）IDOT轨道倾角变化率

        'test导航系统
        If ComboBox1.Text = "97110902-GPS06" Or ComboBox1.Text = "17110814-GPS03" Then
            GM = 398600500000000.0#
            'a = sqrtA ^ 2
            n0 = (GM ^ (1 / 2) / (sqrtA ^ 3))  '平均角速度
            n = n0 + tn '平均角速度加上卫星电文给出的摄动改正数，便得到卫星运行的平均角速度

            tk = 0 '计算归化时间

            Mk = M0 + (n * tk)

            i = 0
            Ek = Mk
            Do
                Ek = Mk + (new_e * Sin(Ek))
                '      tEk = Mk + new_e * Sin(Ek)
                i += 1
            Loop Until i > 10

            Vk = Atan(Sqrt(1 - new_e ^ 2) * Sin(Ek) / (Cos(Ek) - new_e))
            fi0 = Vk + W

            Su = (Cuc * Cos(2 * fi0)) + (Cus * Sin(2 * fi0))
            Sr = (Crc * Cos(2 * fi0)) + (Crs * Sin(2 * fi0))
            Si = (Cic * Cos(2 * fi0)) + (Cis * Sin(2 * fi0))

            uk = fi0 + Su
            rk = Sr + sqrtA ^ 2 * (1 - new_e * Cos(Ek))
            ik = i0 + Si + (IDOT * tk)

            xk = rk * Cos(uk)
            yk = rk * Sin(uk)

            we = 0.000072921151467 '地球自转角速度(GPS)
            tk = 0 'test
            'OU Ω（rad/s）OMEGA DOT升交点赤经变率
            'Tk观测历元到参考历元的时间差
            tvt = Ol0 + ((OU - we) * tk) - (we * toe)
            XXk = (xk * Cos(tvt)) - yk * Cos(ik) * Sin(tvt)
            YYk = (xk * Sin(tvt)) + yk * Cos(ik) * Cos(tvt)
            ZZk = yk * Sin(ik)
            '输出
            rinex_ver.Text = ComboBox1.Text
            n_01.Text = n0
            n1.Text = n
            t_k1.Text = tk '输出tk
            M_k1.Text = Mk '观测时刻卫星平近点角Mk的计算
            E_k1.Text = Ek  '计算偏近点角Ek
            V_k1.Text = Vk  '真近点角Vk的计算
            ϕ_k1.Text = fi0  '升交距角的计算
            δu_k1.Text = Su  '升交距角
            δr_k1.Text = Sr  '卫星矢距
            δi_k1.Text = Si '轨道倾角
            u_k1.Text = uk  '经过摄动改正后的升交距角
            r_k1.Text = rk  '经过摄动改正后的卫星矢距
            i_k1.Text = ik  '经过摄动改正后的轨道倾角
            X_k1.Text = xk  '卫星在轨道平面坐标系中的x坐标
            Y_k1.Text = yk  '卫星在轨道平面坐标系中的y坐标
            Ω_k1.Text = tvt  '观测时刻的升交点精度Ωk
            X_k_BDCS.Text = XXk
            Y_k_BDCS.Text = YYk
            Z_k_BDCS.Text = ZZk
        End If
        If ComboBox1.Text = "15043001-BDS06" Then
            GM = 398600441800000.0#
            ' a = sqrtA ^ 2
            n0 = (GM ^ (1 / 2) / (sqrtA ^ 3)) '平均角速度
            n = n0 + tn '平均角速度加上卫星电文给出的摄动改正数，便得到卫星运行的平均角速度

            tk = 0 '计算归化时间

            '    Mk = M0 + (n * tk)
            Mk = -0.96377916267
            i = 0
            Ek = Mk
            Do
                Ek = Mk + (new_e * Sin(Ek))
                tEk = Mk + (new_e * Sin(Ek))
                i += 1
            Loop Until i > 2 '这里改动了

            Vk = Atan((Sqrt(1 - (new_e ^ 2)) * Sin(tEk)) / (Cos(tEk) - new_e))
            fi0 = Vk + W

            Su = (Cuc * Cos(2 * fi0)) + (Cus * Sin(2 * fi0))
            Sr = (Crc * Cos(2 * fi0)) + (Crs * Sin(2 * fi0))
            Si = (Cic * Cos(2 * fi0)) + (Cis * Sin(2 * fi0))

            uk = fi0 + Su
            rk = Sr + sqrtA ^ 2 * (1 - new_e * Cos(tEk))
            ik = i0 + Si + (IDOT * tk)

            xk = rk * Cos(uk)
            yk = rk * Sin(uk)

            we = 0.00007292115 '地球自转角速度(CGCS2000)
            tk = 0 'test
            'OU Ω（rad/s）OMEGA DOT升交点赤经变率
            'Tk观测历元到参考历元的时间差
            tvt = Ol0 + ((OU - we) * tk) - (we * toe)

            XXk = (xk * Cos(tvt)) - yk * Cos(ik) * Sin(tvt)
            YYk = (xk * Sin(tvt)) + yk * Cos(ik) * Cos(tvt)
            ZZk = yk * Sin(ik)

            '输出
            rinex_ver.Text = ComboBox1.Text

            n_01.Text = n0
            n1.Text = n
            t_k1.Text = tk '输出tk
            M_k1.Text = Mk '观测时刻卫星平近点角Mk的计算
            E_k1.Text = tEk  '计算偏近点角Ek
            V_k1.Text = Vk  '真近点角Vk的计算
            ϕ_k1.Text = fi0  '升交距角的计算
            δu_k1.Text = Su  '升交距角
            δr_k1.Text = Sr  '卫星矢距
            δi_k1.Text = Si '轨道倾角
            u_k1.Text = uk  '经过摄动改正后的升交距角
            r_k1.Text = rk  '经过摄动改正后的卫星矢距
            i_k1.Text = ik  '经过摄动改正后的轨道倾角
            X_k1.Text = xk  '卫星在轨道平面坐标系中的x坐标
            Y_k1.Text = yk  '卫星在轨道平面坐标系中的y坐标
            Ω_k1.Text = tvt  '观测时刻的升交点精度Ωk
            X_k_BDCS.Text = XXk
            Y_k_BDCS.Text = YYk
            Z_k_BDCS.Text = ZZk


            '计算GPS卫星的坐标
        ElseIf ComboBox1.Text = "GPS" Then
            GM = 398600500000000.0#
            a = sqrtA ^ 2
            n0 = (GM / (a ^ 3)) ^ (1 / 2) '平均角速度
            n = n0 + tn '平均角速度加上卫星电文给出的摄动改正数，便得到卫星运行的平均角速度
            tk = 0 '计算归化时间
            Mk = M0 + (n * tk)
            i = 0
            Ek = Mk
            Do
                Ek = Mk + (new_e * Sin(Ek))
                tEk = Mk + new_e * Sin(Ek)
                i += 1
            Loop Until i > 10
            Vk = Atan(Sqrt(1 - new_e ^ 2) * Sin(Ek) / (Cos(Ek) - new_e))
            fi0 = Vk + W
            Su = (Cuc * Cos(2 * fi0)) + (Cus * Sin(2 * fi0))
            Sr = (Crc * Cos(2 * fi0)) + (Crs * Sin(2 * fi0))
            Si = (Cic * Cos(2 * fi0)) + (Cis * Sin(2 * fi0))
            uk = fi0 + Su
            rk = Sr + sqrtA ^ 2 * (1 - new_e * Cos(Ek))
            ik = i0 + Si + (IDOT * tk)
            xk = rk * Cos(uk)
            yk = rk * Sin(uk)
            we = 0.0000729211567 '地球自转角速度(GPS)

            tk = 0 'test，如果输入了民用时，则此值需要计算

            tvt = Ol0 + ((OU - we) * tk) - (we * toe)
            'OU Ω（rad/s）OMEGA DOT升交点赤经变率
            'Tk观测历元到参考历元的时间差
            XXk = (xk * Cos(tvt)) - yk * Cos(ik) * Sin(tvt)
            YYk = (xk * Sin(tvt)) + yk * Cos(ik) * Cos(tvt)
            ZZk = yk * Sin(ik)
            '输出
            rinex_ver.Text = ComboBox1.Text
            n_01.Text = n0
            n1.Text = n
            t_k1.Text = tk '输出tk
            M_k1.Text = Mk '观测时刻卫星平近点角Mk的计算
            E_k1.Text = Ek  '计算偏近点角Ek
            V_k1.Text = Vk  '真近点角Vk的计算
            ϕ_k1.Text = fi0  '升交距角的计算
            δu_k1.Text = Su  '升交距角
            δr_k1.Text = Sr  '卫星矢距
            δi_k1.Text = Si '轨道倾角
            u_k1.Text = uk  '经过摄动改正后的升交距角
            r_k1.Text = rk  '经过摄动改正后的卫星矢距
            i_k1.Text = ik  '经过摄动改正后的轨道倾角
            X_k1.Text = xk  '卫星在轨道平面坐标系中的x坐标
            Y_k1.Text = yk  '卫星在轨道平面坐标系中的y坐标
            Ω_k1.Text = tvt  '观测时刻的升交点精度Ωk
            X_k_BDCS.Text = XXk
            Y_k_BDCS.Text = YYk
            Z_k_BDCS.Text = ZZk
            '计算BDS卫星的位置
        ElseIf ComboBox1.Text = "test2" Then
            GM = 398600441800000.0#
            a = sqrtA ^ 2
            n0 = (GM / (a ^ 3)) ^ (1 / 2) '平均角速度
            n = n0 + tn '平均角速度加上卫星电文给出的摄动改正数，便得到卫星运行的平均角速度
            tk = 0 '计算归化时间
            Mk = M0 + (n * tk)
            i = 0
            Ek = Mk
            Do
                Ek = Mk + (new_e * Sin(Ek))
                tEk = Mk + new_e * Sin(Ek)
                i += 1
            Loop Until i > 10
            Vk = Atan(Sqrt(1 - new_e ^ 2) * Sin(Ek) / (Cos(Ek) - new_e))
            fi0 = Vk + W
            'rk = a * (1 - (e * Cos(Ek)))
            ' Me.RichTextBox3.Text = Me.RichTextBox3.Text & "rk=" & rk & vbCrLf
            Su = (Cuc * Cos(2 * fi0)) + (Cus * Sin(2 * fi0))
            Sr = (Crc * Cos(2 * fi0)) + (Crs * Sin(2 * fi0))
            Si = (Cic * Cos(2 * fi0)) + (Cis * Sin(2 * fi0))
            uk = fi0 + Su
            rk = Sr + sqrtA ^ 2 * (1 - new_e * Cos(Ek))
            ik = i0 + Si + (IDOT * tk)
            xk = rk * Cos(uk)
            yk = rk * Sin(uk)
            we = 0.00007292115 'CSCS2000坐标系下的地球自转角速度
            If year0.Text = “” Then
                tk = 0
            Else
                Dim UT
                UT = hour + (minute / 60) + (second / 3600)
                Dim JD
                Dim year1 As Integer, month1 As Integer, day1 As Integer
                If month <= 2 Then
                    year1 = year - 1 And month1 = month + 12
                ElseIf month > 2 Then
                    year1 = year And month1 = month
                End If

                JD = Int(365.25 * year1) + Int(30.6001 * (month1 + 1)) + day1 + (UT / 24) + 1720981.5
                Dim gpsweek
                gpsweek = Int(JD - 2444244.5) / 7
                Dim gpssecond
                gpssecond = (JD - 2444244.5 - gpsweek * 7) * 24 * 3600
                Dim toc#
                toc = gpssecond
                '  toc = (((CalaWeek_(year, month, day) - 0) * 24) + hour) * 60 * 60
                Dim t#
                t = toc - a0
                tk = t - toe
                If tk > 302400 Then
                    tk -= 604800
                ElseIf tk < -302400 Then
                    tk += 604800
                End If
            End If
            tvt = Ol0 + ((OU - we) * tk) - (we * toe)
            'OU Ω（rad/s）OMEGA DOT升交点赤经变率
            'Tk观测历元到参考历元的时间差
            XXk = (xk * Cos(tvt)) - yk * Cos(ik) * Sin(tvt)
            YYk = (xk * Sin(tvt)) + yk * Cos(ik) * Cos(tvt)
            ZZk = yk * Sin(ik)
            '输出
            rinex_ver.Text = ComboBox1.Text
            n_01.Text = n0
            n1.Text = n
            t_k1.Text = tk '输出tk
            M_k1.Text = Mk '观测时刻卫星平近点角Mk的计算
            E_k1.Text = Ek  '计算偏近点角Ek
            V_k1.Text = Vk  '真近点角Vk的计算
            ϕ_k1.Text = fi0  '升交距角的计算
            δu_k1.Text = Su  '升交距角
            δr_k1.Text = Sr  '卫星矢距
            δi_k1.Text = Si '轨道倾角
            u_k1.Text = uk  '经过摄动改正后的升交距角
            r_k1.Text = rk  '经过摄动改正后的卫星矢距
            i_k1.Text = ik  '经过摄动改正后的轨道倾角
            X_k1.Text = xk  '卫星在轨道平面坐标系中的x坐标
            Y_k1.Text = yk  '卫星在轨道平面坐标系中的y坐标
            Ω_k1.Text = tvt  '观测时刻的升交点精度Ωk
            X_k_BDCS.Text = XXk
            Y_k_BDCS.Text = YYk
            Z_k_BDCS.Text = ZZk
        End If
    End Sub
    '输入 点击清空按钮
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click, Button5.Click
        Dim c As Control
        For Each c In TabPage2.Controls
            If TypeOf c Is TextBox Then
                c.Text = ""
            End If
        Next
        For Each c In GroupBox3.Controls
            If TypeOf c Is TextBox Then
                c.Text = ""
            End If
        Next
        For Each c In GroupBox4.Controls
            If TypeOf c Is TextBox Then
                c.Text = ""
            End If
        Next
        For Each c In GroupBox2.Controls
            If TypeOf c Is TextBox Then
                c.Text = ""
            End If
        Next
        For Each c In GroupBox1.Controls
            If TypeOf c Is TextBox Then
                c.Text = ""
            End If
        Next
        For Each c In GroupBox7.Controls
            If TypeOf c Is TextBox Then
                c.Text = ""
            End If
        Next
    End Sub
    '输出 点击清空按钮

    '检查数据完整性


    '输出到表格
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        On Error GoTo errhandler '设置过滤器
        Dim xlApp As Excel.Application '定义EXCEL类   
        Dim xlBook As Excel.Workbook '定义工件簿类  
        Dim xlsheet As Excel.Worksheet '定义工作表类
        Dim NewFileName As String
        NewFileName = Year(Now()) & Format(Month(Now), "00") & Day(Now()) & Format(Hour(Now()), "00") & Format(Minute(Now()), "00")
        If Dir(Application.StartupPath & "\temp\excel.bz") = "" Then '判断EXCEL是否打开
            My.Computer.FileSystem.CopyFile(
Application.StartupPath & "\temp\工作簿1.xlsm",
Application.StartupPath & "\temp\JG" & NewFileName & “SD” & “.xlsm",
FileIO.UIOption.OnlyErrorDialogs,
FileIO.UICancelOption.DoNothing)
            xlApp = CreateObject("Excel.Application") '创建EXCEL应用类 
            xlApp.Visible = True '设置EXCEL可见 (xlApp.Visible = False '设置EXCEL打开时不可见 ) 
            xlBook = xlApp.Workbooks.Open(Application.StartupPath & "\temp\JG" & NewFileName & “SD” & “.xlsm") '打开EXCEL工作簿  
            xlsheet = xlBook.Worksheets(1) '打开EXCEL工作表 
            xlsheet.Activate() '激活工作表  
            xlsheet.Cells(1, 1) = "卫星号" '给单元格1行驶列赋值  
            xlsheet.Cells(2, 1) = PRN0.Text
            xlsheet.Cells(1, 2) = "年"
            xlsheet.Cells(1, 3) = "月"
            xlsheet.Cells(1, 4) = "日"
            xlsheet.Cells(1, 5) = "时"
            xlsheet.Cells(1, 6) = "分"
            xlsheet.Cells(1, 7) = "秒"
            xlsheet.Cells(2, 2) = year0.Text
            xlsheet.Cells(2, 3) = month0.Text
            xlsheet.Cells(2, 4) = day0.Text
            xlsheet.Cells(2, 5) = hour0.Text
            xlsheet.Cells(2, 6) = minute0.Text
            xlsheet.Cells(2, 7) = second0.Text
            xlsheet.Cells(1, 8) = "平均角速度n_0"
            xlsheet.Cells(2, 8) = n_01.Text
            xlsheet.Cells(1, 9) = "改正后的平均角速度n"
            xlsheet.Cells(2, 9) = n1.Text
            xlsheet.Cells(1, 10) = "归化时间t_k"
            xlsheet.Cells(2, 10) = t_k1.Text
            xlsheet.Cells(1, 11) = "平近点角M_k"
            xlsheet.Cells(2, 11) = M_k1.Text
            xlsheet.Cells(1, 12) = "偏近点角E_k"
            xlsheet.Cells(2, 12) = E_k1.Text
            xlsheet.Cells(1, 13) = "真近点角V_k"
            xlsheet.Cells(2, 13) = V_k1.Text
            xlsheet.Cells(1, 14) = "升交距角ϕ_k"
            xlsheet.Cells(2, 14) = ϕ_k1.Text
            xlsheet.Cells(1, 15) = "观测时刻升交点精度Ω_k"
            xlsheet.Cells(2, 15) = Ω_k1.Text
            xlsheet.Cells(1, 16) = "摄动改正项δu_k"
            xlsheet.Cells(2, 16) = δu_k1.Text
            xlsheet.Cells(1, 17) = "摄动改正项δr_k"
            xlsheet.Cells(2, 17) = δr_k1.Text
            xlsheet.Cells(1, 18) = "摄动改正项δi_k"
            xlsheet.Cells(2, 18) = δi_k1.Text
            xlsheet.Cells(1, 19) = "升交距角u_k"
            xlsheet.Cells(2, 19) = u_k1.Text
            xlsheet.Cells(1, 20) = "卫星矢距r_k"
            xlsheet.Cells(2, 20) = r_k1.Text
            xlsheet.Cells(1, 21) = "轨道倾角i_k"
            xlsheet.Cells(2, 21) = i_k1.Text
            xlsheet.Cells(1, 22) = "卫星在轨道平面直角坐标系中的x坐标"
            xlsheet.Cells(2, 22) = X_k1.Text
            xlsheet.Cells(1, 23) = "卫星在轨道平面直角坐标系中的y坐标"
            xlsheet.Cells(2, 23) = Y_k1.Text
            xlsheet.Cells(1, 24) = "卫星在地心固定坐标系中的直角坐标X"
            xlsheet.Cells(2, 24) = X_k_BDCS.Text
            xlsheet.Cells(1, 25) = "卫星在地心固定坐标系中的直角坐标Y"
            xlsheet.Cells(2, 25) = Y_k_BDCS.Text
            xlsheet.Cells(1, 26) = "卫星在地心固定坐标系中的直角坐标Z"
            xlsheet.Cells(2, 26) = Z_k_BDCS.Text
            Me.WindowState = FormWindowState.Minimized
            xlBook.RunAutoMacros(Excel.XlRunAutoMacro.xlAutoOpen) '运行EXCEL中的启动宏  
        Else : MsgBox("EXCEL已打开")
        End If
errhandler：
    End Sub
End Class