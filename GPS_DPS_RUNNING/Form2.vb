Option Explicit On
Imports System.Math
Imports System.IO
Imports Microsoft.Office.Interop
Imports Microsoft.VisualBasic.DateAndTime
Imports System.Windows.Forms.Clipboard
Public Class Form2
    Dim xlApp As Excel.Application '定义EXCEL类   
    Dim xlBook As Excel.Workbook '定义工件簿类  
    Dim xlsheet As Excel.Worksheet '定义工作表类
    ''***************************头文件***********************************''
    Dim Data() As String                '存放原参数字符
    Private key_Head%                       '存放查找头文件结束符的返回值
    Private key_L%                          '存放头文件结束符所在的行号
    Dim n#, n0#, tn#, GM#, a#
    Dim toe#, tk#, toc#, t#, a0#, a1#, a2#
    Dim Mk#, M0#
    Dim Ek#, e#, tEk#
    Dim Vk#
    Dim fi0#, W#
    Dim rk#
    Dim Su#, Cuc#, Cus#
    Dim Sr#, Crc#, Crs#
    Dim Si#, Cic#, Cis#
    Dim uk#
    Dim ik#, i0#, IDOT#
    Dim xk#, yk#
    Dim tvt#, OL#, OU#, we#
    Dim XXk#, YYk#, ZZk#
    Dim i%

    Dim IDOE#
    Dim sqrtA#
    Dim Ol0#
    Dim L2Data#, PS#, L2P#

    Private Sub 帮助ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 帮助ToolStripMenuItem.Click
        Dim strPath As String
        strPath = Application.StartupPath & "\temp\GPS和BDS卫星位置计算程序设计.pdf"
        System.Diagnostics.Process.Start("IExplore.exe", strPath)
    End Sub
    Dim ST_PRC#, ST_HEL#, TGD#, IDOC#
    Dim Send_Time#, Countin_h#, ps_1#, ps_2#
    Dim prn
    ''***************************头文件***********************************''
    Private Sub DefException(ByVal data)
        If data Is Nothing Then
            Dim MyEx As New Exception("应先选择历书文件") With {
            .Source = "DefException"
        }
            Throw MyEx 'throw err
        End If
    End Sub
    'Rinex 2.10
    Public Function Data_operate_1(data, j)
        Dim year%, month%, day%, hour%, minute%
        Dim second#
        If j Is Nothing Then
            Throw New ArgumentNullException(NameOf(j)) 'mach fix
        End If
        ''*******************************取参数********************************''
        Dim prn
        prn = Val(Microsoft.VisualBasic.Left(data(key_L + 1), 2)) '卫星号
        year = Val(Mid(data(key_L + 1), 4, 2)) '年
        month = Val(Mid(data(key_L + 1), 7, 2)) '月
        day = Val(Mid(data(key_L + 1), 10, 2)) '日
        hour = Val(Mid(data(key_L + 1), 13, 2)) '时
        minute = Val(Mid(data(key_L + 1), 16, 2)) '分
        second = Val(Mid(data(key_L + 1), 18, 5)) '秒
        a0 = Val(Mid(data(key_L + 1), 24, 19)) '(s)卫星时钟偏差
        a1 = Val(Mid(data(key_L + 1), 43, 19)) '(s/s)卫星时钟漂移
        a2 = Val(Mid(data(key_L + 1), 62, 19)) '(s/(s^(1/2)))卫星时钟漂移率
        ''************************************************************************''
        ''第2行
        IDOE = Val(Mid(data(key_L + 2), 5, 19)) '星历数据的有效期龄
        Crs = Val(Mid(data(key_L + 2), 24, 19)) 'crs(m)轨道半径正弦改正项
        tn = Val(Mid(data(key_L + 2), 43, 19)) '△n(rad/s)平均运动修正量
        M0 = Val(Mid(data(key_L + 2), 62, 19)) 'M0(rad)M0toe时的平近点角
        ''************************************************************************''
        ''第3行
        Cuc = Val(Mid(data(key_L + 3), 5, 19)) 'Cue(rad)纬度幅角余弦改正项
        e = Val(Mid(data(key_L + 3), 24, 19)) 'e卫星轨道偏心率
        Cus = Val(Mid(data(key_L + 3), 43, 19)) 'Cus(radians)纬度幅角正弦改正项
        sqrtA = Val(Mid(data(key_L + 3), 62, 19)) 'sqrt(A)(m^1/2)轨道长半径平根
        ''************************************************************************''
        ''第4行
        toe = Val(Mid(data(key_L + 4), 5, 19)) 'toe星历的基准时间（GPS周内的秒数）
        Cic = Val(Mid(data(key_L + 4), 24, 19)) 'Cic(rad)轨道倾角余弦调和项
        Ol0 = Val(Mid(data(key_L + 4), 43, 19)) 'Ω(rad)升交点赤经
        Cis = Val(Mid(data(key_L + 4), 62, 19)) 'Cis(rad)轨道倾角正弦项
        ''************************************************************************''
        ''第5行
        i0 = Val(Mid(data(key_L + 5), 5, 19)) 'i0(rad)轨道倾角
        Crc = Val(Mid(data(key_L + 5), 24, 19)) 'Crc(m)轨道半径余弦调和项
        W = Val(Mid(data(key_L + 5), 43, 19)) 'w（rad/s）近地点角距
        OU = Val(Mid(data(key_L + 5), 62, 19)) 'Ω（rad/s）OMEGA DOT升交点赤经变率
        ''************************************************************************''
        ''第6行
        IDOT = Val(Mid(data(key_L + 6), 5, 19)) 'i（rad/s）IDOT轨道倾角变化率
        L2Data = Val(Mid(data(key_L + 6), 24, 19)) 'L2上的码
        PS = Val(Mid(data(key_L + 6), 43, 19)) 'GPS星期数（与TOE一同表示时间），为连续计数，不是1021的余数
        L2P = Val(Mid(data(key_L + 6), 62, 19)) 'L2P码数据标志
        ''************************************************************************''
        ''第7行
        ST_PRC = Val(Mid(data(key_L + 7), 5, 19)) '卫星精度（m）
        ST_HEL = Val(Mid(data(key_L + 7), 24, 19)) '卫星健康（MSB第1子帧第3字第17~22位)
        TGD = Val(Mid(data(key_L + 7), 43, 19)) 'TGD(Sec)
        IDOC = Val(Mid(data(key_L + 7), 62, 19)) 'IODC种的数据龄期
        ''************************************************************************''
        ''第8行
        Send_Time = Val(Mid(data(key_L + 8), 5, 19)) '电文发送时刻（单位为GPS周的秒，通过交换字（HOW)中的Z计数得出）
        Countin_h = Val(Mid(data(key_L + 8), 24, 19)) '拟合区间（h），如未知则为零
        ps_1 = Val(Mid(data(key_L + 8), 43, 19)) '备用
        ps_2 = Val(Mid(data(key_L + 8), 62, 19)) '备用
        ''*************************输出各参数至text2*************************************''
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "卫星号PRN =" & prn & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "----" & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "年，月，日y,m,d=" & " " & year & " " & month & " " & day & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "时，分，秒h,min,sec=" & hour & " " & minute & " " & second & " " & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "卫星时钟偏差a0=" & a0 & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "卫星时钟漂移a1=" & a1 & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "卫星时钟漂移率a2=" & a2 & vbCrLf
        ''*************************************************************************''
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "星历数据的有效期龄IDOE=" & IDOE & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "平均运动修正量tn=" & tn & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "按参考历元计算的平近点角M0=" & M0 & vbCrLf
        ''*************************************************************************''
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "纬度幅角的余弦调和项改正的振幅Cuc=" & Cuc & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "轨道偏心率e=" & e & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "纬度幅角的正弦调和项改正的振幅Cus=" & Cus & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "轨道长半径的平方根SqrtA=" & sqrtA & vbCrLf
        ''*************************************************************************''
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "星历的基准时间（GPS周内的秒数）toe=" & toe & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "轨道倾角的余弦调和项改正的振幅Cic=" & Cic & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "升交点赤经Ω=" & Ol0 & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "轨道倾角的正弦调和项改正的振幅Cis=" & Cis & vbCrLf
        ''*************************************************************************''
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "按参考历元计算的轨道倾角i0=" & i0 & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "轨道半径的余弦调和项改正的振幅Crc=" & Crc & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "近地点角距w=" & W & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "升交点赤经变率OMEGA DOT=" & OU & vbCrLf
        ''*************************************************************************''
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "轨道倾角变化率IDOT=" & IDOT & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "L2上的码L2Data=" & L2Data & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "GPS星期数ps=" & PS & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "码数据标志L2P=" & L2P & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "******************************************" & vbCrLf
        ''********************************************************************************************************''
        ''******************************计算**************************************''
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "卫星号PRN=" & prn & vbCrLf
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "-----------------------------" & vbCrLf
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & year & month & day & hour & minute & second
        GM = 398600500000000.0#
        a = sqrtA ^ 2
        n0 = (GM / (a ^ 3)) ^ (1 / 2) '平均角速度
        n = n0 + tn '平均角速度加上卫星电文给出的摄动改正数，便得到卫星运行的平均角速度
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "长半轴a=" & a & vbCrLf
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "卫星运行的平均角速度n=" & n & vbCrLf
        ''*************************************************************************''
        toc = (((CalaWeek_(year, month, day) - 0) * 24) + hour) * 60 * 60
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "参考时刻toc=" & toc & vbCrLf '参考时刻toc
        t = toc - a0
        tk = t - toe
        If tk > 302400 Then
            tk -= 604800
        ElseIf tk < -302400 Then
            tk += 604800
        End If
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "归化时间tk=" & tk & vbCrLf '计算归化时间
        ''*************************************************************************''
        Mk = M0 + (n * tk)
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "观测时刻的卫星平近点角Mk=" & Mk & vbCrLf '观测时刻卫星平近点角Mk的计算
        ''*************************************************************************''
        i = 0
        Ek = Mk
        Do
            Ek = Mk + (e * Sin(Ek))
            tEk = Mk + e * Sin(Ek)
            i += 1
        Loop Until i > 10
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "偏近点角Ek = " & Ek & vbCrLf '计算偏近点角Ek
        ''***********************************************************************''
        Vk = Atan(Sqrt(1 - e ^ 2) * Sin(Ek) / (Cos(Ek) - e))
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "真近点角Vk=" & Vk & vbCrLf '真近点角Vk的计算
        ''***********************************************************************''
        fi0 = Vk + W
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "升交距角fi0=" & fi0 & vbCrLf '升交距角的计算
        ''***********************************************************************''
        'rk = a * (1 - (e * Cos(Ek)))
        ' Me.RichTextBox3.Text = Me.RichTextBox3.Text & "rk=" & rk & vbCrLf
        ''***********************************************************************''
        Su = (Cuc * Cos(2 * fi0)) + (Cus * Sin(2 * fi0))
        Sr = (Crc * Cos(2 * fi0)) + (Crs * Sin(2 * fi0))
        Si = (Cic * Cos(2 * fi0)) + (Cis * Sin(2 * fi0))
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "升交距角Su=" & Su & vbCrLf '升交距角
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "卫星矢距Sr=" & Sr & vbCrLf '卫星矢距
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "轨道倾角Si=" & Si & vbCrLf '轨道倾角
        ''***********************************************************************''
        uk = fi0 + Su
        rk += Sr
        ik = i0 + Si + (IDOT * tk)
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "经过摄动改正后的升交距角uk=" & uk & vbCrLf '经过摄动改正后的升交距角
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "经过摄动改正后的卫星矢距rk=" & rk & vbCrLf '经过摄动改正后的卫星矢距
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "经过摄动改正后的轨道倾角ik=" & ik & vbCrLf '经过摄动改正后的轨道倾角
        ''***********************************************************************''
        xk = rk * Cos(uk)
        yk = rk * Sin(uk)
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "卫星在轨道平面坐标系中的x坐标xk=" & xk & vbCrLf '卫星在轨道平面坐标系中的x坐标
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "卫星在轨道平面坐标系中的y坐标yk=" & yk & vbCrLf '卫星在轨道平面坐标系中的y坐标
        ''***********************************************************************''
        we = 0.0000729211567 '地球自转的速率
        tvt = OL + ((OU - we) * tk) - (we * toe)
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "观测时刻的升交点精度Ωk=" & tvt & vbCrLf '观测时刻的升交点精度Ωk
        ''***********************************************************************''
        XXk = (xk * Cos(tvt)) - yk * Cos(ik) * Sin(tvt)
        YYk = (xk * Sin(tvt)) + yk * Cos(ik) * Cos(tvt)
        ZZk = yk * Sin(ik)
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & “卫星在地心固定坐标系中的直角坐标 ：” & vbCrLf & "Xk=" & XXk & vbCrLf
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "Yk=" & YYk & vbCrLf
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "ZK=" & ZZk & vbCrLf
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "********************************************" & vbCrLf
        ''*******************************************************************************************************''

    End Function
    'Rinex 2.11
    Public Function Data_operate_2(data, j)
        Dim year%, month%, day%, hour%, minute%
        Dim second#
        If j Is Nothing Then
            Throw New ArgumentNullException(NameOf(j)) 'mach fix
        End If
        ''*******************************取参数********************************''
        Dim prn
        prn = Val(Microsoft.VisualBasic.Left(data(key_L + 1), 2)) '卫星号
        year = Val(Mid(data(key_L + 1), 4, 2)) '年
        month = Val(Mid(data(key_L + 1), 7, 2)) '月
        day = Val(Mid(data(key_L + 1), 10, 2)) '日
        hour = Val(Mid(data(key_L + 1), 13, 2)) '时
        minute = Val(Mid(data(key_L + 1), 16, 2)) '分
        second = Val(Mid(data(key_L + 1), 18, 5)) '秒
        a0 = Val(Mid(data(key_L + 1), 24, 19)) '(s)卫星时钟偏差
        a1 = Val(Mid(data(key_L + 1), 43, 19)) '(s/s)卫星时钟漂移
        a2 = Val(Mid(data(key_L + 1), 62, 19)) '(s/(s^(1/2)))卫星时钟漂移率
        ''************************************************************************''
        ''第2行
        IDOE = Val(Mid(data(key_L + 2), 5, 19)) '星历数据的有效期龄
        Crs = Val(Mid(data(key_L + 2), 24, 19)) 'crs(m)轨道半径正弦改正项
        tn = Val(Mid(data(key_L + 2), 43, 19)) '△n(rad/s)平均运动修正量
        M0 = Val(Mid(data(key_L + 2), 62, 19)) 'M0(rad)M0toe时的平近点角
        ''************************************************************************''
        ''第3行
        Cuc = Val(Mid(data(key_L + 3), 5, 19)) 'Cue(rad)纬度幅角余弦改正项
        e = Val(Mid(data(key_L + 3), 24, 19)) 'e卫星轨道偏心率
        Cus = Val(Mid(data(key_L + 3), 43, 19)) 'Cus(radians)纬度幅角正弦改正项
        sqrtA = Val(Mid(data(key_L + 3), 62, 19)) 'sqrt(A)(m^1/2)轨道长半径平根
        ''************************************************************************''
        ''第4行
        toe = Val(Mid(data(key_L + 4), 5, 19)) 'toe星历的基准时间（GPS周内的秒数）
        Cic = Val(Mid(data(key_L + 4), 24, 19)) 'Cic(rad)轨道倾角余弦调和项
        Ol0 = Val(Mid(data(key_L + 4), 43, 19)) 'Ω(rad)升交点赤经
        Cis = Val(Mid(data(key_L + 4), 62, 19)) 'Cis(rad)轨道倾角正弦项
        ''************************************************************************''
        ''第5行
        i0 = Val(Mid(data(key_L + 5), 5, 19)) 'i0(rad)轨道倾角
        Crc = Val(Mid(data(key_L + 5), 24, 19)) 'Crc(m)轨道半径余弦调和项
        W = Val(Mid(data(key_L + 5), 43, 19)) 'w（rad/s）近地点角距
        OU = Val(Mid(data(key_L + 5), 62, 19)) 'Ω（rad/s）OMEGA DOT升交点赤经变率
        ''************************************************************************''
        ''第6行
        IDOT = Val(Mid(data(key_L + 6), 5, 19)) 'i（rad/s）IDOT轨道倾角变化率
        L2Data = Val(Mid(data(key_L + 6), 24, 19)) 'L2上的码
        PS = Val(Mid(data(key_L + 6), 43, 19)) 'GPS星期数（与TOE一同表示时间），为连续计数，不是1021的余数
        L2P = Val(Mid(data(key_L + 6), 62, 19)) 'L2P码数据标志
        ''************************************************************************''
        ''第7行
        ST_PRC = Val(Mid(data(key_L + 7), 5, 19)) '卫星精度（m）
        ST_HEL = Val(Mid(data(key_L + 7), 24, 19)) '卫星健康（MSB第1子帧第3字第17~22位)
        TGD = Val(Mid(data(key_L + 7), 43, 19)) 'TGD(Sec)
        IDOC = Val(Mid(data(key_L + 7), 62, 19)) 'IODC种的数据龄期
        ''************************************************************************''
        ''第8行
        Send_Time = Val(Mid(data(key_L + 8), 5, 19)) '电文发送时刻（单位为GPS周的秒，通过交换字（HOW)中的Z计数得出）
        Countin_h = Val(Mid(data(key_L + 8), 24, 19)) '拟合区间（h），如未知则为零
        ps_1 = Val(Mid(data(key_L + 8), 43, 19)) '备用
        ps_2 = Val(Mid(data(key_L + 8), 62, 19)) '备用
        ''*************************输出各参数至text2*************************************''
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "卫星号PRN =" & prn & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "----" & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "年，月，日y,m,d=" & " " & year & " " & month & " " & day & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "时，分，秒h,min,sec=" & hour & " " & minute & " " & second & " " & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "卫星时钟偏差a0=" & a0 & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "卫星时钟漂移a1=" & a1 & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "卫星时钟漂移率a2=" & a2 & vbCrLf
        ''*************************************************************************''
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "星历数据的有效期龄IDOE=" & IDOE & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "平均运动修正量tn=" & tn & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "按参考历元计算的平近点角M0=" & M0 & vbCrLf
        ''*************************************************************************''
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "纬度幅角的余弦调和项改正的振幅Cuc=" & Cuc & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "轨道偏心率e=" & e & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "纬度幅角的正弦调和项改正的振幅Cus=" & Cus & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "轨道长半径的平方根SqrtA=" & sqrtA & vbCrLf
        ''*************************************************************************''
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "星历的基准时间（GPS周内的秒数）toe=" & toe & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "轨道倾角的余弦调和项改正的振幅Cic=" & Cic & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "升交点赤经Ω=" & Ol0 & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "轨道倾角的正弦调和项改正的振幅Cis=" & Cis & vbCrLf
        ''*************************************************************************''
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "按参考历元计算的轨道倾角i0=" & i0 & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "轨道半径的余弦调和项改正的振幅Crc=" & Crc & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "近地点角距w=" & W & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "升交点赤经变率OMEGA DOT=" & OU & vbCrLf
        ''*************************************************************************''
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "轨道倾角变化率IDOT=" & IDOT & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "L2上的码L2Data=" & L2Data & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "GPS星期数ps=" & PS & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "码数据标志L2P=" & L2P & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "******************************************" & vbCrLf
        ''********************************************************************************************************''
        ''******************************计算**************************************''
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "卫星号PRN=" & prn & vbCrLf
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "-----------------------------" & vbCrLf
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & year & month & day & hour & minute & second
        GM = 398600500000000.0#
        a = sqrtA ^ 2
        n0 = (GM / (a ^ 3)) ^ (1 / 2) '平均角速度
        n = n0 + tn '平均角速度加上卫星电文给出的摄动改正数，便得到卫星运行的平均角速度
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "长半轴a=" & a & vbCrLf
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "平均角速度n0=" & n0 & vbCrLf
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "卫星运行的平均角速度n=" & n & vbCrLf
        ''*************************************************************************''
        toc = (((CalaWeek_(year, month, day) - 0) * 24) + hour) * 60 * 60
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "参考时刻toc=" & toc & vbCrLf '参考时刻toc
        t = toc - a0
        tk = t - toe
        If tk > 302400 Then
            tk -= 604800
        ElseIf tk < -302400 Then
            tk += 604800
        End If
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "归化时间tk=" & tk & vbCrLf '计算归化时间
        ''*************************************************************************''
        Mk = M0 + (n * tk)
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "观测时刻的卫星平近点角Mk=" & Mk & vbCrLf '观测时刻卫星平近点角Mk的计算
        ''*************************************************************************''
        i = 0
        Ek = Mk
        Do
            Ek = Mk + (e * Sin(Ek))
            tEk = Mk + e * Sin(Ek)
            i += 1
        Loop Until i > 10
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "偏近点角Ek = " & Ek & vbCrLf '计算偏近点角Ek
        ''***********************************************************************''
        Vk = Atan(Sqrt(1 - e ^ 2) * Sin(Ek) / (Cos(Ek) - e))
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "真近点角Vk=" & Vk & vbCrLf '真近点角Vk的计算
        ''***********************************************************************''
        fi0 = Vk + W
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "升交距角fi0=" & fi0 & vbCrLf '升交距角的计算
        ''***********************************************************************''
        'rk = a * (1 - (e * Cos(Ek)))
        ' Me.RichTextBox3.Text = Me.RichTextBox3.Text & "rk=" & rk & vbCrLf
        ''***********************************************************************''
        Su = (Cuc * Cos(2 * fi0)) + (Cus * Sin(2 * fi0))
        Sr = (Crc * Cos(2 * fi0)) + (Crs * Sin(2 * fi0))
        Si = (Cic * Cos(2 * fi0)) + (Cis * Sin(2 * fi0))
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "升交距角Su=" & Su & vbCrLf '升交距角
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "卫星矢距Sr=" & Sr & vbCrLf '卫星矢距
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "轨道倾角Si=" & Si & vbCrLf '轨道倾角
        ''***********************************************************************''
        uk = fi0 + Su
        rk += Sr
        ik = i0 + Si + (IDOT * tk)
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "经过摄动改正后的升交距角uk=" & uk & vbCrLf '经过摄动改正后的升交距角
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "经过摄动改正后的卫星矢距rk=" & rk & vbCrLf '经过摄动改正后的卫星矢距
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "经过摄动改正后的轨道倾角ik=" & ik & vbCrLf '经过摄动改正后的轨道倾角
        ''***********************************************************************''
        xk = rk * Cos(uk)
        yk = rk * Sin(uk)
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "卫星在轨道平面坐标系中的x坐标xk=" & xk & vbCrLf '卫星在轨道平面坐标系中的x坐标
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "卫星在轨道平面坐标系中的y坐标yk=" & yk & vbCrLf '卫星在轨道平面坐标系中的y坐标
        ''***********************************************************************''
        we = 0.0000729211567 '地球自转的速率
        tvt = OL + ((OU - we) * tk) - (we * toe)
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "观测时刻的升交点精度Ωk=" & tvt & vbCrLf '观测时刻的升交点精度Ωk
        ''***********************************************************************''
        XXk = (xk * Cos(tvt)) - yk * Cos(ik) * Sin(tvt)
        YYk = (xk * Sin(tvt)) + yk * Cos(ik) * Cos(tvt)
        ZZk = yk * Sin(ik)
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & “卫星在地心固定坐标系中的直角坐标 ：” & vbCrLf & "Xk=" & XXk & vbCrLf
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "Yk=" & YYk & vbCrLf
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "ZK=" & ZZk & vbCrLf
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "********************************************" & vbCrLf
        ''*******************************************************************************************************''
    End Function
    'Rinex 3.02
    Public Function Data_operate_3(data, j)
        Dim year%, month%, day%, hour%, minute%
        Dim second#
        If j Is Nothing Then
            Throw New ArgumentNullException(NameOf(j)) 'mach fix
        End If
        ''*******************************取参数********************************''
        Dim prn
        prn = Val(Microsoft.VisualBasic.Left(data(key_L + 1), 3)) '卫星号
        year = Val(Mid(data(key_L + 1), 5, 4)) '年
        month = Val(Mid(data(key_L + 1), 10, 2)) '月
        day = Val(Mid(data(key_L + 1), 13, 2)) '日
        hour = Val(Mid(data(key_L + 1), 16, 2)) '时
        minute = Val(Mid(data(key_L + 1), 19, 2)) '分
        second = Val(Mid(data(key_L + 1), 21, 2)) '秒
        a0 = Val(Mid(data(key_L + 1), 24, 19)) '(s)卫星时钟偏差
        a1 = Val(Mid(data(key_L + 1), 43, 19)) '(s/s)卫星时钟漂移
        a2 = Val(Mid(data(key_L + 1), 62, 19)) '(s/(s^(1/2)))卫星时钟漂移率
        ''************************************************************************''
        ''第2行
        IDOE = Val(Mid(data(key_L + 2), 6, 19)) '星历数据的有效期龄
        Crs = Val(Mid(data(key_L + 2), 24, 19)) 'crs(m)轨道半径正弦改正项
        tn = Val(Mid(data(key_L + 2), 43, 19)) '△n(rad/s)平均运动修正量
        M0 = Val(Mid(data(key_L + 2), 62, 19)) 'M0(rad)M0toe时的平近点角
        ''************************************************************************''
        ''第3行
        Cuc = Val(Mid(data(key_L + 3), 6, 19)) 'Cue(rad)纬度幅角余弦改正项
        e = Val(Mid(data(key_L + 3), 24, 19)) 'e卫星轨道偏心率
        Cus = Val(Mid(data(key_L + 3), 43, 19)) 'Cus(radians)纬度幅角正弦改正项
        sqrtA = Val(Mid(data(key_L + 3), 62, 19)) 'sqrt(A)(m^1/2)轨道长半径平根
        ''************************************************************************''
        ''第4行
        toe = Val(Mid(data(key_L + 4), 5, 19)) 'toe星历的基准时间（GPS周内的秒数）
        Cic = Val(Mid(data(key_L + 4), 24, 19)) 'Cic(rad)轨道倾角余弦调和项
        Ol0 = Val(Mid(data(key_L + 4), 43, 19)) 'Ω(rad)升交点赤经
        Cis = Val(Mid(data(key_L + 4), 62, 19)) 'Cis(rad)轨道倾角正弦项
        ''************************************************************************''
        ''第5行
        i0 = Val(Mid(data(key_L + 5), 5, 10)) 'i0(rad)轨道倾角
        Crc = Val(Mid(data(key_L + 5), 24, 19)) 'Crc(m)轨道半径余弦调和项
        W = Val(Mid(data(key_L + 5), 43, 19)) 'w（rad/s）近地点角距
        OU = Val(Mid(data(key_L + 5), 62, 19)) 'Ω（rad/s）OMEGA DOT升交点赤经变率
        ''************************************************************************''
        ''第6行
        IDOT = Val(Mid(data(key_L + 6), 5, 19)) 'i（rad/s）IDOT轨道倾角变化率
        L2Data = Val(Mid(data(key_L + 6), 24, 19)) 'L2上的码
        PS = Val(Mid(data(key_L + 6), 43, 19)) 'GPS星期数（与TOE一同表示时间），为连续计数，不是1021的余数
        L2P = Val(Mid(data(key_L + 6), 62, 19)) 'L2P码数据标志
        ''************************************************************************''
        ''第7行
        ST_PRC = Val(Mid(data(key_L + 7), 5, 19)) '卫星精度（m）
        ST_HEL = Val(Mid(data(key_L + 7), 24, 19)) '卫星健康（MSB第1子帧第3字第17~22位)
        TGD = Val(Mid(data(key_L + 7), 43, 19)) 'TGD(Sec)
        IDOC = Val(Mid(data(key_L + 7), 62, 19)) 'IODC种的数据龄期
        ''************************************************************************''
        ''第8行
        Send_Time = Val(Mid(data(key_L + 8), 5, 19)) '电文发送时刻（单位为GPS周的秒，通过交换字（HOW)中的Z计数得出）
        Countin_h = Val(Mid(data(key_L + 8), 24, 19)) '拟合区间（h），如未知则为零
        ps_1 = Val(Mid(data(key_L + 8), 43, 19)) '备用
        ps_2 = Val(Mid(data(key_L + 8), 62, 19)) '备用

        ''*************************输出各参数至text2*************************************''
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "卫星号PRN =" & prn & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "----" & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "年，月，日y,m,d=" & " " & year & " " & month & " " & day & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "时，分，秒h,min,sec=" & hour & " " & minute & " " & second & " " & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "卫星钟差a_0=" & a0 & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "卫星钟速a_1=" & a1 & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "卫星钟速变率a_2=" & a2 & vbCrLf
        ''*************************************************************************''
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "星历数据的有效期龄IDOE=" & IDOE & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & “轨道半径的正弦调和项改正的振幅C_rs=” & Crs & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "平均运动修正量△_n=" & tn & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & “参考时间的平近点角M_0=" & M0 & vbCrLf
        ''*************************************************************************''
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "纬度幅角的余弦调和项改正的振幅C_uc=" & Cuc & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "轨道偏心率e=" & e & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "纬度幅角的正弦调和项改正的振幅C_us=" & Cus & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "轨道长半径的平方根SqrtA=" & sqrtA & vbCrLf
        ''*************************************************************************''
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "星历参考时间t_oe=" & toe & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "轨道倾角的余弦调和项改正的振幅C_ic=" & Cic & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "升交点赤经Ω_0=" & Ol0 & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "轨道倾角的正弦调和项改正的振幅C_is=" & Cis & vbCrLf
        ''*************************************************************************''
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "参考时间的轨道倾角i_0=" & i0 & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "轨道半径的余弦调和项改正的振幅C_rc=" & Crc & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "近地点幅角ω=" & W & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "升交点赤经变率(Ω_0)/s=" & OU & vbCrLf
        ''*************************************************************************''
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "轨道倾角变化率IDOT=" & IDOT & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "L2上的码L2Data=" & L2Data & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "GPS星期数ps=" & PS & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "码数据标志L2P=" & L2P & vbCrLf
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "******************************************" & vbCrLf
        ''********************************************************************************************************''
        ''******************************计算**************************************''
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "卫星号PRN=" & prn & vbCrLf
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "-----------------------------" & vbCrLf
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "时间" & year & month & day & hour & minute & second & vbCrLf
        GM = 398600500000000.0#
        a = sqrtA ^ 2
        n0 = (GM / (a ^ 3)) ^ (1 / 2) '平均角速度
        n = n0 + tn '平均角速度加上卫星电文给出的摄动改正数，便得到卫星运行的平均角速度
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "长半轴a=" & a & vbCrLf
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "卫星平均角速度n_0=" & n0 & vbCrLf
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "改正后的平均角速度n=" & n & vbCrLf
        ''*************************************************************************''
        ' Dim UT
        ' UT = hour + (minute / 60) + (second / 3600)
        ' Dim JD
        ' Dim year1, month1, day1
        ' If month <= 2 Then
        '      year1 = year - 1 And month1 = month + 12
        '  ElseIf month > 2 Then
        '      year1 = year And month1 = month
        ' End If

        '  JD = Int(365.25 * year1) + Int(30.6001 * (month1 + 1)) + day1 + (UT / 24) + 1720981.5
        '  Dim gpsweek
        '  gpsweek = Int(JD - 2444244.5) / 7
        '  Dim gpssecond
        ' gpssecond = (JD - 2444244.5 - gpsweek * 7) * 24 * 3600

        '  toc = gpssecond
        ''toc = (((CalaWeek_(year, month, day) - 0) * 24) + hour) * 60 * 60
        '  Me.RichTextBox3.Text = Me.RichTextBox3.Text & "参考时刻toc=" & toc & vbCrLf '参考时刻toc
        '  t = toc - a0
        ' tk = t - toe
        '  If tk > 302400 Then
        '  tk -= 604800
        '  ElseIf tk < -302400 Then
        ' tk += 604800
        '  End If
        tk = 0
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "归化时间tk=" & tk & vbCrLf '计算归化时间
        ''*************************************************************************''
        Mk = M0 + (n * tk)
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "观测时刻的卫星平近点角Mk=" & Mk & vbCrLf '观测时刻卫星平近点角Mk的计算
        ''*************************************************************************''
        i = 0
        Ek = Mk
        Do
            Ek = Mk + (e * Sin(Ek))
            tEk = Mk + e * Sin(Ek)
            i += 1
        Loop Until i > 10
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "偏近点角Ek = " & Ek & vbCrLf '计算偏近点角Ek
        ''***********************************************************************''
        Vk = Atan(Sqrt(1 - e ^ 2) * Sin(Ek) / (Cos(Ek) - e))
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "真近点角Vk=" & Vk & vbCrLf '真近点角Vk的计算
        ''***********************************************************************''
        fi0 = Vk + W
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "升交距角fi0=" & fi0 & vbCrLf '升交距角的计算
        ''***********************************************************************''
        Su = (Cuc * Cos(2 * fi0)) + (Cus * Sin(2 * fi0))
        Sr = (Crc * Cos(2 * fi0)) + (Crs * Sin(2 * fi0))
        Si = (Cic * Cos(2 * fi0)) + (Cis * Sin(2 * fi0))
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "升交距角Su=" & Su & vbCrLf '升交距角
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "卫星矢距Sr=" & Sr & vbCrLf '卫星矢距
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "轨道倾角Si=" & Si & vbCrLf '轨道倾角
        ''***********************************************************************''
        uk = fi0 + Su
        rk = Sr + sqrtA ^ 2 * (1 - e * Cos(Ek))
        ik = i0 + Si + (IDOT * tk)
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "经过摄动改正后的升交距角uk=" & uk & vbCrLf '经过摄动改正后的升交距角
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "经过摄动改正后的卫星矢距rk=" & rk & vbCrLf '经过摄动改正后的卫星矢距
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "经过摄动改正后的轨道倾角ik=" & ik & vbCrLf '经过摄动改正后的轨道倾角
        ''***********************************************************************''
        xk = rk * Cos(uk)
        yk = rk * Sin(uk)
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "卫星在轨道平面坐标系中的x坐标xk=" & xk & vbCrLf '卫星在轨道平面坐标系中的x坐标
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "卫星在轨道平面坐标系中的y坐标yk=" & yk & vbCrLf '卫星在轨道平面坐标系中的y坐标
        ''***********************************************************************''
        we = 0.0000729211567 'BDCS坐标系下的地球自转角速度
        tk = 0 'test
        'OU Ω（rad/s）OMEGA DOT升交点赤经变率
        'Tk观测历元到参考历元的时间差
        tvt = Ol0 + ((OU - we) * tk) - (we * toe)
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "观测时刻的升交点精度Ωk=" & tvt & vbCrLf '观测时刻的升交点精度Ωk
        ''***********************************************************************''
        XXk = (xk * Cos(tvt)) - yk * Cos(ik) * Sin(tvt)
        YYk = (xk * Sin(tvt)) + yk * Cos(ik) * Cos(tvt)
        ZZk = yk * Sin(ik)
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & “卫星在地心固定坐标系中的直角坐标 ：” & vbCrLf & "Xk=" & XXk & vbCrLf
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "Yk=" & YYk & vbCrLf
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "ZK=" & ZZk & vbCrLf
        Me.RichTextBox3.Text = Me.RichTextBox3.Text & "********************************************" & vbCrLf
        ''*******************************************************************************************************''
    End Function
    '关闭Excel（未启用）
    Private Sub Button7_Click(sender As Object, e As EventArgs)
        If Dir(Application.StartupPath & "\temp\excel.bz") <> "" Then '由VB关闭EXCEL   
            xlApp.Application.DisplayAlerts = False '关闭EXCEL的警告提示，不然用程序关闭时会有警报提示，还要手动去确定。
            xlBook.RunAutoMacros(Excel.XlRunAutoMacro.xlAutoClose) '执行EXCEL关闭宏  
            xlBook.Close(True) '关闭EXCEL工作簿  
            xlApp.Quit() '关闭EXCEL  
        End If
        xlApp = Nothing '释放EXCEL对象  
    End Sub
    '保存卫星参数为txt格式
    Public Sub 保存卫星参数为txt格式(sender As Object, e As EventArgs) Handles SaveCsAsText.Click
        Dim NewFileName As String
        NewFileName = Year(Now()) & Format(Month(Now), "00") & Format(Day(Now), "00") & Format(Hour(Now()), "00") & Format(Minute(Now()), "00")
        With SaveFileDialog1
            .DefaultExt = "txt"
            .FileName = "CS" & NewFileName
            .Filter = "文本文件(*.txt)|*.txt|所有文件(*.*)|*.*"
            .FilterIndex = 1
            .InitialDirectory = Application.StartupPath & "\temp"
            .OverwritePrompt = True
            .Title = "Save File Dialog"
        End With
        Dim tip As Integer
        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim strfilename As String
            strfilename = SaveFileDialog1.FileName
            Dim objwriter As StreamWriter = New StreamWriter(strfilename, False)
            objwriter.Write(RichTextBox2.Text)
            objwriter.Close()
        ElseIf SaveFileDialog1.ShowDialog() = DialogResult.Cancel Then
            RichTextBox4.Text = RichTextBox4.Text & vbCrLf & "用户取消了参数输出为txt"
        End If
    End Sub
    '**********************************************************解算数据入口***************************************************************
    '开始解算
    Private Sub 解算历书文件ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 解算历书文件ToolStripMenuItem.Click， 解算历书文件ToolStripMenuItem1.Click
        ''调用数据处理过程，处理所有卫星的位置
        Try
            DefException(UBound(Data))
        Catch ex As Exception
            MsgBox("发生了异常：" & ex.ToString & vbCrLf & "异常来源于：" & ex.Source & vbCrLf & "提示信息：“ & ex.Message)
            RichTextBox4.Text = RichTextBox4.Text & "未读取数据，请选择感兴趣的历书文件"
        End Try

        For j = key_L To UBound(Data) - 1 Step 8 'ubound,求数组某维上界函数 step 这里步长为8，即中间隔7个数组
            If ToolStripTextBox1.Text = "2.10" Then Call Data_operate_1(Data, j)
            If ToolStripTextBox1.Text = "2.11" Then Call Data_operate_2(Data, j)
            If ToolStripTextBox1.Text = "3.02" Then Call Data_operate_3(Data, j)
        Next j

        If ToolStripTextBox1.Text = "2.10" Then
            RichTextBox4.Text = RichTextBox4.Text & vbCrLf & "文件：" & OpenFileDialog1.FileName & "解算成功”
        ElseIf ToolStripTextBox1.Text = "2.11" Then
            RichTextBox4.Text = RichTextBox4.Text & vbCrLf & "文件：" & OpenFileDialog2.FileName & "解算成功”
        ElseIf ToolStripTextBox1.Text = "3.02" Then
            RichTextBox4.Text = RichTextBox4.Text & vbCrLf & "文件：" & OpenFileDialog3.FileName & "解算成功”
        End If
        'for next循环：for 循环变量=初值to终值[step 步长]
        '循环体
        'next 循环变量
    End Sub
    Private Sub ToolStripProgressBar1_TextChanged(sender As Object, e As EventArgs) Handles ToolStripProgressBar1.TextChanged
        If ToolStripProgressBar1.Text = "2.10" Then RichTextBox4.Text = RichTextBox4.Text & vbCrLf & "文件：" & OpenFileDialog1.FileName & "解算成功”
        If ToolStripProgressBar1.Text = "2.11" Then RichTextBox4.Text = RichTextBox4.Text & vbCrLf & "文件：" & OpenFileDialog2.FileName & "解算成功”
        If ToolStripProgressBar1.Text = "3.02" Then RichTextBox4.Text = RichTextBox4.Text & vbCrLf & "文件：" & OpenFileDialog3.FileName & "解算成功”
    End Sub
    '输出参数到表格
    Private Sub 表格ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveCsAsXlsm.Click

        ''调用数据处理过程，处理所有卫星的位置
        Try
            DefException(UBound(Data))
        Catch ex As Exception
            MsgBox("发生了异常：" & ex.ToString & vbCrLf & "异常来源于：" & ex.Source & vbCrLf & "提示信息：“ & ex.Message)
        End Try

        Dim NewFileName As String
        NewFileName = Year(Now()) & Format(Month(Now), "00") & Format(Day(Now), "00") & Format(Hour(Now()), "00") & Format(Minute(Now()), "00")
        If Dir(Application.StartupPath & "\temp\excel.bz") = "" Then '判断EXCEL是否打开
            My.Computer.FileSystem.CopyFile(
Application.StartupPath & "\temp\工作簿1.xlsm",
Application.StartupPath & "\temp\CS" & NewFileName & “.xlsm",
FileIO.UIOption.OnlyErrorDialogs,
FileIO.UICancelOption.DoNothing)
            xlApp = CreateObject("Excel.Application") '创建EXCEL应用类 
            xlApp.Visible = True '设置EXCEL可见 (xlApp.Visible = False '设置EXCEL打开时不可见 ) 
            xlBook = xlApp.Workbooks.Open(Application.StartupPath & "\temp\CS" & NewFileName & “.xlsm") '打开EXCEL工作簿  
            xlsheet = xlBook.Worksheets(1) '打开EXCEL工作表 
            xlsheet.Activate() '激活工作表  
            xlsheet.Cells(1, 1) = "卫星号PRN"
            xlsheet.Cells(1, 2) = "年Y"
            xlsheet.Cells(1, 3) = "月M"
            xlsheet.Cells(1, 4) = "日D"
            xlsheet.Cells(1, 5) = "时H"
            xlsheet.Cells(1, 6) = "分Min"
            xlsheet.Cells(1, 7) = "秒Sec"
            xlsheet.Cells(1, 8) = "卫星时钟偏差a0"
            xlsheet.Cells(1, 9) = "卫星时钟漂移a1"
            xlsheet.Cells(1, 10) = "卫星时钟漂移率a2"
            xlsheet.Cells(1, 11) = "星历数据的有效期龄IDOE"
            xlsheet.Cells(1, 12) = "平均运动修正量tn"
            xlsheet.Cells(1, 13) = "按参考历元计算的平近点角M0"
            xlsheet.Cells(1, 14) = "纬度幅角的余弦调和项改正的振幅Cuc"
            xlsheet.Cells(1, 15) = "轨道偏心率e"
            xlsheet.Cells(1, 16) = "纬度幅角的正弦调和项改正的振幅Cus"
            xlsheet.Cells(1, 17) = "轨道长半径的平方根SqrtA"
            xlsheet.Cells(1, 18) = "星历的基准时间（GPS周内的秒数）toe"
            xlsheet.Cells(1, 19) = "轨道倾角的余弦调和项改正的振幅Cic"
            xlsheet.Cells(1, 20) = "升交点赤经Ω"
            xlsheet.Cells(1, 21) = "轨道倾角的正弦调和项改正的振幅Cis"
            xlsheet.Cells(1, 22) = "按参考历元计算的轨道倾角i0"
            xlsheet.Cells(1, 23) = "轨道半径的余弦调和项改正的振幅Crc"
            xlsheet.Cells(1, 24) = "近地点角距w"
            xlsheet.Cells(1, 25) = "升交点赤经变率OMEGA DOT"
            xlsheet.Cells(1, 26) = "轨道倾角变化率IDOT"
            xlsheet.Cells(1, 27) = "L2上的码L2Data"
            xlsheet.Cells(1, 28) = "GPS星期数ps"
            xlsheet.Cells(1, 29) = "码数据标志L2P"
        Else : MsgBox("EXCEL已打开")
        End If
        xlBook.RunAutoMacros(Excel.XlRunAutoMacro.xlAutoOpen) '运行EXCEL中的启动宏  
        Me.WindowState = FormWindowState.Minimized

        For j = key_L To UBound(Data) - 1 Step 8 'ubound,求数组某维上界函数 step 这里步长为8，即中间隔7个数组

            If ToolStripTextBox1.Text = "2.10" Or ToolStripTextBox1.Text = "2.11" Then Call Data_operate_save_CS210(Data, j)
            If ToolStripTextBox1.Text = "3.02" Then Call Data_operate_save_CS302(Data, j)

        Next j
        'for next循环：for 循环变量=初值to终值[step 步长]
        '循环体
        'next 循环变量
    End Sub
    '保存210和211的参数
    Public Function Data_operate_save_CS210(data, j)
        Dim year0%, month0%, day0%, hour0%, minute0%
        Dim second0#
        If j Is Nothing Then
            Throw New ArgumentNullException(NameOf(j)) 'mach fix
        End If
        ''*******************************取参数********************************''
        Dim prn
        prn = Val(Microsoft.VisualBasic.Left(data(key_L + 1), 2)) '卫星号
        year0 = Val(Mid(data(key_L + 1), 4, 2)) '年
        month0 = Val(Mid(data(key_L + 1), 7, 2)) '月
        day0 = Val(Mid(data(key_L + 1), 10, 2)) '日
        hour0 = Val(Mid(data(key_L + 1), 13, 2)) '时
        minute0 = Val(Mid(data(key_L + 1), 16, 2)) '分
        second0 = Val(Mid(data(key_L + 1), 18, 5)) '秒
        a0 = Val(Mid(data(key_L + 1), 24, 19)) '(s)卫星时钟偏差
        a1 = Val(Mid(data(key_L + 1), 43, 19)) '(s/s)卫星时钟漂移
        a2 = Val(Mid(data(key_L + 1), 62, 19)) '(s/(s^(1/2)))卫星时钟漂移率
        ''************************************************************************''
        ''第2行
        IDOE = Val(Mid(data(key_L + 2), 5, 19)) '星历数据的有效期龄
        Crs = Val(Mid(data(key_L + 2), 24, 19)) 'crs(m)轨道半径正弦改正项
        tn = Val(Mid(data(key_L + 2), 43, 19)) '△n(rad/s)平均运动修正量
        M0 = Val(Mid(data(key_L + 2), 62, 19)) 'M0(rad)M0toe时的平近点角
        ''************************************************************************''
        ''第3行
        Cuc = Val(Mid(data(key_L + 3), 5, 19)) 'Cue(rad)纬度幅角余弦改正项
        e = Val(Mid(data(key_L + 3), 24, 19)) 'e卫星轨道偏心率
        Cus = Val(Mid(data(key_L + 3), 43, 19)) 'Cus(radians)纬度幅角正弦改正项
        sqrtA = Val(Mid(data(key_L + 3), 62, 19)) 'sqrt(A)(m^1/2)轨道长半径平根
        ''************************************************************************''
        ''第4行
        toe = Val(Mid(data(key_L + 4), 5, 19)) 'toe星历的基准时间（GPS周内的秒数）
        Cic = Val(Mid(data(key_L + 4), 24, 19)) 'Cic(rad)轨道倾角余弦调和项
        Ol0 = Val(Mid(data(key_L + 4), 43, 19)) 'Ω(rad)升交点赤经
        Cis = Val(Mid(data(key_L + 4), 62, 19)) 'Cis(rad)轨道倾角正弦项
        ''************************************************************************''
        ''第5行
        i0 = Val(Mid(data(key_L + 5), 5, 19)) 'i0(rad)轨道倾角
        Crc = Val(Mid(data(key_L + 5), 24, 19)) 'Crc(m)轨道半径余弦调和项
        W = Val(Mid(data(key_L + 5), 43, 19)) 'w（rad/s）近地点角距
        OU = Val(Mid(data(key_L + 5), 62, 19)) 'Ω（rad/s）OMEGA DOT升交点赤经变率
        ''************************************************************************''
        ''第6行
        IDOT = Val(Mid(data(key_L + 6), 5, 19)) 'i（rad/s）IDOT轨道倾角变化率
        L2Data = Val(Mid(data(key_L + 6), 24, 19)) 'L2上的码
        PS = Val(Mid(data(key_L + 6), 43, 19)) 'GPS星期数（与TOE一同表示时间），为连续计数，不是1021的余数
        L2P = Val(Mid(data(key_L + 6), 62, 19)) 'L2P码数据标志
        ''************************************************************************''
        ''第7行
        ST_PRC = Val(Mid(data(key_L + 7), 5, 19)) '卫星精度（m）
        ST_HEL = Val(Mid(data(key_L + 7), 24, 19)) '卫星健康（MSB第1子帧第3字第17~22位)
        TGD = Val(Mid(data(key_L + 7), 43, 19)) 'TGD(Sec)
        IDOC = Val(Mid(data(key_L + 7), 62, 19)) 'IODC种的数据龄期
        ''************************************************************************''
        ''第8行
        Send_Time = Val(Mid(data(key_L + 8), 5, 19)) '电文发送时刻（单位为GPS周的秒，通过交换字（HOW)中的Z计数得出）
        Countin_h = Val(Mid(data(key_L + 8), 24, 19)) '拟合区间（h），如未知则为零
        ps_1 = Val(Mid(data(key_L + 8), 43, 19)) '备用
        ps_2 = Val(Mid(data(key_L + 8), 62, 19)) '备用
        Dim jj As Integer
        Dim zz As Integer
        jj = j
        If jj = 4 Then
            jj -= 3
            zz = jj
        ElseIf ((j - 4) Mod 8) = 0 Then
            zz = (j - 4) / 8
        ElseIf ((j - 4) Mod 8) <> 0 Then
            MsgBox(“tip”, "j:" & j & “zz:” & zz)
        End If
        Dim z As Integer
        z = 1
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "卫星号PRN =" & prn & vbCrLf
        xlsheet.Cells(zz + 1, z) = prn

        z += 1
        xlsheet.Cells(zz + 1, z) = year0
        z += 1
        xlsheet.Cells(zz + 1, z) = month0
        z += 1
        xlsheet.Cells(zz + 1, z) = day0
        z += 1
        xlsheet.Cells(zz + 1, z) = hour0
        z += 1
        xlsheet.Cells(zz + 1, z) = minute0
        z += 1
        xlsheet.Cells(zz + 1, z) = second0
        z += 1
        xlsheet.Cells(zz + 1, z) = a0
        z += 1
        xlsheet.Cells(zz + 1, z) = a1
        z += 1
        xlsheet.Cells(zz + 1, z) = a2
        z += 1
        xlsheet.Cells(zz + 1, z) = IDOE
        z += 1
        xlsheet.Cells(zz + 1, z) = tn
        z += 1
        xlsheet.Cells(zz + 1, z) = M0
        z += 1
        xlsheet.Cells(zz + 1, z) = Cuc
        z += 1
        xlsheet.Cells(zz + 1, z) = e
        z += 1
        xlsheet.Cells(zz + 1, z) = Cus
        z += 1
        xlsheet.Cells(zz + 1, z) = sqrtA
        z += 1
        xlsheet.Cells(zz + 1, z) = toe
        z += 1
        xlsheet.Cells(zz + 1, z) = Cic
        z += 1
        xlsheet.Cells(zz + 1, z) = Cic
        z += 1
        xlsheet.Cells(zz + 1, z) = Cic
        z += 1
        xlsheet.Cells(zz + 1, z) = i0
        z += 1
        xlsheet.Cells(zz + 1, z) = Crc
        z += 1
        xlsheet.Cells(zz + 1, z) = W
        z += 1
        xlsheet.Cells(zz + 1, z) = OU
        z += 1
        xlsheet.Cells(zz + 1, z) = IDOT
        z += 1
        xlsheet.Cells(zz + 1, z) = L2Data
        z += 1
        xlsheet.Cells(zz + 1, z) = PS
        z += 1
        xlsheet.Cells(zz + 1, z) = L2P
    End Function
    Public Function Data_operate_save_CS302(data, j)
        Dim year0%, month0%, day0%, hour0%, minute0%
        Dim second0#
        If j Is Nothing Then
            Throw New ArgumentNullException(NameOf(j)) 'mach fix
        End If
        ''*******************************取参数********************************''
        Dim prn
        prn = Val(Microsoft.VisualBasic.Left(data(key_L + 1), 2)) '卫星号
        year0 = Val(Mid(data(key_L + 1), 4, 2)) '年
        month0 = Val(Mid(data(key_L + 1), 7, 2)) '月
        day0 = Val(Mid(data(key_L + 1), 10, 2)) '日
        hour0 = Val(Mid(data(key_L + 1), 13, 2)) '时
        minute0 = Val(Mid(data(key_L + 1), 16, 2)) '分
        second0 = Val(Mid(data(key_L + 1), 18, 5)) '秒
        a0 = Val(Mid(data(key_L + 1), 24, 19)) '(s)卫星时钟偏差
        a1 = Val(Mid(data(key_L + 1), 43, 19)) '(s/s)卫星时钟漂移
        a2 = Val(Mid(data(key_L + 1), 62, 19)) '(s/(s^(1/2)))卫星时钟漂移率
        ''************************************************************************''
        ''第2行
        IDOE = Val(Mid(data(key_L + 2), 5, 19)) '星历数据的有效期龄
        Crs = Val(Mid(data(key_L + 2), 24, 19)) 'crs(m)轨道半径正弦改正项
        tn = Val(Mid(data(key_L + 2), 43, 19)) '△n(rad/s)平均运动修正量
        M0 = Val(Mid(data(key_L + 2), 62, 19)) 'M0(rad)M0toe时的平近点角
        ''************************************************************************''
        ''第3行
        Cuc = Val(Mid(data(key_L + 3), 5, 19)) 'Cue(rad)纬度幅角余弦改正项
        e = Val(Mid(data(key_L + 3), 24, 19)) 'e卫星轨道偏心率
        Cus = Val(Mid(data(key_L + 3), 43, 19)) 'Cus(radians)纬度幅角正弦改正项
        sqrtA = Val(Mid(data(key_L + 3), 62, 19)) 'sqrt(A)(m^1/2)轨道长半径平根
        ''************************************************************************''
        ''第4行
        toe = Val(Mid(data(key_L + 4), 5, 19)) 'toe星历的基准时间（GPS周内的秒数）
        Cic = Val(Mid(data(key_L + 4), 24, 19)) 'Cic(rad)轨道倾角余弦调和项
        Ol0 = Val(Mid(data(key_L + 4), 43, 19)) 'Ω(rad)升交点赤经
        Cis = Val(Mid(data(key_L + 4), 62, 19)) 'Cis(rad)轨道倾角正弦项
        ''************************************************************************''
        ''第5行
        i0 = Val(Mid(data(key_L + 5), 5, 19)) 'i0(rad)轨道倾角
        Crc = Val(Mid(data(key_L + 5), 24, 19)) 'Crc(m)轨道半径余弦调和项
        W = Val(Mid(data(key_L + 5), 43, 19)) 'w（rad/s）近地点角距
        OU = Val(Mid(data(key_L + 5), 62, 19)) 'Ω（rad/s）OMEGA DOT升交点赤经变率
        ''************************************************************************''
        ''第6行
        IDOT = Val(Mid(data(key_L + 6), 5, 19)) 'i（rad/s）IDOT轨道倾角变化率
        L2Data = Val(Mid(data(key_L + 6), 24, 19)) 'L2上的码
        PS = Val(Mid(data(key_L + 6), 43, 19)) 'GPS星期数（与TOE一同表示时间），为连续计数，不是1021的余数
        L2P = Val(Mid(data(key_L + 6), 62, 19)) 'L2P码数据标志
        ''************************************************************************''
        ''第7行
        ST_PRC = Val(Mid(data(key_L + 7), 5, 19)) '卫星精度（m）
        ST_HEL = Val(Mid(data(key_L + 7), 24, 19)) '卫星健康（MSB第1子帧第3字第17~22位)
        TGD = Val(Mid(data(key_L + 7), 43, 19)) 'TGD(Sec)
        IDOC = Val(Mid(data(key_L + 7), 62, 19)) 'IODC种的数据龄期
        ''************************************************************************''
        ''第8行
        Send_Time = Val(Mid(data(key_L + 8), 5, 19)) '电文发送时刻（单位为GPS周的秒，通过交换字（HOW)中的Z计数得出）
        Countin_h = Val(Mid(data(key_L + 8), 24, 19)) '拟合区间（h），如未知则为零
        ps_1 = Val(Mid(data(key_L + 8), 43, 19)) '备用
        ps_2 = Val(Mid(data(key_L + 8), 62, 19)) '备用
        Dim jj As Integer
        Dim zz As Integer
        jj = j
        If jj = 4 Then
            jj -= 3
            zz = jj
        ElseIf ((j - 4) Mod 8) = 0 Then
            zz = (j - 4) / 8
        ElseIf ((j - 4) Mod 8) <> 0 Then
            MsgBox(“tip”, "j:" & j & vbCrLf & “zz:” & zz)
        End If
        Dim z As Integer
        z = 1
        Me.RichTextBox2.Text = Me.RichTextBox2.Text & "卫星号PRN =" & prn & vbCrLf
        xlsheet.Cells(zz + 1, z) = prn

        z += 1
        xlsheet.Cells(zz + 1, z) = year0
        z += 1
        xlsheet.Cells(zz + 1, z) = month0
        z += 1
        xlsheet.Cells(zz + 1, z) = day0
        z += 1
        xlsheet.Cells(zz + 1, z) = hour0
        z += 1
        xlsheet.Cells(zz + 1, z) = minute0
        z += 1
        xlsheet.Cells(zz + 1, z) = second0
        z += 1
        xlsheet.Cells(zz + 1, z) = a0
        z += 1
        xlsheet.Cells(zz + 1, z) = a1
        z += 1
        xlsheet.Cells(zz + 1, z) = a2
        z += 1
        xlsheet.Cells(zz + 1, z) = IDOE
        z += 1
        xlsheet.Cells(zz + 1, z) = tn
        z += 1
        xlsheet.Cells(zz + 1, z) = M0
        z += 1
        xlsheet.Cells(zz + 1, z) = Cuc
        z += 1
        xlsheet.Cells(zz + 1, z) = e
        z += 1
        xlsheet.Cells(zz + 1, z) = Cus
        z += 1
        xlsheet.Cells(zz + 1, z) = sqrtA
        z += 1
        xlsheet.Cells(zz + 1, z) = toe
        z += 1
        xlsheet.Cells(zz + 1, z) = Cic
        z += 1
        xlsheet.Cells(zz + 1, z) = Cic
        z += 1
        xlsheet.Cells(zz + 1, z) = Cic
        z += 1
        xlsheet.Cells(zz + 1, z) = i0
        z += 1
        xlsheet.Cells(zz + 1, z) = Crc
        z += 1
        xlsheet.Cells(zz + 1, z) = W
        z += 1
        xlsheet.Cells(zz + 1, z) = OU
        z += 1
        xlsheet.Cells(zz + 1, z) = IDOT
        z += 1
        xlsheet.Cells(zz + 1, z) = L2Data
        z += 1
        xlsheet.Cells(zz + 1, z) = PS
        z += 1
        xlsheet.Cells(zz + 1, z) = L2P
    End Function
    '退出主界面
    Private Sub 退出ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 退出ToolStripMenuItem.Click
        Me.Close()
    End Sub
    '打开Rinex2.10文件
    Public Function Rinex201ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Rinex210ToolStripMenuItem.Click， Rinex210ToolStripMenuItem1.Click
        ToolStripProgressBar1.Value = 0
        ToolStripProgressBar1.Minimum = 0
        ToolStripProgressBar1.Maximum = 5
        ' Threading.Thread.Sleep(100)
        On Error GoTo errhandler '设置过滤器
        OpenFileDialog1.Filter = "Renix2.10(*.97N)|*.97N" '指定过滤器
        OpenFileDialog1.FilterIndex = 1 '显示打开对话框
        OpenFileDialog1.ShowDialog() '调用打开文件的过程
        Dim FileNo_
        FileNo_ = FreeFile()
        'open OpenFileDialog.FileName For Input As FileNo_
        FileOpen(FileNo_, OpenFileDialog1.FileName, OpenMode.Input) 'fileopen（文件号，文件名，模式）
        '模式分三种：  OpenMode.Output : 对文件进行写操作。若文件已经存在，则文件中所有内容将被删除
        'openmode.input：对文件进行读操作
        'openmode.append:在文件末尾追加记录
        Dim count_1 As Integer
        Dim linedata As String
        count_1 = 1                                 ''数组0下标弃用
        Do While Not EOF(1) '读到文件尾
            ' LineInput #FileNo_, linedata
            linedata = LineInput(FileNo_)
            ReDim Preserve Data(count_1) 'As String
            Data(count_1) = linedata                ''文件内容存放在data一维数组中
            RichTextBox1.Text = RichTextBox1.Text + Data(count_1) + vbCrLf
            count_1 += 1
        Loop

        FileClose(FileNo_) '关闭fileNo_文件
        ''*************************************************************************''
        ''找到头文件截至行
        ''判断 end of header 并提取出行编号 然后对之后的数据以每8行为一颗卫星提取数据
        For i = 1 To 8
            ToolStripProgressBar1.Value = i
            key_Head = InStr(Data(i), "END OF HEADER") 'instr(start,String1,String2,Compare)
            'Start 可选项。数值表达式，设置每个搜索的起始位置。如果省略该参数，则从第一个字符位置开始搜索。起始索引从一开始。
            'String1 必选项。在哪找。
            'String2 必选项。找啥。
            'compare 参数设置包括：
            ' 常量  值 说明
            'Binary 0  执行二进制比较
            'Text   1  执行文本比较
            If key_Head <> 0 Then
                key_L = i
            End If
        Next i
        ''***********************************************''
errhandler:
    End Function
    '打开Rinex2.11文件
    Public Function Renix211ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Rinex211ToolStripMenuItem.Click, Rinex211ToolStripMenuItem1.Click

        ToolStripProgressBar1.Value = 0
        ToolStripProgressBar1.Minimum = 0
        ToolStripProgressBar1.Maximum = 5
        System.Threading.Thread.Sleep(100)

        On Error GoTo errhandler '设置过滤器
        OpenFileDialog2.Filter = "Renix2.11(*.19N)|*.19N" '指定过滤器
        OpenFileDialog2.FilterIndex = 1 '显示打开对话框
        OpenFileDialog2.ShowDialog() '调用打开文件的过程
errhandler:'用户点击取消按钮

        Dim FileNo_
        FileNo_ = FreeFile()
        'open OpenFileDialog.FileName For Input As FileNo_
        FileOpen(FileNo_, OpenFileDialog2.FileName, OpenMode.Input) 'fileopen（文件号，文件名，模式）
        '模式分三种：  OpenMode.Output : 对文件进行写操作。若文件已经存在，则文件中所有内容将被删除
        'openmode.input：对文件进行读操作
        'openmode.append:在文件末尾追加记录
        Dim count_1 As Integer
        Dim linedata As String
        count_1 = 1                                 ''数组0下标弃用
        Do While Not EOF(1) '读到文件尾
            ' LineInput #FileNo_, linedata
            linedata = LineInput(FileNo_)
            ReDim Preserve Data(count_1) 'As String
            Data(count_1) = linedata                ''文件内容存放在data一维数组中
            RichTextBox1.Text = RichTextBox1.Text + Data(count_1) + vbCrLf
            count_1 += 1
        Loop
        Me.RichTextBox4.Text = Me.RichTextBox4.Text & vbCrLf & "文件：" & OpenFileDialog1.FileName & “读取成功”
        FileClose(FileNo_) '关闭fileNo_文件
        ''*************************************************************************''
        ''找到头文件截至行
        ''判断 end of header 并提取出行编号 然后对之后的数据以每8行为一颗卫星提取数据
        For i = 1 To 8
            ToolStripProgressBar1.Value = i
            key_Head = InStr(Data(i), "END OF HEADER") 'instr(start,String1,String2,Compare)
            'Start 可选项。数值表达式，设置每个搜索的起始位置。如果省略该参数，则从第一个字符位置开始搜索。起始索引从一开始。
            'String1 必选项。在哪找。
            'String2 必选项。找啥。
            'compare 参数设置包括：
            ' 常量  值 说明
            'Binary 0  执行二进制比较
            'Text   1  执行文本比较
            If key_Head <> 0 Then
                key_L = i
            End If
        Next i
        ''***********************************************''
    End Function
    'TEST课本实例
    Public Function Renix302ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Rinex302ToolStripMenuItem.Click， Rinex302ToolStripMenuItem1.Click
        'ToolStripProgressBar1.Value = 0
        'ToolStripProgressBar1.Minimum = 0
        'ToolStripProgressBar1.Maximum = 5
        ' System.Threading.Thread.Sleep(100)
        On Error GoTo errhandler '设置过滤器
        OpenFileDialog3.Filter = "Renix3.02(*.17P)|*.17P|Renix2.10(*.97N)|*.97N" '指定过滤器
        OpenFileDialog3.FilterIndex = 1 '显示打开对话框
        OpenFileDialog3.ShowDialog() '调用打开文件的过程
        Dim FileNo_
        FileNo_ = FreeFile()
        'open OpenFileDialog.FileName For Input As FileNo_
        FileOpen(FileNo_, OpenFileDialog3.FileName, OpenMode.Input) 'fileopen（文件号，文件名，模式）
        '模式分三种：  OpenMode.Output : 对文件进行写操作。若文件已经存在，则文件中所有内容将被删除
        'openmode.input：对文件进行读操作
        'openmode.append:在文件末尾追加记录
        Dim count_1 As Integer
        Dim linedata As String
        count_1 = 1                                 ''数组0下标弃用
        Do While Not EOF(1) '读到文件尾
            ' LineInput #FileNo_, linedata
            linedata = LineInput(FileNo_)
            ReDim Preserve Data(count_1) 'As String
            Data(count_1) = linedata                ''文件内容存放在data一维数组中
            RichTextBox1.Text = RichTextBox1.Text + Data(count_1) + vbCrLf
            count_1 += 1
        Loop
        Me.RichTextBox4.Text = Me.RichTextBox4.Text & vbCrLf & "文件：" & OpenFileDialog3.FileName & “读取成功”
        FileClose(FileNo_) '关闭fileNo_文件
        ''*************************************************************************''
        ''找到头文件截至行
        ''判断 end of header 并提取出行编号 然后对之后的数据以每8行为一颗卫星提取数据
        For i = 1 To 8
            'Data() As String                '存放原参数字符
            'key_Head%                       '存放查找头文件结束符的返回值
            'key_L%                          '存放头文件结束符所在的行号

            key_Head = InStr(Data(i), "END OF HEADER") 'instr(start,String1,String2,Compare)
            'Start 可选项。数值表达式，设置每个搜索的起始位置。如果省略该参数，则从第一个字符位置开始搜索。起始索引从一开始。
            'String1 必选项。在哪找。
            'String2 必选项。找啥。
            'compare 参数设置包括：
            ' 常量  值 说明
            'Binary 0  执行二进制比较
            'Text   1  执行文本比较
            If key_Head <> 0 Then
                key_L = i
            End If
        Next i
        ''***********************************************''
errhandler:'用户点击取消按钮
    End Function
    '保存计算结果到文本文档
    Private Sub 文本ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveRuAsText.Click
        With SaveFileDialog1
            .DefaultExt = "txt"
            .FileName = "计算结果"
            .Filter = "文本文件(*.txt)|*.txt|所有文件(*.*)|*.*"
            .FilterIndex = 1
            .InitialDirectory = "d:\卫星位置计算1.1"
            .OverwritePrompt = True
            .Title = "Save File Dialog"
        End With
        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim strfilename As String
            strfilename = SaveFileDialog1.FileName
            Dim objwriter As StreamWriter = New StreamWriter(strfilename, False)
            objwriter.Write(Me.RichTextBox3.Text)
            objwriter.Close()
        End If
    End Sub
    '呼出“手动输入界面”
    Private Sub 手动输入卫星参数ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 手动输入卫星参数ToolStripMenuItem.Click， 手动输入卫星参数ToolStripMenuItem1.Click
        Form3.Show()
        Form3.StartPosition = 1
    End Sub

    '保存结果到表格
    Private Sub SaveRuAsXlsm_Click(sender As Object, e As EventArgs) Handles SaveRuAsXlsm.Click
        '   On Error GoTo errhandler '设置过滤器
        Dim NewFileName As String
        NewFileName = Year(Now()) & Format(Month(Now), "00") & Day(Now()) & Format(Hour(Now()), "00") & Format(Minute(Now()), "00")
        If Dir(Application.StartupPath & "\temp\excel.bz") = "" Then '判断EXCEL是否打开
            My.Computer.FileSystem.CopyFile(
Application.StartupPath & "\temp\工作簿1.xlsm",
Application.StartupPath & "\temp\JG" & NewFileName & “.xlsm",
FileIO.UIOption.OnlyErrorDialogs,
FileIO.UICancelOption.DoNothing)
            xlApp = CreateObject("Excel.Application") '创建EXCEL应用类 
            xlApp.Visible = True '设置EXCEL可见 (xlApp.Visible = False '设置EXCEL打开时不可见 ) 
            xlBook = xlApp.Workbooks.Open(Application.StartupPath & "\temp\JG" & NewFileName & “.xlsm") '打开EXCEL工作簿  
            xlsheet = xlBook.Worksheets(1) '打开EXCEL工作表 
            xlsheet.Activate() '激活工作表  
            'xlsheet.Cells(1, 1) = "卫星号" '给单元格1行驶列赋值  
            'xlsheet.Cells(1, 2) = "长半轴"
            'xlsheet.Cells(1, 3) = "平均角速度"
            ' xlsheet.Cells(1, 4) = "卫星运行的平均角速度"
            'xlsheet.Cells(1, 5) = "参考时刻"
            'xlsheet.Cells(1, 6) = "归化时间"
            ' xlsheet.Cells(1, 7) = "观测时刻的卫星平近点角"
            ' xlsheet.Cells(1, 8) = "平近点角"
            ' xlsheet.Cells(1, 9) = "真近点角"
            ' xlsheet.Cells(1, 10) = "升交距角"
            'xlsheet.Cells(1, 11) = "升交距角"   '此处有尚未解决的问题
            'xlsheet.Cells(1, 12) = "卫星矢距"
            ' xlsheet.Cells(1, 13) = "轨道倾角"
            ' xlsheet.Cells(1, 14) = "经过摄动改正的升交距角"
            ' xlsheet.Cells(1, 15) = "经过摄动改正后的卫星矢距"
            ' xlsheet.Cells(1, 15) = "经过摄动改正后的轨道倾角"
            'xlsheet.Cells(1, 15) = "卫星在轨道平面坐标系中的x坐标"
            'xlsheet.Cells(1, 15) = "卫星在轨道平面坐标系中的y坐标"
            'xlsheet.Cells(1, 15) = "观测时刻的升交点精度"
            'xlsheet.Cells(1, 15) = "卫星在地心固定坐标系中的直角坐标X"
            'xlsheet.Cells(1, 15) = "卫星在地心固定坐标系中的直角坐标Y"
            ' xlsheet.Cells(1, 15) = "卫星在地心固定坐标系中的直角坐标Z"
            Clipboard.Clear()                                     '清空剪贴板
            Dim clipb As String = RichTextBox3.Text
            Clipboard.SetText(clipb)

            'xlsheet.Cells(2, 1) = RichTextBox3.Text
            xlBook.RunAutoMacros(Excel.XlRunAutoMacro.xlAutoOpen) '运行EXCEL中的启动宏  
        Else : MsgBox("EXCEL已打开")
        End If
        'errhandler：
        'MsgBox("EXCEL已经被错误地打开" & vbCrLf & "关闭当前所有的Excel文件可能解决问题")
    End Sub
    '清空屏幕
    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        RichTextBox1.Text = "文件数据：" & vbCrLf
        RichTextBox2.Text = "卫星参数：" & vbCrLf
        RichTextBox3.Text = "计算结果：" & vbCrLf
    End Sub
    '打开文件成功后，更新RINEX状态栏
    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        ToolStripTextBox1.Text = "2.10"
        Me.RichTextBox4.Text = Me.RichTextBox4.Text & vbCrLf & "文件：" & OpenFileDialog1.FileName & “读取成功”
    End Sub
    Private Sub OpenFileDialog2_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog2.FileOk
        ToolStripTextBox1.Text = "2.11"
    End Sub
    Private Sub OpenFileDialog3_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog3.FileOk
        ToolStripTextBox1.Text = "3.02"
    End Sub
    '消息框:更新保存成功情况
    Private Sub SaveFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles SaveFileDialog1.FileOk
        RichTextBox4.Text = RichTextBox4.Text & vbCrLf & "参数已保存到：" & SaveFileDialog1.FileName
    End Sub
    Private Sub SaveFileDialog2_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles SaveFileDialog2.FileOk
        RichTextBox4.Text = RichTextBox4.Text & vbCrLf & "参数已保存到：" & SaveFileDialog2.FileName
    End Sub
    Private Sub SaveFileDialog3_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles SaveFileDialog3.FileOk
        RichTextBox4.Text = RichTextBox4.Text & vbCrLf & "结果已保存到" & SaveFileDialog3.FileName
    End Sub
    Private Sub SaveFileDialog4_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles SaveFileDialog4.FileOk
        RichTextBox4.Text = RichTextBox4.Text & vbCrLf & "结果已保存到" & SaveFileDialog4.FileName
    End Sub
End Class
