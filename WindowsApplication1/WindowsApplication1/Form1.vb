Public Class Form1
    Dim counter As Integer
    Dim banjin As Integer = 230
    Dim zhijin As Integer = 460
    Dim bj As Single = 230.0
    Dim zj As Single = 460.0
    Dim flag As Integer
    Dim xc, yc As Single
    Dim xa, ya, xb, yb As Single
    Dim g As Graphics
    Dim d As Single
    Dim jpg As Bitmap
    Dim a, b, c, aa, bb, delta, f, h As Single
    Dim Isforstop As Single
    Dim Isforchange As Single
    Dim Rin, Xin, Gin, Bin As Single
    Dim Arin, Axin, Brin, Bxin As Single
    Dim R_ax, R_ar As Single
    Dim R_reflection, R_rin, R_xin As Single
    Dim angle, range As Single
    Dim theta, theta1, theta2, theta3, theta4, theta_y1p, theta_par1, theta_par2, theta__ As Single
    Dim VSWR As Single
    Dim D_sera, L_sera, D_serb, L_serb, arg, lm, l3, l4 As Single    '串联单支节电长度及短截线长
    Dim D_para, L_para, D_parb, L_parb As Single    '并联单支节点电长度及短截线长
    Dim D_ser_A, D_ser_B, D_ser, L_A, L_B, L_ser, LD_ser, l1, l2, angle_A, angle_D, R_yin As Single              '并联电长度及短截线长
    Dim L_y1p, Gr_y1p, Gi_y1p, R_y1p, X_y1p, Gr_y1, Gi_y1, R_y1, X_y1, R_y2p, X_y2p, Gr_y2p, Gi_y2p, X_par1 As Single
    Dim Gr_y11, Gi_y11, R_y11, X_y11, R_y22p, X_y22p, Gr_y22p, Gi_y22p, X_par11 As Single
    Dim L_par1, L_par2 As Single '并联两个枝节的长度
    Dim L_par11, L_par22 As Single '并联两个枝节的长度
    Dim Zin_par, Yin_ser As Single
    Dim L_C, LD_C As Single         '阻抗电长度  导纳电长度
    Const pi = 3.1415926
    Dim D_par, L_par As Single
    Dim alpha As Single '电抗圆弧弧度
    Dim r1, r2, beta As Single '史密斯圆图的半径
    ' Dim cnt As Single

    Private Sub PictureBox1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseMove
        '单位圆的半径是200像素，但是最终会归一化
        xc = e.X / banjin - 1 '将系统坐标转换为用户坐标
        yc = 1 - e.Y / banjin


        Rin = ((1 - xc * xc - yc * yc) / ((1 - xc) * (1 - xc) + yc * yc))   '计算r
        Xin = (2 * yc / ((1 - xc) * (1 - xc) + yc * yc))                    '计算x

        R_reflection = (xc ^ 2 + yc ^ 2) ^ 0.5 '等反射系数圆半径

        R_rin = 1 / (Rin + 1)  '电阻圆半径
        R_xin = 1 / Xin        '电抗圆半径

        Gin = Rin / (Rin ^ 2 + Xin ^ 2)  '电导圆半径
        Bin = -Xin / (Rin ^ 2 + Xin ^ 2) '电纳圆半径

        range = R_reflection '反射系数模值

        VSWR = (1 + range) / (1 - range) '电压驻波比

        If yc < 0 Then               '负载点与X夹角    -180到180度
            angle = -(Math.Acos(xc / range) * 360 / 2 / pi)
        Else
            angle = (Math.Acos(xc / range) * 360 / 2 / pi)
        End If

        L_C = (180 - angle) * 0.5 / 360  '求阻抗电长度
        If L_C < 0.25 Then
            LD_C = L_C + 0.25 '求导纳电长度
        Else
            LD_C = L_C - 0.25 '求导纳电长度
        End If

        '求匹配坐标
        xa = R_reflection ^ 2 '相当于r*cos(arg(A))
        xb = xa
        yb = -((R_reflection ^ 2 - xa ^ 2) ^ 0.5)
        ya = ((R_reflection ^ 2 - xa ^ 2) ^ 0.5) '相当于r*sin(arg(A))
        Arin = ((1 - xa * xa - ya * ya) / ((1 - xa) * (1 - xa) + ya * ya))   '计算Ar
        Axin = (2 * ya / ((1 - xa) * (1 - xa) + ya * ya))                    '计算Ax
        Brin = Arin
        Bxin = -Axin
        angle_A = Math.Acos(R_reflection) * 180 / pi '求出A点与x轴夹角
        L_A = (180 - angle_A) * 0.5 / 360   'A点电长度
        L_B = (180 + angle_A) * 0.5 / 360   'B点电长度

        '-------------------------------------------------------------
        '--------------------------串联单支节-------------------------
        '-------------------------------A-----------------------------
        '几何法求串联短截线在传输线上的位置d,A点
        If L_C < L_A Then
            D_sera = L_A - L_C
        Else
            D_sera = 0.5 + L_A - L_C
        End If
        '计算串联枝节长度L
        If (2 * Bxin) / (Bxin ^ 2 - 1) > 0 Then
            theta1 = (pi - Math.Atan((2 * Bxin) / (Bxin ^ 2 - 1))) * 180 / pi
        Else
            theta1 = (-Math.Atan((2 * Bxin) / (Bxin ^ 2 - 1))) * 180 / pi
        End If
        L_sera = 0.25 + theta1 * 0.5 / 360

        '-------------------------------B-----------------------------
        '几何法求串联短截线在传输线上的位置d,B点
        If L_C < L_B Then
            D_serb = L_B - L_C
        Else
            D_serb = 0.5 + L_B - L_C
        End If
        '计算串联枝节长度L
        If (2 * Axin) / (Axin ^ 2 - 1) > 0 Then
            theta2 = (Math.Atan((2 * Axin) / (Axin ^ 2 - 1))) * 180 / pi
        Else
            theta2 = (pi + Math.Atan((2 * Axin) / (Axin ^ 2 - 1))) * 180 / pi
        End If
        L_serb = (180 - theta2) * 0.5 / 360


        '-------------------------------------------------------------
        '--------------------------并联单支节-------------------------
        '-------------------------------A-----------------------------
        '几何法求并联短截线在传输线上的位置d, A点
        If LD_C < L_A Then
            D_para = L_A - LD_C
        Else
            D_para = 0.5 + L_A - LD_C
        End If
        '计算并联枝节长度L
        If (2 * Bxin) / (Bxin ^ 2 - 1) > 0 Then
            theta2 = (pi - Math.Atan((2 * Bxin) / (Bxin ^ 2 - 1))) * 180 / pi
        Else
            theta2 = (-Math.Atan((2 * Bxin) / (Bxin ^ 2 - 1))) * 180 / pi
        End If
        L_para = theta2 * 0.5 / 360

        '-------------------------------B-----------------------------
        '几何法求并联短截线在传输线上的位置d,B点
        If LD_C < L_B Then
            D_parb = L_B - LD_C
        Else
            D_parb = 0.5 + L_B - LD_C
        End If
        '计算并联枝节长度L
        If (2 * Axin) / (Axin ^ 2 - 1) > 0 Then
            theta2 = (Math.Atan((2 * Axin) / (Axin ^ 2 - 1))) * 180 / pi
        Else
            theta2 = (pi + Math.Atan((2 * Axin) / (Axin ^ 2 - 1))) * 180 / pi
        End If
        L_parb = 0.25 + (180 - theta2) * 0.5 / 360

        '-----------------------------------------------------------------------------
        '----------------------------------并联双支节---------------------------------
        '几何法求并联短截线在传输线上的位置（两组解）
        If LD_C < L_A Then
            D_ser_A = L_A - LD_C
        Else
            D_ser_A = 0.5 - LD_C + L_A
        End If
        If LD_C < L_B Then
            D_ser_B = L_B - LD_C
        Else
            D_ser_B = 0.5 - LD_C + L_B
        End If

        d = 0.15

calculate:
        L_y1p = (LD_C + d) Mod 0.5
        theta_y1p = pi - L_y1p * 2 * pi / 0.5
        Gr_y1p = range * Math.Cos(theta_y1p)
        Gi_y1p = range * Math.Sin(theta_y1p)
        R_y1p = (1 - Gr_y1p ^ 2 - Gi_y1p ^ 2) / ((1 - Gr_y1p) ^ 2 + Gi_y1p ^ 2)
        X_y1p = 2 * Gi_y1p / ((1 - Gr_y1p) ^ 2 + Gi_y1p ^ 2)
        aa = R_y1p / (1 + R_y1p)
        bb = 1 / (1 + R_y1p)
        a = 4 * aa ^ 2 + 1
        b = -2 * aa * (1 + 2 * (aa ^ 2 - bb ^ 2))
        c = (aa ^ 2 - bb ^ 2) * (1 + aa ^ 2 - bb ^ 2)
        delta = b ^ 2 - 4 * a * c
        If (Not (delta < 0)) Then
            flag = 1
            Gr_y1 = (-b + Math.Sqrt(delta)) / (2 * a)
            Gi_y1 = 2 * aa * Gr_y1 + bb * bb - aa * aa
            R_y1 = (1 - Gr_y1 ^ 2 - Gi_y1 ^ 2) / ((1 - Gr_y1) ^ 2 + Gi_y1 ^ 2)
            X_y1 = 2 * Gi_y1 / ((1 - Gr_y1) ^ 2 + Gi_y1 ^ 2)
            '错误已改正''''''''''''''''''''
            X_par1 = X_y1 - X_y1p
            If X_par1 > 0 Then
                If 2 * X_par1 / (X_par1 ^ 2 - 1) < 0 Then
                    theta_par1 = pi + Math.Atan(2 * X_par1 / (X_par1 ^ 2 - 1))
                Else
                    theta_par1 = Math.Atan(2 * X_par1 / (X_par1 ^ 2 - 1))
                End If
                L_par1 = 0.25 + (pi - theta_par1) * 0.5 / 2 / pi
            Else
                If 2 * X_par1 / (X_par1 ^ 2 - 1) < 0 Then
                    theta_par1 = -Math.Atan(2 * X_par1 / (X_par1 ^ 2 - 1))
                Else
                    theta_par1 = pi - Math.Atan(2 * X_par1 / (X_par1 ^ 2 - 1))
                End If
                L_par1 = theta_par1 * 0.5 / 2 / pi
            End If
            Gr_y2p = Gi_y1
            Gi_y2p = -Gr_y1
            R_y2p = (1 - Gr_y2p ^ 2 - Gi_y2p ^ 2) / ((1 - Gr_y2p) ^ 2 + Gi_y2p ^ 2)
            X_y2p = 2 * Gi_y2p / ((1 - Gr_y2p) ^ 2 + Gi_y2p ^ 2)
            h = -X_y2p
            If h > 0 Then
                If (2 * h / (h ^ 2 - 1)) < 0 Then
                    theta_par2 = pi + Math.Atan(2 * h / (h * h - 1))
                Else
                    theta_par2 = Math.Atan(2 * h / (h * h - 1))
                End If
                L_par2 = 0.25 + (pi - theta_par2) * 0.5 / 2 / pi
            Else
                If (2 * f / (1 - f * f)) < 0 Then
                    theta_par2 = -Math.Atan(2 * h / (h * h - 1))
                Else
                    theta_par2 = pi - Math.Atan(2 * h / (h * h - 1))
                End If
                L_par2 = theta_par2 * 0.5 / 2 / pi
            End If



            '-----------------------------------------------------
            Gr_y11 = -(b + Math.Sqrt(delta)) / (2 * a)
            Gi_y11 = 2 * aa * Gr_y11 + bb * bb - aa * aa
            R_y11 = (1 - Gr_y11 ^ 2 - Gi_y11 ^ 2) / ((1 - Gr_y11) ^ 2 + Gi_y11 ^ 2)
            X_y11 = 2 * Gi_y11 / ((1 - Gr_y11) ^ 2 + Gi_y11 ^ 2)
            '错误已改正''''''''''''''''''''
            X_par11 = X_y11 - X_y1p
            If X_par11 > 0 Then
                If 2 * X_par11 / (X_par11 ^ 2 - 1) < 0 Then
                    theta_par1 = pi + Math.Atan(2 * X_par11 / (X_par11 ^ 2 - 1))
                Else
                    theta_par1 = Math.Atan(2 * X_par11 / (X_par11 ^ 2 - 1))
                End If
                L_par11 = 0.25 + (pi - theta_par1) * 0.5 / 2 / pi
            Else
                If 2 * X_par11 / (X_par11 ^ 2 - 1) < 0 Then
                    theta_par1 = -Math.Atan(2 * X_par11 / (X_par11 ^ 2 - 1))
                Else
                    theta_par1 = pi - Math.Atan(2 * X_par11 / (X_par11 ^ 2 - 1))
                End If
                L_par11 = theta_par1 * 0.5 / 2 / pi
            End If

            Gr_y22p = Gi_y11
            Gi_y22p = -Gr_y11
            R_y22p = (1 - Gr_y22p ^ 2 - Gi_y22p ^ 2) / ((1 - Gr_y22p) ^ 2 + Gi_y22p ^ 2)
            X_y22p = 2 * Gi_y22p / ((1 - Gr_y22p) ^ 2 + Gi_y22p ^ 2)
            h = -X_y22p
            If h > 0 Then
                If (2 * h / (h ^ 2 - 1)) < 0 Then
                    theta_par2 = pi + Math.Atan(2 * h / (h * h - 1))
                Else
                    theta_par2 = Math.Atan(2 * h / (h * h - 1))
                End If
                L_par22 = 0.25 + (pi - theta_par2) * 0.5 / 2 / pi
            Else
                If (2 * h / (h ^ 2) - 1) < 0 Then
                    theta_par2 = -Math.Atan(2 * h / (h * h - 1))
                Else
                    theta_par2 = pi - Math.Atan(2 * h / (h * h - 1))
                End If
                L_par22 = theta_par2 * 0.5 / 2 / pi
            End If
        Else
            flag = 0
            'MsgBox("进入匹配盲区，无法完成双支节匹配！", vbOKOnly, "温馨提示")
        End If
        Call Draw()
        Call Shownumber()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text.Trim = "" Or TextBox2.Text.Trim = "" Or TextBox3.Text.Trim = "" Then
            MsgBox("请输入完整的负载和特征阻抗！", vbOKOnly, "警告！")
            GoTo click_end
        End If
        If Not (IsNumeric(TextBox1.Text) And IsNumeric(TextBox2.Text) And IsNumeric(TextBox3.Text)) Then
            MsgBox("请输入数字！", vbOKOnly, "警告！")
            GoTo click_end
        End If
        Rin = TextBox1.Text / TextBox3.Text
        Xin = TextBox2.Text / TextBox3.Text


        xc = (Rin * Rin + Xin * Xin - 1) / ((Rin + 1) * (Rin + 1) + Xin * Xin)
        yc = 2 * Xin / ((Rin + 1) * (Rin + 1) + Xin * Xin)

        Rin = ((1 - xc * xc - yc * yc) / ((1 - xc) * (1 - xc) + yc * yc))   '计算r
        Xin = (2 * yc / ((1 - xc) * (1 - xc) + yc * yc))                    '计算x

        R_reflection = (xc ^ 2 + yc ^ 2) ^ 0.5 '等反射系数圆半径

        R_rin = 1 / (Rin + 1)  '电阻圆半径
        R_xin = 1 / Xin        '电抗圆半径

        Gin = Rin / (Rin ^ 2 + Xin ^ 2)  '电导圆半径
        Bin = -Xin / (Rin ^ 2 + Xin ^ 2) '电纳圆半径

        range = R_reflection '反射系数模值

        VSWR = (1 + range) / (1 - range) '电压驻波比

        If yc < 0 Then               '负载点与X夹角    -180到180度
            angle = -(Math.Acos(xc / range) * 360 / 2 / pi)
        Else
            angle = (Math.Acos(xc / range) * 360 / 2 / pi)
        End If

        L_C = (180 - angle) * 0.5 / 360  '求阻抗电长度
        If L_C < 0.25 Then
            LD_C = L_C + 0.25 '求导纳电长度
        Else
            LD_C = L_C - 0.25 '求导纳电长度
        End If

        '求匹配坐标
        xa = R_reflection ^ 2 '相当于r*cos(arg(A))
        xb = xa
        yb = -((R_reflection ^ 2 - xa ^ 2) ^ 0.5)
        ya = ((R_reflection ^ 2 - xa ^ 2) ^ 0.5) '相当于r*sin(arg(A))
        Arin = ((1 - xa * xa - ya * ya) / ((1 - xa) * (1 - xa) + ya * ya))   '计算Ar
        Axin = (2 * ya / ((1 - xa) * (1 - xa) + ya * ya))                    '计算Ax
        Brin = Arin
        Bxin = -Axin
        angle_A = Math.Acos(R_reflection) * 180 / pi '求出A点与x轴夹角
        L_A = (180 - angle_A) * 0.5 / 360   'A点电长度
        L_B = (180 + angle_A) * 0.5 / 360   'B点电长度

        '-------------------------------------------------------------
        '--------------------------串联单支节-------------------------
        '-------------------------------A-----------------------------
        '几何法求串联短截线在传输线上的位置d,A点
        If L_C < L_A Then
            D_sera = L_A - L_C
        Else
            D_sera = 0.5 + L_A - L_C
        End If
        '计算串联枝节长度L
        If (2 * Bxin) / (Bxin ^ 2 - 1) > 0 Then
            theta1 = (pi - Math.Atan((2 * Bxin) / (Bxin ^ 2 - 1))) * 180 / pi
        Else
            theta1 = (-Math.Atan((2 * Bxin) / (Bxin ^ 2 - 1))) * 180 / pi
        End If
        L_sera = 0.25 + theta1 * 0.5 / 360

        '-------------------------------B-----------------------------
        '几何法求串联短截线在传输线上的位置d,B点
        If L_C < L_B Then
            D_serb = L_B - L_C
        Else
            D_serb = 0.5 + L_B - L_C
        End If
        '计算串联枝节长度L
        If (2 * Axin) / (Axin ^ 2 - 1) > 0 Then
            theta2 = (Math.Atan((2 * Axin) / (Axin ^ 2 - 1))) * 180 / pi
        Else
            theta2 = (pi + Math.Atan((2 * Axin) / (Axin ^ 2 - 1))) * 180 / pi
        End If
        L_serb = (180 - theta2) * 0.5 / 360


        '-------------------------------------------------------------
        '--------------------------并联单支节-------------------------
        '-------------------------------A-----------------------------
        '几何法求并联短截线在传输线上的位置d, A点
        If LD_C < L_A Then
            D_para = L_A - LD_C
        Else
            D_para = 0.5 + L_A - LD_C
        End If
        '计算并联枝节长度L
        If (2 * Bxin) / (Bxin ^ 2 - 1) > 0 Then
            theta2 = (pi - Math.Atan((2 * Bxin) / (Bxin ^ 2 - 1))) * 180 / pi
        Else
            theta2 = (-Math.Atan((2 * Bxin) / (Bxin ^ 2 - 1))) * 180 / pi
        End If
        L_para = theta2 * 0.5 / 360

        '-------------------------------B-----------------------------
        '几何法求并联短截线在传输线上的位置d,B点
        If LD_C < L_B Then
            D_parb = L_B - LD_C
        Else
            D_parb = 0.5 + L_B - LD_C
        End If
        '计算并联枝节长度L
        If (2 * Axin) / (Axin ^ 2 - 1) > 0 Then
            theta2 = (Math.Atan((2 * Axin) / (Axin ^ 2 - 1))) * 180 / pi
        Else
            theta2 = (pi + Math.Atan((2 * Axin) / (Axin ^ 2 - 1))) * 180 / pi
        End If
        L_parb = 0.25 + (180 - theta2) * 0.5 / 360

        '-----------------------------------------------------------------------------
        '----------------------------------并联双支节---------------------------------
        '几何法求并联短截线在传输线上的位置（两组解）
        If LD_C < L_A Then
            D_ser_A = L_A - LD_C
        Else
            D_ser_A = 0.5 - LD_C + L_A
        End If
        If LD_C < L_B Then
            D_ser_B = L_B - LD_C
        Else
            D_ser_B = 0.5 - LD_C + L_B
        End If

        If TextBox4.Text.Trim <> "" Then
            If IsNumeric(TextBox4.Text) Then
                d = TextBox4.Text
            Else
                MsgBox("请输入数字！", vbOKOnly, "警告！")
                GoTo click_end
            End If

        Else
            MsgBox("未输入并联双枝节匹配的d1,将按照0.15λ计算！", vbOKOnly, "温馨提示")
            d = 0.15
        End If

calculate:
        L_y1p = (LD_C + d) Mod 0.5
        theta_y1p = pi - L_y1p * 2 * pi / 0.5
        Gr_y1p = range * Math.Cos(theta_y1p)
        Gi_y1p = range * Math.Sin(theta_y1p)
        R_y1p = (1 - Gr_y1p ^ 2 - Gi_y1p ^ 2) / ((1 - Gr_y1p) ^ 2 + Gi_y1p ^ 2)
        X_y1p = 2 * Gi_y1p / ((1 - Gr_y1p) ^ 2 + Gi_y1p ^ 2)
        aa = R_y1p / (1 + R_y1p)
        bb = 1 / (1 + R_y1p)
        a = 4 * aa ^ 2 + 1
        b = -2 * aa * (1 + 2 * (aa ^ 2 - bb ^ 2))
        c = (aa ^ 2 - bb ^ 2) * (1 + aa ^ 2 - bb ^ 2)
        delta = b ^ 2 - 4 * a * c
        If (Not (delta < 0)) Then
            flag = 1
            Gr_y1 = (-b + Math.Sqrt(delta)) / (2 * a)
            Gi_y1 = 2 * aa * Gr_y1 + bb * bb - aa * aa
            R_y1 = (1 - Gr_y1 ^ 2 - Gi_y1 ^ 2) / ((1 - Gr_y1) ^ 2 + Gi_y1 ^ 2)
            X_y1 = 2 * Gi_y1 / ((1 - Gr_y1) ^ 2 + Gi_y1 ^ 2)
            '错误已改正''''''''''''''''''''
            X_par1 = X_y1 - X_y1p
            L_par1 = Length(X_par1)
            Gr_y2p = Gi_y1
            Gi_y2p = -Gr_y1
            R_y2p = (1 - Gr_y2p ^ 2 - Gi_y2p ^ 2) / ((1 - Gr_y2p) ^ 2 + Gi_y2p ^ 2)
            X_y2p = 2 * Gi_y2p / ((1 - Gr_y2p) ^ 2 + Gi_y2p ^ 2)
            h = -X_y2p
            L_par2 = Length(h)



            '-----------------------------------------------------
            Gr_y11 = -(b + Math.Sqrt(delta)) / (2 * a)
            Gi_y11 = 2 * aa * Gr_y11 + bb * bb - aa * aa
            R_y11 = (1 - Gr_y11 ^ 2 - Gi_y11 ^ 2) / ((1 - Gr_y11) ^ 2 + Gi_y11 ^ 2)
            X_y11 = 2 * Gi_y11 / ((1 - Gr_y11) ^ 2 + Gi_y11 ^ 2)
            '错误已改正''''''''''''''''''''
            X_par11 = X_y11 - X_y1p
            L_par11 = Length(X_par11)

            Gr_y22p = Gi_y11
            Gi_y22p = -Gr_y11
            R_y22p = (1 - Gr_y22p ^ 2 - Gi_y22p ^ 2) / ((1 - Gr_y22p) ^ 2 + Gi_y22p ^ 2)
            X_y22p = 2 * Gi_y22p / ((1 - Gr_y22p) ^ 2 + Gi_y22p ^ 2)
            h = -X_y22p
            L_par22 = Length(h)
        Else
            MsgBox("进入匹配盲区，无法完成双支节匹配！", vbOKOnly, "温馨提示")
            flag = 0
        End If

        Call Draw()
        Call Shownumber()
click_end:
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        banjin = 300
        zhijin = 2 * banjin
        bj = banjin
        zj = 2 * bj
        jpg = New Bitmap(zhijin + 1, zhijin + 1)
        g = Graphics.FromImage(jpg)
        Dim mPen As Pen = New Pen(Color.Red, 1)
        mPen.DashStyle = Drawing.Drawing2D.DashStyle.DashDotDot
        g.DrawEllipse(mPen, New Rectangle(banjin / 2, 0, banjin, banjin)) '画匹配圆
        g.DrawLine(New Pen(Color.Black, 1), New Point(0, banjin), New Point(zhijin, banjin)) '画归一化X轴
        g.DrawLine(New Pen(Color.Black, 1), New Point(banjin, 0), New Point(banjin, zhijin)) '画归一化y轴
        g.DrawEllipse(New Pen(Color.Red, 1), New Rectangle(banjin, banjin / 2, banjin, banjin)) '画匹配圆
        g.DrawEllipse(New Pen(Color.Black, 1), New Rectangle(0, 0, zhijin, zhijin)) '画外围归一化
        Call Draw_Smith()
        PictureBox1.Image = jpg
        Label41.Text = "          （2）"
        Label38.Text = "并联双支节（1）"
        Label37.Text = ""
        Label36.Text = ""
        Label35.Text = "请输入双支节匹配时的d1(默认0.15λ)"
        Label34.Text = "          （2）"
        Label33.Text = "并联单支节（1）"
        Label32.Text = "          （2）"
        Label31.Text = "串联单支节（1）"
        Label30.Text = "B点电长度"
        Label29.Text = "B点坐标"
        Label28.Text = "A点电长度"
        Label27.Text = "A点坐标"
        Label26.Text = "驻波比ρ"
        Label25.Text = "负载电长度"
        Label24.Text = "反射系数Γ"
        Label23.Text = "归一化导纳（S）"
        Label22.Text = "归一化阻抗(Ω)"
        Label21.Text = "归一化坐标"
        Label20.Text = "请输入负载阻抗(Ω)"
        Label19.Text = "+"
        Label18.Text = "j"
        Label17.Text = "请输入特征阻抗(Ω)"
        Label14.Text = "λ"
        Label13.Text = ""
        Label12.Text = ""
        Label11.Text = ""
        Label10.Text = ""
        Label9.Text = ""
        Label8.Text = ""
        Label7.Text = ""
        Label6.Text = ""
        Label5.Text = ""
        Label4.Text = ""
        Label3.Text = ""
        Label2.Text = ""
        Label1.Text = ""
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
    End Sub

    Public Sub Draw_Smith()
        '------------------------------------Draw Smith Chart----------------------------------------------------------------------
        For i = 0.01 To 0.08 Step 0.02
            r1 = 1 / (i + 1)  '电阻圆半径
            r2 = 1 / i        '电抗圆半径
            g.DrawEllipse(New Pen(Color.LightBlue, 1), New Rectangle(zhijin - zhijin * r1, banjin - banjin * r1, zhijin * r1, zhijin * r1))
            beta = 2 * Math.Atan(1 / r2) * 360 / 2 / pi
            g.DrawArc(New Pen(Color.LightGreen, 1), New Rectangle(zhijin - banjin * r2, banjin - zhijin * r2, zhijin * r2, zhijin * r2), 90, beta)
            g.DrawArc(New Pen(Color.LightGreen, 1), New Rectangle(zhijin - banjin * r2, banjin, zhijin * r2, zhijin * r2), 270 - beta, beta)
        Next i
        For i = 0.1 To 0.95 Step 0.05
            r1 = 1 / (i + 1)  '电阻圆半径
            r2 = 1 / i        '电抗圆半径
            If (i Mod 0.1 = 0) Then
                g.DrawEllipse(New Pen(Color.LightBlue, 1.5), New Rectangle(zhijin - zhijin * r1, banjin - banjin * r1, zhijin * r1, zhijin * r1))
                beta = 2 * Math.Atan(1 / r2) * 360 / 2 / pi
                g.DrawArc(New Pen(Color.LightGreen, 1.5), New Rectangle(zhijin - banjin * r2, banjin - zhijin * r2, zhijin * r2, zhijin * r2), 90, beta)
                g.DrawArc(New Pen(Color.LightGreen, 1.5), New Rectangle(zhijin - banjin * r2, banjin, zhijin * r2, zhijin * r2), 270 - beta, beta)
            Else
                g.DrawEllipse(New Pen(Color.LightBlue, 1), New Rectangle(zhijin - zhijin * r1, banjin - banjin * r1, zhijin * r1, zhijin * r1))
                beta = 2 * Math.Atan(1 / r2) * 360 / 2 / pi
                g.DrawArc(New Pen(Color.LightGreen, 1), New Rectangle(zhijin - banjin * r2, banjin - zhijin * r2, zhijin * r2, zhijin * r2), 90, beta)
                g.DrawArc(New Pen(Color.LightGreen, 1), New Rectangle(zhijin - banjin * r2, banjin, zhijin * r2, zhijin * r2), 270 - beta, beta)
            End If
        Next i
        For i = 1.1 To 1.9 Step 0.1
            r1 = 1 / (i + 1)  '电阻圆半径
            r2 = 1 / i        '电抗圆半径
            If (i Mod 0.2 = 0) Then
                g.DrawEllipse(New Pen(Color.LightBlue, 1.5), New Rectangle(zhijin - zhijin * r1, banjin - banjin * r1, zhijin * r1, zhijin * r1))
                beta = 2 * Math.Atan(1 / r2) * 360 / 2 / pi
                g.DrawArc(New Pen(Color.LightGreen, 1.5), New Rectangle(zhijin - banjin * r2, banjin - zhijin * r2, zhijin * r2, zhijin * r2), 90, beta)
                g.DrawArc(New Pen(Color.LightGreen, 1.5), New Rectangle(zhijin - banjin * r2, banjin, zhijin * r2, zhijin * r2), 270 - beta, beta)
            Else
                g.DrawEllipse(New Pen(Color.LightBlue, 1), New Rectangle(zhijin - zhijin * r1, banjin - banjin * r1, zhijin * r1, zhijin * r1))
                beta = 2 * Math.Atan(1 / r2) * 360 / 2 / pi
                g.DrawArc(New Pen(Color.LightGreen, 1), New Rectangle(zhijin - banjin * r2, banjin - zhijin * r2, zhijin * r2, zhijin * r2), 90, beta)
                g.DrawArc(New Pen(Color.LightGreen, 1), New Rectangle(zhijin - banjin * r2, banjin, zhijin * r2, zhijin * r2), 270 - beta, beta)
            End If

        Next i
        For i = 2 To 3 Step 0.2
            r1 = 1 / (i + 1)  '电阻圆半径
            r2 = 1 / i        '电抗圆半径
            If (i = 2 Or i = 3) Then
                g.DrawEllipse(New Pen(Color.LightBlue, 1.5), New Rectangle(zhijin - zhijin * r1, banjin - banjin * r1, zhijin * r1, zhijin * r1))
                beta = 2 * Math.Atan(1 / r2) * 360 / 2 / pi
                g.DrawArc(New Pen(Color.LightGreen, 1.5), New Rectangle(zhijin - banjin * r2, banjin - zhijin * r2, zhijin * r2, zhijin * r2), 90, beta)
                g.DrawArc(New Pen(Color.LightGreen, 1.5), New Rectangle(zhijin - banjin * r2, banjin, zhijin * r2, zhijin * r2), 270 - beta, beta)
            Else
                g.DrawEllipse(New Pen(Color.LightBlue, 1), New Rectangle(zhijin - zhijin * r1, banjin - banjin * r1, zhijin * r1, zhijin * r1))
                beta = 2 * Math.Atan(1 / r2) * 360 / 2 / pi
                g.DrawArc(New Pen(Color.LightGreen, 1), New Rectangle(zhijin - banjin * r2, banjin - zhijin * r2, zhijin * r2, zhijin * r2), 90, beta)
                g.DrawArc(New Pen(Color.LightGreen, 1), New Rectangle(zhijin - banjin * r2, banjin, zhijin * r2, zhijin * r2), 270 - beta, beta)
            End If

        Next i
        For i = 3 To 5 Step 1
            r1 = 1 / (i + 1)  '电阻圆半径
            r2 = 1 / i        '电抗圆半径
            g.DrawEllipse(New Pen(Color.LightBlue, 1.5), New Rectangle(zhijin - zhijin * r1, banjin - banjin * r1, zhijin * r1, zhijin * r1))
            beta = 2 * Math.Atan(1 / r2) * 360 / 2 / pi
            g.DrawArc(New Pen(Color.LightGreen, 1.5), New Rectangle(zhijin - banjin * r2, banjin - zhijin * r2, zhijin * r2, zhijin * r2), 90, beta)
            g.DrawArc(New Pen(Color.LightGreen, 1.5), New Rectangle(zhijin - banjin * r2, banjin, zhijin * r2, zhijin * r2), 270 - beta, beta)
        Next i
        For i = 10 To 20 Step 10
            r1 = 1 / (i + 1)  '电阻圆半径
            r2 = 1 / i        '电抗圆半径
            g.DrawEllipse(New Pen(Color.LightBlue, 1), New Rectangle(zhijin - zhijin * r1, banjin - banjin * r1, zhijin * r1, zhijin * r1))
            beta = 2 * Math.Atan(1 / r2) * 360 / 2 / pi
            g.DrawArc(New Pen(Color.LightGreen, 1), New Rectangle(zhijin - banjin * r2, banjin - zhijin * r2, zhijin * r2, zhijin * r2), 90, beta)
            g.DrawArc(New Pen(Color.LightGreen, 1), New Rectangle(zhijin - banjin * r2, banjin, zhijin * r2, zhijin * r2), 270 - beta, beta)
        Next i
        For i = 50 To 100 Step 50
            r1 = 1 / (i + 1)  '电阻圆半径
            r2 = 1 / i        '电抗圆半径
            g.DrawEllipse(New Pen(Color.LightBlue, 1), New Rectangle(zhijin - zhijin * r1, banjin - banjin * r1, zhijin * r1, zhijin * r1))
            beta = 2 * Math.Atan(1 / r2) * 360 / 2 / pi
            g.DrawArc(New Pen(Color.LightGreen, 1), New Rectangle(zhijin - banjin * r2, banjin - zhijin * r2, zhijin * r2, zhijin * r2), 90, beta)
            g.DrawArc(New Pen(Color.LightGreen, 1), New Rectangle(zhijin - banjin * r2, banjin, zhijin * r2, zhijin * r2), 270 - beta, beta)
        Next i
        '--------------------------------------------------------------------------------------------------------------------------------------------------------
    End Sub

    Public Function Length(h As Single) As Single
        If h > 0 Then
            If (2 * h / (h ^ 2 - 1)) < 0 Then
                theta__ = pi + Math.Atan(2 * h / (h * h - 1))
            Else
                theta__ = Math.Atan(2 * h / (h * h - 1))
            End If
            Length = 0.25 + (pi - theta__) * 0.5 / 2 / pi
        Else
            If (2 * f / (1 - f * f)) < 0 Then
                theta__ = -Math.Atan(2 * h / (h * h - 1))
            Else
                theta__ = pi - Math.Atan(2 * h / (h * h - 1))
            End If
            Length = theta__ * 0.5 / 2 / pi
        End If
    End Function

    Public Sub Draw()
        '作图
        jpg = New Bitmap(zhijin + 1, zhijin + 1)
        g = Graphics.FromImage(jpg)
        g.DrawLine(New Pen(Color.Black, 1), New Point(0, banjin), New Point(zhijin, banjin)) '画归一化X轴
        g.DrawLine(New Pen(Color.Black, 1), New Point(banjin, 0), New Point(banjin, zhijin)) '画归一化y轴
        Dim mPen As Pen = New Pen(Color.Red, 1)
        mPen.DashStyle = Drawing.Drawing2D.DashStyle.Dot
        g.DrawEllipse(mPen, New Rectangle(banjin / 2, 0, banjin, banjin)) '画辅助圆
        g.DrawEllipse(New Pen(Color.Red, 1), New Rectangle(banjin, banjin / 2, banjin, banjin)) '画辅助圆
        g.DrawEllipse(New Pen(Color.Black, 1), New Rectangle(0, 0, zhijin, zhijin)) '画外围归一化
        Call Draw_Smith()

        If (Not (xc ^ 2 + yc ^ 2 > 1)) Then
            '反射系数圆
            g.DrawEllipse(New Pen(Color.Gold, 2), New Rectangle(banjin - banjin * R_reflection, banjin - banjin * R_reflection, zhijin * R_reflection, zhijin * R_reflection))
            '画电阻圆
            g.DrawEllipse(New Pen(Color.Green, 2), New Rectangle(zhijin - zhijin * R_rin, banjin - banjin * R_rin, zhijin * R_rin, zhijin * R_rin))
            '显示负载点
            g.DrawEllipse(New Pen(Color.Green, 2), New RectangleF(banjin * (1 + xc) - 2, banjin * (1 - yc) - 2, 4, 4))
            If xa < 1 Then
                g.DrawEllipse(New Pen(Color.Blue, 2), New RectangleF(banjin * (1 + xa) - 2, banjin * (1 - ya) - 2, 4, 4)) '显示匹配点A
            End If
            If xb < 1 Then
                g.DrawEllipse(New Pen(Color.Blue, 2), New RectangleF(banjin * (1 + xb) - 2, banjin * (1 - yb) - 2, 4, 4)) '显示匹配点B
            End If
            If Xin = 0 Then
                g.DrawLine(New Pen(Color.Green, 2), New Point(0, banjin), New Point(zhijin, banjin)) '负载电抗为0，画归一化X轴
            ElseIf R_xin < 0 Then '画电抗圆
                R_xin = -R_xin
                alpha = -2 * Math.Atan(Xin) * 360 / 2 / pi
                g.DrawArc(New Pen(Color.Green, 2), New Rectangle(zhijin - banjin * R_xin, banjin, zhijin * R_xin, zhijin * R_xin), 270 - alpha, alpha)
            Else
                alpha = 2 * Math.Atan(Xin) * 360 / 2 / pi
                g.DrawArc(New Pen(Color.Green, 2), New Rectangle(zhijin - banjin * R_xin, banjin - zhijin * R_xin, zhijin * R_xin, zhijin * R_xin), 90, alpha)
            End If

            '画匹配点电抗圆（A,B两点）
            If Axin = 0 Then
                g.DrawLine(New Pen(Color.Green, 2), New Point(0, banjin), New Point(zhijin, banjin)) '负载电抗为0，画归一化X轴
            Else
                R_ax = 1 / Axin
                alpha = 2 * Math.Atan(Axin) * 360 / 2 / pi
                g.DrawArc(New Pen(Color.MediumBlue, 2), New Rectangle(zhijin - banjin * R_ax, banjin - zhijin * R_ax, zhijin * R_ax, zhijin * R_ax), 90, alpha)
                g.DrawArc(New Pen(Color.MediumBlue, 2), New Rectangle(zhijin - banjin * R_ax, banjin, zhijin * R_ax, zhijin * R_ax), 270 - alpha, alpha)

            End If
        End If
        PictureBox1.Image = jpg
    End Sub

    Public Sub Shownumber()
        '显示数值
        Label1.Text = "X: " + xc.ToString("F5") + "   Y: " + yc.ToString("F5")  '显示归一化坐标
        Label2.Text = Rin.ToString("F5") + "   +  " + Xin.ToString("F5") + "j"  '显示阻抗大小
        Label3.Text = Gin.ToString("F5") + "   +  " + Bin.ToString("F5") + "j"   '显示导纳
        Label4.Text = range.ToString("F5") + "∠" + angle.ToString("F5") + "°"  '显示反射系数
        Label5.Text = "(阻)" + L_C.ToString("F5") + "λ" + "(导)" + LD_C.ToString("F5") + "λ"   '显示电长度
        Label6.Text = VSWR.ToString("F5")    '显示驻波系数
        Label7.Text = "(" + xa.ToString("F5") + "," + ya.ToString("F5") + "）"  'A点坐标及电长度
        Label36.Text = L_A.ToString("F5") + "λ"     'A点电长度
        Label8.Text = "(" + xb.ToString("F5") + "," + yb.ToString("F5") + ")"   'B点坐标及电长度
        Label37.Text = L_B.ToString("F5") + "λ"
        Label9.Text = "dA：" + D_sera.ToString("F5") + "λ" + "        LA：" + L_sera.ToString("F5") + "λ"
        Label10.Text = "dB：" + D_serb.ToString("F5") + "λ" + "        LB：" + L_serb.ToString("F5") + "λ"
        Label12.Text = "dA：" + D_para.ToString("F5") + "λ" + "        LA：" + L_para.ToString("F5") + "λ"
        Label13.Text = "dB：" + D_parb.ToString("F5") + "λ" + "        LB：" + L_parb.ToString("F5") + "λ"
        Label11.Text = "L1：" + L_par1.ToString("F5") + "λ" + "        L2：" + L_par2.ToString("F5") + "λ"
        Label14.Text = "λ"
        Label15.Text = "x_y1p:" + X_y1p.ToString("F5") + " x_y11:" + X_y11.ToString("F5") + " x_y22p" + X_y22p.ToString("F5") + _
                     " Gr_y11:" + Gr_y11.ToString("F5") + " Gi_y11:" + Gi_y11.ToString("F5") + " X_y1:" + X_y1.ToString("F5")
        Label39.Text = " Gr_y1:" + Gr_y1.ToString("F5") + " Gi_y1:" + Gi_y1.ToString("F5") + _
                        " LD_C:" + LD_C.ToString("F5") + " Brin:" + Brin.ToString("F5") + " Bxin:" + Bxin.ToString("F5") + _
                        " Axin:" + Axin.ToString("F5")
        Label40.Text = " X_y2p:" + X_y2p.ToString("F5") + " R_y2p:" + R_y2p.ToString("F5") + " X_par1:" + X_par1.ToString("F5") + _
                        " R_y11:" + R_y11.ToString("F5") + " theta_par2:" + theta_par2.ToString("F5")
        Label16.Text = "L1：" + L_par11.ToString("F5") + "λ" + "        L2：" + L_par22.ToString("F5") + "λ"
        Label20.Text = "请输入负载阻抗(Ω)"
        Label19.Text = "+"
        Label18.Text = "j"
        Label17.Text = "请输入特征阻抗(Ω)"
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub
End Class