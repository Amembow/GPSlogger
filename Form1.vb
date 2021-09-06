Option Strict Off
Option Explicit On
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.ComponentModel
Imports System.IO
Imports System.Text

Imports System.Diagnostics
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Collections
Imports System.Windows.Automation

Imports System.Drawing

Public Class GPSLogger

    'API～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～

    Private Declare Function AllocConsole Lib "kernel32.dll" Alias "AllocConsole" () As Boolean
    Public Const ATTACH_PARENT_PROCESS As UInteger = UInteger.MaxValue
    Private Declare Function AttachConsole Lib "kernel32.dll" Alias "AttachConsole" (Optional ByVal dwProcessId As UInteger = ATTACH_PARENT_PROCESS) As Boolean
    Private Declare Function FreeConsole Lib "kernel32.dll" Alias "FreeConsole" () As Boolean

    Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hwndParent As Integer, ByVal hwndChildAfter As Integer, ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
    Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Integer, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As String) As Integer
    Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
    Public Declare Function GetDlgItem Lib "user32.dll" Alias "GetDlgItem" (ByVal hDlg As Integer, ByVal nIDDlgItem As Integer) As Integer
    Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hwnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As IntPtr) As IntPtr


    '～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～

    '定数～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～

    Const GRS80_A = 6378137.0#
    Const GRS80_E2 = 0.00669438002301188
    Const GRS80_MNUM = 6335439.32708317
    Const PAI = 3.14159265358979

    Private uiAuto As UIAutomationClient.CUIAutomation
    Private Const UIA_WindowControlTypeId As Integer = 50032
    Private Const UIA_ControlTypePropertyId As Integer = 30003
    Private Const UIA_ClassNamePropertyId As Integer = 30012
    Private Const UIA_TextControlTypeId As Integer = 50020
    Private Const TreeScope_Subtree As Integer = 7
    Private Const UIA_InvokePatternId As Integer = 10000
    Private Const UIA_AutomationPropertyID As Integer = 30011
    Private Const UIA_NamePropertyID As Integer = 30005
    Private Const UIA_ValuePatternId As Integer = 10002

    '～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～～

    Private cts As CancellationTokenSource
    Dim SPWKT(,) As String '計測計画線のノード座標配列
    Dim JGD As String '座標系　テキストボックス１に入れる





    Sub Main()

        JGD = TextBox1.Text

        If JGD = "" Then
            MsgBox("座標系を入力してください")
            Exit Sub
        End If


        On Error GoTo ErrSub

        Dim i As Double

        Dim lat As String
        Dim lon As String



        Dim elmPrevious As UIAutomationClient.IUIAutomationElement

        uiAuto = New UIAutomationClient.CUIAutomation
        elmPrevious = uiAuto.GetRootElement

        Dim cndCoreWindow As UIAutomationClient.IUIAutomationCondition
        Dim elmCoreWindow As UIAutomationClient.IUIAutomationElement

        Dim AutomationID As String


        For i = 1 To 13



            Select Case i
                Case 1
                    AutomationID = "QgisApp"
                Case 2
                    AutomationID = "QgisApp.GPSInformation"
                Case 3
                    AutomationID = "QgisApp.GPSInformation.QgsRendererWidgetContainerBase"
                Case 4
                    AutomationID = "QgisApp.GPSInformation.QgsRendererWidgetContainerBase.scrollArea"
                Case 5
                    AutomationID = "QgisApp.GPSInformation.QgsRendererWidgetContainerBase.scrollArea.qt_scrollarea_viewport"
                Case 6
                    AutomationID = "QgisApp.GPSInformation.QgsRendererWidgetContainerBase.scrollArea.qt_scrollarea_viewport.scrollAreaWidgetContents"
                Case 7
                    AutomationID = "QgisApp.GPSInformation.QgsRendererWidgetContainerBase.scrollArea.qt_scrollarea_viewport.scrollAreaWidgetContents.mStackedWidget"
                Case 8
                    AutomationID = "QgisApp.GPSInformation.QgsRendererWidgetContainerBase.scrollArea.qt_scrollarea_viewport.scrollAreaWidgetContents.mStackedWidget.QgsGpsInformationWidgetBase"
                Case 9
                    AutomationID = "QgisApp.GPSInformation.QgsRendererWidgetContainerBase.scrollArea.qt_scrollarea_viewport.scrollAreaWidgetContents.mStackedWidget.QgsGpsInformationWidgetBase.mStackedWidget"
                Case 10
                    AutomationID = "QgisApp.GPSInformation.QgsRendererWidgetContainerBase.scrollArea.qt_scrollarea_viewport.scrollAreaWidgetContents.mStackedWidget.QgsGpsInformationWidgetBase.mStackedWidget.stackedWidgetPage1"
                Case 11
                    AutomationID = "QgisApp.GPSInformation.QgsRendererWidgetContainerBase.scrollArea.qt_scrollarea_viewport.scrollAreaWidgetContents.mStackedWidget.QgsGpsInformationWidgetBase.mStackedWidget.stackedWidgetPage1.scrollArea_2"
                Case 12
                    AutomationID = "QgisApp.GPSInformation.QgsRendererWidgetContainerBase.scrollArea.qt_scrollarea_viewport.scrollAreaWidgetContents.mStackedWidget.QgsGpsInformationWidgetBase.mStackedWidget.stackedWidgetPage1.scrollArea_2.qt_scrollarea_viewport"
                Case 13
                    AutomationID = "QgisApp.GPSInformation.QgsRendererWidgetContainerBase.scrollArea.qt_scrollarea_viewport.scrollAreaWidgetContents.mStackedWidget.QgsGpsInformationWidgetBase.mStackedWidget.stackedWidgetPage1.scrollArea_2.qt_scrollarea_viewport.scrollAreaWidgetContents_2"

            End Select

            cndCoreWindow = uiAuto.CreatePropertyCondition(UIA_AutomationPropertyID, AutomationID)
            elmCoreWindow = elmPrevious.FindFirst(TreeScope_Subtree, cndCoreWindow)

            If elmCoreWindow Is Nothing Then
                Console.WriteLine(i & "番目取得失敗")
                Exit Sub
            End If

            elmPrevious = elmCoreWindow

        Next




        Dim cndCoreWindowlat As UIAutomationClient.IUIAutomationCondition
        Dim elmCoreWindowlat As UIAutomationClient.IUIAutomationElement


        cndCoreWindowlat = uiAuto.CreatePropertyCondition(UIA_AutomationPropertyID, "QgisApp.GPSInformation.QgsRendererWidgetContainerBase.scrollArea.qt_scrollarea_viewport.scrollAreaWidgetContents.mStackedWidget.QgsGpsInformationWidgetBase.mStackedWidget.stackedWidgetPage1.scrollArea_2.qt_scrollarea_viewport.scrollAreaWidgetContents_2.mTxtLatitude")
        elmCoreWindowlat = elmPrevious.FindFirst(TreeScope_Subtree, cndCoreWindowlat)
        If elmCoreWindowlat Is Nothing Then
            Console.WriteLine("lat取得失敗")
            Exit Sub
        End If

        Dim ptnVal As UIAutomationClient.IUIAutomationValuePattern = elmCoreWindowlat.GetCurrentPattern(UIA_ValuePatternId)
        lat = ptnVal.CurrentValue


        Dim cndCoreWindowlon As UIAutomationClient.IUIAutomationCondition
        Dim elmCoreWindowlon As UIAutomationClient.IUIAutomationElement

        cndCoreWindowlon = uiAuto.CreatePropertyCondition(UIA_AutomationPropertyID, "QgisApp.GPSInformation.QgsRendererWidgetContainerBase.scrollArea.qt_scrollarea_viewport.scrollAreaWidgetContents.mStackedWidget.QgsGpsInformationWidgetBase.mStackedWidget.stackedWidgetPage1.scrollArea_2.qt_scrollarea_viewport.scrollAreaWidgetContents_2.mTxtLongitude")
        elmCoreWindowlon = elmPrevious.FindFirst(TreeScope_Subtree, cndCoreWindowlon)
        If elmCoreWindowlon Is Nothing Then
            Console.WriteLine("lot取得失敗")
            Exit Sub
        End If
        ptnVal = elmCoreWindowlon.GetCurrentPattern(UIA_ValuePatternId)
        lon = ptnVal.CurrentValue

        If lon = "" Or lat = "" Then
            MsgBox("座標が取得できていません。" & vbCrLf & "衛星受信中につき少々お待ちください。")
            cts.Cancel()
            Exit Sub
        End If


        Dim Answer() As String
        Answer = Split(WKT_Analyzer(calc_x(lat, lon, CInt(JGD)), calc_y(lat, lon, CInt(JGD))), ",")

        AttachConsole()

        If Answer(1) = "Out Of Range" Then

            Console.ForegroundColor = ConsoleColor.Red
        Else


            If CDbl(Answer(1)) > 1.5 Then '計画線からどれだけ離れていたらアラートを出すか
                Console.ForegroundColor = ConsoleColor.Red
            ElseIf CDbl(Answer(1)) > 1 Then
                Console.ForegroundColor = ConsoleColor.Yellow
            Else
                Console.ForegroundColor = ConsoleColor.Green
            End If

        End If

        Console.WriteLine(DateTime.Now & "," & calc_y(lat, lon, CInt(JGD)) & "," & calc_x(lat, lon, CInt(JGD)) & "," & WKT_Analyzer(calc_x(lat, lon, CInt(JGD)), calc_y(lat, lon, CInt(JGD))))
        Console.ForegroundColor = ConsoleColor.White


        'ここからインジケータのコントロール~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

        Dim SP() As String
        Dim fig As String = ""
        Dim Deg As Single

        i = 0
        SP = Split(WKT_Analyzer(calc_x(lat, lon, CInt(JGD)), calc_y(lat, lon, CInt(JGD))), ",")

        Label5.Text = SP(0)　'左右
        Label8.Text = SP(1)　'誤差
        Label10.Text = CDbl(SP(2))  '方位

        Dim Magni As Double = CDbl(TextBox3.Text)
        '以下Test___________________________________________________________________________________________________________________________


        If SP(0) = "Right" Then
            Deg = SP(2) * Magni
            Do
                i = 0.1 + i
                fig = "■" & fig
            Loop While i < CDbl(SP(1))

            If i < 0.5 Then
                Label12.ForeColor = Color.Green
            ElseIf i < 1.0 Then
                Label12.ForeColor = Color.Yellow
            Else
                Label12.ForeColor = Color.Red
            End If

            Label11.Text = ""
            Label11.Visible = False

            Label12.Visible = True
            Label12.Text = fig

        ElseIf SP(0) = "Left" Then
            Deg = SP(2) * -1 * Magni

            Do
                i = 0.1 + i
                fig = "■" & fig
            Loop While i < CDbl(SP(1))

            If i < 0.5 Then
                Label11.ForeColor = Color.Green
            ElseIf i < 1.0 Then
                Label11.ForeColor = Color.Yellow
            Else
                Label11.ForeColor = Color.Red
            End If

            Label12.Text = ""
            Label12.Visible = False

            Label11.Visible = True
            Label11.Text = fig

        Else
            Deg = 180
            Label11.Text = ""
            Label11.Visible = False

            Label12.Text = ""
            Label12.Visible = False


        End If

        '方向インジケータ＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿＿








        '描画先とするImageオブジェクトを作成する
        Dim canvas As New Bitmap(PictureBox1.Width, PictureBox1.Height)
        'ImageオブジェクトのGraphicsオブジェクトを作成する
        Dim g As Graphics = Graphics.FromImage(canvas)

        '描画する画像のBitmapオブジェクトを作成
        Dim img As Bitmap = New Bitmap(My.Resources.A2)

        '新しい座標位置を計算する

        Dim D1() As String = Split(Affine((img.Width) / 2 * -1, img.Height / 2, Deg), ",")
        Dim D2() As String = Split(Affine((img.Width) / 2, img.Height / 2, Deg), ",")
        Dim D3() As String = Split(Affine((img.Width) / 2 * -1, img.Height / 2 * -1, Deg), ",")

        Dim x As Single = D1(0)
        Dim y As Single = D1(1)
        Dim x1 As Single = D2(0)
        Dim y1 As Single = D2(1)
        Dim x2 As Single = D3(0)
        Dim y2 As Single = D3(1)

        'PointF配列を作成
        Dim destinationPoints() As PointF = {New PointF(x, y), New PointF(x1, y1), New PointF(x2, y2)}

        '画像を表示
        g.DrawImage(img, destinationPoints)

        'リソースを解放する
        img.Dispose()
        g.Dispose()

        'PictureBox1に表示する
        PictureBox1.Image = canvas





        Exit Sub

ErrSub:
        MsgBox("UIAutomation取得エラー")
        cts.Cancel()
    End Sub




    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '計測計画線読み込み

        JGD = TextBox1.Text

        If JGD = "" Then
            MsgBox("座標系を入力してください")
            Exit Sub
        End If


        ReDim SPWKT(0, 0)

        Dim fso As Object
        fso = CreateObject("Scripting.FileSystemObject")

        Dim WKTPath As String
        Dim i As Long
        Dim SP() As String

        With OpenFileDialog1
            .Multiselect = False
            .Title = "WKTファイルの入力"
            .Filter = "WKTファイル|*.csv"
        End With

        If OpenFileDialog2.ShowDialog() = Windows.Forms.DialogResult.OK Then
            WKTPath = OpenFileDialog2.FileName

            Dim inputFile As Object
            inputFile = fso.OpenTextFile(WKTPath)

            SP = Split(inputFile.ReadLine, ",")
            Dim z() As String

            ReDim SPWKT(SP.Length - 1, 1)

            AllocConsole()
            AttachConsole()
            Console.WriteLine("計測計画線ノード　読み取り開始")

            For i = 0 To SP.Length - 1
                z = Split(SP(i), " ")

                SPWKT(i, 0) = calc_y(z(1), z(0), CInt(JGD))
                SPWKT(i, 1) = calc_x(z(1), z(0), CInt(JGD))

            Next

            Console.WriteLine("計測計画線ノード　読み取り完了")

            Label1.Text = "[" & IO.Path.GetFileName(WKTPath) & "]"

        Else
            Exit Sub

        End If

    End Sub


    Function WKT_Analyzer(y As Double, x As Double) As String

        Dim Distance As Double = 1000
        Dim Nearest As Double = 10000
        Dim Direction As String = "Out Of Range"
        Dim i As Long, y1 As Double, y2 As Double, x1 As Double, x2 As Double, x01 As Double, x02 As Double, y01 As Double, y02 As Double
        Dim TempDis As Double
        Dim a As Double, b As Double, c As Double
        Dim Nearest1 As Double, Nearest2 As Double
        Dim Phi As Double, Theta As Double

        For i = 0 To SPWKT.Length / 2 - 1 '計画線ノードは２次元配列なので、ノード数の半分だけ繰り返す。配列数なので１を引く。

            If i <> SPWKT.Length / 2 - 1 Then　'計画線ノードの隣り合う２点を取得しているので、配列の最終ノードは端数として切り捨てる。

                '計画線ノードの隣り合う２点を取得
                x1 = SPWKT(i, 0)
                x2 = SPWKT(i + 1, 0)
                y1 = SPWKT(i, 1)
                y2 = SPWKT(i + 1, 1)

                Nearest1 = Math.Sqrt((x1 - x) ^ 2 + (y1 - y) ^ 2)　'隣り合う其々の計画線ノードと自己位置の距離を求める。
                Nearest2 = Math.Sqrt((x2 - x) ^ 2 + (y2 - y) ^ 2)



                If Nearest > Nearest1 + Nearest2 Then　'隣り合う其々の計画線ノードと自己位置の距離を足し合わせたものを、繰り返しの中で比較していく。自己位置に対して最少となるものを取得する。

                    '垂線長の求め方
                    '点(x0,y0)から、直線 ax + by + c = 0 へも垂線の長さは
                    ' Math.Abs(a * x0 + b * y0 + c))/Math.Sqrt(a^2 + b^2)

                    '２点を通る１次方程式の求め方
                    '点(x1,y1)と点(x2,y2)を通る直線の１次方程式は
                    '(y1 - y2) * x - (x1 - x2) * y + (x1 - x2) * y1 - (y1 - y2) * x1 = 0

                    a = y1 - y2
                    b = (x1 - x2) * -1
                    c = (x1 - x2) * y1 - (y1 - y2) * x1


                    If y <= (b / a) * x + y2 - (b / a) * x2 And y >= (b / a) * x + y1 - (b / a) * x1 Then　'隣り合う２点の計画線ノードを結ぶ直線に直交し、其々の計画線ノードを通る２直線で囲まれてる範囲に自己位置があった場合。2パターンある。

                        TempDis = Math.Abs(a * x + b * y + c) / Math.Sqrt(a ^ 2 + b ^ 2)　'条件を満たした場合、隣り合う２点の計画線ノードを結ぶ直線に対して、自己位置から垂直線を下ろし、交点から自己位置までの長さを求める。

                        Phi = Math.Atan((x2 - x1) / (y2 - y1))　'隣り合う２点の計画線ノードのうち、最初のノードと自己位置を結ぶ直線でなす角と、隣り合う２点の計画線ノードを結ぶ直線でなす角ををれぞれ算出する。
                        Theta = Math.Atan((x - x1) / (y - y1))

                        If Phi = Theta Then　'上記で求めた角度を比較し場合分けする。
                            Direction = "Centre"
                        ElseIf Phi > Theta Then
                            Direction = "Right"
                        Else
                            Direction = "Left"
                        End If


                    ElseIf y >= (b / a) * x + y2 - (b / a) * x2 And y <= (b / a) * x + y1 - (b / a) * x1 Then　'２つめのパターン。条件が違うだけで、処理は同じ。

                        TempDis = Math.Abs(a * x + b * y + c) / Math.Sqrt(a ^ 2 + b ^ 2)

                        Phi = Math.Atan((x2 - x1) / (y2 - y1))
                        Theta = Math.Atan((x - x1) / (y - y1))

                        If Phi = Theta Then
                            Direction = "Centre"
                        ElseIf Phi > Theta Then
                            Direction = "Right"
                        Else
                            Direction = "Left"
                        End If


                    Else　'隣り合う２点の計画線ノードを結ぶ直線に直交し、其々の計画線ノードを通る２直線で囲まれてる範囲に自己位置がなかった場合、大きな距離を仮置きする。
                        TempDis = 1000

                    End If



                    If Distance > Math.Min(Nearest1, Math.Min(Nearest2, TempDis)) Then　'隣り合う其々の計画線ノードからの自己位置の距離と、上記の垂線長の３つのうち、最も短いものを最近傍として保持する。

                        Distance = Math.Min(Nearest1, Math.Min(Nearest2, TempDis))
                        x01 = x1
                        x02 = x2
                        y01 = y1
                        y02 = y2

                    End If

                    Nearest = Nearest1 + Nearest2

                End If

            End If

        Next i

        Dim Er As String = "Out Of Range"

        If Distance = 1000 Then
            Return Er & "," & Er & ",_____________________________," & Er & "," & Er & "," & Er & "," & Er
        Else
            Return Direction & "," & CStr(Distance) & "," & CStr(Math.Abs(Phi - Theta)) & "," & CStr(x01) & "," & CStr(y01) & "," & CStr(x02) & "," & CStr(y02)
        End If

    End Function


    Private Async Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'QGISから座標取得
        Dim Sec As String
        Sec = TextBox2.Text

        If Sec = "" Then
            MsgBox("更新間隔を入力してください")
            Exit Sub
        End If

        cts = New CancellationTokenSource

        AllocConsole()
        AttachConsole()
        Console.WriteLine("現在時刻,現在地(x),現在地(y),計画線に対しての左右,最近傍予測距離(m),計画線に対しての角度,計画線ノード(x1),計画線ノード(y1),計画線ノード(x2),計画線ノード(y2)")



        Await Task.Run(Sub()

                           While cts.IsCancellationRequested = False

                               Call Main()

                               Thread.Sleep(CDbl(Sec) * 1000)

                           End While

                       End Sub)



    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        cts.Cancel()

    End Sub


    Function Deg2rad(deg As Double) As Double

        Deg2rad = deg * PAI / 180.0#

    End Function

    '***********************************************************************
    '　[説　明]　2地点間の直線距離を求める関数
    '　[引　数]　lat1:地点1緯度　lng1:地点1経度　lat2:地点2緯度　lng2:地点2経度
    '　[戻り値]　なし
    '***********************************************************************
    Function CalcHubeny(lat1 As Double, lng1 As Double, lat2 As Double, lng2 As Double)

        Dim my As Double
        Dim dy As Double
        Dim dx As Double
        Dim sin As Double
        Dim w As Double
        Dim m As Double
        Dim n As Double
        Dim dym As Double
        Dim dxncos As Double

        my = Deg2rad((lat1 + lat2) / 2.0#)
        dy = Deg2rad(lat1 - lat2)
        dx = Deg2rad(lng1 - lng2)

        sin = Math.Sin(my)
        w = Math.Sqrt(1.0# - GRS80_E2 * sin * sin)
        m = GRS80_MNUM / (w * w * w)
        n = GRS80_A / w

        dym = dy * m
        dxncos = dx * n * Math.Cos(my)

        'メートルで取得するので、キロメートルに変換し小数第一位までにする
        CalcHubeny = Math.Sqrt(dym * dym + dxncos * dxncos)

    End Function

    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    Dim NaN As Double

    '系番号から座標系x原点とy原点を取得するための配列
    Const TargetX As String = "0,33,33,36,33,36,36,36,36,36,40,44,44,44,26,26,26,26,20,26"
    Const TargetY As String = "0,129.5,131,132.16666666666667,133.5,134.33333333333333,136,137.16666666666667,138.5,139.83333333333333,140.83333333333333,140.25,142.25,144.25,142,127.5,124,131,136,154"

    '定数 (a, F: 世界測地系-測地基準系1980（GRS80）楕円体)
    Dim m0 As Double
    Dim a As Double
    Dim F As Double

    'セル関数として呼ばれる--平面直角座標の緯度を返す
    Public Function calc_x(phi_deg As Double, lambda_deg As Double, groupNum As Integer) As Double
        '    Attribute calc_x.VB_Description = "十進法緯度・経度を指定された平面直角座標の緯度（Ｘ座標）に変換します。"
        '    Attribute calc_x.VB_ProcData.VB_Invoke_Func = " \n14"
        calc_x = Split(calc_xy(phi_deg, lambda_deg, CDbl(Split(TargetX, ",")(groupNum)), CDbl(Split(TargetY, ",")(groupNum))), ",")(0)
    End Function

    ''セル関数として呼ばれる--平面直角座標の経度を返す
    Public Function calc_y(phi_deg As Double, lambda_deg As Double, groupNum As Integer) As Double
        '    Attribute calc_y.VB_Description = "十進法緯度・経度を指定された平面直角座標の経度（Y座標）に変換します。"
        '    Attribute calc_y.VB_ProcData.VB_Invoke_Func = " \n14"
        calc_y = Split(calc_xy(phi_deg, lambda_deg, CDbl(Split(TargetX, ",")(groupNum)), CDbl(Split(TargetY, ",")(groupNum))), ",")(1)
    End Function
    '
    Private Function calc_xy(phi_deg As Double, lambda_deg As Double, phi0_deg As Double, lambda0_deg As Double) As String
        '緯度経度を平面直角座標に変換する
        'input:
        ' (phi_deg, lambda_deg): 変換したい緯度・経度[ 十進法度]（分・秒でなく小数）
        ' (phi0_deg, lambda0_deg): 平面直角座標系原点の緯度・経度[ 十進法度]（分・秒でなく小数）
        'output:
        ' x: 変換後の平面直角座標[m]
        ' y: 変換後の平面直角座標[m]
        '緯度経度・平面直角座標系原点をラジアンに直す
        '
        NaN = CDbl(999999999)
        m0 = CDbl(0.9999)
        a = CDbl(6378137)
        F = CDbl(298.257222101)
        '
        Dim phi_rad As Double, lambda_rad As Double, phi0_rad As Double, lambda0_rad As Double
        '
        phi_rad = fnradians(phi_deg)
        lambda_rad = fnradians(lambda_deg)
        phi0_rad = fnradians(phi0_deg)
        lambda0_rad = fnradians(lambda0_deg)
        '
        '① n, A_i, alpha_iの計算
        Dim n As Double
        Dim A_array() As Double, alpha_array() As Double
        n = 1.0# / (2.0# * F - 1.0#)
        A_array = fnA_array(n)
        alpha_array = fnalpha_array(n)
        '
        '②S_, A_の計算
        Dim A_ As Double, S_ As Double
        A_ = ((m0 * a) / (1.0# + n)) * A_array(0)  'ｍ
        S_ = ((m0 * a) / (1.0# + n)) * (A_array(0) * phi0_rad + np_dot(A_array, np_sin(2.0# * phi0_rad, np_arange(1, 5)))) 'ｍ
        '
        '③ lambda_c, lambda_sの計算
        Dim lambda_c As Double, lambda_s As Double
        lambda_c = Math.Cos(lambda_rad - lambda0_rad)
        lambda_s = Math.Sin(lambda_rad - lambda0_rad)
        '
        '④t, t_の計算
        Dim t As Double, t_ As Double
        t = fnsinh(fnarctanh(Math.Sin(phi_rad)) - ((2.0# * Math.Sqrt(n)) / (1.0# + n)) * fnarctanh(((2.0# * Math.Sqrt(n)) / (1.0# + n)) * Math.Sin(phi_rad)))
        t_ = Math.Sqrt(1.0# + t * t)
        '
        ' ⑤ xi', eta'の計算
        Dim xi2 As Double, eta2 As Double
        xi2 = Math.Atan(t / lambda_c)    ' [rad]
        eta2 = fnarctanh(lambda_s / t_)
        '
        '⑥ x, yの計算
        Dim x As Double, y As Double
        x = A_ * (xi2 + np_sum(np_multiply(alpha_array, np_multiply(np_sin(2.0# * xi2, np_arange(1, 5)), np_cosh(2.0# * eta2, np_arange(1, 5)))))) - S_ ' ｍ
        y = A_ * (eta2 + np_sum(np_multiply(alpha_array, np_multiply(np_cos(2.0# * xi2, np_arange(1, 5)), np_sinh(2.0# * eta2, np_arange(1, 5))))))     ' ｍ
        '
        calc_xy = x & "," & y

    End Function
    '
    Function fnA_array(dn As Double) As Double()
        Dim Ar(5) As Double
        Ar(0) = 1.0# + (dn ^ 2) / 4.0# + (dn ^ 4) / 64.0#
        Ar(1) = -(3.0# / 2.0#) * (dn - (dn ^ 3) / 8.0# - (dn ^ 5) / 64.0#)
        Ar(2) = (15.0# / 16.0#) * (dn ^ 2 - (dn ^ 4) / 4.0#)
        Ar(3) = -(35.0# / 48.0#) * (dn ^ 3 - (5.0# / 16.0#) * (dn ^ 5))
        Ar(4) = (315.0# / 512.0#) * (dn ^ 4)
        Ar(5) = -(693.0# / 1280.0#) * (dn ^ 5)
        '
        fnA_array = Ar
    End Function
    '
    Function fnalpha_array(dn As Double) As Double()
        Dim Ar(5) As Double
        Ar(0) = NaN  'NaN
        Ar(1) = (1.0# / 2.0#) * dn - (2.0# / 3.0#) * (dn ^ 2) + (5.0# / 16.0#) * (dn ^ 3) + (41.0# / 180.0#) * (dn ^ 4) - (127.0# / 288.0#) * (dn ^ 5)
        Ar(2) = (13.0# / 48.0#) * (dn ^ 2) - (3.0# / 5.0#) * (dn ^ 3) + (557.0# / 1440.0#) * (dn ^ 4) + (281.0# / 630.0#) * (dn ^ 5)
        Ar(3) = (61.0# / 240.0#) * (dn ^ 3) - (103.0# / 140.0#) * (dn ^ 4) + (15061.0# / 26880.0#) * (dn ^ 5)
        Ar(4) = (49561.0# / 161280.0#) * (dn ^ 4) - (179.0# / 168.0#) * (dn ^ 5)
        Ar(5) = (34729.0# / 80640.0#) * (dn ^ 5)
        '
        fnalpha_array = Ar
    End Function
    '
    Function np_dot(da1() As Double, da2() As Double, Optional sta As Integer = 1) As Double
        Dim rt As Double
        Dim i As Integer
        rt = 0#
        For i = sta To UBound(da1)
            If da1(i) <> NaN And da2(i) <> NaN Then rt = rt + da1(i) * da2(i)
        Next i
        '
        np_dot = rt
    End Function
    '
    Function np_sum(da1() As Double, Optional sta As Integer = 1) As Double
        Dim rt As Double
        Dim i As Integer
        rt = 0#
        For i = sta To UBound(da1)
            If da1(i) <> NaN Then rt = rt + da1(i)
        Next i
        '
        np_sum = rt
    End Function
    '
    Function np_multiply(da1() As Double, da2() As Double, Optional sta As Integer = 1) As Double()
        Dim i As Integer
        Dim rt() As Double
        ReDim rt(UBound(da1))
        For i = sta To UBound(da1)
            If da1(i) <> NaN And da2(i) <> NaN Then rt(i) = da1(i) * da2(i)
        Next i
        '
        np_multiply = rt
    End Function
    '
    Function np_cos(dn As Double, da1() As Double, Optional sta As Integer = 1) As Double()
        Dim i As Integer
        Dim rt() As Double
        ReDim rt(UBound(da1))
        For i = sta To UBound(da1)
            If da1(i) <> NaN Then rt(i) = Math.Cos(dn * da1(i))
        Next i
        '
        np_cos = rt
    End Function
    '
    Function np_sin(dn As Double, da1() As Double, Optional sta As Integer = 1) As Double()
        Dim i As Integer
        Dim rt() As Double
        ReDim rt(UBound(da1))
        For i = sta To UBound(da1)
            If da1(i) <> NaN Then rt(i) = Math.Sin(dn * da1(i))
        Next i
        '
        np_sin = rt
    End Function
    '
    Function np_cosh(dn As Double, da1() As Double, Optional sta As Integer = 1) As Double()
        Dim i As Integer
        Dim rt() As Double
        ReDim rt(UBound(da1))
        For i = sta To UBound(da1)
            If da1(i) <> NaN Then rt(i) = fncosh(dn * da1(i))
        Next i
        '
        np_cosh = rt
    End Function
    '
    Function np_sinh(dn As Double, da1() As Double, Optional sta As Integer = 1) As Double()
        Dim i As Integer
        Dim rt() As Double
        ReDim rt(UBound(da1))
        For i = sta To UBound(da1)
            If da1(i) <> NaN Then rt(i) = fnsinh(dn * da1(i))
        Next i
        '
        np_sinh = rt
    End Function
    '
    Function np_arange(sta As Integer, sto As Integer, Optional ste As Integer = 1) As Double()
        Dim i As Integer
        Dim rt() As Double
        ReDim rt(sto)
        For i = sta To sto Step ste
            rt(i) = CDbl(i)
        Next i
        '
        np_arange = rt
    End Function
    '
    Function fnradians(dn As Double) As Double
        fnradians = dn / 45.0# * Math.Atan(1.0#)
    End Function
    '
    Function fncosh(dn As Double) As Double
        fncosh = (Math.Exp(dn) + Math.Exp(-dn)) / 2.0#
    End Function
    '
    Function fnsinh(dn As Double) As Double
        fnsinh = (Math.Exp(dn) - Math.Exp(-dn)) / 2.0#
    End Function
    '
    Function fnarctanh(dn As Double) As Double
        fnarctanh = Math.Log((1.0# + dn) / (1.0# - dn)) / 2.0#
    End Function

    Public Function Affine(ByVal x1 As Double, ByVal y1 As Double, ByVal p As Double) As String
        Dim angle As Double, x As Double, y As Double

        angle = Math.PI * p / 180
        x = x1 * Math.Cos(angle) - y1 * Math.Sin(angle) + 50
        y = x1 * Math.Sin(angle) + y1 * Math.Cos(angle) + 50

        Return x.ToString("F6") & "," & y.ToString("F6")
    End Function


End Class
