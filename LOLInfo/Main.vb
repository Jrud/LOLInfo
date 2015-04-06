Imports System.Speech.Recognition
Imports GenericDataAccessClass
Imports System.Net
Imports System.IO
Imports System.Text

Public Class Main
    Implements SCOUT.IRecognitionPlugin


    Public Function AddedFunctionality() As System.Collections.Generic.List(Of String) Implements SCOUT.IRecognitionPlugin.AddedFunctionality
        Dim SL As New List(Of String)

        SL.Add("I can give you quick League of Legends statistics!")

        Return SL
    End Function

    Private DB As GenericDataAccessClass.DBMC
    Private SE As SCOUT.SpeechEngine
    Private S As SCOUT.SaySomething

    Public Sub Initialize(ByVal oDB As GenericDataAccessClass.DBMC, ByVal oSE As SCOUT.SpeechEngine, ByVal oS As SCOUT.SaySomething) Implements SCOUT.IRecognitionPlugin.Initialize
        DB = oDB
        SE = oSE
        S = oS

        Try
            Dim DSData As DataSet = DB.DBSelect("SELECT KeyVal FROM tblMaint WHERE DataKey = ""SummonerName""")

            If DSData.Tables(0).Rows.Count = 0 Then
                DB.DBInsert("INSERT INTO tblMaint (DataKey, KeyVal) VALUES (""SummonerName"", """")")
            End If
        Catch ex As Exception
            DB.DBInsert("CREATE TABLE tblMaint (MaintID AUTOINCREMENT PRIMARY KEY, DataKey Text, KeyVal Text);")

            DB.DBInsert("INSERT INTO tblMaint (DataKey, KeyVal) VALUES (""SummonerName"", """")")
        End Try

        Try
            Dim DSData As DataSet = DB.DBSelect("SELECT Name FROM tblLOLChamps")
        Catch ex As Exception
            DB.DBInsert("CREATE TABLE tblLOLChamps (ChampID INT PRIMARY KEY, Name Text);")
        End Try
    End Sub

    Dim LOLGrammarList As New List(Of System.Speech.Recognition.Grammar)
    Private SetSummoner As Grammar = New Grammar(New Choices(New String() {"Set Summoner Name", "Set My Summoner Name", "Set The Summoner Name", "I Would Like To Set My Summoner Name", "Lets Set My Summoner Name", "Set My Summoner Name Please", "Set Summoner Name Please", "Lets Set My Summoner Name Please", "Give Me A Tactical Analysis Of The Enemy Team", "Let's Have A Tactical Analysis"}))
    Private GetGameInfo As Grammar = New Grammar(New Choices(New String() {"Get Basic Game Info", "Get Current Match Info", "Get Current Match Stats", "Match Info", "Match Stats", "Match Stats Please", "Give Me The Match Stats Please", "Current Match Stats", "Current Match Info"}))

    Public Function MainMenuGrammars() As System.Collections.Generic.List(Of System.Speech.Recognition.Grammar) Implements SCOUT.IRecognitionPlugin.MainMenuGrammars

        LOLGrammarList.Add(SetSummoner)
        LOLGrammarList.Add(GetGameInfo)

        Return LOLGrammarList
    End Function

    Public Function NextRecognized(ByVal Sender As Object, ByVal e As System.Speech.Recognition.SpeechRecognizedEventArgs, ByRef State As String) As Boolean Implements SCOUT.IRecognitionPlugin.NextRecognized
        Return False
    End Function

    Public Function SpeechRecognized(ByVal Sender As Object, ByVal e As System.Speech.Recognition.SpeechRecognizedEventArgs, ByRef State As String) As Boolean Implements SCOUT.IRecognitionPlugin.SpeechRecognized
        If e.Result.Grammar Is SetSummoner Then
            RequestSummonerName()

            Return True
        ElseIf e.Result.Grammar Is GetGameInfo Then
            S.Say("Please wait while I collect the information")

            GetCurrentMatchStats()
        End If

        Return False
    End Function

    Private Sub GetCurrentMatchStats()

        Dim SummonerID As Long = WSGetSummonerID()

        Dim L As List(Of Player) = WSGetPlayerList(SummonerID)

        Dim MyTeam As Integer = 0

        For Each P As Player In L
            If P.SummonerID = SummonerID Then
                MyTeam = P.Team
            End If
        Next

        For Each P As Player In L
            If Not P.Team = MyTeam Then
                Dim SBInfo As New StringBuilder(ChampIDToName(P.ChampionID) + " is in " + P.Tier + " " + P.Division.ToString())

                If P.KDA > 3 Then
                    If P.WinLoss < 60 Then
                        SBInfo.Append(", and has a K.D.A. of " + P.KDA.ToString())
                    Else
                        SBInfo.Append(", has a K.D.A. of " + P.KDA.ToString())
                    End If
                End If

                If P.WinLoss > 60 Then
                    SBInfo.Append(", and has a win rate of " + P.WinLoss.ToString() + "%")
                End If

                S.Say(SBInfo.ToString())
            End If
        Next
    End Sub

    Private Function WSGetPlayerList(ByVal SummonerID As Long) As List(Of Player)
        Dim L As New List(Of Player)

        Dim WR As HttpWebRequest = HttpWebRequest.Create("https://na.api.pvp.net/observer-mode/rest/consumer/getSpectatorGameInfo/NA1/" + SummonerID.ToString() + "?api_key=920b706e-c903-441c-8e5a-ddc8654f8404")

        Dim Resp As WebResponse = WR.GetResponse()

        Dim Temp As String = (New StreamReader(Resp.GetResponseStream())).ReadToEnd()

        Temp = Temp.Substring(Temp.IndexOf("""participants"":[") + 16)

        For Each S As String In Temp.Split(New String() {"{""teamId"":"}, StringSplitOptions.RemoveEmptyEntries)
            If Not S = "" Then
                L.Add(New Player(S))
            End If
        Next

        Return L
    End Function

    Private Function DBGetSummonerName() As String
        Return DB.DBSelect("SELECT KeyVal FROM tblMaint WHERE DataKey = ""SummonerName""").Tables(0).Rows(0).Item(0).ToString()
    End Function

    Private Function WSGetSummonerID() As Integer
        Dim WR As HttpWebRequest = HttpWebRequest.Create("https://na.api.pvp.net/api/lol/na/v1.4/summoner/by-name/" + DBGetSummonerName() + "?api_key=920b706e-c903-441c-8e5a-ddc8654f8404")

        Dim Resp As WebResponse = WR.GetResponse()

        Dim Temp As String = (New StreamReader(Resp.GetResponseStream())).ReadToEnd()

        Temp = Temp.Substring(Temp.IndexOf("""id"":") + 5)
        Return CLng(Temp.Substring(0, Temp.IndexOf(",")))
    End Function

    Private Sub RequestSummonerName()
        S.Say("Please type in your summoner name")

        DB.DBUpdate("UPDATE tblMaint SET KeyVal = """ + InputBox("What is your summoner name?", "Input Summoner Name").Replace("""", """""") + """ WHERE DataKey = ""SummonerName""")
    End Sub

    Private Function ChampIDToName(ByVal ChampID As Integer) As String
        Dim dsName As DataSet = DB.DBSelect("SELECT Name FROM tblLOLChamps WHERE ChampID = " + ChampID.ToString())

        If dsName.Tables(0).Rows.Count = 0 Then
            Dim WR As HttpWebRequest = HttpWebRequest.Create("https://global.api.pvp.net/api/lol/static-data/na/v1.2/champion/" + ChampID.ToString() + "?champData=info&api_key=920b706e-c903-441c-8e5a-ddc8654f8404")

            Dim Resp As WebResponse = WR.GetResponse()

            Dim Temp As String = (New StreamReader(Resp.GetResponseStream())).ReadToEnd()

            Temp = Temp.Substring(Temp.IndexOf("""name"":""") + 8)
            Dim name As String = Temp.Substring(0, Temp.IndexOf(""""))

            DB.DBInsert("INSERT INTO tblLOLChamps (Name, ChampID) VALUES (""" + name + """, " + ChampID.ToString() + ")")

            Return name
        Else
            Return dsName.Tables(0).Rows(0).Item(0).ToString()
        End If
    End Function
End Class

Public Class Player
    Public Name As String
    Public SummonerID As Long
    Public ChampionID As Integer
    Public Team As Integer

    Private lGamesPlayed As Long = -1
    Private dKDA As Double = -1
    Private dKD As Double = -1
    Private dWinLoss As Double = -1

    Public ReadOnly Property GamesPlayed As Long
        Get
            If lGamesPlayed = -1 Then
                AcquireRankedDetails(ChampionID)
            End If

            Return lGamesPlayed
        End Get
    End Property
    Public ReadOnly Property KDA As Double
        Get
            If dKDA = -1 Then
                AcquireRankedDetails(ChampionID)
            End If

            Return dKDA
        End Get
    End Property
    Public ReadOnly Property KD As Long
        Get
            If dKD = -1 Then
                AcquireRankedDetails(ChampionID)
            End If

            Return dKD
        End Get
    End Property
    Public ReadOnly Property WinLoss As Long
        Get
            If dWinLoss = -1 Then
                AcquireRankedDetails(ChampionID)
            End If

            Return dWinLoss
        End Get
    End Property


    Private iDivision As Integer = -1
    Private sTier As String = -1

    Public ReadOnly Property Division As Integer
        Get
            If iDivision = -1 Then
                AcquireLeagueDetails()
            End If

            Return iDivision
        End Get
    End Property
    Public ReadOnly Property Tier As String
        Get
            If sTier = -1 Then
                AcquireLeagueDetails()
            End If

            Return sTier
        End Get
    End Property

    Public Sub New(ByVal sName As String, ByVal iTeam As Integer, ByVal lSummonerID As Long, ByVal iChampionID As Integer)
        Name = sName
        If iTeam > 5 Then
            Team = iTeam / 100
        Else
            Team = iTeam
        End If
        SummonerID = lSummonerID
        ChampionID = iChampionID
    End Sub

    Public Sub New(ByVal ParticipantXML As String)
        Dim Temp As String = ParticipantXML.Substring(ParticipantXML.IndexOf("""summonerName"":""") + 16)
        Name = Temp.Substring(0, Temp.IndexOf(""","))

        Temp = ParticipantXML.Substring(ParticipantXML.IndexOf("""summonerId"":") + 13)
        SummonerID = CLng(Temp.Substring(0, Temp.IndexOf(",")))

        Temp = ParticipantXML.Substring(ParticipantXML.IndexOf("""championId"":") + 13)
        ChampionID = CLng(Temp.Substring(0, Temp.IndexOf(",")))

        Team = CInt(ParticipantXML.Substring(0, 1))
    End Sub

    Private Sub AcquireLeagueDetails()
        Dim WR As HttpWebRequest = HttpWebRequest.Create("https://na.api.pvp.net/api/lol/na/v2.5/league/by-summoner/" + SummonerID.ToString() + "/entry?api_key=920b706e-c903-441c-8e5a-ddc8654f8404")

        Dim Resp As WebResponse = WR.GetResponse()

        Dim LeagueXML As String = (New StreamReader(Resp.GetResponseStream())).ReadToEnd()

        Dim Temp As String = LeagueXML.Substring(LeagueXML.IndexOf("""tier"":""") + 8)
        sTier = Temp.Substring(0, Temp.IndexOf(""""))

        Temp = LeagueXML.Substring(LeagueXML.IndexOf("""division"":""") + 12)
        Dim tDivision As String = Temp.Substring(0, Temp.IndexOf(""""))

        If tDivision = "I" Then
            iDivision = 1
        ElseIf tDivision = "II" Then
            iDivision = 2
        ElseIf tDivision = "III" Then
            iDivision = 3
        ElseIf tDivision = "IV" Then
            iDivision = 4
        ElseIf tDivision = "V" Then
            iDivision = 5
        End If
    End Sub

    Private Sub AcquireRankedDetails(ByVal ChampID As Integer)
        Dim WR As HttpWebRequest

        Dim Resp As WebResponse
        Dim Success As Boolean = False
        While Not Success
            Try
                WR = HttpWebRequest.Create("https://na.api.pvp.net/api/lol/na/v1.3/stats/by-summoner/" + SummonerID.ToString() + "/ranked?season=SEASON2015&api_key=920b706e-c903-441c-8e5a-ddc8654f8404&randthing=" + Guid.NewGuid().ToString())
                Resp = WR.GetResponse()
                Success = True
            Catch ex As Exception
                Threading.Thread.Sleep(11000)
            End Try
        End While
        Dim LeagueXML As String = (New StreamReader(Resp.GetResponseStream())).ReadToEnd()
        LeagueXML = LeagueXML.Substring(LeagueXML.IndexOf("""id"":" + ChampID.ToString()))
        LeagueXML = LeagueXML.Substring(0, LeagueXML.IndexOf("}"))

        Dim Temp As String = LeagueXML.Substring(LeagueXML.IndexOf("""totalSessionsPlayed"":") + 22)
        lGamesPlayed = CInt(Temp.Substring(0, Temp.IndexOf(",")))

        Temp = LeagueXML.Substring(LeagueXML.IndexOf("""totalDeathsPerSession"":") + 24)
        Dim Deaths As Long = CInt(Temp.Substring(0, Temp.IndexOf(",")))

        Temp = LeagueXML.Substring(LeagueXML.IndexOf("""totalChampionKills"":") + 21)
        Dim Kills As Long = CInt(Temp.Substring(0, Temp.IndexOf(",")))

        Temp = LeagueXML.Substring(LeagueXML.IndexOf("""totalAssists"":") + 15)
        Dim Assists As Long = CInt(Temp.Substring(0, Temp.IndexOf(",")))

        dKD = Kills / Deaths
        dKDA = (Kills + Assists) / Deaths

        Temp = LeagueXML.Substring(LeagueXML.IndexOf("""totalSessionsWon"":") + 19)
        Dim Wins As Long = CInt(Temp.Substring(0, Temp.IndexOf(",")))

        dWinLoss = (Wins / GamesPlayed) * 100
    End Sub
End Class
