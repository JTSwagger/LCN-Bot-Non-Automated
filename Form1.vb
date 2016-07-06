Imports System.Runtime.InteropServices
Imports System.Speech
Imports System.Speech.Recognition
Imports NDde.Client
Imports mshtml
Imports OpenQA.Selenium.Firefox
Imports OpenQA.Selenium.Chrome
Imports OpenQA.Selenium
Imports System.Threading
Imports System.ComponentModel
Imports Microsoft.ProjectOxford.SpeechRecognition
Imports System.Collections.Generic
Imports OpenQA.Selenium.Support.UI
Imports Newtonsoft.Json

Public Class Form1
    Dim VehicleList(0) As Vehicle
    Public Sub LoadVehicles()

        Dim Data As IO.StreamReader = New IO.StreamReader(Application.StartupPath & "\Vehicles.csv")
        Dim NewLine(3) As String
        While Not Data.EndOfStream
            ReDim Preserve VehicleList(VehicleList.Length + 1)
            NewLine = Data.ReadLine.Split(",")
            VehicleList(VehicleList.Length - 1).Year = NewLine(0)
            VehicleList(VehicleList.Length - 1).Make = NewLine(1)
            VehicleList(VehicleList.Length - 1).model = NewLine(2)
            Console.WriteLine("CAR: " & VehicleList(VehicleList.Length - 1).Year & " " & VehicleList(VehicleList.Length - 1).Make & " " & VehicleList(VehicleList.Length - 1).model)
        End While
    End Sub
    Public Function Search(text As String) As String
        Dim startMake As Integer = 0
        Dim startmodel As Integer = 0

        For y As Integer = 0 To VehicleList.Length - 1
            If text.Contains(VehicleList(y).Year) Then
                For x As Integer = y To VehicleList.Length - 1
                    If text.Contains(VehicleList(x).Make) Then
                        For z As Integer = x To VehicleList.Length - 1
                            If text.Contains(VehicleList(z).model) Then
                                Return (VehicleList(z).model)

                            End If
                        Next
                    End If
                Next
            End If
        Next
        Return "Not Found"
    End Function

    Public Structure Vehicle
        Dim Year As String
        Dim Make As String
        Dim model As String
    End Structure
    Public WithEvents m As MicrophoneRecognitionClient
    Private demoThread As Thread = Nothing
    Private WithEvents BakgroundWorker As BackgroundWorker

    <DllImport("User32.dll")>
    Public Shared Function RegisterHotKey(ByVal hwnd As IntPtr,
                  ByVal id As Integer, ByVal fsModifiers As Integer,
                  ByVal vk As Integer) As Integer
    End Function
    Dim push As Boolean = False

    <DllImport("User32.dll")>
    Public Shared Function UnregisterHotKey(ByVal hwnd As IntPtr, ByVal id As Integer) As Integer

    End Function
    Dim sResidenceType As String
    Dim waitCheck As Integer
    Dim theSpouseName As String
    Dim Names As New List(Of String)
    Dim stillthere As Integer = 0
    Dim counter As Integer = 0
    Dim numCounter As Integer = 0
    Dim counter2 As Integer = 0
    Dim Driver As FirefoxDriver
    Dim Already_Handled As Boolean = False

    Dim selectElement As SelectElement
    Public Sub onChange(sender As Object, e As MicrophoneEventArgs) Handles m.OnMicrophoneStatus
        Recording_status = e.Recording
        Me.BeginInvoke(New Action(AddressOf updateLabel))
    End Sub

    Sub handlepartialquestion()
        Console.WriteLine("looking through partial questions for: " & Part)
        Console.WriteLine("reps: " & quest)

        Try
            If Currently_Rebuttaling = False Then
                Select Case True
                    Case Part.Contains("who is this"), Part.Contains("who are you"), Part.Contains("who is calling"), Part.Contains("who's this"), Part.Contains("who's calling"), Part.Contains("who do you represent")
                        Currently_Rebuttaling = True
                        clipType = "Objection"
                        If CurrentQ = 3 Then
                            CurrentQ = 0
                        End If
                        rolltheclipThread("c:\soundboard\cheryl\INTRO\CHERYLCALLING.mp3")
                        Timer2.Enabled = True
                    Case Part.Contains("what is this"), Part.Contains("what's this"), Part.Contains("what is the nature of this call"), Part.Contains("what are you calling about"), Part.Contains("what is purpose of this call")
                        Currently_Rebuttaling = True
                        clipType = "Objection"
                        If CurrentQ = 3 Then
                            CurrentQ = 0
                        End If
                        rolltheclipThread("c:\soundboard\cheryl\INTRO\THISISTOGIVENEWQUOTE.mp3")
                        Already_Handled = True
                        Timer2.Enabled = True
                    Case Part.Contains("what is lcn"), Part.Contains("what is elsieanne"), Part.Contains("about your company"), s.Contains("lcn")
                        Currently_Rebuttaling = True
                        clipType = "Objection"
                        If CurrentQ = 3 Then
                            CurrentQ = 0
                        End If
                        rolltheclipThread("c:\soundboard\cheryl\Rebuttals\What's LCN.mp3")
                        Already_Handled = True
                        Timer2.Enabled = True

                    Case Part.Contains("why are you calling")
                        Currently_Rebuttaling = True
                        clipType = "Objection"
                        rolltheclipThread("c:\soundboard\cheryl\INTRO\THISISTOGIVENEWQUOTE.mp3")
                        Already_Handled = True
                        Timer2.Enabled = True

                    Case Part.Contains("how did you get my info"), Part.Contains("where did you get my info")
                        Currently_Rebuttaling = True
                        clipType = "Objection"
                        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\Where Did You get My info.mp3")
                        Already_Handled = True
                        Timer2.Enabled = True
                End Select
            End If
        Catch ex As Exception
            Console.WriteLine("problem with part question")
        End Try


    End Sub 'Checks for questions in the p
    Dim NumWords As Integer = 0
    Dim totalInbetween As Integer = 0
    Public Sub SomeSpeech(ByVal sender As Object, ByVal e As Microsoft.ProjectOxford.SpeechRecognition.PartialSpeechResponseEventArgs) Handles m.OnPartialResponseReceived
        NumWords += 1
        totalInbetween += theSilence / 1000
        SilenceCap = (totalInbetween / NumWords) + 2
        theSilence = 0
        Part = e.PartialResult
        Me.BeginInvoke(New Action(AddressOf handlePartialObjection))
    End Sub
    Dim Currently_Rebuttaling As Boolean = False
    Public Sub handlePartialObjection()
        Console.WriteLine("looking through partial objections for: " & Part)
        txtSpeech.Text = "The bot heard:  " & Part
        inBetween = False
        If Currently_Rebuttaling = False Then
            Select Case True
                Case Part.Contains("is this a real person"), Part.Contains("is this a recording"), s.Contains("robot"), s.Contains("automated")
                    Currently_Rebuttaling = True
                    Part = ""
                    clipType = "Objection"
                    rolltheclipThread("C:\Soundboard\Cheryl\REACTIONS\Loud-laugh.mp3")
                    Timer2.Enabled = True
                    NICount += 1
                Case Part.Contains("no vehicle"), Part.Contains("sold the car"), Part.Contains("sold my car"), Part.Contains("no car"), Part.Contains("don't have a vehicle"), Part.Contains("don't") And Part.Contains("have a car"), Part.Contains("don't have an automobile"), Part.Contains("dont't have my own car"), Part.Contains("doesn't have a car")
                    Currently_Rebuttaling = True
                    Part = ""
                    newobjection = False
                    Console.WriteLine("THEY DON'T HAVE A CAR")
                    rolltheclipThread("C:/Soundboard/Cheryl/WRAPUP/have a great day.mp3")
                    cmbDispo.Text = "No Car"
                    CurrentQ = 31
                    Timer2.Enabled = True
                    counter2 = 0
                Case Part.Contains("not interested"), Part.Contains("don't need a quote"), Part.Contains("i'm fine"), Part.Contains("not really interested"), Part.Contains("not in arrested"), Part.Contains("that's okay thank you"), Part.Contains("no interest"), Part.Contains("stop calling"), Part.Contains("i'm good"), Part.Contains("all set"), Part.Contains("don't want it"), Part.Contains("not changing"), Part.Contains("i'm happy with"), Part.Contains("very happy"), Part.Contains("no thank you"), Part.Contains("not looking"), Part.Contains("don't wanna change"), Part.Contains("no thank you"), Part.Contains("don't need insurance"), Part.Contains("won't change") 'NI
                    Currently_Rebuttaling = True
                    Part = ""

                    clipType = "Objection"
                    newobjection = False
                    Console.WriteLine("NOT INTERESTED")
                    If CurrentQ = 3 Then
                        CurrentQ = 0
                    End If
                    Select Case NICount
                        Case 0
                            rolltheclipThread("C:\soundboard\cheryl\INTRO\THISISTOGIVENEWQUOTE.mp3")
                            NICount += 1
                            If CurrentQ = 3 Then
                                CurrentQ = 0
                            End If
                            Timer2.Enabled = True
                        Case 1
                            rolltheclipThread("C:\SoundBoard\Cheryl\REBUTTALS\REBUTTAL1.mp3")
                            numbreps += 1
                            If CurrentQ = 3 Then
                                CurrentQ = 0
                            End If
                            Timer2.Enabled = True
                    End Select


                Case Part.Contains("busy"), Part.Contains("at work"), Part.Contains("driving"), Part.Contains("can't talk"), Part.Contains("call me back"), Part.Contains("could you call back"), Part.Contains("call back another time"), Part.Contains("call later"), Part.Contains("working right now")
                    Currently_Rebuttaling = True
                    Part = ""
                    newobjection = False
                    clipType = "Objection"
                    If CurrentQ = 3 Then
                        CurrentQ = 0
                    End If
                    Select Case counter
                        Case 0
                            rolltheclipThread("C:\SoundBoard\Cheryl\REBUTTALS\THIS WILL BE REAL QUICK.mp3")
                            Timer2.Enabled = True
                            NICount += 1
                        Case Else
                            rolltheclipThread("C:\SoundBoard\Cheryl\Birthday\questions 5-4-16\questions 5-4-16\Im busy.MP3")
                            Timer2.Enabled = True
                            NICount += 1
                            counter = 0
                    End Select
                Case Part.Contains("wrong number"), Part.Contains("by that name"), Part.Contains("wrong phone number")
                    Currently_Rebuttaling = True
                    Part = ""
                    newobjection = False
                    rolltheclipThread("c:\soundboard\cheryl\Rebuttals\SORRY.mp3")
                    cmbDispo.Text = "Wrong Number"
                    CurrentQ = 31
                    Timer2.Enabled = True

                Case Part.Contains("already have"), Part.Contains("already have insurance"), Part.Contains("already got insurance"), Part.Contains("happy with"), Part.Contains("i have insurance"), Part.Contains("i got insurance")
                    Currently_Rebuttaling = True
                    Part = ""
                    clipType = "Objection"
                    rolltheclipThread("C:\SoundBoard\Cheryl\Birthday\questions 5-4-16\questions 5-4-16\i have insurance.mp3")
                    Timer2.Enabled = True
                    NICount += 1
                Case Part.Contains("take me off your list"), Part.Contains("name off your list"), Part.Contains("number off your list"), Part.Contains("take me off"), Part.Contains("take me off your call list"), Part.Contains("no call list"), Part.Contains("take this number off the list"), Part.Contains("do not call list"), Part.Contains("remove me from the list"), Part.Contains("taken off his collar"), Part.Contains("remove me from your calling list"), Part.Contains("call list"), Part.Contains("calling list")
                    Currently_Rebuttaling = True
                    Part = ""
                    newobjection = False

                    rolltheclipThread("C:\SoundBoard\Cheryl\REBUTTALS\DNC.mp3")
                    cmbDispo.Text = "Do Not Call"
                    CurrentQ = 31
                    Timer2.Enabled = True
            End Select
        End If
        handlepartialquestion()
    End Sub                         ' Checks for Objections from partial speech received.

    Public Function getVehicleInfo(vehiclenum As Integer) As Boolean
        Dim cont As Boolean = False
        If getYear(vehiclenum) Then
            cont = True
        End If
        Return False
    End Function



    Public Function getMake(vehiclenum As Integer) As Boolean 'currentq for this is 8
        If secondPass = False Then
            ModelHolder = s
        End If
        Timer2.Enabled = False
        Dim X As Integer
        Console.WriteLine(s)
        For X = 1 To MAKELIST.Length - 1
            If s.Contains(LCase(MAKELIST(X))) Then
                vMake(vehiclenum) = UCase(MAKELIST(X).Replace(" ", "%20"))
                Console.WriteLine(vMake(vehiclenum))
                Exit For
            End If
        Next
        If s.Contains("chevy") Then
            vMake(vehiclenum) = "CHEVROLET"
            Console.WriteLine("It's a Chevy")
        End If
        If s.Contains("folks wagon") Then
            vMake(vehiclenum) = "VOLKSWAGEN"
            Console.WriteLine("It's a VOLKSWAGEN")
        End If
        If vMake(vehiclenum) <> "" Then

            If vehiclenum = 1 Then

                Try
                    selectElement = New SelectElement(local_browser.FindElementById("vehicle-make"))
                    selectElement.SelectByText(vMake(vehiclenum))
                    Return True
                Catch ex As Exception
                    Do Until local_browser.FindElementById("vehicle-make").GetAttribute("class").Contains("hide") = False
                        Console.WriteLine("whoopwhoopwhoop shoopdawhoop")
                    Loop
                    selectElement = New SelectElement(local_browser.FindElementById("vehicle-make"))
                    selectElement.SelectByText(vMake(vehiclenum))
                    Return True
                End Try

            Else
                selectElement = New SelectElement(local_browser.FindElementById("vehicle" & vehiclenum & "-make"))
                selectElement.SelectByText(vMake(vehiclenum))
                Return True
            End If
        Else
            Console.WriteLine("-----MAKE NOT FOUND-----")
            secondPass = True
            Timer2.Enabled = False
            CurrentQ = 8
            rolltheclipThread("C:\SoundBoard\Cheryl\VEHICLE INFO\WHO MAKES THAT VEHICLE.MP3")
            isQuestion = True
        End If
        Return False
    End Function 'GETS THE MAKE OF THE VEHICLE
    Dim callPos As String = ""
    Const Intro As String = "Intro"
    Const Insurance_Provider As String = "Insurance Provider"
    Const Policy_Expiration As String = "Policy Expiration"
    Const Policy_Start As String = "Policy Start"
    Const Number_Of_Vehicles As String = "Number of Vehicles"
    Const Year_Make_Model As String = "Year Make Model"                             ' This Block of constants are used to check
    Const Driver_Birthday As String = "Driver Birthday"                             ' what should be done with the recognized speech
    Const Finalize_BDAY As String = "Finalize BDAY"                                 ' after it is checked against objections & questions
    Const Marital_Status As String = "Marital Status"                               ' They are assigned to CallPos
    Const Spouse_Name As String = "Spouse Name"
    Const Spouse_DOB As String = "Spouse DOB"
    Const Own_Rent As String = "Own or Rent"
    Const Home_Type As String = "Residence Type"
    Const Their_Address As String = "Their Address"
    Const Finalize_Address As String = "Finalize Address"
    Const Email_Address As String = "Email Address"
    Const Credit As String = "Credit"
    Const Phone_Type As String = "Phone Type"
    Const Last_Name As String = "Last Name"
    Const TCPA_Wrap As String = "TCPA"

    Public Function NumtoMonth(Month As String) As String
        Select Case Month
            Case "01", "1"
                Return "Jan"
            Case "02", "2"
                Return "Feb"
            Case "03", "3"
                Return "mar"
            Case "04", "4"
                Return "Apr"
            Case "05", "5"
                Return "May"
            Case "06", "6"
                Return "Jun"
            Case "07", "7"
                Return "Jul"
            Case "08", "8"
                Return "Aug"
            Case "09", "9"
                Return "Sep"
            Case "10"
                Return "Oct"
            Case "11"
                Return "Nov"
            Case "12"
                Return "Dec"
        End Select
        Return False
    End Function


    Public Sub handleResponse()
        tmrSilence.Enabled = False
        Dim req As Net.WebRequest
        Dim resp As Net.WebResponse

        If clipType = "Question" Then
            If waveOut.PlaybackState = 0 Then
                txtSpeech.Text = "In response to asking " & callPos & ": " & vbNewLine & "They said: " & s
                txtSpeech.Text += vbNewLine & callPos

                Select Case callPos

                    Case Intro
                        clipType = ""
                        callPos = Insurance_Provider
                        s = ""
                        If FullAuto.Checked Then
                            CurrentQ = 3
                            Timer2.Enabled = True
                        End If

                    Case Insurance_Provider
                        Console.WriteLine("Checking Insurance Provider")
                        If CheckForCompany() Then
                            clipType = ""
                            callPos = Policy_Expiration
                            s = ""
                            If FullAuto.Checked Then
                                CurrentQ = 4
                                Timer2.Enabled = True
                            End If
                        Else
                            clipType = "Question"
                        End If
                    Case Policy_Expiration
                        Console.WriteLine("Checking Insurance Expiration Date")
                        If checkExpiration() Then
                            clipType = ""
                            callPos = Policy_Start
                            s = ""
                            If FullAuto.Checked Then
                                CurrentQ = 5
                                Timer2.Enabled = True
                            End If
                        Else
                            s = ""

                        End If
                    Case Policy_Start
                        Console.WriteLine("Checking Insurance Start Date")
                        If CheckHowLong() Then
                            clipType = ""
                            callPos = Number_Of_Vehicles
                            s = ""
                            If FullAuto.Checked Then
                                CurrentQ = 6
                                Timer2.Enabled = True
                            End If
                        End If
                    Case Number_Of_Vehicles
                        Console.WriteLine("Checking Number of Vehicles")
                        If checkForNumVehicles() Then
                            clipType = ""
                            callPos = Year_Make_Model
                            s = ""
                            If FullAuto.Checked Then
                                CurrentQ = 7
                                Timer2.Enabled = True
                            End If
                        End If

                    Case Year_Make_Model

                        Console.WriteLine("Checking Year Make and model of Vehicle Number  " & VehicleNum & " out of " & NumberOfVehicles)
                        If getYear(VehicleNum) Then
                            clipType = ""
                            CurrentQ = 8
                            If getMake(VehicleNum) Then
                                CurrentQ = 9
                                'req = Net.WebRequest.Create("http://212.47.241.129:5000/search/" & VYear(VehicleNum) & "/" & vMake(VehicleNum) & "/" & s)
                                'resp = req.GetResponse
                                'Dim webReader As New IO.StreamReader(resp.GetResponseStream)
                                ''Dim results As String = webReader.ReadToEnd()
                                'resp.Close()
                                'Dim jsonresults As responser = JsonConvert.DeserializeObject(Of responser)(results)
                                'If jsonresults.response = "true" Then
                                'vmodel(VehicleNum) = jsonresults.model
                                'Console.WriteLine(vmodel(VehicleNum))
                                If getModel(VehicleNum) Then
                                    '  local_browser.FindElementById("vehicle-model").SendKeys(vmodel(VehicleNum)
                                    If NumberOfVehicles > 1 And VehicleNum < NumberOfVehicles Then
                                        VehicleNum += 1
                                        CurrentQ = 7
                                        s = ""
                                        callPos = Year_Make_Model
                                        Timer2.Enabled = True

                                    Else
                                        clipType = ""
                                        callPos = Driver_Birthday
                                        s = ""
                                        If FullAuto.Checked Then
                                            CurrentQ = 10
                                            Timer2.Enabled = True
                                        End If
                                    End If
                                Else
                                    clipType = "Question"
                                End If
                            Else
                                Console.WriteLine("there's a problem with getmake")
                                clipType = "Question"
                            End If
                        Else
                            Console.WriteLine("there's a problem with getyear")
                            clipType = "Question"
                        End If

                    Case Driver_Birthday
                        If getBirthdaWAV() Then
                            If GetBirthday() Then
                                clipType = ""
                                callPos = Marital_Status
                                If FullAuto.Checked Then
                                    CurrentQ = 11
                                    Timer2.Enabled = True
                                End If
                            Else
                                rolltheclipThread("C: \Soundboard\Cheryl\DRIVER INFO\DOB1.mp3")
                                Already_Handled = False
                            End If
                        Else
                            If getSpouseBDAY(False) Then
                                clipType = ""
                                callPos = Finalize_BDAY
                                If FullAuto.Checked Then
                                    CurrentQ = 14
                                    Timer2.Enabled = True
                                End If
                            End If
                        End If
                    Case Finalize_BDAY
                        If finalizeSpouseBDay(False) Then
                            clipType = ""
                            callPos = Marital_Status
                            If FullAuto.Checked Then
                                CurrentQ = 11
                                Timer2.Enabled = True
                            End If
                        ElseIf finalizeSpouseBDay(True) Then
                            clipType = ""
                            callPos = Own_Rent
                            If FullAuto.Checked Then
                                CurrentQ = 15
                                Timer2.Enabled = True
                            End If
                        End If
                    Case Marital_Status
                        If checkMaritalStatus() Then
                            If maritalStatus = "Married" Then
                                callPos = Spouse_Name
                                clipType = ""
                                If FullAuto.Checked Then
                                    CurrentQ = 12
                                    Timer2.Enabled = True
                                End If
                            Else
                                callPos = Own_Rent
                                clipType = ""
                                If FullAuto.Checked Then
                                    CurrentQ = 15
                                    Timer2.Enabled = True
                                End If
                            End If
                        End If
                    Case Spouse_Name
                        If checkForSpouseName() Then
                            clipType = ""
                            callPos = Spouse_DOB
                            If FullAuto.Checked Then
                                CurrentQ = 13
                                Timer2.Enabled = True
                            End If
                        End If
                    Case Spouse_DOB
                        If getSpouseBDAY(True) Then
                            clipType = ""
                            callPos = Finalize_BDAY
                            If FullAuto.Checked Then
                                CurrentQ = 14
                                Timer2.Enabled = True
                            End If
                        End If
                    Case Own_Rent
                        If getHomeType() Then
                            clipType = ""
                            callPos = Home_Type
                            If FullAuto.Checked Then
                                CurrentQ = 16
                                Timer2.Enabled = True
                            End If
                        End If
                    Case Home_Type
                        If getResType() Then
                            clipType = ""
                            callPos = Their_Address
                            If FullAuto.Checked Then
                                CurrentQ = 17
                                Timer2.Enabled = True
                            End If
                        End If
                    Case Their_Address
                        If doaddressstuff() Then
                            clipType = ""
                            callPos = Finalize_Address
                            If FullAuto.Checked Then
                                CurrentQ = 18
                                Timer2.Enabled = True
                            End If
                        End If
                    Case Finalize_Address
                        If finalizeAddress() Then
                            clipType = ""
                            callPos = Email_Address
                            If FullAuto.Checked Then
                                CurrentQ = 19
                                Timer2.Enabled = True
                            End If
                        End If

                    Case Email_Address
                        If getEmail() Then
                            clipType = ""
                            callPos = Credit
                            If FullAuto.Checked Then
                                CurrentQ = 21
                                Timer2.Enabled = True
                            End If
                        End If
                    Case Credit
                        If getCredit() Then
                            clipType = ""
                            callPos = Phone_Type
                            If FullAuto.Checked Then
                                CurrentQ = 22
                                Timer2.Enabled = True
                            End If
                        End If
                    Case Phone_Type
                        If getPhoneType() Then
                            clipType = ""
                            callPos = Last_Name
                            If FullAuto.Checked Then
                                CurrentQ = 23
                                Timer2.Enabled = True
                            End If
                        End If
                    Case Last_Name
                        If getLastName() Then
                            clipType = ""
                            callPos = TCPA_Wrap
                            If FullAuto.Checked Then
                                CurrentQ = 27
                                Timer2.Enabled = True
                            End If
                        End If
                    Case TCPA_Wrap
                        If handleTCPA() Then
                            Console.WriteLine("we haz lead yay")
                            moo()
                        Else
                            Console.WriteLine(":(")
                            Console.WriteLine("sad panda")
                        End If
                        Timer2.Enabled = True
                End Select
            End If
        End If
    End Sub                                                 ' Handles calling data parsing functions based on CallPos

    Public Sub GotSpeech(ByVal sender As Object, ByVal e As Microsoft.ProjectOxford.SpeechRecognition.SpeechResponseEventArgs) Handles m.OnResponseReceived
        Console.WriteLine(e.PhraseResponse.RecognitionStatus)
        If e.PhraseResponse.Results.Length > 0 Then
            s += LCase(e.PhraseResponse.Results(0).DisplayText)
        End If
        If e.PhraseResponse.RecognitionStatus = RecognitionStatus.InitialSilenceTimeout Then
            m.EndMicAndRecognition()
        End If
    End Sub


    Dim MAKELIST() As String =
{"Acura",
"Alfa Romeo",
"AM General",
"AMC",
"Aston Martin",
"Audi",
"Bentley",
"BMW",
"Bugatti",
"Buick",
"Cadillac",
"Chevrolet",
"Chrysler",
"Daewoo",
"Daihatsu",
"Datsun",
"Dodge",
"Eagle",
"Ferrari",
"FIAT",
"Fisker",
"Ford",
"Geo",
"GMC",
"Honda",
"HUMMER",
"Hyundai",
"Infiniti",
"Isuzu",
"Jaguar",
"Jeep",
"Kia",
"Lamborghini",
"Land Rover",
"Lexus",
"Lincoln",
"Lotus",
"Maserati",
"Maybach",
"Mazda",
"McLaren",
"Mercedes-Benz",
"Mercury",
"Merkur",
"MINI",
"Mitsubishi",
"Nissan",
"Oldsmobile",
"Panoz",
"Peugeot",
"Plymouth",
"Pontiac",
"Porsche",
"Ram",
"Renault",
"Rolls-Royce",
"Saab",
"Saturn",
"Scion",
"Smart",
"Sterling",
"Subaru",
"Suzuki",
"Tesla",
"Toyota",
"Volkswagen",
"Volvo",
"Yugo"}
    Dim stringarray2() As String
    Dim LifeQual As Boolean = False
    Dim HomeQual As Boolean = False
    Dim rentQual As Boolean = False
    Dim HealthQual As Boolean = False
    Dim renterQual As Boolean = False
    Dim Mediqual As Boolean = False
    'FORM VARIABLES
    Dim TempV(2) As String
    Dim VehicleNum As Integer = 1
    Dim InsuranceCarrier As String
    Dim Expiration(1) As String
    Dim Start(1) As String
    Dim theVehicle(3, 2) As String
    Dim theBDAY As String
    Dim gender As String
    Dim maritalStatus As String
    Dim spouseGender As String
    Dim spouseBDAY As String
    Dim Address As String                   'USED TO AUTO ENTER AT THE END OF THE CALL
    Dim zipcode As String
    Dim residence As String
    Dim residenceType As String
    Dim creditRating As String
    Dim email As String
    Dim phoneType As String
    Dim lastName As String
    Dim HomeExtra As Boolean
    Dim lifeExtra As Boolean
    Dim healthextra As Boolean
    Dim medicareextra As Boolean
    Dim renters As Boolean
    Dim theyearBuilt As String
    Dim theSqFt As String
    Dim ppc As String
    Dim introBday As Boolean
    Dim insurancePass As Boolean
    Dim FirstNameField As String
    Dim TEST As String
    Dim BDayCounter As Integer
    Dim notReady As Boolean
    Dim IProvider As String = ""
    Dim HumanCounter As Integer = 1
    Dim tempStr() As String
    Dim displayTime As String
    Dim newcall As Boolean = True
    Dim isecond As Double = 0
    Dim WordBuffer() As String
    Dim BDay As String
    Dim BMonth As String
    Dim BYear As String
    Dim finalCompare() As String
    Dim theMonth As String
    Dim theYear As String
    Dim writtenMonth As String
    Dim Au As New AutoItX3Lib.AutoItX3
    Dim vars() As String
    Dim dde As New DdeClient("Firefox", "WWW_GetWindowInfo")
    Dim GramBuild As GrammarBuilder
    Dim headElement As HtmlElement
    Dim scriptElement As HtmlElement
    Dim xTypecast As HTMLElementCollection
    Dim element As IHTMLScriptElement
    Dim theurl As String = ""
    Dim DickGram As Recognition.DictationGrammar
    Dim objection As Integer
    Dim WithEvents waveOut As New NAudio.Wave.WaveOut()
    Dim WithEvents waveOut2 As New NAudio.Wave.WaveOut()
    Dim SECONDARIES As Boolean
    Dim CURRENTQUESTION(31) As String
    Dim HOMEOWNER As Boolean = False
    Dim isplaying As Boolean
    Dim INSCO As String
    Dim POLSTART As String
    Dim POLEnd As String    ' hehe, Poland.
    Dim CurrentQ As Integer
    Dim LastCustomer
    Dim CustomerName As String
    Dim globalFile As String
    Dim globalFile2 As String
    Dim clipnum(50) As Integer
    Dim globalfile3 As String
    Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Integer, ByVal lpWindowText As String, ByVal cch As Integer) As Integer
    Public Delegate Function CallBack(ByVal hwnd As Integer, ByVal lParam As Integer) As Boolean
    Public Declare Function EnumWindows Lib "user32" (ByVal Adress As CallBack, ByVal y As Integer) As Integer
    Public Declare Function IsWindowVisible Lib "user32.dll" (ByVal hwnd As IntPtr) As Boolean
    Private Const SW_HIDE As Integer = 0
    Private Const SW_RESTORE As Integer = 9
    Private hWnd As Integer
    Public selectedIndex As Integer
    Private ActiveWindows As New System.Collections.ObjectModel.Collection(Of IntPtr)
    Public Const MOD_ALT As Integer = &H1 'Alt key
    Public Const MOD_CONTROL As Integer = &H2
    Public Const WM_HOTKEY As Integer = &H312
    Dim deviceNum1 As Integer = 0
    Dim DeviceNum2 As Integer = 0
    Dim totalCalls As Integer
    Dim totalLeads As Integer
    Dim NewMod As String = ""
    Dim anobjection As Boolean = False
    Dim isCompany As Boolean = False
    Dim timesAsking As Integer = 0
    Dim NICount As Integer = 0
    Dim VYear(3) As String
    Dim vMake(3) As String
    Dim vmodel(3) As String
    Dim secondPass As Boolean = False
    Dim PRIMARY_KEY As String = "ce43e8a4d7a844b1be7950b260d6b8bd"
    Dim timeout As Integer = 30000
    Dim numRepeats As Integer = 0
    Dim Newvar As String = ""
    Dim ParsedInsurance As String
    Dim isoption As Boolean = False
    Dim partialMatch As Integer
    Dim playcounter As Integer = 0
    Dim Playlist(1) As String
    Dim bday1 As String
    Dim bmonth1 As String
    Dim byear1 As String
    Dim AutoClip As String
    Dim stringArray() As String
    Dim Years() As String = {"81", "82", "83", "84", "85", "86", "87", "88", "89", "90", "91", "92", "93", "94", "95", "96", "97", "98", "99", "2000", "2001", "2002", "2003", "2004", "2005", "2006", "2007",
        "2008", "2009", "2010", "2011", "2012", "2013", "2014", "2015", "2016", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "16", "17"}
    Dim whatGotSaid As String = ""
    Dim s As String = ""
    Dim NEWEMAIL As String
    Dim StringSplit() As String
    Dim Clarify As Integer = 0
    Dim emailPart1 As String
    Dim fullEmail As String
    Dim StreetSpelling As String = ""
    Dim rawAddress As String = ""
    Dim NewAddress As String = ""
    Dim zip As String
    Dim theOldWindowTitle As String = "OldWindow"
    Dim theWindowTitle As String
    Dim CustName(1) As String
    Dim F As New Form
    Dim oldCust(1) As String

    Public Function getModel(ByRef VehicleNum As Integer) As Boolean
        Thread.Sleep(750)
        If local_browser.FindElementById("vehicle-make").Text <> vMake(VehicleNum) Then
            local_browser.FindElementById("vehicle-make").SendKeys(vMake(VehicleNum))
        End If
        Console.WriteLine("Getting model: ")
        selectElement = New SelectElement(local_browser.FindElementById("vehicle-model"))
        Dim Model_List As New List(Of String)
        For i As Integer = 0 To selectElement.Options.Count - 1
            Model_List.Add(selectElement.Options(i).Text)
        Next
        Dim split_speech() As String = s.Split()
        For z As Integer = 0 To split_speech.Length - 1
            split_speech(z) = split_speech(z).Replace("?", "")
            split_speech(z) = UCase(split_speech(z).Replace(".", " "))
        Next
        For x As Integer = 0 To Model_List.Count - 1
            For Y As Integer = 0 To split_speech.Length - 1
                Console.WriteLine(Model_List.Item(x) & ">>>>" & split_speech(Y))
                If Model_List.Item(x).Contains(split_speech(Y)) And split_speech(Y).Length > 2 Then
                    Console.WriteLine("Found it bitch!")
                    selectElement.SelectByText(Model_List.Item(x))
                    Return True
                End If
            Next
        Next
        Return False
    End Function  '

    Function getBirthdaWAV() As Boolean
        Try
            If local_browser.FindElementById("frmDOB_Month").GetAttribute("value") <> "" And local_browser.FindElementById("frmDOB_Day").GetAttribute("value") <> "" And local_browser.FindElementById("frmDOB_Year").GetAttribute("value") <> "" Then
                bmonth1 = local_browser.FindElementById("frmDOB_Month").GetAttribute("value")
                bday1 = local_browser.FindElementById("frmDOB_Day").GetAttribute("value")
                If bday1.Length < 2 Then
                    bday1 = bday1.Insert(0, "0")
                End If
                byear1 = local_browser.FindElementById("frmDOB_Year").GetAttribute("value")
                txtDOB.Text = bmonth1 & "/" & bday1 & "/" & byear1.Substring(2)
                isQuestion = True
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            Return False

        End Try
    End Function 'Checks to see if the birthday exists in the autoform so it can verify, if not it returns false to ask
    Function getBDayValues(text As String) As String()
        Dim tempArray() As String = text.Split("/")
        Select Case tempArray(0)
            Case "1", "01"
                tempArray(0) = "Jan"
            Case "2", "02"
                tempArray(0) = "Feb"
            Case "3", "03"
                tempArray(0) = "Mar"
            Case "4", "04"
                tempArray(0) = "Apr"
            Case "5", "05"
                tempArray(0) = "May"
            Case "6", "06"
                tempArray(0) = "Jun"
            Case "7", "07"
                tempArray(0) = "July"
            Case "8", "08"
                tempArray(0) = "Aug"
            Case "9", "09"
                tempArray(0) = "Sep"
            Case "10"
                tempArray(0) = "Oct"
            Case "11"
                tempArray(0) = "Nov"
            Case "12"
                tempArray(0) = "Dec"
        End Select
        Try
            tempArray(2) = tempArray(2).Insert(0, "19")
        Catch
            Console.WriteLine("Error: Birthday not entered")

        End Try

        Return tempArray
    End Function


    Dim Recording_status As Boolean
    Sub updateLabel()
        lblRecording.Text = callPos & ":       RECORDING: " & Recording_status
        If Recording_status = False And newcall = False Then
            m.StartMicAndRecognition()
        End If
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim i As Integer = 0
            cmbDispo.Text = "Not Available"
            CURRENTQUESTION(0) = "Hello"
            CURRENTQUESTION(1) = "Hello"
            CURRENTQUESTION(2) = CustomerName
            CURRENTQUESTION(3) = "INTRO/INSURANCE PROVIDER"
            CURRENTQUESTION(4) = "Insurance Expiration"
            CURRENTQUESTION(5) = "Insurance Start"
            CURRENTQUESTION(6) = "How Many Vehicles"
            CURRENTQUESTION(7) = "Vehicle Year/Make/Model"
            CURRENTQUESTION(8) = "DOB"
            CURRENTQUESTION(9) = "Marital Status"
            CURRENTQUESTION(10) = "Spouse Name"
            CURRENTQUESTION(13) = "Spouse DOB"
            CURRENTQUESTION(14) = "Own/Rent"
            CURRENTQUESTION(15) = "Home Type"
            CURRENTQUESTION(16) = "Address"
            CURRENTQUESTION(17) = "Email"
            CURRENTQUESTION(18) = "Credit"
            CURRENTQUESTION(19) = "Phone Type"
            CURRENTQUESTION(20) = "Spelling of Last Name"
            CURRENTQUESTION(21) = "Secondaries"
            CURRENTQUESTION(22) = "YEAR BUILT"
            CURRENTQUESTION(23) = "SQARE FEET"
            CURRENTQUESTION(24) = "TCPA"
            CURRENTQUESTION(25) = "TCPA"
            CURRENTQUESTION(26) = ""
            CURRENTQUESTION(27) = ""
            CURRENTQUESTION(28) = ""
            CURRENTQUESTION(29) = "TCPA"
            CURRENTQUESTION(30) = "TCPA"
            CURRENTQUESTION(31) = "DISPO"


            Dim DeviceCount As Integer = NAudio.Wave.WaveOut.DeviceCount()              'Gets The number of audio devices on the machine
            Dim SDevice As String = Nothing
            Dim RDevice As String = Nothing
            cmbMoreVehicles.SelectedIndex = 0
            CurrentQ = 1
            For i = 0 To DeviceCount - 1                                            'This loop fills the audiodevices into the primary and secondary audio comboboxes
                SDevice = NAudio.Wave.WaveOut.GetCapabilities(i).ProductName
                Primary.Items.Add(SDevice)
            Next
            Primary.SelectedIndex() = 0
            deviceNum1 = 0
            Register()

        Catch ex As Exception
        End Try
        Const Key As String = "ce43e8a4d7a844b1be7950b260d6b8bd"
        Const Key2 As String = "0d2797650c8648d18474399744512f17"
        m = SpeechRecognitionServiceFactory.CreateMicrophoneClient(SpeechRecognitionMode.LongDictation, "en-us", Key, Key2)
    End Sub

    Dim happytreefriends As FirefoxBinary = New FirefoxBinary(Application.StartupPath & "\core\firefox.exe")

    Dim prof As FirefoxProfile = New FirefoxProfile()


    Public local_browser As ChromeDriver


    Public Sub Unregister()

        UnregisterHotKey(Me.Handle, 101)
        UnregisterHotKey(Me.Handle, 201)
        UnregisterHotKey(Me.Handle, 301)
        UnregisterHotKey(Me.Handle, 102)
        UnregisterHotKey(Me.Handle, 103)
        UnregisterHotKey(Me.Handle, 104)

        'Register TIE INSystem.Windows.Forms.Keys.
        UnregisterHotKey(Me.Handle, 105)
        UnregisterHotKey(Me.Handle, 106)
        UnregisterHotKey(Me.Handle, 206)
        UnregisterHotKey(Me.Handle, 306)
        UnregisterHotKey(Me.Handle, 107)
        UnregisterHotKey(Me.Handle, 207)
        UnregisterHotKey(Me.Handle, 307)
        UnregisterHotKey(Me.Handle, 108)
        UnregisterHotKey(Me.Handle, 208)
        UnregisterHotKey(Me.Handle, 308)

        'Register REBUTTALSystem.Windows.Forms.Keys.
        UnregisterHotKey(Me.Handle, 109)
        UnregisterHotKey(Me.Handle, 110)
        UnregisterHotKey(Me.Handle, 111)
        UnregisterHotKey(Me.Handle, 112)

        UnregisterHotKey(Me.Handle, 169)
        UnregisterHotKey(Me.Handle, 170)
        UnregisterHotKey(Me.Handle, 171)
        UnregisterHotKey(Me.Handle, 172)
        UnregisterHotKey(Me.Handle, 173)

    End Sub                      'Unregisters global hotkeys
    Public Sub rolltheclip()

        StopThatClip()
        waveOut = New NAudio.Wave.WaveOut()
        If My.Computer.FileSystem.FileExists(clipname) Then
            Dim mp3File As New NAudio.Wave.Mp3FileReader(clipname)
            waveOut.DeviceNumber = deviceNum1
            waveOut.Init(mp3File)
            waveOut.Play()
        Else
            Console.WriteLine(clipname & " not available")
        End If
    End Sub        'Plays sound clips through whatever audio outs are selected
    Sub SpeechtoVar(speech As String) 'to break up month/year'
        Select Case speech
            Case "January", "Next January", "This January"
                theMonth = 1
                writtenMonth = "Jan"
                BMonth = "01"
            Case "February", "Next February", "This February"
                theMonth = 2
                writtenMonth = "Feb"
                BMonth = "02"
            Case "March", "Next March", "This March"
                theMonth = 3
                BMonth = "03"
                writtenMonth = "Mar"
            Case "April", "Next April", "This April"
                theMonth = 4
                BMonth = "04"
                writtenMonth = "Apr"
            Case "May", "Next May", "This May"
                theMonth = 5
                BMonth = "05"
                writtenMonth = "May"
            Case "June", "Next June", "This June"
                theMonth = 6
                BMonth = "06"
                writtenMonth = "Jun"
            Case "July", "Next July", "This July"
                theMonth = 7
                BMonth = "07"
                writtenMonth = "Jul"
            Case "August", "Next August", "This August"
                theMonth = 8
                BMonth = "08"
                writtenMonth = "Aug"
            Case "September", "Next September", "This September"
                theMonth = 9
                BMonth = "09"
                writtenMonth = "Sep"
            Case "October"
                theMonth = 10
                BMonth = "10"
                writtenMonth = "Oct"

            Case "November"
                theMonth = 11
                BMonth = "11"
                writtenMonth = "Nov"
            Case "December"
                writtenMonth = "Dec"
                theMonth = 12
                BMonth = "12"
        End Select
        Select Case speech
            Case "January", "February", "March"
                theYear = "2017"
            Case "Next January", "Next February", "Next March", "Next April", "Next May", "Next June", "Next July", "Next August", "Next September", "Next October", "Next November", "Next December"
                theYear = "2017"
            Case "April", "May", "June", "July", "August", "September", "October", "November", "December"
                theYear = "2016"
        End Select
    End Sub
    Public Function WordToNum(phrase As String) As String
        If phrase.Contains("one") Then
            If phrase.Contains("twenty one") Then
                Return "21"
            ElseIf phrase.Contains("thirty one") Then
                Return "31"
            ElseIf phrase.Contains("forty one") Then
                Return ("41")
            ElseIf phrase.Contains("fifty one") Then
                Return "51"
            ElseIf phrase.Contains("sixty one") Then
                Return ("61")
            ElseIf phrase.Contains("seventy one") Then
                Return "71"
            ElseIf phrase.Contains("eighty one") Then
                Return ("81")
            ElseIf phrase.Contains("ninety one") Then
                Return "91"
            Else
                Return "1"
            End If
        ElseIf phrase.Contains("two") Then
            If phrase.Contains("twenty two") Then
                Return "22"
            ElseIf phrase.Contains("thirty two") Then
                Return "32"
            ElseIf phrase.Contains("forty two") Then
                Return ("42")
            ElseIf phrase.Contains("fifty two") Then
                Return "52"
            ElseIf phrase.Contains("sixty two") Then
                Return ("62")
            ElseIf phrase.Contains("seventy two") Then
                Return "72"
            ElseIf phrase.Contains("eighty two") Then
                Return ("82")
            ElseIf phrase.Contains("ninety two") Then
                Return "92"
            Else
                Return "2"
            End If
        ElseIf phrase.Contains("three") Then
            If phrase.Contains("twenty three") Then
                Return "23"
            ElseIf phrase.Contains("thirty three") Then
                Return "33"
            ElseIf phrase.Contains("forty three") Then
                Return ("43")
            ElseIf phrase.Contains("fifty three") Then
                Return "53"
            ElseIf phrase.Contains("sixty three") Then
                Return ("63")
            ElseIf phrase.Contains("seventy three") Then
                Return "73"
            ElseIf phrase.Contains("eighty three") Then
                Return ("83")
            ElseIf phrase.Contains("ninety three") Then
                Return "93"
            Else
                Return "3"
            End If
        ElseIf phrase.Contains("four") Or phrase.Contains("For") Then
            If phrase.Contains("fourteen") Then
                Return ("14")
            ElseIf phrase.Contains("twenty four") Then
                Return "24"
            ElseIf phrase.Contains("thirty four") Then
                Return "34"
            ElseIf phrase.Contains("forty four") Then
                Return ("44")
            ElseIf phrase.Contains("fifty four") Then
                Return "54"
            ElseIf phrase.Contains("sixty four") Then
                Return ("64")
            ElseIf phrase.Contains("seventy four") Then
                Return "74"
            ElseIf phrase.Contains("eighty four") Then
                Return ("84")
            ElseIf phrase.Contains("ninety four") Then
                Return "94"
            Else
                Return "4"
            End If

        ElseIf phrase.Contains("five") Then
            If phrase.Contains("twenty five") Then
                Return "25"
            ElseIf phrase.Contains("thirty five") Then
                Return "35"
            ElseIf phrase.Contains("forty five") Then
                Return ("45")
            ElseIf phrase.Contains("fifty five") Then
                Return "55"
            ElseIf phrase.Contains("sixty five") Then
                Return ("65")
            ElseIf phrase.Contains("seventy five") Then
                Return "75"
            ElseIf phrase.Contains("eighty five") Then
                Return ("85")
            ElseIf phrase.Contains("ninety five") Then
                Return "95"
            Else
                Return "5"
            End If
        ElseIf phrase.Contains("six") Then
            If phrase.Contains("sixteen") Then
                Return "16"

            ElseIf phrase.Contains("twenty six") Then
                Return "26"
            ElseIf phrase.Contains("thirty six") Then
                Return "36"
            ElseIf phrase.Contains("forty six") Then
                Return ("46")
            ElseIf phrase.Contains("fifty six") Then
                Return "56"
            ElseIf phrase.Contains("sixty six") Then
                Return ("66")
            ElseIf phrase.Contains("seventy six") Then
                Return "76"
            ElseIf phrase.Contains("eighty six") Then
                Return ("86")
            ElseIf phrase.Contains("ninety six") Then
                Return "96"
            Else
                Return "6"
            End If
        ElseIf phrase.Contains("seven") Then
            If phrase.Contains("seventeen") Then
                Return "17"
            ElseIf phrase.Contains("twenty seven") Then
                Return "27"
            ElseIf phrase.Contains("thirty seven") Then
                Return "37"
            ElseIf phrase.Contains("forty seven") Then
                Return ("47")
            ElseIf phrase.Contains("fifty seven") Then
                Return "57"
            ElseIf phrase.Contains("sixty seven") Then
                Return ("67")
            ElseIf phrase.Contains("seventy seven") Then
                Return "77"
            ElseIf phrase.Contains("eighty seven") Then
                Return ("87")
            ElseIf phrase.Contains("ninety seven") Then
                Return "97"
            Else
                Return "7"
            End If
        ElseIf phrase.Contains("eight") Then
            If phrase.Contains("eighteen") Then
                Return "18"
            ElseIf phrase.Contains("twenty eight") Then
                Return "28"
            ElseIf phrase.Contains("thirty eight") Then
                Return "38"
            ElseIf phrase.Contains("forty eight") Then
                Return ("48")
            ElseIf phrase.Contains("fifty eight") Then
                Return "58"
            ElseIf phrase.Contains("sixty eight") Then
                Return ("68")
            ElseIf phrase.Contains("seventy eight") Then
                Return "78"
            ElseIf phrase.Contains("eighty eight") Then
                Return ("88")
            ElseIf phrase.Contains("ninety eight") Then
                Return "98"
            Else
                Return "8"
            End If
        ElseIf phrase.Contains("nine") Then
            If phrase.Contains("nineteen") Then
                Return "19"

            ElseIf phrase.Contains("twenty nine") Then
                Return "29"
            ElseIf phrase.Contains("thirty nine") Then
                Return "39"
            ElseIf phrase.Contains("forty nine") Then
                Return ("49")
            ElseIf phrase.Contains("fifty nine") Then
                Return "59"
            ElseIf phrase.Contains("sixty nine") Then
                Return ("69")
            ElseIf phrase.Contains("seventy nine") Then
                Return "79"
            ElseIf phrase.Contains("eighty nine") Then
                Return ("89")
            ElseIf phrase.Contains("ninety nine") Then
                Return "99"
            Else
                Return "9"
            End If

        ElseIf phrase.Contains("twenty") Then
            If phrase.Contains("twenty one") Then
                Return "21"
            ElseIf phrase.Contains("twenty two") Then
                Return "22"
            ElseIf phrase.Contains("twenty three") Then
                Return "23"
            ElseIf phrase.Contains("twenty four") Then
                Return "24"
            ElseIf phrase.Contains("twenty five") Then
                Return "25"
            ElseIf phrase.Contains("twenty six") Then
                Return "26"
            ElseIf phrase.Contains("twenty seven") Then
                Return "27"
            ElseIf phrase.Contains("twenty eight") Then
                Return "28"
            ElseIf phrase.Contains("twenty nine") Then
                Return "29"
            Else
                Return "20"
            End If
        ElseIf phrase.Contains("thirty") Then
            If phrase.Contains("thirty one") Then
                Return "31"
            ElseIf phrase.Contains("thirty two") Then
                Return "33"
            ElseIf phrase.Contains("thirty three") Then
                Return "33"
            ElseIf phrase.Contains("thirty four") Then
                Return "34"
            ElseIf phrase.Contains("thirty five") Then
                Return "35"
            ElseIf phrase.Contains("thirty six") Then
                Return "36"
            ElseIf phrase.Contains("thirty seven") Then
                Return "37"
            ElseIf phrase.Contains("thirty eight") Then
                Return "38"
            ElseIf phrase.Contains("thirty nine") Then
                Return "39"
            Else
                Return "30"
            End If
        ElseIf phrase.Contains("forty") Then
            If phrase.Contains("forty one") Then
                Return "41"
            ElseIf phrase.Contains("forty two") Then
                Return "44"
            ElseIf phrase.Contains("forty three") Then
                Return "43"
            ElseIf phrase.Contains("forty four") Then
                Return "44"
            ElseIf phrase.Contains("forty five") Then
                Return "45"
            ElseIf phrase.Contains("forty six") Then
                Return "46"
            ElseIf phrase.Contains("forty seven") Then
                Return "47"
            ElseIf phrase.Contains("forty eight") Then
                Return "48"
            ElseIf phrase.Contains("forty nine") Then
                Return "49"
            Else
                Return "40"
            End If
        ElseIf phrase.Contains("fifty") Then
            If phrase.Contains("fifty one") Then
                Return "51"
            ElseIf phrase.Contains("fifty two") Then
                Return "55"
            ElseIf phrase.Contains("fifty three") Then
                Return "53"
            ElseIf phrase.Contains("fifty four") Then
                Return "54"
            ElseIf phrase.Contains("fifty five") Then
                Return "55"
            ElseIf phrase.Contains("fifty six") Then
                Return "56"
            ElseIf phrase.Contains("fifty seven") Then
                Return "57"
            ElseIf phrase.Contains("fifty eight") Then
                Return "58"
            ElseIf phrase.Contains("fifty nine") Then
                Return "59"
            Else
                Return "50"
            End If
        ElseIf phrase.Contains("sixty") Then
            If phrase.Contains("sixty one") Then
                Return "61"
            ElseIf phrase.Contains("sixty two") Then
                Return "66"
            ElseIf phrase.Contains("sixty three") Then
                Return "63"
            ElseIf phrase.Contains("sixty four") Then
                Return "64"
            ElseIf phrase.Contains("sixty five") Then
                Return "65"
            ElseIf phrase.Contains("sixty six") Then
                Return "66"
            ElseIf phrase.Contains("sixty seven") Then
                Return "67"
            ElseIf phrase.Contains("sixty eight") Then
                Return "68"
            ElseIf phrase.Contains("sixty nine") Then
                Return "69"
            Else
                Return "60"
            End If
        ElseIf phrase.Contains("seventy") Then
            If phrase.Contains("seventy one") Then
                Return "71"
            ElseIf phrase.Contains("seventy two") Then
                Return "77"
            ElseIf phrase.Contains("seventy three") Then
                Return "73"
            ElseIf phrase.Contains("seventy four") Then
                Return "74"
            ElseIf phrase.Contains("seventy five") Then
                Return "75"
            ElseIf phrase.Contains("seventy six") Then
                Return "76"
            ElseIf phrase.Contains("seventy seven") Then
                Return "77"
            ElseIf phrase.Contains("seventy eight") Then
                Return "78"
            ElseIf phrase.Contains("seventy nine") Then
                Return "79"
            Else
                Return "70"
            End If
        ElseIf phrase.Contains("eighty") Then
            If phrase.Contains("eighty one") Then
                Return "81"
            ElseIf phrase.Contains("eighty two") Then
                Return "88"
            ElseIf phrase.Contains("eighty three") Then
                Return "83"
            ElseIf phrase.Contains("eighty four") Then
                Return "84"
            ElseIf phrase.Contains("eighty five") Then
                Return "85"
            ElseIf phrase.Contains("eighty six") Then
                Return "86"
            ElseIf phrase.Contains("eighty seven") Then
                Return "87"
            ElseIf phrase.Contains("eighty eight") Then
                Return "88"
            ElseIf phrase.Contains("eighty nine") Then
                Return "89"
            Else
                Return "80"
            End If
        ElseIf phrase.Contains("ninety") Then
            If phrase.Contains("ninety one") Then
                Return "91"
            ElseIf phrase.Contains("ninety two") Then
                Return "99"
            ElseIf phrase.Contains("ninety three") Then
                Return "93"
            ElseIf phrase.Contains("ninety four") Then
                Return "94"
            ElseIf phrase.Contains("ninety five") Then
                Return "95"
            ElseIf phrase.Contains("ninety six") Then
                Return "96"
            ElseIf phrase.Contains("ninety seven") Then
                Return "97"
            ElseIf phrase.Contains("ninety eight") Then
                Return "98"
            ElseIf phrase.Contains("ninety nine") Then
                Return "99"
            Else
                Return "90"
            End If
        ElseIf phrase.Contains("ten") Then
            Return ("10")
        ElseIf phrase.Contains("eleven") Then
            Return "11"
        ElseIf phrase.Contains("twelve") Then
            Return ("12")
        ElseIf phrase.Contains("thirteen") Then
            Return ("13")
        ElseIf phrase.Contains("fifteen") Then
            Return ("15")
        Else
            Return "null"

        End If

    End Function
    Public Function OnlyNumbers(speech As String) As String
        Dim x As Integer = 0
        Dim newstring As String = ""
        Dim tempString() As String = Split(speech)
        For x = 0 To tempString.Length - 1
            Select Case tempString(x)
                Case "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten", "elven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen", "twenty", "thirty", "fourty", "fifty", "sixty", "seventy", "eighty", "ninety"
                    newstring += (tempString(x)) & " "
                Case Else

            End Select
        Next
        Return Trim(newstring)
    End Function
    Public Function getMonth(monthnum As String) As String
        Select Case monthnum
            Case "01", "1"
                Return "Jan"
            Case "02", "2"
                Return "Feb"

            Case "03", "3"
                Return "Mar"
            Case "04", "4"
                Return "Apr"
            Case "05", "5"
                Return "May"
            Case "06", "6"
                Return "June"
            Case "07", "7"
                Return "Jul"
            Case "08", "8"
                Return "Aug"
            Case "09", "9"
                Return "Sep"
            Case "10"
                Return "Oct"
            Case "11"
                Return "Nov"
            Case "12"
                Return "Dec"
            Case Else
                Return "null"
        End Select
    End Function
    Sub DeriveBDay(speech As String)
        Dim counter As Integer = 0
        Dim x As Integer
        Dim Y As Integer
        Dim Bdays() As String = speech.Split(" ")
        If speech.Contains("January") Or speech.Contains("February") Or speech.Contains("March") Or speech.Contains("April") Or speech.Contains("May") Or speech.Contains("June") Or speech.Contains("July") Or
                               speech.Contains("August") Or speech.Contains("September") Or speech.Contains("October") Or speech.Contains("November") Or speech.Contains("December") Then
            BMonth = getMonth(Bdays(0))
            Console.WriteLine(BMonth)
            If Bdays(1).Contains(", ") Then
                BDay = Bdays(1).Substring(0, Bdays(1).Length - 1)
                BYear = Bdays(2)
            Else
                BDay = Bdays(1)
                BYear = Bdays(Bdays.Length - 1)
            End If
        Else
            Dim stringArray() As String = (OnlyNumbers(BDayString).Split)




            If stringArray(stringArray.Length - 1) = "one" Or stringArray(stringArray.Length - 1) = "two" Or stringArray(stringArray.Length - 1) = "three" Or stringArray(stringArray.Length - 1) = "four" Or stringArray(stringArray.Length - 1) = "five" Or stringArray(stringArray.Length - 1) = "six" Or stringArray(stringArray.Length - 1) = "seven" Or stringArray(stringArray.Length - 1) = "eight" Or stringArray(stringArray.Length - 1) = "nine" Then
                BYear = stringArray(stringArray.Length - 2) & " " & stringArray(stringArray.Length - 1)
                BYear = "19" & WordToNum(BYear)
                Console.WriteLine(BYear)
                Y = stringArray.Length - 3
                ReDim finalCompare(Y)
                For counter = 1 To stringArray.Length - 2
                    Try
                        finalCompare(counter) = stringArray(counter)
                    Catch
                    End Try
                Next
            Else
                BYear = "19" & WordToNum(stringArray(stringArray.Length - 1))
                For counter = 1 To stringArray.Length - 1
                    Try
                        finalCompare(counter) = stringArray(counter)
                    Catch
                    End Try
                Next
            End If
            BMonth = getMonth(WordToNum(stringArray(0)))
            If finalCompare.Length > 1 Then
                If finalCompare(finalCompare.Length - 1) = "nineteen" Then
                    Try
                        For x = 0 To finalCompare.Length - 2
                            BDay += finalCompare(x)
                            BDay += " "
                        Next
                        BDay = WordToNum(Trim(BDay))
                    Catch
                    End Try
                Else
                    For x = 0 To finalCompare.Length - 1
                        BDay += finalCompare(x)
                        BDay += " "
                    Next
                    BDay = WordToNum(Trim(BDay))
                End If
            Else
                BDay = WordToNum(finalCompare(finalCompare.Length - 1))
            End If
        End If
    End Sub
    Sub VerifyModel(speech As String, vehicleNum As Integer)
        isoption = False
        partialMatch = 0
        Dim VehicleID As String = ""

        If vehicleNum = 1 Then
            VehicleID = "vehicle-model"
        Else
            VehicleID = "vehicle" & CStr(vehicleNum) & "-model"
        End If
        Dim v As Integer = 0
        '    For v = 0 To 'LeadForm.Document.GetElementById(VehicleID).GetAttribute("length")
        ''LeadForm.Document.GetElementById(VehicleID).SetAttribute("selectedIndex", v)

        ' If speech = 'LeadForm.Document.GetElementById(VehicleID).GetAttribute("value") Then
        '  isoption = True

        '   Exit For
        '       End If
        '   Next
    End Sub
    Public Function getYear(VehicleNum As Integer) As Boolean
        temperstring = s
        Dim X As Integer
        For X = 1 To Years.Length - 1
            If s.Contains(Years(X)) Then
                If CStr(Years(X)) > 80 And CStr(Years(X)) < 100 Then
                    VYear(VehicleNum) = "19" & Years(X)
                    Exit For
                ElseIf CStr(Years(X)) > 6 And CStr(Years(X)) < 17 Then
                    VYear(VehicleNum) = "20" & Years(X)
                    Exit For
                ElseIf CStr(Years(X)) > 1980 And CStr(Years(X)) < 2017 Then
                    VYear(VehicleNum) = Years(X)
                    Exit For

                End If
            End If
        Next
        If VYear(VehicleNum) <> "" Then
            If VehicleNum = 1 Then
                local_browser.FindElementById("vehicle-year").SendKeys(VYear(VehicleNum))
                local_browser.Keyboard.PressKey(Keys.Return)
                CurrentQ += 1
                Return True
            Else
                local_browser.FindElementById("vehicle" & VehicleNum & "-year").SendKeys(VYear(VehicleNum))
                local_browser.Keyboard.PressKey(Keys.Return)
                CurrentQ += 1
                Return True
            End If
        Else
            isQuestion = True
            Return False
        End If
    End Function 'GETS THE VEHICLE YEAR
    Dim BADINFOCOUNTER As Integer = 0


    Dim ModelHolder As String = ""
    Dim numbreps As Integer = 0

    Public Function isMachine()
        Select Case True
            Case s.Contains("leave a message"), s.Contains("unable to take your call"), s.Contains("after the beep"), s.Contains("after the tone"), s.Contains("at the tone"), s.Contains("leave your Name"), s.Contains("mailbox is full")
                cmbDispo.Text = "Not Available"
                CurrentQ = 31
                DispositionCall()
                Return True
            Case Else
                Return False

        End Select

    End Function                       'checks to see if the initial speech received can confirm an answering machine
    Dim clipType As String = ""
    Public Function HandlePartObjection() As Boolean
        isQuestion = True
        Part = LCase(Part)
        If Part <> "" Then
            Console.WriteLine("CHECKING AGAINST PARTIAL OBJECTIONS")
            Console.WriteLine(Part)

            Try

                Select Case True
                    Case Part.Contains("is this a real person"), Part.Contains("is this a recording"), s.Contains("robot"), s.Contains("automated")

                        rolltheclipThread("C:\Soundboard\Cheryl\REACTIONS\Loud-laugh.mp3")
                        Timer2.Enabled = True
                        NICount += 1
                        Return True
                    Case Part.Contains("no vehicle"), Part.Contains("sold the car"), Part.Contains("sold my car"), Part.Contains("no car"), Part.Contains("don't have a vehicle"), Part.Contains("don't") And Part.Contains("have a car"), Part.Contains("don't have an automobile"), Part.Contains("dont't have my own car"), Part.Contains("doesn't have a car"),
                        newobjection = False

                        Console.WriteLine("THEY DON'T HAVE A CAR")
                        rolltheclipThread("C:/Soundboard/Cheryl/WRAPUP/have a great day.mp3")
                        cmbDispo.Text = "No Car"
                        CurrentQ = 31
                        Timer2.Enabled = True
                        counter2 = 0
                        Return True
                    Case Part.Contains("not interested"), Part.Contains("don't need a quote"), Part.Contains("i'm fine"), Part.Contains("not really interested"), Part.Contains("not in arrested"), Part.Contains("that's okay thank you"), Part.Contains("no interest"), Part.Contains("stop calling"), Part.Contains("i'm good"), Part.Contains("all set"), Part.Contains("don't want it"), Part.Contains("not changing"), Part.Contains("i'm happy with"), Part.Contains("very happy"), Part.Contains("no thank you"), Part.Contains("not looking"), Part.Contains("don't wanna change"), Part.Contains("no thank you"), Part.Contains("don't need insurance") 'NI

                        newobjection = False
                        Console.WriteLine("NOT INTERESTED")
                        If CurrentQ = 3 Then
                            CurrentQ = 0
                        End If

                        Select Case NICount
                            Case 0
                                If counter2 < 2 Then
                                    Playlist(0) = "c:\soundboard\cheryl\REBUTTALS\I COMPLETELY UNDERSTAND.mp3"
                                    Playlist(1) = ("C:\soundboard\cheryl\INTRO\THISISTOGIVENEWQUOTE.mp3")
                                    NICount += 1
                                    counter2 += 1
                                    If CurrentQ = 3 Then
                                        CurrentQ = 0
                                    End If
                                    Timer2.Enabled = True
                                    Return True
                                Else
                                    rolltheclipThread("C:/Soundboard/Cheryl/WRAPUP/have a great day.mp3")
                                    cmbDispo.Text = "Not Interested"
                                    CurrentQ = 31
                                    Timer2.Enabled = True
                                End If

                            Case 1

                                If counter2 < 2 Then

                                    Playlist(0) = ("C:\SoundBoard\Cheryl\reactions\I get that 3.mp3 ")
                                    Playlist(1) = ("C:\SoundBoard\Cheryl\REBUTTALS\REBUTTAL1.mp3")
                                    numbreps += 1
                                    If CurrentQ = 3 Then
                                        CurrentQ = 0
                                    End If
                                    counter2 += 1


                                    Return True
                                Else
                                    rolltheclipThread("C:/Soundboard/Cheryl/WRAPUP/have a great day.mp3")
                                    cmbDispo.Text = "Not Interested"
                                    CurrentQ = 31
                                    Timer2.Enabled = True
                                    counter2 = 0
                                End If
                            Case Else
                                rolltheclipThread("C:/Soundboard/Cheryl/WRAPUP/have a great day.mp3")
                                cmbDispo.Text = "Not Interested"
                                CurrentQ = 31
                                Timer2.Enabled = True
                                counter2 = 0
                        End Select

                    Case Part.Contains("busy"), Part.Contains("at work"), Part.Contains("driving"), Part.Contains("can't talk"), Part.Contains("call me back"), Part.Contains("could you call back"), Part.Contains("call back another time"), Part.Contains("call later"), Part.Contains("working right now")
                        newobjection = False

                        If CurrentQ = 3 Then
                            CurrentQ = 0
                        End If

                        Select Case counter
                            Case 0
                                rolltheclipThread("C:\SoundBoard\Cheryl\REBUTTALS\THIS WILL BE REAL QUICK.mp3")
                                Timer2.Enabled = True
                                NICount += 1
                                Return True
                            Case Else
                                rolltheclipThread("C:\SoundBoard\Cheryl\Birthday\questions 5-4-16\questions 5-4-16\Im busy.MP3")
                                Timer2.Enabled = True
                                NICount += 1
                                counter = 0
                                Return True

                        End Select
                    Case Part.Contains("wrong number"), Part.Contains("by that name"), Part.Contains("wrong phone number")
                        newobjection = False

                        rolltheclipThread("c:\soundboard\cheryl\Rebuttals\SORRY.mp3")
                        cmbDispo.Text = "Wrong Number"
                        CurrentQ = 31
                        Timer2.Enabled = True
                        Return True

                    Case Part.Contains("already have"), Part.Contains("already have insurance"), Part.Contains("already got insurance"), Part.Contains("happy with"), Part.Contains("i have insurance"), Part.Contains("i got insurance")

                        rolltheclipThread("C:\SoundBoard\Cheryl\Birthday\questions 5-4-16\questions 5-4-16\i have insurance.mp3")
                        Timer2.Enabled = True
                        NICount += 1
                        Return True


                    Case Part.Contains("take me off your list"), Part.Contains("name off your list"), Part.Contains("number off your list"), Part.Contains("take me off"), Part.Contains("take me off your call list"), Part.Contains("no call list"), Part.Contains("take this number off the list"), Part.Contains("do not call list"), Part.Contains("remove me from the list"), Part.Contains("taken off his collar"), Part.Contains("remove me from your calling list"), Part.Contains("call list"), Part.Contains("calling list")
                        newobjection = False

                        rolltheclipThread("C:\SoundBoard\Cheryl\REBUTTALS\DNC.mp3")
                        cmbDispo.Text = "Do Not Call"
                        CurrentQ = 31
                        Timer2.Enabled = True
                    Case Else



                End Select

                If CurrentQ = 2 Then
                    If CheckWhoseTalking2() Then
                        CurrentQ = 3

                        Timer2.Enabled = True
                    End If
                End If

            Catch EX As Exception
                Console.WriteLine("error with objection")

            End Try
        End If
        Return False
    End Function  'Handles Objection from the partial returned speech
    Dim dontKnowCount As Integer = 0
    Public Function HandleObjection(obj As String, ByRef numReps As Integer) As Boolean
        Console.WriteLine("CHECKING AGAINST OBJECTIONS")
        Console.WriteLine("reps:" & numReps)
        clipType = "Objection"

        Select Case True

            Case obj.Contains("no vehicle"), obj.Contains("no car"), obj.Contains("don't have a vehicle"), obj.Contains("don't have a car")

                Console.WriteLine("THEY DON'T HAVE A CAR")
                rolltheclipThread("C:/Soundboard/Cheryl/WRAPUP/have a great day.mp3")
                cmbDispo.Text = "No Car"
                Return True
                CurrentQ = 31
                Timer2.Enabled = True
                counter2 = 0



            Case obj.Contains("not interested"), obj.Contains("no interest"), obj.Contains("stop calling me"), obj.Contains("i'm good"), obj.Contains("all set"), obj.Contains("don't want it") 'NI
                Console.WriteLine("NOT INTERESTED")
                If CurrentQ = 3 Then
                    CurrentQ = 0
                End If
                Select Case numReps
                    Case 0
                        If counter2 < 2 Then
                            Playlist(0) = "c:\soundboard\cheryl\REBUTTALS\I COMPLETELY UNDERSTAND.mp3"
                            Playlist(1) = ("C:\soundboard\cheryl\INTRO\THISISTOGIVENEWQUOTE.mp3")
                            numReps += 1
                            counter2 += 1
                            If CurrentQ = 3 Then
                                CurrentQ = 0
                            End If
                            Timer2.Enabled = True

                            Return True
                        Else
                            rolltheclipThread("C:/Soundboard/Cheryl/WRAPUP/have a great day.mp3")
                            cmbDispo.Text = "Not Interested"
                            CurrentQ = 31
                            Timer2.Enabled = True
                        End If

                    Case 1

                        If counter2 < 2 Then

                            Playlist(0) = ("C:\soundboard\cheryl\reactions\I Get that 1.mp3")
                            Playlist(1) = ("C:\SoundBoard\Cheryl\REBUTTALS\REBUTTAL1.mp3")
                            numReps += 1
                            If CurrentQ = 3 Then
                                CurrentQ = 0
                            End If
                            counter2 += 1
                            Timer2.Enabled = True

                            Return True
                        Else
                            rolltheclipThread("C:/Soundboard/Cheryl/WRAPUP/have a great day.mp3")
                            cmbDispo.Text = "Not Interested"
                            CurrentQ = 31
                            Timer2.Enabled = True
                            counter2 = 0
                        End If
                    Case Else
                        rolltheclipThread("C:/Soundboard/Cheryl/WRAPUP/have a great day.mp3")
                        cmbDispo.Text = "Not Interested"
                        CurrentQ = 31
                        Timer2.Enabled = True
                        counter2 = 0
                End Select
            Case obj.Contains("don't know"), obj.Contains("no idea"), obj.Contains("no clue"), obj.Contains("not sure"), obj.Contains("couldn't tell you"), Part.Contains("you'd have to talk to")
                Console.WriteLine("THEY DON'T KNOW")
                tmrSilence.Enabled = True
                If CurrentQ = 3 Then
                    isQuestion = True
                    rolltheclipThread("c:\soundboard\cheryl\PUSHONS\allstategeicostatefarm.mp3")
                    Return True
                ElseIf CurrentQ = 4 Then
                    rolltheclipThread("C:/SOUNDBOARD/CHERYL/REBUTTALS/JANUARY FEB MARCH APRIL.mp3")
                    Return True
                ElseIf CurrentQ = 7 Then
                    rolltheclipThread("c:\soundboard\cheryl\PUSHONS\chevyfordgmc.mp3")
                    Return True
                Else
                    rolltheclipThread("C:\SoundBoard\Cheryl\TIE INS\Great What's Your Best Guess.mp3")
                    Return True
                End If
                If CurrentQ = 3 Then
                    CurrentQ = 0
                End If
            Case obj.Contains("busy"), obj.Contains("at work"), obj.Contains("driving"), obj.Contains("can't talk"), obj.Contains("no time"), obj.Contains("don't have time")
                If CurrentQ = 3 Then
                    CurrentQ = 0
                End If
                Select Case counter
                    Case >= 0
                        rolltheclipThread("C:\SoundBoard\Cheryl\REBUTTALS\THIS WILL BE REAL QUICK.mp3")
                        Timer2.Enabled = True
                        numReps += 1
                        Return True


                    Case Else
                        Return True

                End Select
            Case obj.Contains("wrong number"), obj.Contains("by that name"), obj.Contains("wrong phone number")
                rolltheclipThread("c:\soundboard\cheryl\Rebuttals\SORRY.mp3")
                cmbDispo.Text = "Wrong Number"
                CurrentQ = 31
                Timer2.Enabled = True
                Return True

            Case Else
                Return False

        End Select
        Return False
    End Function  'Handles Objection based on full returned result
    Dim quest As Integer = 1
    Public Function HandleQuestion(obj As String) As Boolean
        Console.WriteLine("CHECKING AGAINST QUESTIONS")
        Console.WriteLine("reps:" & quest)

        Select Case quest
            Case 1
                Playlist(0) = "c:\soundboard\cheryl\REBUTTALS\THAT'S A GREAT QUESTION.mp3"
                quest += 1
            Case 2
                Playlist(0) = "C:\SoundBoard\Cheryl\Birthday\questions 5-4-16\questions 5-4-16\whatta great question.mp3"
                quest = 1
        End Select

        Select Case True
            Case obj.Contains("who is this"), obj.Contains("who are you"), obj.Contains("who is calling"), obj.Contains("who's this"), obj.Contains("who's calling"), obj.Contains("who do you represent")
                If CurrentQ = 3 Then
                    CurrentQ = 0
                End If
                Playlist(1) = "c:\soundboard\cheryl\INTRO\CHERYLCALLING.mp3"
                Timer2.Enabled = True
                Return True
            Case obj.Contains("who makes it")
                rolltheclipThread("c:\soundboard\cheryl\REACTIONS\YES.mp3")
                Return True
            Case obj.Contains("what is this"), obj.Contains("what's this"), obj.Contains("what is the nature of this call"), obj.Contains("what are you calling about"), obj.Contains("what is purpose of this call")
                If CurrentQ = 3 Then
                    CurrentQ = 0
                End If
                Playlist(1) = "c:\soundboard\cheryl\INTRO\THISISTOGIVENEWQUOTE.mp3"
                Timer2.Enabled = True
                Return True
            Case obj.Contains("what is lcn"), obj.Contains("what is elsieanne")
                If CurrentQ = 3 Then
                    CurrentQ = 0
                End If
                Playlist(1) = "c:\soundboard\cheryl\Rebuttals\What's LCN.mp3"
                Timer2.Enabled = True
                Return True
            Case obj.Contains("why are you calling")
                Playlist(1) = "c:\soundboard\cheryl\INTRO\THISISTOGIVENEWQUOTE.mp3"
                Timer2.Enabled = True
                Return True
            Case obj.Contains("how did you get my information")
                Playlist(0) = "c:\soundboard\cheryl\REBUTTALS\Where Did You get My info.mp3"
                Timer2.Enabled = True
                Return True
            Case Else
                Return False
        End Select

    End Function 'Handles Question
    Public Function CheckWhoseTalking() As Boolean
        Select Case True
            Case s.Contains("this is"), s.Contains("speaking"), s.Contains("you've got him"), s.Contains("you've got her"), s.Contains("yes"), s.Contains("yeah"), s.Contains("what's up?"), s.Contains("how can i help you"), s.Contains("hey"), s.Contains("what do you want"), s.Contains("hello"), s.Contains("hi"), s.Contains("his spouse"), s.Contains("her spouse"), s.Contains("his wife"), s.Contains("her husband")
                Return True

            Case s.Contains("not home"), s.Contains("he isn't"), s.Contains("not available"), s.Contains("he's not"), s.Contains(" a message"), s.Contains("he's working"), s.Contains("not here"), s.Contains("not right now")
                rolltheclipThread("C:/Soundboard/Cheryl/WRAPUP/have a great day.mp3")
                CurrentQ = 31
                Timer2.Enabled = True
                Return False

            Case Else

                Return False
        End Select
    End Function


    Public Function CheckWhoseTalking2() As Boolean
        Select Case True
            Case Part.Contains("this is"), Part.Contains("speaking"), Part.Contains("you've got him"), Part.Contains("you've got her"), Part.Contains("yes"), Part.Contains("yeah"), Part.Contains("what's up?"), Part.Contains("how can i help you"), Part.Contains("hey"), Part.Contains("what do you want"), Part.Contains("hello"), Part.Contains("hi")
                Return True
            Case Part.Contains("not home"), Part.Contains("he isn't"), Part.Contains("not available"), Part.Contains("he's not"), Part.Contains(" a message"), Part.Contains("he's working"), Part.Contains("not here"), Part.Contains("not right now")
                rolltheclipThread("C:/Soundboard/Cheryl/WRAPUP/have a great day.mp3")
                CurrentQ = 31

                Timer2.Enabled = True
                Return False
            Case Else

                Return False
        End Select
    End Function
    Public Function GetBirthday()
        Select Case True
            Case s.Contains("true"), s.Contains("correct"), s.Contains("yes"), s.Contains("you got it"), s.Contains("that's right")
                Return True
            Case Else
                Return False
        End Select
    End Function
    Dim NumberOfVehicles As Integer = 0


    Function checkForNumVehicles() As Boolean
        Select Case True
            Case s.Contains("one"), s.Contains("1"), s.Contains("won"), s = "1", s = "1.", s.Contains("just one"), s.Contains("want")
                NumberOfVehicles = 1
                Return True
            Case s.Contains("two"), s.Contains("2"), s.Contains("too"), s = "2", s = "2."
                NumberOfVehicles = 2
                Return True
            Case s.Contains("three"), s.Contains("3"), s = "3", s = "3."
                NumberOfVehicles = 3
                Return True
            Case s.Contains("four"), s.Contains("4"), s = "4", s = "4."
                NumberOfVehicles = 4
                Return True
            Case s.Contains("five"), s.Contains("six"), s.Contains("seven"), s.Contains("eight"), s.Contains("nine")
                rolltheclipThread("C:\SoundBoard\Cheryl\VEHICLE INFO\only4spots.mp3")
                NumberOfVehicles = 4
                Return True
            Case Else
                NumberOfVehicles = 4
                Return False
        End Select

    End Function
    Public Function handleTCPA() As Boolean
        Dim response As String = s
        Select Case True
            Case response.Contains("yes"), response.Contains("okay"), response.Contains("ok"), response.Contains("sounds good"), response.Contains("affirmative"), response.Contains("alright")
                cmbDispo.Text = "Auto Lead"
                CurrentQ = 31
                rolltheclipThread("C:/Soundboard/Cheryl/WRAPUP/ENDCALL.mp3")
                'Timer2.Enabled = True
                Return True
            Case Else
                rolltheclipThread("C:/Soundboard/Cheryl/WRAPUP/have a great day.mp3")
                cmbDispo.Text = "Lost On Wrap Up"
                CurrentQ = 31
                rolltheclipThread("C:/Soundboard/Cheryl/WRAPUP/ENDCALL.mp3")
                'Timer2.Enabled = True
                Return False
        End Select
    End Function
    Public Function getLastName() As Boolean
        Dim response As String = s
        Try
            local_browser.FindElementById("frmLastName").SendKeys(response)
            Return True
        Catch ex As Exception
            moo()
            Console.WriteLine("PANDA EXCEPTION!")
            Return False
        End Try

        Return False
    End Function
    Public Function getPhoneType() As Boolean
        Dim response As String = s
        selectElement = New SelectElement(local_browser.FindElementById("frmPhoneType1"))
        Select Case True
            Case response.Contains("cell"), response.Contains("mobile")
                moo()
                selectElement.SelectByText("Mobile/Cell")
                Return True
            Case response.Contains("home")
                moo()
                selectElement.SelectByText("Home")
                Return True
            Case response.Contains("work")
                selectElement.SelectByText("Work")
                Return True
        End Select

        Return False
    End Function
    Public Function getCredit() As Boolean
        selectElement = New SelectElement(local_browser.FindElementById("frmCreditRating"))
        Dim response As String = s
        Select Case True
            Case response.Contains("excellent")
                selectElement.SelectByText("Excellent")
                Return True
            Case response.Contains("good")
                moo()
                selectElement.SelectByText("Good")
                Return True
            Case response.Contains("fair")
                selectElement.SelectByText("Fair")
                Return True
        End Select

        Return False
    End Function
    Public Function getEmail() As Boolean
        Dim response As String = s
        response = response.Replace(" ", "")
        Dim atIndex As Integer = response.LastIndexOf(" at ")
        Dim domain As String = response.Substring(atIndex)

        Dim endlength As Integer = response.Length - domain.Length

        Dim nombre As String = response.Substring(0, endlength)
        domain = domain.Replace("at", "@")

        Dim derEmail As String = nombre + domain

        Try
            local_browser.FindElementById("frmEmailAddress").SendKeys(derEmail)
            Return True
        Catch ex As Exception
            Return False
        End Try
        Return False
    End Function

    Public Function finalizeAddress() As Boolean
        NewAddress += " " & s
        Console.WriteLine(NewAddress)
        local_browser.FindElementById("frmAddress").SendKeys(NewAddress)
        Return True
    End Function
    Public Function getAddressNum() As Boolean
        NewAddress = ""
        Dim x As Integer = 0
        Try
            Do Until s.Substring(x, 1) = " " Or x = s.Length
                Select Case True
                    Case s.Substring(x, 1) = "1", s.Substring(x, 1) = "2", s.Substring(x, 1) = "3", s.Substring(x, 1) = "4", s.Substring(x, 1) = "5", s.Substring(x, 1) = "6", s.Substring(x, 1) = "7", s.Substring(x, 1) = "8", s.Substring(x, 1) = "9", s.Substring(x, 1) = "0"
                        NewAddress += s.Substring(x, 1)
                End Select
                x = x + 1
            Loop
        Catch ex As Exception
            Console.WriteLine("Problem with address")
        End Try
        If NewAddress <> "" Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Function getResType() As Boolean
        Select Case True
            Case s.Contains("single family"), s.Contains("house")
                sResidenceType = "Single Family"

            Case s.Contains("apartment")
                sResidenceType = "Apartment"

            Case s.Contains("duplex")
                sResidenceType = "Duplex"

            Case s.Contains("condominum")
                sResidenceType = "Condominium"

            Case s.Contains("town home")
                sResidenceType = "Townhome"

            Case s.Contains("mobile home")
                sResidenceType = "Mobile Home"

        End Select
        If residenceType <> "" Then
            selectElement = New SelectElement(local_browser.FindElementById("frmDwellingType"))
            selectElement.SelectByText(sResidenceType)
            Return True
        Else

            Return False
        End If
    End Function
    Public Function getHomeType() As Boolean
        Select Case True
            Case s.Contains("own")
                residenceType = "Own"

            Case s.Contains("rent")
                residenceType = "Rent"

            Case s.Contains("other")
                residenceType = "Other"

        End Select
        If residenceType <> "" Then
            selectElement = New SelectElement(local_browser.FindElementById("frmResidenceType"))
            selectElement.SelectByText(residenceType)
            Return True
        Else

            Return False
        End If
    End Function
    Dim custBday As Boolean
    Public Function getSpouseBDAY(isspouse As Boolean) As Boolean
        Dim x As Integer = 0
        Dim str As String = ""
        For x = 0 To s.Length - 1
            Select Case True
                Case s.Substring(x, 1) = "1", s.Substring(x, 1) = "2", s.Substring(x, 1) = "3", s.Substring(x, 1) = "4", s.Substring(x, 1) = "5", s.Substring(x, 1) = "6", s.Substring(x, 1) = "7", s.Substring(x, 1) = "8", s.Substring(x, 1) = "9", s.Substring(x, 1) = "0"
                    str += s.Substring(x, 1)
                    Console.WriteLine(str)
            End Select
        Next
        Console.WriteLine(str)
        If Not isspouse Then
            custBday = True
        End If
        CurrentQ = 14
        spouseBDAY = str
        If spouseBDAY.Length >= 4 Then
            Return True
        Else
            Return False
        End If

    End Function

    Dim spousebdaymonth As String


    Public Function finalizeSpouseBDay(isspouse As Boolean) As Boolean

        Select Case True
            Case s.Contains("january")
                BMonth = "JAN"
                spouseBDAY = spouseBDAY.Substring(1)
            Case s.Contains("february")
                BMonth = "FEB"
                spouseBDAY = spouseBDAY.Substring(1)
            Case s.Contains("march")
                BMonth = "MAR"
                spouseBDAY = spouseBDAY.Substring(1)
            Case s.Contains("april")
                BMonth = "APR"
                spouseBDAY = spouseBDAY.Substring(1)
            Case s.Contains("may")
                BMonth = "MAY"
                spouseBDAY = spouseBDAY.Substring(1)
            Case s.Contains("june")
                BMonth = "JUN"
                spouseBDAY = spouseBDAY.Substring(1)
            Case s.Contains("july")
                BMonth = "JUL"
                spouseBDAY = spouseBDAY.Substring(1)
            Case s.Contains("august")
                BMonth = "AUG"
                spouseBDAY = spouseBDAY.Substring(1)
            Case s.Contains("september")
                BMonth = "SEP"
                spouseBDAY = spouseBDAY.Substring(1)
            Case s.Contains("october")
                BMonth = "OCT"
                spouseBDAY = spouseBDAY.Substring(2)
            Case s.Contains("november")
                BMonth = "NOV"
                spouseBDAY = spouseBDAY.Substring(2)
            Case s.Contains("december")
                BMonth = "DEC"
                spouseBDAY = spouseBDAY.Substring(2)
        End Select
        Console.WriteLine("BMONTH IS: " & BMonth)
        If spouseBDAY.Length > 4 Then
            Console.WriteLine(spouseBDAY)
            BDay = spouseBDAY.Substring(0, spouseBDAY.Length - 4)
            Console.WriteLine(spouseBDAY)
            BYear = spouseBDAY.Substring(spouseBDAY.Length - 4)
            Console.WriteLine(spouseBDAY)
            Console.WriteLine(BYear)
        Else
            Console.WriteLine(BDay)
            BDay = spouseBDAY.Substring(0, spouseBDAY.Length - 2)
            Console.WriteLine(BDay)
            BYear = "19" & spouseBDAY.Substring(spouseBDAY.Length - 2)
            Console.WriteLine(BYear)
        End If
        If isspouse Then
            If BDay <> "" And BMonth <> "" And BYear <> "" Then
                local_browser.FindElementById("frmSpouseDOB_Month").SendKeys(BMonth)
                local_browser.FindElementById("frmSpouseDOB_Day").SendKeys(BDay)
                local_browser.FindElementById("frmSpouseDOB_Year").SendKeys(BYear)
                BDay = ""
                BMonth = ""
                BYear = ""
                Return True
            Else
                Return False
            End If
        Else

            If BDay <> "" And BMonth <> "" And BYear <> "" Then
                local_browser.FindElementById("frmDOB_Month").SendKeys(BMonth)
                local_browser.FindElementById("frmDOB_Day").SendKeys(BDay)
                local_browser.FindElementById("frmDOB_Year").SendKeys(BYear)
                BDay = ""
                BMonth = ""
                BYear = ""
                CurrentQ = 11
                Return True

            Else
                Return False
            End If
        End If

    End Function

    Sub LoadSpouseNames()
        Try
            Dim x As Integer = 1

            Dim LastName As String = ""
            Dim tempName As String = ""
            For Each foundFile As String In My.Computer.FileSystem.GetFiles("C:/SoundBoard/Cheryl/Names/")
                tempName = LCase(foundFile.Replace(".mp3", ""))
                tempName = tempName.Replace("c:\soundboard\cheryl\names\", "")
                tempName = tempName.Replace(" 1", "")
                tempName = tempName.Replace(" 2", "")
                tempName = tempName.Replace(" 3", "")
                If Not Names.Contains(tempName) Then
                    Names.Add(tempName)
                    Console.WriteLine(x & ": " & tempName)

                    x += 1
                End If
            Next
        Catch ex As Exception

        End Try


    End Sub

    Public Function checkForSpouseName()
        Dim item As String
        For Each item In Names
            If s.Contains(item) Then
                Console.WriteLine(item)
                theSpouseName = item
                Exit For
            End If
        Next

        If theSpouseName <> "" Then
            local_browser.FindElementById("frmSpouseFirstName").SendKeys(theSpouseName)
            local_browser.FindElementById("frmSpouseLastName").SendKeys(local_browser.FindElementById("frmLastName").Text)
            Return True
        Else


            Return False
        End If
    End Function

    Public Function checkMaritalStatus() As Boolean
        Select Case True
            Case s.Contains("married")
                maritalStatus = "Married"
            Case s.Contains("single")
                maritalStatus = "Single"
            Case s.Contains("divorced")
                maritalStatus = "Divorced"
            Case s.Contains("separated")
                maritalStatus = "Separated"
            Case s.Contains("widowed")
                maritalStatus = "Widowed"
            Case s.Contains("domestic partner")
                maritalStatus = "Domestic Partner"
            Case Else

        End Select
        If maritalStatus <> "" Then
            selectElement = New SelectElement(local_browser.FindElementById("frmMaritalStatus"))
            selectElement.SelectByText(maritalStatus)
        Else
            Return False
        End If

        Return False
    End Function
    Public Sub updateSpeechText()
        txtSpeech.Text += "::SPEECH ENDED::" & vbNewLine
    End Sub 'so speech text can be done crossthreaded

    Dim UnsureAboutCompany As Integer = 0
    Public Function CheckForCompany() As Boolean

        If s.Contains("don't know") Or s.Contains("not sure") Or s.Contains("not certain") Then
            Select Case UnsureAboutCompany
                Case 0
                    rolltheclipThread("C:\SoundBoard\Cheryl\PUSHONS\allstategeicostatefarm.mp3")
                    UnsureAboutCompany += 1
                Case 1
                    IProvider = "Progressive"

                    UnsureAboutCompany = 0
            End Select

        End If

        Select Case True
            Case s.Contains("none"), s.Contains("no insurance"), s.Contains("don't have"), s.Contains("don't have insurance"), s.Contains("not insured")
                IProvider = "No current insurance"
            Case s.Contains("twenty first"), s.Contains("21st"), s.Contains("twenty first century"), s.Contains("21st century"), s.Contains("twenty first century insurance"), s.Contains("21st century insurance"), s.Contains("first century")
                IProvider = "21st Century Insurance"
            Case s.Contains("AAA"), s.Contains("triple A"), s.Contains("triple a"), s.Contains("aaa")
                IProvider = "AAA Insurance Co."
            Case s.Contains("aarp"), s.Contains("AARP")
                IProvider = "AARP"
            Case s.Contains("etna"), s.Contains("edna"), s.Contains("aetna"), s.Contains("AETNA"), s.Contains("Edna")
                IProvider = "AETNA"
            Case s.Contains("aflac"), s.Contains("affleck"), s.Contains("afleck"), s.Contains("AFLAC")
                moo()
                IProvider = "AFLAC"
            Case s.Contains("aig"), s.Contains("AIG")
                IProvider = "AIG"
            Case s.Contains("AIU"), s.Contains("eye you")
                IProvider = "AIU"
            Case s.Contains("allied"), s.Contains("ally")
                IProvider = "Allied"
            Case s.Contains("allstate"), s.Contains("all state"), s.Contains("ball state"), s.Contains("mall state")
                IProvider = "Allstate Insurance"
            Case s.Contains("american")
                IProvider = "American Insurance"
            Case s.Contains("ameriplan")
                IProvider = "AmeriPlan"
            Case s.Contains("amica"), s.Contains("amiga"), s.Contains("amigo"), s.Contains("omika"), s.Contains("amika")
                IProvider = "Amica Insurance"
            Case s.Contains("answer"), s.Contains("answer financial")
                IProvider = "Answer Financial"
            Case s.Contains("arbella"), s.Contains("bella")
                IProvider = "Arbella"
            Case s.Contains("associated")
                IProvider = "Associated Indemnity"
            Case s.Contains("atlanta casualty")
                IProvider = "Atlanta Casualty"
            Case s.Contains("atlantic"), s.Contains("atlantic indemnity")
                IProvider = "Atlantic Indemnity"
            Case s.Contains("auto club")
                IProvider = "Auto Club Insurance Company"
            Case s.Contains("auto owners")
                IProvider = "Auto-Owners Insurance"
            Case s.Contains("axa"), s.Contains("axe")
                moo()
                IProvider = "AXA Advisors"
            Case s.Contains("bankers")
                IProvider = "Bankers Life and Casualty"
            Case s.Contains("banner")
                IProvider = "Banner Life"
            Case s.Contains("best")
                IProvider = "Best Agency USA"
            Case s.Contains("blue cross"), s.Contains("blue cross blue shield"), s.Contains("bcbs")
                IProvider = "Blue Cross and Blue Shield"
            Case s.Contains("brooke")
                IProvider = "Brooke Insurance"
            Case s.Contains("cal farm"), s.Contains("call form"), s.Contains("cal form"), s.Contains("call farm")
                IProvider = "Cal Farm Insurance"
            Case s.Contains("california state"), s.Contains("cal state")
                IProvider = "California State Automobile Association"
            Case s.Contains("chub"), s.Contains("chubb")
                IProvider = "Chubb"
            Case s.Contains("citizen"), s.Contains("citizens")
                IProvider = "Citizens"
            Case s.Contains("clarendon")
                IProvider = "Clarendon American Insurance"
            Case s.Contains("cna"), s.Contains("see na")
                IProvider = "CNA"
            Case s.Contains("colonial")
                IProvider = "Colonial Insurance"
            Case s.Contains("comparison")
                IProvider = "Comparison Market"
            Case s.Contains("continental")
                IProvider = "Continental Insurance"
            Case s.Contains("cotton"), s.Contains("cotton states")
                IProvider = "Cotton States Insurance"
            Case s.Contains("country insurance")
                IProvider = "Country Insurance and Financial Services"
            Case s.Contains("countrywide")
                IProvider = "Countrywide Insurance"
            Case s.Contains("cse"), s.Contains("easy")
                IProvider = "CSE Insurance Group"
            Case s.Contains("dairy")
                IProvider = "Dairyland Insurance"
            Case s.Contains("e health"), s.Contains("ehealth")
                IProvider = "eHealthInsurance Services"
            Case s.Contains("electric")
                IProvider = "Electric Insurance"
            Case s.Contains("erie"), s.Contains("eerie"), s.Contains("hear ye")
                IProvider = "Erie Insurance Company"
            Case s.Contains("esurance")
                IProvider = "Esurance"
            Case s.Contains("farm bureau")
                IProvider = "Farm Bureau/Farm Family/Rural"
            Case s.Contains("farmers")
                moo() ' okay, this one's actually funny
                IProvider = "Farmers Insurance"
            Case s.Contains("finance box")
                IProvider = "FinanceBox.com"
            Case s.Contains("fire and casualty")
                IProvider = "Fire and Casualty Insurance Co of CT"
            Case s.Contains("fireman"), s.Contains("firemens"), s.Contains("firemens fund")
                IProvider = "Fireman's Fund"
            Case s.Contains("foremost")
                IProvider = "Foremost"
            Case s.Contains("forester")
                IProvider = "Foresters"
            Case s.Contains("frank"), s.Contains("frankenstein"), s.Contains("frankenmuth")
                IProvider = "Frankenmuth Insurance"
            Case s.Contains("geico"), s.Contains("gecko"), s.Contains("i go")
                IProvider = "Geico General Insurance"
            Case s.Contains("gmac"), s.Contains("gee em")
                IProvider = "GMAC Insurance"
            Case s.Contains("golden rule")
                IProvider = "Golden Rule Insurance"
            Case s.Contains("government")
                IProvider = "Government Employees Insurance"
            Case s.Contains("i am a panda hear me moo")
                IProvider = "Progressive"
                Console.WriteLine("I AM A PANDA, HEAR ME MOO")
                Console.WriteLine("moooooooooo")
            Case s.Contains("guaranty"), s.Contains("guarantee")
                IProvider = "Guaranty National Insurance"
            Case s.Contains("guide"), s.Contains("guide one")
                IProvider = "Guide One Insurance"
            Case s.Contains("hanover"), s.Contains("lloyd")
                IProvider = "Hanover Lloyd's Insurance Company"
            Case s.Contains("hartford")
                IProvider = "Hartfod Insurance Co of the Southeast"
            Case s.Contains("hastings"), s.Contains("hasting mutual")
                IProvider = "Hastings Mutual Insurance Company"
            Case s.Contains("health benefits")
                IProvider = "Health Benefits Direct"
            Case s.Contains("health plus")
                IProvider = "Health Plus of America"
            Case s.Contains("health share")
                IProvider = "HealthShare American"
            Case s.Contains("human"), s.Contains("humana")
                IProvider = "Humana"
            Case s.Contains("ifa"), s.Contains("eye f a")
                IProvider = "IFA Auto Insurance"
            Case s.Contains("igf")
                IProvider = "IGF Insurance"
            Case s.Contains("infinity"), s.Contains("infinite")
                IProvider = "Infinity Insurance"
            Case s.Contains("insurance insight")
                IProvider = "Insurance Insight"
            Case s.Contains("insurance dot com"), s.Contains("insurance.com")
                IProvider = "Insurance.com"
            Case s.Contains("insurance leads dot com"), s.Contains("insuranceleads.com")
                IProvider = "InsuranceLeaders.com"
            Case s.Contains("insweb"), s.Contains("in web")
                IProvider = "Insweb"
            Case s.Contains("integon"), s.Contains("pentagon")
                IProvider = "Integon"
            Case s.Contains("hancock"), s.Contains("john hancock")  ' i should really make it match 'John Footpenis', but like... we'd get sued
                IProvider = "John Hancock"
            Case s.Contains("kaiser"), s.Contains("kayser"), s.Contains("permanent")
                IProvider = "Kaiser Permanente"
            Case s.Contains("kemper"), s.Contains("camper"), s.Contains("lloyd")
                IProvider = "Kemper Lloyds Insurance"
            Case s.Contains("landmark")
                IProvider = "Landmark American Insurance"
            Case s.Contains("leader")
                IProvider = "Leader National Insurance"
            Case s.Contains("liberty"), s.Contains("liberty mutual")
                IProvider = "Liberty Mutual Insurance"
            Case s.Contains("lumber")
                IProvider = "Lumbermens Mutual"
            Case s.Contains("maryland")
                IProvider = "Maryland Casualty"
            Case s.Contains("mass mutual")
                IProvider = "Mass Mutual"
            Case s.Contains("mega"), s.Contains("midwest"), s.Contains("mega midwest")
                IProvider = "Mega/Midwest"
            Case s.Contains("mercury")
                IProvider = "Mercury"
            Case s.Contains("met life"), s.Contains("metlife")
                IProvider = "MetLife Auto and Home"
            Case s.Contains("metropolitan")
                IProvider = "Metropolitan Insurance Co."
            Case s.Contains("mid century")
                IProvider = "Mid Century Insurance"
            Case s.Contains("mid continent")
                IProvider = "Mid-Continent Casualty"
            Case s.Contains("middlesex")
                IProvider = "Middlesex Insurance"
            Case s.Contains("midland national"), s.Contains("midland")
                IProvider = "Midland National Life"
            Case s.Contains("mutual of new york")
                IProvider = "Mutual of New York"
            Case s.Contains("mutual of omaha"), s.Contains("omaha")
                IProvider = "Mutual of Omaha"
            Case s.Contains("national ben franklin"), s.Contains("ben franklin")
                IProvider = "National Ben Franklin Insurance"
            Case s.Contains("national casualty")
                IProvider = "National Casualty"
            Case s.Contains("national continental")
                IProvider = "National Continental Insurance"
            Case s.Contains("national fire")
                IProvider = "National Fire Insurance Company of Hartford"
            Case s.Contains("national health")
                IProvider = "National Health Insurance"
            Case s.Contains("national indemnity")
                IProvider = "National Indemnity"
            Case s.Contains("national union fire of los angeles"), s.Contains("national union fire insurance of los angeles"), s.Contains("national union fire of la"), s.Contains("national union fire insurance of los angeles")
                IProvider = "National Union Fire Insurance of LA"
            Case s.Contains("national union fire of pennsylvania"), s.Contains("national union fire insurance of pennsylvania")
                IProvider = "National Union Fire Insurance of PA"
            Case s.Contains("nationwide")
                IProvider = "Nationwood Insurance Company"
            Case s.Contains("new england financial")
                IProvider = "New England Financial"
            Case s.Contains("new york life")
                IProvider = "New York Life Insurance"
            Case s.Contains("northwestern")
                IProvider = "Northwestern Mutual Life"
            Case s.Contains("omni")
                IProvider = "Omni Insruance"
            Case s.Contains("orion")
                IProvider = "Orion Insurance"
            Case s.Contains("pacific insurance")
                IProvider = "Pacific Insurance"
            Case s.Contains("pafco")
                IProvider = "Pafco General Insurance"
            Case s.Contains("patriot")
                IProvider = "Patriot General Insurance"
            Case s.Contains("peak property")
                IProvider = "Peak Property and Casualty Insurance"
            Case s.Contains("pemco")
                IProvider = "PEMCO Insurance"
            Case s.Contains("physicians")
                IProvider = "Physicians"
            Case s.Contains("pioneer")
                IProvider = "Pioneer State Mutual Insurance Company"
            Case s.Contains("preferred")
                IProvider = "Preferred Mutual"
            Case s.Contains("progressive"), s.Contains("progress"), s.Contains("aggressive")
                moo()
                IProvider = "Progressive"
            Case s.Contains("prudential")
                IProvider = "Prudential Insurance Co."
            Case s.Contains("reliance"), s.Contains("reliant")
                IProvider = "Reliance Insurance"
            Case s.Contains("response")
                IProvider = "Response Insurance"
            Case s.Contains("safeco"), s.Contains("safe co")
                IProvider = "SAFECO"
            Case s.Contains("safeway"), s.Contains("safe way")
                IProvider = "Safeway Insurance"
            Case s.Contains("security insurance")
                IProvider = "Security Insurance Co of Hartford"
            Case s.Contains("security national")
                IProvider = "Security National Insurance Co of FL"
            Case s.Contains("sentinel")
                IProvider = "Sentinel Insurance"
            Case s.Contains("sentry")
                moo()
                IProvider = "Sentry Insurance Group"
            Case s.Contains("shelter")
                IProvider = "Shelter Insurance Co."
            Case s.Contains("saint paul"), s.Contains("st. paul")
                IProvider = "St. Paul"
            Case s.Contains("standard fire")
                IProvider = "Standard Fire Insurance Company"
            Case s.Contains("state and county")
                IProvider = "State and County Mutual Fire Insurance"
            Case s.Contains("state farm"), s.Contains("statefarm"), s.Contains("haystack"), s.Contains("stay farm"), s.Contains("stayfarm")
                IProvider = "State Farm General"
            Case s.Contains("state fund")
                IProvider = "State Fund"
            Case s.Contains("state national")
                IProvider = "State National Insurance"
            Case s.Contains("superior")
                IProvider = "Superior Insurance"
            Case s.Contains("sure health"), s.Contains("sure")
                IProvider = "Sure Health Plans"
            Case s.Contains("abe group"), s.Contains("ah be group")
                IProvider = "The Ahbe Group"
            Case s.Contains("general"), s.Contains("the general")
                IProvider = "The General"
            Case s.Contains("tico"), s.Contains("tye co"), s.Contains("tie co")
                IProvider = "TICO Insruance"
            Case s.Contains("tig countrywide"), s.Contains("tig")
                IProvider = "TIG Countrywide Insurance"
            Case s.Contains("the hartford")
                IProvider = "The Hartford"
            Case s.Contains("titan")
                IProvider = "Titan"
            Case s.Contains("trans"), s.Contains("transamerica"), s.Contains("trans america")
                moo()
                IProvider = "TransAmerica"
            Case s.Contains("travelers")
                IProvider = "Travelers Insurance Company" ' pandas are cool
            Case s.Contains("tri-state"), s.Contains("tri state")
                IProvider = "Tri-State Consumer Insurance"
            Case s.Contains("twin city")
                IProvider = "Twin City Fire Insurance"
            Case s.Contains("unicare"), s.Contains("unicorn")
                IProvider = "UniCare"
            Case s.Contains("united american")
                IProvider = "United American/Farm and Ranch"
            Case s.Contains("united pacific")
                IProvider = "United Pacific Insurance"
            Case s.Contains("united security")
                IProvider = "United Security"
            Case s.Contains("united services")
                IProvider = "United Serviecs Automobile Association"
            Case s.Contains("unitrin")
                IProvider = "Unitrin Direct"
            Case s.Contains("universal")
                IProvider = "Universal Underwriters Insurance"
            Case s.Contains("US financial")
                IProvider = "US Financial"
            Case s.Contains("USA")
                IProvider = "USA Benefits/Continental General"
            Case s.Contains("USAA"), s.Contains("USA")
                IProvider = "USAA"
            Case s.Contains("usf and g")
                IProvider = "USF and G"
            Case s.Contains("viking")
                moo()
                IProvider = "Viking Insurance Co of WI"
            Case s.Contains("western and")
                IProvider = "Western and Southern Life"
            Case s.Contains("western")
                IProvider = "Western Mutual"
            Case s.Contains("windsor")
                IProvider = "Windsor Insurance"
            Case s.Contains("woodland")
                IProvider = "Woolands Financial Group"
            Case s.Contains("zurich")
                IProvider = "Zurich North America"
            Case Else
                If IProvider = "Farmers Insurance" And Not s.Contains("farmers") Then
                    Console.WriteLine("I am the walrus")
                End If
                If s.Length > 7 Then
                    IProvider = "Progressive"
                Else

                    isQuestion = True
                End If

        End Select
        Console.WriteLine("Detected " & IProvider & " as insurance")
        If IProvider <> "" Then
            Try
                'local_browser.Navigate.GoToUrl("https://forms.lead.co/auto/?agent_name=Justin+Theriault&lead_id=421&lead_guid=7af28e93-bfdf-43d0-8e81-742cbdf34ad2&import_id=13395")
                selectElement = New SelectElement(local_browser.FindElementById("frmInsuranceCarrier"))
                selectElement.SelectByText(IProvider)
            Catch ex As Exception
                Console.WriteLine(ex.Message)
                Console.WriteLine(ex.Source)
                Console.WriteLine(ex.InnerException)
                Console.WriteLine(ex.StackTrace)
            End Try

            Return True
        Else
            Return False
        End If
    End Function
    Public Sub RandomHumanism()
        isQuestion = False
        Console.WriteLine("HUMAN EXPRESSION: " & HumanCounter)
        Select Case HumanCounter
            Case 1
                rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\wonderful.mp3")
                HumanCounter += 1
                Timer2.Enabled = True
            Case 2
                rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\excellent 2.mp3")
                HumanCounter += 1
                Timer2.Enabled = True
            Case 3
                rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\great 2.mp3")
                HumanCounter += 1
                Timer2.Enabled = True
            Case 4
                rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\okGreat.mp3")
                HumanCounter += 1
                Timer2.Enabled = True
            Case 5
                HumanCounter += 1
                rolltheclipThread("C:\SoundBoard\Cheryl\Birthday\questions 5-4-16\questions 5-4-16\almost done.mp3")
                Timer2.Enabled = True
            Case 6
                HumanCounter += 1
                ' rolltheclipThread("C:\SoundBoard\Cheryl\reactions\doing an excellent job.wav")
                rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\OK.mp3")
                Timer2.Enabled = True
            Case 7
                rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\okGreat.mp3")
                HumanCounter += 1
                Timer2.Enabled = True
            Case 8
                HumanCounter += 1
                rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\OK2.MP3")
                Timer2.Enabled = True
            Case 9
                HumanCounter = 1
                rolltheclipThread("C:\SoundBoard\Cheryl\Birthday\questions 5-4-16\questions 5-4-16\almost out hair.mp3")
                Timer2.Enabled = True
        End Select

    End Sub
    Public Function checkExpiration() As Boolean
        Dim monthnow As String = Date.Now.Month
        Select Case True
            Case s.Contains("don't know"), s.Contains("not even sure"), s.Contains("not sure")
                If numRepeats = 0 Then
                    rolltheclipThread("C:\SoundBoard\Cheryl\REBUTTALS\JANUARY FEB MARCH APRIL.MP3")
                    isQuestion = True
                    numRepeats += 1
                Else
                    theMonth = Now.Month
                    theYear = Now.Year
                End If

            Case s.Contains("month to month") Or s.Contains("month by month")
                theYear = Now.Year
                theMonth = Now.Month + 1
            Case s.Contains("next year")
                theYear = Date.Now.Year + 1
                theMonth = Date.Now.Month
            Case s.Contains("just renewed"), s.Contains("been renewed")

                theMonth = Date.Now.Month + 6
                If CInt(theMonth) < Date.Now.Month Then
                    theYear = Date.Now.Year + 1
                Else
                    theYear = Date.Now.Year
                End If

                CurrentQ = 6
                RandomHumanism()

            Case s.Contains("january")

                theMonth = 1
                If secondPass = False Then
                    If monthnow < theMonth Then
                        theYear = Date.Now.Year
                    Else
                        theYear = Date.Now.Year + 1
                    End If
                Else

                End If

                CurrentQ = 5
            Case s.Contains("february")
                theMonth = 2
                If secondPass = False Then
                    If monthnow < theMonth Then
                        theYear = Date.Now.Year
                    Else
                        theYear = Date.Now.Year + 1
                    End If
                Else
                End If

                CurrentQ = 5

            Case s.Contains("march")
                theMonth = 3
                If secondPass = False Then
                    If monthnow < theMonth Then
                        theYear = Date.Now.Year
                    Else
                        theYear = Date.Now.Year + 1
                    End If
                Else
                End If

                CurrentQ = 5
            Case s.Contains("april")
                theMonth = 4
                If secondPass = False Then
                    If monthnow < theMonth Then
                        theYear = Date.Now.Year
                    Else
                        theYear = Date.Now.Year + 1
                    End If
                Else

                End If

                CurrentQ = 5
            Case s.Contains("may")
                theMonth = 5
                If secondPass = False Then
                    If monthnow < theMonth Then
                        theYear = Date.Now.Year
                    Else
                        theYear = Date.Now.Year + 1
                    End If
                Else
                End If

                CurrentQ = 5
            Case s.Contains("june")
                theMonth = 6
                If secondPass = False Then
                    If monthnow < theMonth Then
                        theYear = Date.Now.Year
                    Else
                        theYear = Date.Now.Year + 1
                    End If
                Else
                End If

                CurrentQ = 5
            Case s.Contains("july")
                theMonth = 7
                If secondPass = False Then
                    If monthnow < theMonth Then
                        theYear = Date.Now.Year
                    Else
                        theYear = Date.Now.Year + 1
                    End If
                Else
                End If

                CurrentQ = 5
            Case s.Contains("august")
                theMonth = 8
                If secondPass = False Then
                    If monthnow < theMonth Then
                        theYear = Date.Now.Year
                    Else
                        theYear = Date.Now.Year + 1
                    End If
                Else
                End If

                CurrentQ = 5
            Case s.Contains("september")
                theMonth = 9
                If secondPass = False Then
                    If monthnow < theMonth Then
                        theYear = Date.Now.Year
                    Else
                        theYear = Date.Now.Year + 1
                    End If
                Else
                End If

                CurrentQ = 5
            Case s.Contains("october")
                theMonth = 10
                If secondPass = False Then
                    If monthnow < theMonth Then
                        theYear = Date.Now.Year
                    Else
                        theYear = Date.Now.Year + 1
                    End If
                Else
                End If

                CurrentQ = 5
            Case s.Contains("november")
                theMonth = 11
                If secondPass = False Then
                    If monthnow < theMonth Then
                        theYear = Date.Now.Year
                    Else
                        theYear = Date.Now.Year + 1
                    End If
                Else
                End If

                CurrentQ = 5
                theMonth = 12
                If secondPass = False Then
                    If monthnow < theMonth Then
                        theYear = Date.Now.Year
                    Else
                        theYear = Date.Now.Year + 1
                    End If
                Else
                End If

                CurrentQ = 5
            Case s.Contains("month")
                Select Case True
                    Case s.Contains("next")
                        theMonth = Date.Now.Month + 1
                        If theMonth > 12 Then
                            theMonth = theMonth - 12
                        End If
                        theYear = Date.Now.Year

                        CurrentQ = 5
                    Case s.Contains("this")
                        theMonth = Date.Now.Month
                        theYear = Date.Now.Year

                        CurrentQ = 5

                    Case s.Contains("3")
                        theMonth = Date.Now.Month + 3
                        If theMonth > 12 Then
                            theMonth = theMonth - 12
                        End If
                        If theMonth < Date.Now.Month Then
                            theYear = Date.Now.Year + 1
                        Else
                            theYear = Date.Now.Year
                        End If

                        CurrentQ = 5
                    Case s.Contains("4")
                        theMonth = Date.Now.Month + 4
                        If theMonth > 12 Then
                            theMonth = theMonth - 12
                        End If
                        If theMonth < Date.Now.Month Then
                            theYear = Date.Now.Year + 1
                        Else
                            theYear = Date.Now.Year
                        End If

                        CurrentQ = 5
                    Case s.Contains("5")
                        theMonth = Date.Now.Month + 5
                        If theMonth > 12 Then
                            theMonth = theMonth - 12
                        End If
                        If theMonth < Date.Now.Month Then
                            theYear = Date.Now.Year + 1
                        Else
                            theYear = Date.Now.Year
                        End If

                        CurrentQ = 5
                    Case s.Contains("6")
                        theMonth = Date.Now.Month + 6
                        If theMonth > 12 Then
                            theMonth = theMonth - 12
                        End If
                        If theMonth < Date.Now.Month Then
                            theYear = Date.Now.Year + 1
                        Else
                            theYear = Date.Now.Year
                        End If

                        CurrentQ = 5
                    Case s.Contains("7")
                        theMonth = Date.Now.Month + 7
                        If theMonth > 12 Then
                            theMonth = theMonth - 12
                        End If
                        If theMonth < Date.Now.Month Then
                            theYear = Date.Now.Year + 1
                        Else
                            theYear = Date.Now.Year
                        End If

                        CurrentQ = 5
                    Case s.Contains("8")
                        theMonth = Date.Now.Month + 8
                        If theMonth > 12 Then
                            theMonth = theMonth - 12
                        End If
                        If theMonth < Date.Now.Month Then
                            theYear = Date.Now.Year + 1
                        Else
                            theYear = Date.Now.Year
                        End If

                        CurrentQ = 5
                    Case s.Contains("9")
                        theMonth = Date.Now.Month + 9
                        If theMonth > 12 Then
                            theMonth = theMonth - 12
                        End If
                        If theMonth < Date.Now.Month Then
                            theYear = Date.Now.Year + 1
                        Else
                            theYear = Date.Now.Year
                        End If

                        CurrentQ = 5
                    Case s.Contains("10")
                        theMonth = Date.Now.Month + 10
                        If theMonth > 12 Then
                            theMonth = theMonth - 12
                        End If
                        If theMonth < Date.Now.Month Then
                            theYear = Date.Now.Year + 1
                        Else
                            theYear = Date.Now.Year
                        End If
                End Select
            Case s.Contains("week")
                theMonth = Date.Now.Month
                theYear = Date.Now.Year

                CurrentQ = 5

            Case Else
                If numRepeats = 0 Then
                    'rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\CAN YOU JUST VERIFY THE MONTH.mp3")
                    numRepeats += 1
                    isQuestion = True
                    m.StartMicAndRecognition()
                Else
                    theMonth = Now.Month
                    theYear = Now.Year
                End If
        End Select
        If theYear <> "" Then
            numRepeats = 0
            local_browser.FindElementById("frmPolicyExpires_Month").SendKeys(NumtoMonth(theMonth))
            'For i As Integer = 0 To 1000
            '    i += 1
            'Next
            selectElement = New SelectElement(local_browser.FindElementById("frmPolicyExpires_Year"))
            selectElement.SelectByText(CStr(theYear))
            theYear = ""
            Return True
        Else
            Return False
        End If
    End Function
    Public Function CheckHowLong() As Boolean
        Select Case True
            Case s.Contains("don't know"), s.Contains("not even sure"), s.Contains("not sure")
                If numRepeats = 0 Then
                    rolltheclipThread("C:\SoundBoard\Cheryl\TIE INS\Great what's your best guess.mp3")
                    clipType = "Objection"
                    numRepeats += 1
                    Return False
                Else
                    theMonth = Now.Month
                    theYear = Now.Year
                    Return True
                End If
            Case s.Contains("1"), s.Contains(" a "), s.Contains("a year"), s.Contains("one year"), s.Contains("won year")
                theYear = CStr(Date.Now.Year - 1)
            Case s.Contains("2"), s.Contains("two years"), s.Contains("two")
                theYear = CStr(Date.Now.Year - 2)
            Case s.Contains("3"), s.Contains("three years"), s.Contains("three")
                theYear = CStr(Date.Now.Year - 3)
            Case s.Contains("4"), s.Contains("four years"), s.Contains("for years"), s.Contains("four") ' for years might break some stuff. Could ask for a possible year verification here.
                theYear = CStr(Date.Now.Year - 4)
            Case s.Contains("5"), s.Contains("five years"), s.Contains("five")
                theYear = CStr(Date.Now.Year - 5)
            Case s.Contains("6"), s.Contains("six years"), s.Contains("six")
                theYear = CStr(Date.Now.Year - 6)
            Case s.Contains("7"), s.Contains("seven years"), s.Contains("seven")
                theYear = CStr(Date.Now.Year - 7)
            Case s.Contains("8"), s.Contains("eight years"), s.Contains("eight")
                theYear = CStr(Date.Now.Year - 8)
            Case s.Contains("9"), s.Contains("nine years"), s.Contains("ten")
                theYear = CStr(Date.Now.Year - 9)
            Case s.Contains("10"), s.Contains("ten years"), s.Contains("ten")
                theYear = CStr(Date.Now.Year - 10)
            Case s.Contains("11"), s.Contains("eleven years"), s.Contains("eleven")
                theYear = CStr(Date.Now.Year - 11)
            Case s.Contains("12"), s.Contains("twelve years"), s.Contains("twelve")
                theYear = CStr(Date.Now.Year - 12)
            Case s.Contains("13"), s.Contains("thirteen years"), s.Contains("thirteen")
                theYear = CStr(Date.Now.Year - 13)
            Case s.Contains("14"), s.Contains("fourteen years"), s.Contains("fourteen")
                theYear = CStr(Date.Now.Year - 14)
            Case s.Contains("15"), s.Contains("fifteen years"), s.Contains("fifteen")
                theYear = CStr(Date.Now.Year - 15)
            Case s.Contains("16"), s.Contains("sixteen years"), s.Contains("sixteen")
                theYear = CStr(Date.Now.Year - 16)
            Case s.Contains("17"), s.Contains("seventeen years"), s.Contains("seventeen")
                theYear = CStr(Date.Now.Year - 17)
            Case s.Contains("18"), s.Contains("eighteen years"), s.Contains("eighteen")
                theYear = CStr(Date.Now.Year - 18)
            Case s.Contains("19"), s.Contains("nineteen years"), s.Contains("nineteen")
                theYear = CStr(Date.Now.Year - 19)
            Case Else
                theYear = CStr(Date.Now.Year - 1)
        End Select
        If theMonth <> "" And theYear <> "" Then
            local_browser.FindElementById("frmPolicyStart_Month").SendKeys(NumtoMonth(theMonth))
            selectElement = New SelectElement(local_browser.FindElementById("frmPolicyStart_Year"))
            selectElement.SelectByText(CStr(theYear))
            Return True
        Else
            Return False
        End If

    End Function

    Dim newobjection As Boolean = False
    Dim Part As String = ""

    Sub handlepartquestion()
        Console.WriteLine("CHECKING AGAINST PARTIAL QUESTIONS")
        Console.WriteLine("reps: " & quest)

        Try
            Select Case True
                Case Part.Contains("who is this"), Part.Contains("who are you"), Part.Contains("who is calling"), Part.Contains("who's this"), Part.Contains("who's calling"), Part.Contains("who do you represent")
                    Select Case quest
                        Case 1
                            Playlist(0) = "c:\soundboard\cheryl\REBUTTALS\THAT'S A GREAT QUESTION.mp3"
                            quest += 1
                        Case 2
                            Playlist(0) = "C:\SoundBoard\Cheryl\Birthday\questions 5-4-16\questions 5-4-16\whatta great question.mp3"
                            quest = 1
                    End Select
                    If CurrentQ = 3 Then
                        CurrentQ = 0
                    End If
                    Playlist(1) = "c:\soundboard\cheryl\INTRO\CHERYLCALLING.mp3"
                    Timer2.Enabled = True
                Case Part.Contains("who makes it")
                    Select Case quest
                        Case 1
                            Playlist(0) = "c:\soundboard\cheryl\REBUTTALS\THAT'S A GREAT QUESTION.mp3"
                            quest += 1
                        Case 2
                            Playlist(0) = "C:\SoundBoard\Cheryl\Birthday\questions 5-4-16\questions 5-4-16\whatta great question.mp3"
                            quest = 1
                    End Select
                    rolltheclipThread("c:\soundboard\cheryl\REACTIONS\YES.mp3")
                Case Part.Contains("what is this"), Part.Contains("what's this"), Part.Contains("what is the nature of this call"), Part.Contains("what are you calling about"), Part.Contains("what is purpose of this call")
                    Select Case quest
                        Case 1
                            Playlist(0) = "c:\soundboard\cheryl\REBUTTALS\THAT'S A GREAT QUESTION.mp3"
                            quest += 1
                        Case 2
                            Playlist(0) = "C:\SoundBoard\Cheryl\Birthday\questions 5-4-16\questions 5-4-16\whatta great question.mp3"
                            quest = 1
                    End Select

                    If CurrentQ = 3 Then
                        CurrentQ = 0
                    End If
                    Playlist(1) = "c:\soundboard\cheryl\INTRO\THISISTOGIVENEWQUOTE.mp3"
                    Timer2.Enabled = True
                Case Part.Contains("what is lcn"), Part.Contains("what is elsieanne"), Part.Contains("about your company"), s.Contains("lcn")
                    Select Case quest
                        Case 1
                            Playlist(0) = "c:\soundboard\cheryl\REBUTTALS\THAT'S A GREAT QUESTION.mp3"
                            quest += 1
                        Case 2
                            Playlist(0) = "C:\SoundBoard\Cheryl\Birthday\questions 5-4-16\questions 5-4-16\whatta great question.mp3"
                            quest = 1
                    End Select
                    If CurrentQ = 3 Then
                        CurrentQ = 0
                    End If
                    Playlist(1) = "c:\soundboard\cheryl\Rebuttals\What's LCN.mp3"
                    Timer2.Enabled = True
                Case Part.Contains("why are you calling")
                    Select Case quest
                        Case 1
                            Playlist(0) = "c:\soundboard\cheryl\REBUTTALS\THAT'S A GREAT QUESTION.mp3"
                            quest += 1
                        Case 2
                            Playlist(0) = "C:\SoundBoard\Cheryl\Birthday\questions 5-4-16\questions 5-4-16\whatta great question.mp3"
                            quest = 1
                    End Select
                    Playlist(1) = "c:\soundboard\cheryl\INTRO\THISISTOGIVENEWQUOTE.mp3"

                    Timer2.Enabled = True

                Case Part.Contains("how did you get my info"), Part.Contains("where did you get my info")
                    Select Case quest
                        Case 1
                            Playlist(0) = "c:\soundboard\cheryl\REBUTTALS\THAT'S A GREAT QUESTION.mp3"
                            quest += 1
                        Case 2
                            Playlist(0) = "C:\SoundBoard\Cheryl\Birthday\questions 5-4-16\questions 5-4-16\whatta great question.mp3"
                            quest = 1
                    End Select
                    Playlist(0) = "c:\soundboard\cheryl\REBUTTALS\Where Did You get My info.mp3"

                    Timer2.Enabled = True
            End Select
        Catch ex As Exception
            Console.WriteLine("problem with part question")
        End Try
    End Sub 'Checks for questions in the partial speech variable (part) handles them if found

    Public Function doaddressstuff() As Boolean
        ParseAddress(s)
        If getAddressNum() Then
            Return True
        Else
            Return False
        End If

        Return False
    End Function

    Public Function ParseAddress(speech As String) As Boolean
        NewAddress = ""
        Dim x As Integer = 0
        Do Until speech.Substring(x, 1) = " " Or x = speech.Length
            NewAddress += speech.Substring(x, 1)
            x = x + 1
        Loop
        NewAddress += " " & StreetSpelling
        zip = speech.Substring(speech.Length - 5, 5)

        If NewAddress <> "" Then
            If getAddressNum() Then
                Return True
            Else
                Return False
            End If
        End If

        Return False
    End Function
    Public Sub StopThatClip()
        BeginInvoke(New Action(AddressOf waveOut.Dispose))
        BeginInvoke(New Action(AddressOf waveOut2.Dispose))
        newobjection = True

    End Sub 'Stops clip and listens

    <DllImport("User32")> Private Shared Function ShowWindow(ByVal hwnd As Integer, ByVal nCmdShow As Integer) As Integer

    End Function
    Public Sub Register()
        'Register REACTIONSystem.Windows.Forms.Keys.
        RegisterHotKey(Me.Handle, 101, System.Windows.Forms.Keys.None, System.Windows.Forms.Keys.F1)
        RegisterHotKey(Me.Handle, 201, MOD_CONTROL, System.Windows.Forms.Keys.F1)
        RegisterHotKey(Me.Handle, 301, MOD_ALT, System.Windows.Forms.Keys.F1)
        RegisterHotKey(Me.Handle, 102, System.Windows.Forms.Keys.None, System.Windows.Forms.Keys.F2)
        RegisterHotKey(Me.Handle, 103, System.Windows.Forms.Keys.None, System.Windows.Forms.Keys.F3)
        RegisterHotKey(Me.Handle, 104, System.Windows.Forms.Keys.None, System.Windows.Forms.Keys.F4)

        'Register TIE INSystem.Windows.Forms.Keys.
        RegisterHotKey(Me.Handle, 105, System.Windows.Forms.Keys.None, System.Windows.Forms.Keys.F5)
        RegisterHotKey(Me.Handle, 106, System.Windows.Forms.Keys.None, System.Windows.Forms.Keys.F6)
        RegisterHotKey(Me.Handle, 206, MOD_CONTROL, System.Windows.Forms.Keys.F6)
        RegisterHotKey(Me.Handle, 306, MOD_ALT, System.Windows.Forms.Keys.F6)
        RegisterHotKey(Me.Handle, 107, System.Windows.Forms.Keys.None, System.Windows.Forms.Keys.F7)
        RegisterHotKey(Me.Handle, 207, MOD_CONTROL, System.Windows.Forms.Keys.F7)
        RegisterHotKey(Me.Handle, 307, MOD_ALT, System.Windows.Forms.Keys.F7)
        RegisterHotKey(Me.Handle, 108, System.Windows.Forms.Keys.None, System.Windows.Forms.Keys.F8)
        RegisterHotKey(Me.Handle, 208, MOD_ALT, System.Windows.Forms.Keys.F8)
        RegisterHotKey(Me.Handle, 308, MOD_CONTROL, System.Windows.Forms.Keys.F8)

        'Register REBUTTALSystem.Windows.Forms.Keys.
        RegisterHotKey(Me.Handle, 109, System.Windows.Forms.Keys.None, System.Windows.Forms.Keys.F9)
        RegisterHotKey(Me.Handle, 110, System.Windows.Forms.Keys.None, System.Windows.Forms.Keys.F10)
        RegisterHotKey(Me.Handle, 111, System.Windows.Forms.Keys.None, System.Windows.Forms.Keys.F11)
        RegisterHotKey(Me.Handle, 112, System.Windows.Forms.Keys.None, System.Windows.Forms.Keys.F12)

        RegisterHotKey(Me.Handle, 169, System.Windows.Forms.Keys.None, System.Windows.Forms.Keys.Escape)

        RegisterHotKey(Me.Handle, 170, System.Windows.Forms.Keys.None, System.Windows.Forms.Keys.Right)
        RegisterHotKey(Me.Handle, 171, MOD_CONTROL, System.Windows.Forms.Keys.Right)


        RegisterHotKey(Me.Handle, 172, System.Windows.Forms.Keys.None, System.Windows.Forms.Keys.Left)

        'Register END CALLSystem.Windows.Forms.Keys.

        'Register Entry Key
        RegisterHotKey(Me.Handle, 173, MOD_CONTROL, System.Windows.Forms.Keys.E)


    End Sub      'Registers hotkeys
    Public Sub AgeFromProg()
        Try

            Dim dates(2) As String
            'LeadForm.Document.GetElementById("frmDOB_Month").GetAttribute("value")
            'LeadForm.Document.GetElementById("frmDOB_Day").GetAttribute("value")
            'LeadForm.Document.GetElementById("frmDOB_Year").GetAttribute("value")

            Dim theMonth As Integer = dates(0)

            Dim theDay As Integer = dates(1)
            Dim theYear As Integer = dates(2)
            Dim theAge As Integer = Today.Year - theYear
            If theMonth > Today.Month Then
                theAge -= 1
            End If
            Console.WriteLine("CUSTOMER Is " & theAge & " YEARS OLD")
            If theAge <= 80 And theAge >= 25 Then
                LifeQual = True
                LifeCheck.Visible = True
            Else
                LifeQual = False
                LifeCheck.Visible = False

            End If
            If theAge >= 64 Then
                HealthCheck.Visible = False
                HealthQual = False
                Mediqual = True
                MedicareCheck.Visible = True
            Else
                HealthQual = True
                Mediqual = False
                MedicareCheck.Visible = False
                HealthCheck.Visible = True

            End If
        Catch
            Console.WriteLine("BIRTHDAY Not ENTERED")
        End Try

    End Sub
    Public Sub getCurrentAge()
        Try
            Dim age As Integer
            age = Today.Year - BYear
            If BMonth > Today.Month Then
                age -= 1
            End If
            If age <= 80 And age >= 25 Then
                LifeQual = True
            Else
                LifeQual = False
            End If
            If age >= 64 Then
                HealthQual = False
                Mediqual = True
            Else
                HealthQual = True
            End If
        Catch
        End Try
    End Sub
    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        If m.Msg = WM_HOTKEY Then
            Dim id As IntPtr = m.WParam
            Select Case (id.ToString)
                Case "101"
                    If clipnum(0) = 0 Then
                        rolltheclipThread("C:  \Soundboard\Cheryl\REACTIONS\OK.mp3")
                        clipnum(0) += 1
                    ElseIf clipnum(0) = 1 Then
                        rolltheclipThread("C:\Soundboard\Cheryl\REACTIONS\OK2.mp3")
                        clipnum(0) += 1
                    Else
                        rolltheclipThread("C:\Soundboard\Cheryl\REACTIONS\okgreat.mp3")
                        clipnum(0) = 0
                    End If


                Case "201"

                Case "301"
                    rolltheclipThread("C:\Soundboard\Cheryl\REACTIONS\OKGreat.mp3")
                Case "102"
                    If clipnum(2) = 0 Then
                        rolltheclipThread("C:\Soundboard\Cheryl\REACTIONS\EXCELLENT 2.mp3")
                        clipnum(2) += 1
                    ElseIf clipnum(2) = 1 Then
                        rolltheclipThread("c:\soundboard\cheryl\REACTIONS\Wonderful.mp3")
                        clipnum(2) += 1
                    Else
                        rolltheclipThread("C:\Soundboard\Cheryl\REACTIONS\GREAT 2.mp3")
                        clipnum(2) = 0
                    End If
                Case "103"
                    rolltheclipThread("C:\Soundboard\Cheryl\REACTIONS\verygood.mp3")
                Case "104"

                Case "105"

                Case "106"
                    'or
                Case "206"
                    'or2
                Case "306"
                    'or3
                Case "107"
                    rolltheclipThread("C:\Soundboard\Cheryl\TIE INS\SPELLTHAT.mp3")
                Case "207"
                    rolltheclipThread("C:\Soundboard\Cheryl\TIE INS\Could You Please Spell That Out.mp3")
                Case "307"
                    rolltheclipThread("C:\Soundboard\Cheryl\TIE INS\Could You Please Spell That Out 2.mp3")
                Case "108"
                    rolltheclipThread("C:\Soundboard\Cheryl\TIE INS\Take Your Best Guess.mp3")
                Case "208"
                    rolltheclipThread("c:\soundboard\cheryl\TIE INS\Okay What's Your Best Guess.mp3")
                Case "308"
                    rolltheclipThread("c:\soundboard\cheryl\TIE INS\Great What's Your Best Guess.mp3")
                'REBUTTALS
                Case "109"
                    rolltheclipThread("c:\soundboard\cheryl\Rebuttals\Rebuttal1.mp3")
                Case "110"
                    rolltheclipThread("c:\soundboard\cheryl\Rebuttals\Rebuttal2.mp3")
                Case "112"
                    rolltheclipThread("c:\soundboard\cheryl\Rebuttals\Rebuttal3.mp3")
                Case "111"
                    rolltheclipThread("c:\soundboard\cheryl\Rebuttals\Rebuttal4.mp3")

                Case "169"
                    StopThatClip()
                Case "170"                   'RIGHT ARROW

                    CurrentQ += 1
                    AskQuestion(CurrentQ, counter)

                Case "171" 'right arrow key SKIP
                    CurrentQ = CurrentQ + 1
                    Reset()

                    If CurrentQ > 28 Then
                        CurrentQ = 1
                    End If
                    lblQuestion.Text = CURRENTQUESTION(CurrentQ)

                Case "172"

                    CurrentQ = CurrentQ - 1
                    Reset()
                    If CurrentQ < 1 Then
                        CurrentQ = 24
                    End If

                Case "173"

            End Select
        End If
        Try
            MyBase.WndProc(m)
        Catch ex As Exception
            Console.WriteLine("****** EXCEPTION PANDA ********")
            Console.WriteLine(ex.Message)
            Console.WriteLine(ex.Data)
            Console.WriteLine(ex.InnerException)
            Console.WriteLine(ex.StackTrace)
            Console.WriteLine("****** END EXCEPTION PANDA ********")
        End Try
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs)
        rolltheclipThread("c: \soundboard\cheryl\INTRO\CHERY_CALLING_FROM_LCN.mp3")
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles HOMETYPE.Click
        isQuestion = True

        rolltheclipThread("c:\soundboard\cheryl\PERSONAL INFO\HOMETYPE.mp3")
        callPos = Home_Type

    End Sub
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        StopThatClip()
        rolltheclipThread("c:\soundboard\cheryl\PERSONAL INFO\DO YOU OWN Or RENT THE HOME.mp3")
        CurrentQ = 15
        isQuestion = True
        clipType = "Question"
        callPos = Own_Rent


    End Sub
    Private Sub Form1_Click(sender As Object, e As EventArgs) Handles MyBase.Click
        'LoadVehicles()
        'Console.WriteLine(Search("2005 Toyota Camry"))
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs)
        rolltheclipThread("C: /Soundboard/Cheryl/WhoDoYouUSe.mp3")
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs)
        rolltheclipThread("C:/Soundboard/Cheryl/Howlong.mp3")
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs)
        rolltheclipThread("C:/Soundboard/Cheryl/Expiration.mp3")
    End Sub
    Public Function GetActiveWindows() As ObjectModel.Collection(Of IntPtr)
        EnumWindows(AddressOf Enumerator, 0)
        Return ActiveWindows
    End Function
    Private Function Enumerator(ByVal hwnd As IntPtr, ByVal lParam As Integer) As Boolean
        Dim customerName As String = ""

        Dim text As String = Space(Int16.MaxValue)
        Dim i As Integer = 0

        If IsWindowVisible(hwnd) Then
            GetWindowText(hwnd, text, Int16.MaxValue)

            If text.Contains("Auto Insurance") Then
                theWindowTitle = text
                Do While text.Substring(i, 1) <> "|"
                    customerName = customerName & text.Substring(i, 1)
                    i = i + 1
                    If i >= text.Length Then
                        Exit Do
                    End If

                Loop
                Dim name() As String = customerName.Split


            End If
        End If

        Return True

    End Function
    Private Shared Function FindWindowEx(ByVal parentHandle As IntPtr,
                                     ByVal childAfter As IntPtr,
                                     ByVal lclassName As String,
                                     ByVal windowTitle As String) As IntPtr
    End Function
    Private Sub HelloButton_Click(sender As Object, e As EventArgs) Handles btnHello.Click
        rolltheclipThread("c:\soundboard\cheryl\INTRO\HELLO.mp3")
        isQuestion = True
    End Sub
    Private Sub Label1_Click_1(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Primary.SelectedIndexChanged
        deviceNum1 = Primary.SelectedIndex

    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs)


    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs)
        rolltheclipThread("C:/Soundboard/Cheryl/year of the vehicle.mp3")
    End Sub
    Private Sub Button16_Click(sender As Object, e As EventArgs)
        Primary.Items.Clear()

        Dim DeviceCount As Integer = NAudio.Wave.WaveOut.DeviceCount()
        Dim SDevice As String = Nothing
        For i As Integer = 0 To DeviceCount - 1
            SDevice = NAudio.Wave.WaveOut.GetCapabilities(i).ProductName
            Primary.Items.Add(SDevice)

        Next
    End Sub
    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles SpouseDOB.Click
        StopThatClip()
        rolltheclipThread("c:\soundboard\cheryl\DRIVER INFO\SPOUSES DATE OF BIRTH.mp3")
        isQuestion = True
        clipType = "Question"
        callPos = Spouse_DOB
    End Sub
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        isQuestion = True
        If getBirthdaWAV() = True Then
            tbCallOrder.SelectedTab = tbDriverInfo
            clipType = "Question"
            callPos = Driver_Birthday
            'LeadForm.Document.GetElementById("frmDOB_Month").Focus()
            CurrentQ = 10
            Timer2.Enabled = True
        Else
            clipType = "Question"
            callPos = Driver_Birthday
            rolltheclipThread("c:\soundboard\cheryl\DRIVER INFO\DOB1.mp3")
        End If

    End Sub
    Dim OnCall As Boolean = False
    Private Sub Button26_Click(sender As Object, e As EventArgs)
        rolltheclipThread("C:/Soundboard/Cheryl/WRAPUP/additionalquotes.mp3")
    End Sub
    Dim clipname As String
    Public Sub rolltheclipThread(fileName As String)
        s = ""
        Part = ""
        clipname = fileName
        If Me.demoThread IsNot Nothing Then
            Me.demoThread.Abort()
        End If
        Me.demoThread = New Thread(New ThreadStart(AddressOf Me.rolltheclip))
        Me.demoThread.Start()


    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles btnTheirName.Click

        Try
            rolltheclipThread(globalFile2)
            isQuestion = True
        Catch

        End Try
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs)
        rolltheclipThread("C:/Soundboard/Cheryl/Car Make Backup.mp3")
    End Sub
    Private Sub Button18_Click(sender As Object, e As EventArgs)
        rolltheclipThread("C:/Soundboard/Cheryl/Model.mp3")
    End Sub
    Private Sub Button19_Click(sender As Object, e As EventArgs)
        rolltheclipThread("C:/Soundboard/Cheryl/OtherCar.mp3")
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        StopThatClip()
        rolltheclipThread("c:\soundboard\cheryl\DRIVER INFO\MaritalStatus2.mp3")
        clipType = "Question"
        callPos = Marital_Status
        isQuestion = True

    End Sub
    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        isQuestion = True

        rolltheclipThread("C:/Soundboard/Cheryl/PERSONAL INFO/phoneType.mp3")
        callPos = Phone_Type
        clipType = "Question"


    End Sub
    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        isQuestion = True
        rolltheclipThread("C:/Soundboard/Cheryl/PERSONAL INFO/Last Name.mp3")
        clipType = "Question"
        callPos = Last_Name

    End Sub
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles SpouseName.Click
        StopThatClip()
        rolltheclipThread("c:\soundboard\cheryl\DRIVER INFO\SPOUSES FIRST NAME.mp3")
        isQuestion = True
        clipType = "Question"
        callPos = Spouse_Name
    End Sub
    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        isQuestion = True

        If clipnum(0) = 0 Then
            rolltheclipThread("C:\Soundboard\Cheryl\REACTIONS\OK.mp3")
            clipnum(0) += 1
        ElseIf clipnum(0) = 1 Then
            rolltheclipThread("C:\Soundboard\Cheryl\REACTIONS\OK2.mp3")
            clipnum(0) += 1
        Else
            rolltheclipThread("C:\Soundboard\Cheryl\REACTIONS\OKgreat.mp3")
            clipnum(0) = 0
        End If


    End Sub
    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        rolltheclipThread("C:\Soundboard\Cheryl\REACTIONS\verygood.mp3")
    End Sub
    Private Sub Button27_Click(sender As Object, e As EventArgs)
        rolltheclipThread("C:/Soundboard/Cheryl/OKEXCELLENT.mp3")
    End Sub
    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click
        rolltheclipThread("C:/Soundboard/Cheryl/SORRY.mp3")
    End Sub
    Private Sub Button32_Click(sender As Object, e As EventArgs)
        rolltheclipThread("C:/Soundboard/Cheryl/Rebuttal1.mp3")
    End Sub
    Private Sub Button31_Click(sender As Object, e As EventArgs)
        rolltheclipThread("C:/Soundboard/Cheryl/Rebuttal2.mp3")
    End Sub
    Private Sub Button30_Click(sender As Object, e As EventArgs)
        rolltheclipThread("C:/Soundboard/Cheryl/REbuttal3.mp3")
    End Sub
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        isQuestion = True
        rolltheclipThread("c:\soundboard\cheryl\REACTIONS\Could you please verify your address.mp3")
        callPos = Their_Address
        clipType = "Question"
    End Sub
    Private Sub Button14_Click(sender As Object, e As EventArgs)
        rolltheclipThread("C:/Soundboard/Cheryl/PERSONAL INFO/email.mp3")
    End Sub
    Private Sub Button22_Click(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\WRAPUP\TCPA Verbatim.mp3")
    End Sub
    Private Sub Button37_Click(sender As Object, e As EventArgs)
        NICount = 0
        cmbMoreVehicles.SelectedIndex = 0
        rolltheclipThread("C:/Soundboard/Cheryl/WRAPUP/ENDCALL.mp3")
        totalCalls = totalCalls + 1
        totalLeads = totalLeads + 1
        lblCalls.Text = totalCalls
        CurrentQ = 1
        lblQuestion.Text = "HELLO"
        txtInsuranceProvider.Clear()
        txtPolicyExpiration.Clear()
        txtPolicyStart.Clear()
    End Sub
    Public Sub resetBot()
        Dim i As Integer
        For i = 0 To 50
            clipnum(i) = 0
            txtInsuranceProvider.Clear()
            txtPolicyExpiration.Clear()
            txtPolicyStart.Clear()
            cmbSecondaries.SelectedIndex = -1
            txtDOB.Clear()
            cmbGender.SelectedIndex = -1
            cmbMaritalStatus.SelectedIndex = -1
            txtSPOUSENAME.Clear()
            txtSPOUSEDOB.Clear()
            cmbSpouseGender.SelectedIndex = -1
            cmbOwnRent.SelectedIndex = -1
            cmbHomeType.SelectedIndex = -1
            txtAddress.Clear()
            txtEmail.Clear()
            cmbCredit.SelectedIndex = -1
            cmbPhoneType.SelectedIndex = -1
            txtName.Clear()
            cmbSecondaries.SelectedIndex = -1
            txtYearBuilt.Clear()
            txtSqFt.Clear()
            cmbTCPA.SelectedIndex = -1
        Next
    End Sub
    Private Sub Button34_Click(sender As Object, e As EventArgs) Handles Button34.Click


        cmbMoreVehicles.SelectedIndex = 0
        theurl = ""
        NICount = 0
        rolltheclipThread("C:/Soundboard/Cheryl/WRAPUP/have a great day.mp3")
        cmbDispo.Text = "Not Available"
        totalCalls = totalCalls + 1
        lblCalls.Text = totalCalls
        lblQuestion.Text = "HELLO"
        txtInsuranceProvider.Clear()
        txtPolicyExpiration.Clear()
        CurrentQ = 31
        Timer2.Enabled = True

        resetBot()

    End Sub
    Private Sub Button33_Click(sender As Object, e As EventArgs) Handles Button33.Click
        isQuestion = True
        clipType = "Question"
        callPos = Credit
        rolltheclipThread("C:/Soundboard/Cheryl/PERSONAL INFO/Credit.mp3")
    End Sub
    Private Sub Button35_Click(sender As Object, e As EventArgs) Handles btnIntro.Click
        CurrentQ = 3
        tmrSilence.Enabled = True
        rolltheclipThread("c:\soundboard\cheryl\INTRO\INTRO2.MP3")
        clipType = "Question"
        callPos = Insurance_Provider
        m.StartMicAndRecognition()
    End Sub

    Private Sub Button38_Click(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\TIE INS\Okay What's Your Best Guess.mp3")
    End Sub
    Private Sub Button36_Click(sender As Object, e As EventArgs)
        rolltheclipThread("C:/Soundboard/Cheryl/Opener 2.MP3")
    End Sub
    Private Sub Button39_Click(sender As Object, e As EventArgs)
        rolltheclipThread("C:/Soundboard/Cheryl/YES.mp3")
    End Sub
    Private Sub Button41_Click(sender As Object, e As EventArgs) Handles Button41.Click
        If clipnum(2) = 0 Then
            rolltheclipThread("C:\Soundboard\Cheryl\REACTIONS\EXCELLENT 2.mp3")
            clipnum(2) += 1
        ElseIf clipnum(2) = 1 Then
            rolltheclipThread("c:\soundboard\cheryl\REACTIONS\Wonderful.mp3")
            clipnum(2) += 1
        Else
            rolltheclipThread("C:\Soundboard\Cheryl\REACTIONS\GREAT 2.mp3")
            clipnum(2) = 0
        End If

    End Sub
    Private Sub Button29_Click(sender As Object, e As EventArgs)
        rolltheclipThread("C:/Soundboard/Cheryl/REBUTTAL4.mp3")
    End Sub
    Private Sub Button42_Click(sender As Object, e As EventArgs)
        rolltheclipThread("C:/Soundboard/Cheryl/Rebuttal5.mp3")
    End Sub
    Private Sub Button27_Click_1(sender As Object, e As EventArgs) Handles btnRepeatThat.Click
        If clipnum(5) = 0 Then
            rolltheclipThread("C:/Soundboard/Cheryl/reactions/Can You Repeat that.mp3")
            clipnum(5) += 1
        ElseIf clipnum(5) = 1 Then
            rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\repeatagain.mp3")
            clipnum(5) += 1
        Else
            rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\brokenears.mp3")
            clipnum(5) = 0
        End If
    End Sub
    Private Sub Button53_Click(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\Rebuttals\Rebuttal3.mp3")
    End Sub
    Private Sub Button63_Click(sender As Object, e As EventArgs) Handles Button63.Click
        Select Case VehicleNum
            Case 1
                Select Case NumberOfVehicles
                    Case 1
                        rolltheclipThread("C:\SoundBoard\Cheryl\VEHICLE INFO\YMMYV.mp3")
                    Case Else
                        rolltheclipThread("C:\SoundBoard\Cheryl\VEHICLE INFO\First Vehicle.mp3")
                End Select
            Case 2
                rolltheclipThread("C:\SoundBoard\Cheryl\VEHICLE INFO\2nd Vehicle.mp3")
            Case 3
                rolltheclipThread("C:\SoundBoard\Cheryl\VEHICLE INFO\Third Vehicle.mp3")
            Case 4
                rolltheclipThread("C:\SoundBoard\Cheryl\VEHICLE INFO\Fourth Vehicle.mp3")
        End Select
        clipType = "Question"
        callPos = Year_Make_Model

    End Sub
    Private Sub Button64_Click(sender As Object, e As EventArgs) Handles Button64.Click
        rolltheclipThread("C:/SOUNDBOARD/CHERYL/VEHICLE INFO/HOW MANY VEHICLES DO YOU HAVE.MP3")
        isQuestion = True
        clipType = "Question"
        callPos = Number_Of_Vehicles
    End Sub
    Private Sub Button50_Click(sender As Object, e As EventArgs) Handles btnWhoDoYouHave.Click
        StopThatClip()
        rolltheclipThread("c:\soundboard\cheryl\INSURANCE INFO\Who Is The Current Auto INsurance Company that you're with.mp3")
        CurrentQ = 3
        isQuestion = True
        clipType = "Question"
        callPos = Insurance_Provider
    End Sub
    Private Sub Button51_Click(sender As Object, e As EventArgs) Handles btnPolicyStart.Click
        StopThatClip()
        rolltheclipThread("c:\soundboard\cheryl\INSURANCE INFO\HOW MANY YEARS HAVE YOU BEEN WITH THEM 2.mp3")
        CurrentQ = 5
        isQuestion = True
        clipType = "Question"
        callPos = Policy_Start
    End Sub
    Private Sub Button49_Click(sender As Object, e As EventArgs) Handles btnExpiration.Click
        StopThatClip()
        rolltheclipThread("c:\soundboard\cheryl\INSURANCE INFO\EXPIRATION.mp3")
        CurrentQ = 4
        clipType = "Question"
        callPos = Policy_Expiration

    End Sub
    Private Sub Button62_Click(sender As Object, e As EventArgs)
        Select Case cmbMoreVehicles.SelectedIndex
            Case 0
                rolltheclipThread("c:\soundboard\cheryl\VEHICLE INFO\WHAT KIND OF VEHICLE IS THAT 1.mp3")
            Case 1
                rolltheclipThread("c:\soundboard\cheryl\VEHICLE INFO\MAKE OF THE SECOND VEHICLE.mp3")
            Case 2
                rolltheclipThread("c:\soundboard\cheryl\VEHICLE INFO\MAKE OF THE THIRD VEHICLE.mp3")
            Case 3
                rolltheclipThread("c:\soundboard\cheryl\VEHICLE INFO\MAKE OF THE FOURTH VEHICLE.mp3")

        End Select
    End Sub
    Private Sub Button25_Click(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\WRAPUP\YEARBUILT.mp3")
    End Sub
    Private Sub Button15_Click_1(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\WRAPUP\SQUARE FOOTAGE.mp3")
    End Sub
    Private Sub Button44_Click(sender As Object, e As EventArgs)
        rolltheclipThread("C:\Soundboard\Cheryl\TIE INS\and.mp3")
    End Sub
    Private Sub Button36_Click_1(sender As Object, e As EventArgs) Handles Button36.Click
        rolltheclipThread("C:\Soundboard\Cheryl\TIE INS\Could You Please Spell That Out.mp3")
    End Sub
    Private Sub Button55_Click(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\Rebuttals\Rebuttal1.mp3")
    End Sub
    Private Sub Button52_Click(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\Rebuttals\Rebuttal2.mp3")
    End Sub
    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        rolltheclipThread("c:\soundboard\cheryl\Rebuttals\SORRY.mp3")
    End Sub
    Private Sub Button3_Click_1(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\WRAPUP\TCPAWARMUP.mp3")
    End Sub
    Private Sub Button4_Click_1(sender As Object, e As EventArgs)
        Playlist(0) = "c:\soundboard\cheryl\REBUTTALS\THIS WILL BE REAL QUICK.mp3"

    End Sub
    Private Sub Button29_Click_1(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\ADDRESS REBUTTAL.mp3")
    End Sub
    Private Sub Button30_Click_1(sender As Object, e As EventArgs) Handles YEARBUILT.Click
        rolltheclipThread("C:\SoundBoard\Cheryl\WRAPUP\YearBuilt.mp3")
    End Sub
    Private Sub Button56_Click(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\REACTIONS\YES.mp3")
    End Sub
    Private Sub Button61_Click(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\REACTIONS\NO.mp3")
    End Sub
    Private Sub Button42_Click_1(sender As Object, e As EventArgs) Handles Button42.Click

        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\Where Did You get My info.mp3")

    End Sub

    Private Sub Button66_Click(sender As Object, e As EventArgs)
        Select Case cmbMoreVehicles.SelectedIndex
            Case 0
                rolltheclipThread("c:\soundboard\cheryl\VEHICLE INFO\WHAT IS THE MODEL OF THE CAR 1.mp3")
            Case 1
                rolltheclipThread("c:\soundboard\cheryl\VEHICLE INFO\model OF THE SECOND VEHICLE.mp3")
            Case 2
                rolltheclipThread("c:\soundboard\cheryl\VEHICLE INFO\model OF THE THIRD VEHICLE.mp3")
            Case 3
                rolltheclipThread("c:\soundboard\cheryl\VEHICLE INFO\model OF THE FOURTH VEHICLE.mp3")

        End Select
    End Sub
    Private Sub Button67_Click(sender As Object, e As EventArgs) Handles Button67.Click

        rolltheclipThread("c:\soundboard\cheryl\Rebuttals\What's LCN.mp3")

    End Sub
    Private Sub Button43_Click(sender As Object, e As EventArgs) Handles Button43.Click
        cmbMoreVehicles.SelectedIndex = 0
        rolltheclipThread("c:\soundboard\cheryl\Rebuttals\DNC.mp3")
        cmbDispo.Text = "Do Not Call"
        CurrentQ = 31
        Timer2.Enabled = True
        tbCallOrder.SelectedTab = tbIntro
        lblQuestion.Text = CURRENTQUESTION(0)
        resetBot()
    End Sub
    Private Sub Button68_Click(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\INTRO\AUTOOpener 2.MP3")
    End Sub

    Private Sub tcpa_Click(sender As Object, e As EventArgs) Handles tcpa.Click
        rolltheclipThread("c:\soundboard\cheryl\WRAPUP\TCPA.mp3")
    End Sub
    Private Sub Button32_Click_2(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\CAN YOU JUST VERIFY THE MONTH.mp3")
    End Sub
    Private Sub Button39_Click_1(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\CAN YOU JUST VERIFY THE YEAR.mp3")
    End Sub
    Private Sub Button44_Click_1(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\CAN YOU JUST VERIFY THE DAY.mp3")
    End Sub
    Private Sub Button29_Click_2(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\WHERE DID YOU GET MY INFO.mp3")
    End Sub
    Private Sub Button5_Click_1(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\my spouse takes care of that.mp3")
    End Sub
    Private Sub Button19_Click_1(sender As Object, e As EventArgs)
        Playlist(0) = "C:\SoundBoard\Cheryl\REACTIONS\BEST NI REBUTTALS ZIP\BEST NI REBUTTALS\i understand.mp3"

    End Sub
    Private Sub Button57_Click(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\I'm REQUIRED TO HAVE YOU VERIFY IT FIRST.mp3")
    End Sub
    Private Sub Button60_Click(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\why do you need my info.mp3")
    End Sub
    Private Sub Button84_Click(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\january feb march april.mp3")
    End Sub
    Private Sub Button85_Click(sender As Object, e As EventArgs) Handles Button85.Click

        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\I'm REQUIRED TO HAVE YOU VERIFY IT FIRST.mp3")

    End Sub
    Private Sub Button58_Click(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\I'M JUST ABOUT DONE.mp3")
    End Sub
    Private Sub Button77_Click(sender As Object, e As EventArgs)
        Playlist(0) = "c:\soundboard\cheryl\REBUTTALS\I'M JUST ABOUT DONE.mp3"

    End Sub
    Private Sub Button25_Click_1(sender As Object, e As EventArgs)
        Playlist(0) = "C:\SoundBoard\Cheryl\REACTIONS\BEST NI REBUTTALS ZIP\BEST NI REBUTTALS\nothing to be interested in.mp3"

    End Sub
    Private Sub Button68_Click_1(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\IS THIS THE SPOUSE.mp3")
    End Sub
    Private Sub Button26_Click_1(sender As Object, e As EventArgs)
        Playlist(0) = "c:\soundboard\cheryl\REBUTTALS\I already have insurance rebuttal.mp3"

    End Sub
    Private Sub Button87_Click(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\This will be real quick.mp3")
    End Sub
    Private Sub Button88_Click(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\what's LCN.mp3")
    End Sub
    Private Sub Button89_Click(sender As Object, e As EventArgs) Handles Button89.Click
        rolltheclipThread("c:\soundboard\cheryl\INTRO\CHERYLCALLING.mp3")
    End Sub
    Private Sub Button90_Click(sender As Object, e As EventArgs) Handles Button90.Click
        rolltheclipThread("c:\soundboard\cheryl\INTRO\HELLO.mp3")
    End Sub
    Private Sub Button91_Click(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\INTRO\THISISTOGIVENEWQUOTE.mp3")
    End Sub
    Private Sub Button92_Click(sender As Object, e As EventArgs) Handles Button92.Click

        rolltheclipThread("c:\soundboard\cheryl\INTRO\THISISTOGIVENEWQUOTE.mp3")

    End Sub
    Private Sub Button46_Click(sender As Object, e As EventArgs) Handles Button46.Click
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\THAT'S A GREAT QUESTION.mp3")
    End Sub
    Private Sub Button93_Click(sender As Object, e As EventArgs) Handles Button93.Click
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\I COMPLETELY UNDERSTAND.mp3")
    End Sub
    Private Sub Button53_Click_1(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\where did you get my info.mp3")
    End Sub
    Private Sub Button45_Click(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\my spouse takes care of that.mp3")
    End Sub
    Private Sub Button3_Click_2(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\my spouse takes care of that.mp3")
    End Sub
    Private Sub Button81_Click(sender As Object, e As EventArgs) Handles Button81.Click
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\local agents and carriers in your area.mp3")
    End Sub
    Private Sub Button82_Click(sender As Object, e As EventArgs) Handles Button82.Click
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\can they email.mp3")

    End Sub
    Private Sub Button51_Click_1(sender As Object, e As EventArgs) Handles Button51.Click
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\you're not giving me a quote.mp3")
    End Sub
    Private Sub Button83_Click(sender As Object, e As EventArgs) Handles Button83.Click
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\when will they call.mp3")


    End Sub
    Private Sub Button18_Click_1(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\my spouse takes care of that.mp3")
    End Sub

    Private Sub Button86_Click(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\I'M JUST ABOUT DONE.mp3")
    End Sub

    Private Sub Button74_Click(sender As Object, e As EventArgs)
        Playlist(0) = "c:\soundboard\cheryl\Rebuttals\Rebuttal4.mp3"
    End Sub

    Private Sub Button72_Click(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\Rebuttals\Rebuttal4.mp3")

    End Sub

    Private Sub Button76_Click(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\Rebuttals\What's LCN.mp3")

    End Sub
    Private Sub Button78_Click(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\where did you get my info.mp3")

    End Sub
    Private Sub InsuranceProvider_TextChanged(sender As Object, e As EventArgs) Handles txtInsuranceProvider.TextChanged
        If String.IsNullOrEmpty(txtInsuranceProvider.Text) = False Then
            btnExpiration.Visible = True
            txtPolicyExpiration.Visible = True
        Else
            btnExpiration.Visible = False
            txtPolicyExpiration.Visible = False
        End If
    End Sub
    Private Sub CARYEAR_TextChanged(sender As Object, e As EventArgs)
    End Sub

    Private Sub CARMAKE_TextChanged(sender As Object, e As EventArgs)

    End Sub
    Private Sub cmbSecondaries_TextChanged_1(sender As Object, e As EventArgs)
        If Strings.LCase(cmbSecondaries.Text) = "y" Or Strings.LCase(cmbSecondaries.Text) = "yes" Then
            SECONDARIES = True
        Else
            SECONDARIES = False
        End If
    End Sub
    Private Sub cmbOwnRent_TextChanged(sender As Object, e As EventArgs)

    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles txtPolicyExpiration.TextChanged
        If String.IsNullOrEmpty(btnExpiration.Text) = False Then
            btnPolicyStart.Visible = True
            txtPolicyStart.Visible = True
        Else
            txtPolicyStart.Visible = False
            btnPolicyStart.Visible = False
        End If
    End Sub
    Private Sub txtSPOUSENAME_TextChanged(sender As Object, e As EventArgs) Handles txtSPOUSENAME.TextChanged
    End Sub
    Private Sub txtSPOUSEDOB_TextChanged(sender As Object, e As EventArgs) Handles txtSPOUSEDOB.TextChanged
    End Sub
    Private Sub YEARBUILTTEXT_TextChanged(sender As Object, e As EventArgs) Handles txtYearBuilt.TextChanged
        If String.IsNullOrEmpty(txtYearBuilt.Text) = False Then
            SQFT.Visible = True
            txtSqFt.Visible = True
        Else
            SQFT.Visible = False
            txtSqFt.Visible = False
        End If
    End Sub

    Private Sub Button3_Click_3(sender As Object, e As EventArgs) Handles Button3.Click
        rolltheclipThread("c:\soundboard\cheryl\REACTIONS\YES.mp3")
    End Sub

    Private Sub Button7_Click_1(sender As Object, e As EventArgs) Handles Button7.Click
        rolltheclipThread("c:\soundboard\cheryl\REACTIONS\NO.mp3")
    End Sub

    Private Sub Button69_Click(sender As Object, e As EventArgs)
        Playlist(0) = "c:\soundboard\cheryl\REBUTTALS\I ACTUALLY HAVE THIS INFORMATION.mp3"

    End Sub

    Private Sub Button17_Click_1(sender As Object, e As EventArgs) Handles Button17.Click
        Try
            rolltheclipThread(globalFile)
        Catch

        End Try
    End Sub

    Private Sub Button15_Click_2(sender As Object, e As EventArgs) Handles Button15.Click
        Try
            rolltheclipThread(globalfile3)
        Catch
        End Try

    End Sub
    Dim skip As Boolean = False
    Private Sub results_TextChanged(sender As Object, e As EventArgs) Handles results.TextChanged

    End Sub




    Private Sub Button18_Click_3(sender As Object, e As EventArgs)
        rolltheclipThread("C:/SoundBoard/Cheryl/Names/Vanessa Name.mp3")
    End Sub
    Public Sub AskQuestion(ByRef Pos As Integer, ByRef numReps As Integer)

        m.EndMicAndRecognition()
        Console.WriteLine("ASKING QUESTION: " & CurrentQ)
        Console.WriteLine("version:" & numReps)
        clipType = "Question"
        Try
            s = ""
            Part = ""

            newobjection = False
            lblQuestion.Text = CURRENTQUESTION(Pos)
            Select Case Pos
                Case 0
                    m.StartMicAndRecognition()
                    Select Case numReps
                        Case 0
                            rolltheclipThread("c:\soundboard\cheryl\INSURANCE INFO\Ins provider 1.mp3")
                            CurrentQ = 3
                        Case 1
                            rolltheclipThread("c:\soundboard\cheryl\INSURANCE INFO\Ins provider 2.mp3")
                            CurrentQ = 3
                        Case Else
                            rolltheclipThread("c:\soundboard\cheryl\INSURANCE INFO\Ins provider 3.mp3")
                            CurrentQ = 3
                            numReps = 0

                    End Select
                    CurrentQ = 3

                Case 1
                    m.StartMicAndRecognition()
                    rolltheclipThread("c:\soundboard\cheryl\INTRO\HELLO.mp3")
                Case 2
                    m.StartMicAndRecognition()
                    Select Case numReps
                        Case 0
                            rolltheclipThread(globalFile)

                        Case 1
                            rolltheclipThread(globalfile3)
                        Case Else
                            rolltheclipThread(globalFile2)
                    End Select
                Case 3
                    rolltheclipThread("C:\SoundBoard\Cheryl\INTRO\INTRO2.MP3")
                    CurrentQ = 3
                    m.StartMicAndRecognition()
                Case 4
                    Select Case numReps
                        Case 0
                            rolltheclipThread("c:\soundboard\cheryl\INSURANCE INFO\EXPIRATION.mp3")
                        Case 1
                            rolltheclipThread("c:\soundboard\cheryl\INSURANCE INFO\policy exp 1.mp3")
                        Case 2
                            rolltheclipThread("c:\soundboard\cheryl\INSURANCE INFO\policy exp 2.mp3")
                        Case Else
                            rolltheclipThread("c:\soundboard\cheryl\INSURANCE INFO\policy exp 3.mp3")
                    End Select
                Case 5
                    m.StartMicAndRecognition()
                    Console.WriteLine("*** PANDA PANDA PANDA NUMREPS = {0}", numReps)
                    Select Case numReps
                        Case 0
                            rolltheclipThread("c:\soundboard\cheryl\INSURANCE INFO\And How many years have you been with them.mp3")
                        Case 1
                            rolltheclipThread("c:\soundboard\cheryl\INSURANCE INFO\years with 1.mp3")
                        Case 2
                            rolltheclipThread("c:\soundboard\cheryl\INSURANCE INFO\years with 2.mp3")
                        Case Else
                            rolltheclipThread("c:\soundboard\cheryl\INSURANCE INFO\years with 3.mp3")
                    End Select
                Case 6
                    rolltheclipThread("C:/SOUNDBOARD/CHERYL/VEHICLE INFO/HOW MANY VEHICLES DO YOU HAVE.MP3")
                    m.StartMicAndRecognition()
                Case 7
                    m.StartMicAndRecognition()
                    Console.WriteLine("on vehicle: " & VehicleNum)
                    Select Case VehicleNum
                        Case 1
                            Select Case NumberOfVehicles
                                Case 1
                                    rolltheclipThread("C:\SoundBoard\Cheryl\VEHICLE INFO\YMMYV.mp3")
                                Case Else
                                    rolltheclipThread("C:\SoundBoard\Cheryl\VEHICLE INFO\First Vehicle.mp3")
                            End Select
                        Case 2
                            rolltheclipThread("C:\SoundBoard\Cheryl\VEHICLE INFO\2nd Vehicle.mp3")
                        Case 3
                            rolltheclipThread("C:\SoundBoard\Cheryl\VEHICLE INFO\Third Vehicle.mp3")
                        Case 4
                            rolltheclipThread("C:\SoundBoard\Cheryl\VEHICLE INFO\Fourth Vehicle.mp3")
                    End Select

                Case 8
                    rolltheclipThread("C:\SoundBoard\Cheryl\VEHICLE INFO\who makes that vehicle.mp3")
                Case 9
                    rolltheclipThread("C:\SoundBoard\Cheryl\VEHICLE INFO\what is the model of the car 1.mp3")
                Case 10

                    If getBirthdaWAV() = True Then
                        tbCallOrder.SelectedTab = tbDriverInfo
                        rolltheclipThread("C:\Soundboard\Cheryl\Birthday\" & bmonth1 & bday1 & ".mp3")
                        While (waveOut.PlaybackState = 1)
                            Console.WriteLine("Checking Birthday")
                        End While
                        rolltheclipThread("C:\Soundboard\Cheryl\Birthday\" & byear1 & ".mp3")
                    Else
                        rolltheclipThread("C:\Soundboard\Cheryl\DRIVER INFO\DOB1.mp3")
                    End If
                Case 11
                    rolltheclipThread("c:\soundboard\cheryl\DRIVER INFO\MaritalStatus2.mp3")
                Case 12
                    rolltheclipThread("c:\soundboard\cheryl\DRIVER INFO\SPOUSES FIRST NAME.mp3")
                Case 13
                    rolltheclipThread("c:\soundboard\cheryl\DRIVER INFO\SPOUSES DATE OF BIRTH.mp3")
                Case 14
                    rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\CAN YOU JUST VERIFY THE MONTH.mp3")
                Case 15
                    rolltheclipThread("c:\soundboard\cheryl\PERSONAL INFO\DO YOU OWN OR RENT THE HOME.mp3")
                Case 16
                    rolltheclipThread("c:\soundboard\cheryl\PERSONAL INFO\HOMETYPE.mp3")
                Case 17
                    rolltheclipThread("c:\soundboard\cheryl\REACTIONS\COULD YOU PLEASE VERIFY YOUR ADDRESS.mp3")
                Case 18
                    rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\Spell out address.mp3")
                Case 19
                    rolltheclipThread("C:\SoundBoard\Cheryl\PERSONAL INFO\EMAIL.mp3")
                Case 20
                    rolltheclipThread("C:\SoundBoard\Cheryl\REBUTTALS\Email Rebuttal.mp3")
                Case 21
                    rolltheclipThread("c:\soundboard\cheryl\PERSONAL INFO\Credit.mp3")
                Case 22
                    rolltheclipThread("c:\soundboard\cheryl\PERSONAL INFO\PhoneType.mp3")
                Case 23
                    rolltheclipThread("c:\soundboard\cheryl\PERSONAL INFO\LAST NAME.mp3")
                Case 24
                    If True Then 'LeadForm.Document.GetElementById("frmResidenceType").GetAttribute("value") = "Own" Then
                        HomeQual = True
                        rentQual = False
                    ElseIf True Then 'LeadForm.Document.GetElementById("frmResidenceType").GetAttribute("value") = "Rent" Then
                        rentQual = True
                        HomeQual = False
                    End If
                    If HomeQual = True And LifeQual = True And Mediqual = True Then
                        rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\Home Life Medicare.mp3")
                        Timer2.Enabled = False
                    ElseIf renterQual = True And LifeQual = True And Mediqual = True Then
                        rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\Renters and Medicare.mp3")
                        Timer2.Enabled = False

                    ElseIf HomeQual = True And LifeQual = False And Mediqual = False Then
                        rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\Home Pitch.mp3")
                        Timer2.Enabled = False

                    ElseIf HomeQual = True And LifeQual = True And Mediqual = False Then
                        rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\life and home.mp3")
                        Timer2.Enabled = False

                    ElseIf HomeQual = True And LifeQual = False And Mediqual = True Then
                        rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\Home and Medicare.mp3")
                        Timer2.Enabled = False

                    ElseIf HomeQual = False And LifeQual = True And Mediqual = True Then
                        rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\Life and Medicare.mp3")
                        Timer2.Enabled = False

                    ElseIf rentQual = True And LifeQual = False And Mediqual = False Then
                        rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\Rental Insurance Pitch.mp3")
                        Timer2.Enabled = False

                    ElseIf rentQual = True And LifeQual = True And Mediqual = False Then
                        rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\rental and life insurance.mp3")
                        Timer2.Enabled = False

                    ElseIf rentQual = True And LifeQual = False And Mediqual = True Then
                        rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\Renters Health.mp3")
                        Timer2.Enabled = False

                    ElseIf renterQual = True And LifeQual = True And HealthQual = True Then
                        rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\Renters Health and Life.mp3")
                        Timer2.Enabled = False

                    ElseIf HomeQual = True And LifeQual = False And HealthQual = False Then
                        rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\Home Pitch.mp3")
                        Timer2.Enabled = False

                    ElseIf HomeQual = True And LifeQual = True And HealthQual = False Then
                        rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\life and home insurance.mp3")
                        Timer2.Enabled = False

                    ElseIf HomeQual = True And LifeQual = False And HealthQual = True Then
                        rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\Home and Health.mp3")
                        Timer2.Enabled = False

                    ElseIf HomeQual = False And LifeQual = True And HealthQual = True Then
                        rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\Life and Health Insurance Pitch.mp3")
                        Timer2.Enabled = False

                    ElseIf rentQual = True And LifeQual = False And HealthQual = False Then
                        rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\Rental Insurance Pitch.mp3")
                        Timer2.Enabled = False

                    ElseIf rentQual = True And LifeQual = True And HealthQual = False Then
                        rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\rental and life insurance.mp3")
                        Timer2.Enabled = False

                    ElseIf rentQual = True And LifeQual = False And HealthQual = True Then
                        rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\Rental and Health Pitch.mp3")
                        Timer2.Enabled = False
                    End If
                Case 25
                    rolltheclipThread("c:\soundboard\cheryl\WRAPUP\YEARBUILT.mp3")
                    Timer2.Enabled = False
                Case 26
                    rolltheclipThread("c:\soundboard\cheryl\WRAPUP\SQUARE FOOTAGE.mp3")
                    Timer2.Enabled = False
                Case 27
                    rolltheclipThread("c:\soundboard\cheryl\WRAPUP\TCPA.mp3")
                    Timer2.Enabled = False
                Case 28
                    Timer2.Enabled = False
                    NICount = 0
            End Select
            tmrSilence.Enabled = True
        Catch ex As Exception
            Console.WriteLine(ex)
        End Try
        isQuestion = True
        counter += 1
    End Sub     'ASKS THE NEXT QUESTION TO KEEP THE CALL MOVING
    Dim NATotal As Integer
    Dim NITotal As Integer
    Dim DNCTotal As Integer
    Dim WrongNumTotal As Integer
    Dim noCarTotal As Integer


    Public Sub DispositionCall()
        Dim resp As Net.WebResponse
        theSilence = 0
        tmrSilence.Enabled = False
        alreadyLoaded = False
        callPos = ""
        clipType = ""
        m.EndMicAndRecognition()
        StopThatClip()
        NumberOfVehicles = 1
        VehicleNum = 1
        For i As Integer = 0 To 3
            VYear(i) = ""
            vMake(i) = ""
            vmodel(i) = ""
        Next

        numRepeats = 0
        Timer2.Enabled = False

        introHello = True
        Timer2.Enabled = False
        stillthere = 0
        isQuestion = True
        numRepeats = 0
        HumanCounter = 1
        newcall = True
        insurancePass = False
        lblQuestion.Text = CURRENTQUESTION(1)
        tbCallOrder.SelectedTab = tbIntro
        introBday = False
        NICount = 0
        timesAsking = 0
        counter = 0
        counter2 = 0
        CurrentQ = 0

        txtSpeech.Clear()
        SilenceReps = 0
        stillthere = 0
        isQuestion = False
        calltime = 0

        Hangup = Net.WebRequest.Create("http://loudcloud9.ytel.com/x5/api/agent.php?source=test&user=101&pass=API101IEpost&agent_user=" & txtVerifierNum.Text & "&function=external_hangup&value=1")
        resp = Hangup.GetResponse
        resp.Close()
        Thread.Sleep(500)
        Console.WriteLine("***********************")
        Console.WriteLine(cmbDispo.Text)
        Console.WriteLine("***********************")
        Select Case cmbDispo.Text
            Case "Not Available"
                Disposition = Net.WebRequest.Create("http://loudcloud9.ytel.com/x5/api/agent.php?source=test&user=101&pass=API101IEpost&agent_user=" & txtVerifierNum.Text & "&function=external_status&value=" & "NotAvl")
                resp = Disposition.GetResponse
                cmbDispo.SelectedIndex = -1
            Case "Not Interested"
                Disposition = Net.WebRequest.Create("http://loudcloud9.ytel.com/x5/api/agent.php?source=test&user=101&pass=API101IEpost&agent_user=" & txtVerifierNum.Text & "&function=external_status&value=" & "NI")
                resp = Disposition.GetResponse
                cmbDispo.SelectedIndex = -1
            Case "Do Not Call"
                Disposition = Net.WebRequest.Create("http://loudcloud9.ytel.com/x5/api/agent.php?source=test&user=101&pass=API101IEpost&agent_user=" & txtVerifierNum.Text & "&function=external_status&value=" & "DNC")
                resp = Disposition.GetResponse
                cmbDispo.SelectedIndex = -1
            Case "Wrong Number"
                Disposition = Net.WebRequest.Create("http://loudcloud9.ytel.com/x5/api/agent.php?source=test&user=101&pass=API101IEpost&agent_user=" & txtVerifierNum.Text & "&function=external_status&value=" & "Wrong")
                resp = Disposition.GetResponse
                cmbDispo.SelectedIndex = -1
            Case "No Car"
                Disposition = Net.WebRequest.Create("http://loudcloud9.ytel.com/x5/api/agent.php?source=test&user=101&pass=API101IEpost&agent_user=" & txtVerifierNum.Text & "&function=external_status&value=" & "NoCar")
                resp = Disposition.GetResponse
                cmbDispo.SelectedIndex = -1
            Case "Auto Lead"
            Case "No English"
                Disposition = Net.WebRequest.Create("http://loudcloud9.ytel.com/x5/api/agent.php?source=test&user=101&pass=API101IEpost&agent_user=" & txtVerifierNum.Text & "&function=external_status&value=" & "NoEng")
                resp = Disposition.GetResponse
                cmbDispo.SelectedIndex = -1
            Case "Lost On Wrap Up"
                Disposition = Net.WebRequest.Create("http://loudcloud9.ytel.com/x5/api/agent.php?source=test&user=101&pass=API101IEpost&agent_user=" & txtVerifierNum.Text & "&function=external_status&value=" & "LOW")
                resp = Disposition.GetResponse
                cmbDispo.SelectedIndex = -1
            Case "Entering Lead/Low"
        End Select
        resp.Close()
    End Sub                                                 ' DISPOSITIONS THE CALL 
    Dim switch As Boolean = False
    Dim temperstring As String
    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick           'CHECKS TO SEE THAT CHERYL IS NOT TALKING SO THE CALL CAN MOVE ON
        Label3.Text = CurrentQ
        If waveOut.PlaybackState = 0 Then
            If CurrentQ < 30 Then
                AskQuestion(CurrentQ, counter)
                Timer2.Enabled = False
            Else
                DispositionCall()
            End If
        Else
        End If
    End Sub
    Dim endcall As Boolean = False
    Dim DISPO As String

    Dim REJECTCOUNT As Integer

    Dim HypBMonth As String
    Dim HypBDay As String
    Dim HypByear As String
    Dim tempcounter As Integer
    Dim temprec As String

    Dim BDayString As String
    Dim addyzip As String
    Dim AUTOYEAR As String
    Dim AUTOMAKE As String
    Dim AUTOMODEL As String

    Function getLastWord(speech As String) As String
        Dim temp() As String = speech.Split()
        Return temp(temp.Length - 1)
    End Function
    Function isNumber(speech As String) As Boolean
        Select Case speech.ToLower
            Case "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "zero"
                Return True
            Case Else
                Return False
        End Select
    End Function

    Private Sub Button29_Click_3(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\VEHICLE INFO\YEAR OF THE FIRST VEHICLE.mp3")
    End Sub

    Private Sub Button30_Click_3(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\VEHICLE INFO\MAKE OF THE FIRST VEHICLE.mp3")
    End Sub

    Private Sub Button45_Click_1(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\VEHICLE INFO\MODEL OF THE FIRST VEHICLE.mp3")
    End Sub

    Private Sub Button22_Click_1(sender As Object, e As EventArgs) Handles Button22.Click
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\That's okay.mp3")
    End Sub

    Private Sub Button18_Click_4(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\REACTIONS\Wonderful.mp3")
    End Sub

    Private Sub Button47_Click(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\REACTIONS\This info.mp3")
    End Sub

    Private Sub Button53_Click_2(sender As Object, e As EventArgs) Handles Button53.Click
        isQuestion = True
        rolltheclipThread("C:/Soundboard/Cheryl/PERSONAL INFO/email.mp3")
        clipType = "Question"
        callPos = Email_Address

    End Sub

    Private Sub Button31_Click_2(sender As Object, e As EventArgs)
        Playlist(0) = "C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\Address Rebuttal.mp3"

    End Sub

    Private Sub Button55_Click_1(sender As Object, e As EventArgs)
        Playlist(0) = "C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\Address Rebuttal 2.mp3"

    End Sub

    Private Sub Button49_Click_1(sender As Object, e As EventArgs)
        Playlist(0) = "C:\SoundBoard\Cheryl\REACTIONS\More More More Cheryl\More More More Cheryl\Email Rebuttal.mp3"

    End Sub

    Private Sub Button56_Click_1(sender As Object, e As EventArgs)
        Playlist(0) = "C:\SoundBoard\Cheryl\REBUTTALS\P.O Box rebuttal.mp3"

    End Sub

    Private Sub Button57_Click_1(sender As Object, e As EventArgs) Handles Button57.Click
        rolltheclipThread("C:\SoundBoard\Cheryl\REBUTTALS\Disclaimer 2.mp3")
    End Sub

    Dim played As Boolean = False


    Private Sub Button58_Click_1(sender As Object, e As EventArgs)
        Playlist(0) = "C:\SoundBoard\Cheryl\REBUTTALS\New address rebuttal.mp3"

    End Sub

    Private Sub Button59_Click(sender As Object, e As EventArgs)
        Playlist(0) = "C:\SoundBoard\Cheryl\PERSONAL INFO\I would just need an email address that you have access to.mp3"

    End Sub

    Private Sub homerenters_CheckedChanged(sender As Object, e As EventArgs) Handles HomeCheck.CheckedChanged
        If cmbSecondaries.Text = "YES" Then
            If HomeCheck.Visible = True And HomeCheck.Checked Or RenterCheck.Visible = True And RenterCheck.Checked Then
                lblQuestion.Text = "YEAR BUILT"
            Else
                lblQuestion.Text = "TCPA"
            End If
        Else
            lblQuestion.Text = "TCPA"
        End If
    End Sub

    Private Sub Timer5_Tick(sender As Object, e As EventArgs)
        If notReady = True Then
            If waveOut.PlaybackState = 0 Then
                notReady = False


            End If
        End If
    End Sub


    Private Sub SQFT_Click(sender As Object, e As EventArgs) Handles SQFT.Click
        If HomeCheck.Visible = True Then
            rolltheclipThread("C:\SoundBoard\Cheryl\WRAPUP\Square Footage.mp3")
        Else
            rolltheclipThread("C:\SoundBoard\Cheryl\WRAPUP\PPCoverage.mp3")
        End If


    End Sub
    Dim weHaveBDay As Boolean = False
    Private Sub Timer6_Tick(sender As Object, e As EventArgs)
        Select Case CurrentQ
            Case 6
                If clipnum(6) = 0 Then
                    Playlist(1) = "c:\soundboard\cheryl\INSURANCE INFO\WHO Is THE CURRENT AUTO INSURANCE COMPANY THAT YOU'RE WITH.mp3"
                Else
                    Playlist(1) = "c:\soundboard\cheryl\INSURANCE INFO\WHODOYOUUSE.mp3"
                End If
            Case 7
                Playlist(1) = "c:\soundboard\cheryl\INSURANCE INFO\EXPIRATION.mp3"
            Case 8
                Playlist(1) = "c:\soundboard\cheryl\INSURANCE INFO\HOW MANY YEARS HAVE YOU BEEN WITH THEM 2.mp3"
            Case 10
                If cmbMoreVehicles.Text = "YES" Then
                    Playlist(1) = "c:\SoundBoard\Cheryl\VEHICLE INFO\YMMTV.mp3"
                Else

                    Playlist(1) = "c:\soundboard\cheryl\VEHICLE INFO\YMMYV.mp3"
                End If
            Case 11
                Playlist(1) = "c:\soundboard\cheryl\VEHICLE INFO\OTHER VEHICLES ON THAT POLICY.mp3"
            Case 12

            Case 13

            Case 14

                Playlist(1) = "c:\soundboard\cheryl\DRIVER INFO\MARITAL STATUS.mp3"
            Case 15
                Playlist(1) = "c:\soundboard\cheryl\DRIVER INFO\SPOUSES FIRST NAME.mp3"
            Case 16
                Playlist(1) = "c:\soundboard\cheryl\DRIVER INFO\SPOUSES DATE OF BIRTH.mp3"
            Case 17
                Playlist(1) = "c:\soundboard\cheryl\PERSONAL INFO\DO YOU OWN OR RENT THE HOME.mp3"
            Case 18
                Playlist(1) = "c:\soundboard\cheryl\PERSONAL INFO\HOMETYPE.mp3"
            Case 19
                Playlist(1) = "C:\SoundBoard\Cheryl\REACTIONS\could you please verify your address.mp3"
            Case 20
                Playlist(1) = "C:\SoundBoard\Cheryl\PERSONAL INFO\email.mp3"
            Case 21
                Playlist(1) = "c:\soundboard\cheryl\PERSONAL INFO\Credit.mp3"
            Case 22
                Playlist(1) = "c:\soundboard\cheryl\PERSONAL INFO\PhoneType.mp3"
            Case 23
                Playlist(1) = "c:\soundboard\cheryl\PERSONAL INFO\LAST NAME.mp3"
            Case 24
                Playlist(1) = "C:\SoundBoard\Cheryl\REBUTTALS\Disclaimer 2.mp3"
            Case Else
                Playlist(1) = "NULL"
        End Select
        If waveOut.PlaybackState = 0 Then
            If playcounter < 1 Then
                rolltheclipThread(Playlist(playcounter))
                playcounter += 1
                clipnum(6) += 1

            ElseIf playcounter = 1 Then

                If Playlist(1) <> "NULL" Then
                    rolltheclipThread(Playlist(playcounter))
                    playcounter += 1
                    clipnum(6) = 0
                    clipnum(13) += 1
                Else
                    playcounter = 0

                    weHaveBDay = True
                    Playlist(0) = "NULL"
                    Playlist(1) = "NULL"
                    clipnum(13) = 1
                End If
            Else
                playcounter = 0

                weHaveBDay = True
                Playlist(0) = "NULL"
                Playlist(1) = "NULL"
            End If
        End If
    End Sub

    Private Sub Button61_Click_2(sender As Object, e As EventArgs) Handles Button61.Click
        rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\Thank-You.mp3")
    End Sub

    Private Sub Button70_Click(sender As Object, e As EventArgs) Handles Button70.Click
        rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\You're-Welcome.mp3")
    End Sub

    Private Sub Button72_Click_1(sender As Object, e As EventArgs)
        rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\zip.mp3")
    End Sub


    Dim HELLOCOUNT As Integer = 0
    Dim repeatcount As Integer = 0
    Dim CurrentVehicle As Integer = 0

    Private Sub cmbOwnRent_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbOwnRent.SelectedIndexChanged
        If cmbOwnRent.Text = "OWN" Then
            HOMEOWNER = True
            HomeCheck.Visible = True
            RenterCheck.Visible = False

        ElseIf cmbOwnRent.Text = "RENT" Then

            HomeCheck.Visible = False
            RenterCheck.Visible = True
        Else
            HomeCheck.Visible = False
            RenterCheck.Visible = False

        End If
    End Sub

    Private Sub cmbSecondaries_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSecondaries.SelectedIndexChanged
        If cmbSecondaries.Text = "YES" Then
            If HomeCheck.Visible = True And HomeCheck.Checked Or RenterCheck.Visible = True And RenterCheck.Checked Then
                lblQuestion.Text = "YEAR BUILT"
            Else
                lblQuestion.Text = "TCPA"
            End If
        Else
            lblQuestion.Text = "TCPA"
        End If
    End Sub

    Private Sub RenterCheck_CheckedChanged(sender As Object, e As EventArgs) Handles RenterCheck.CheckedChanged
        If cmbSecondaries.Text = "YES" Then
            If HomeCheck.Visible = True And HomeCheck.Checked Or RenterCheck.Visible = True And RenterCheck.Checked Then
                lblQuestion.Text = "YEAR BUILT"
            Else
                lblQuestion.Text = "TCPA"
            End If
        Else
            lblQuestion.Text = "TCPA"
        End If
    End Sub

    Private Sub tbDriverInfo_Click(sender As Object, e As EventArgs) Handles tbDriverInfo.Click

    End Sub

    Private Sub tbVehicleInfo_Click(sender As Object, e As EventArgs) Handles tbVehicleInfo.Click

    End Sub






    Private Sub Button12_Click_1(sender As Object, e As EventArgs) Handles Button12.Click

        Reset()
        cmbMoreVehicles.SelectedIndex = 0
        theurl = ""
        NICount = 0
        rolltheclipThread("C:/Soundboard/Cheryl/WRAPUP/have a great day.mp3")
        cmbDispo.Text = "Not Interested"
        totalCalls = totalCalls + 1
        lblCalls.Text = totalCalls
        CurrentQ = 0
        lblQuestion.Text = "HELLO"
        txtInsuranceProvider.Clear()
        txtPolicyExpiration.Clear()
        txtPolicyStart.Clear()
        CurrentQ = 31
        Timer2.Enabled = True
        resetBot()

    End Sub

    Private Sub Button6_Click_1(sender As Object, e As EventArgs) Handles Button6.Click


        Reset()
        cmbMoreVehicles.SelectedIndex = 0

        theurl = ""
        NICount = 0
        rolltheclipThread("C:/Soundboard/Cheryl/WRAPUP/have a great day.mp3")
        cmbDispo.Text = "Wrong Number"
        totalCalls = totalCalls + 1
        lblCalls.Text = totalCalls
        CurrentQ = 1
        lblQuestion.Text = "HELLO"
        txtInsuranceProvider.Clear()
        txtPolicyExpiration.Clear()
        txtPolicyStart.Clear()
        CurrentQ = 31
        Timer2.Enabled = True
        resetBot()
    End Sub

    Private Sub Button10_Click_1(sender As Object, e As EventArgs) Handles Button10.Click
        Reset()
        cmbMoreVehicles.SelectedIndex = 0
        theurl = ""
        NICount = 0
        rolltheclipThread("C:/Soundboard/Cheryl/WRAPUP/have a great day.mp3")
        cmbDispo.Text = "No Car"
        CurrentQ = 31
        Timer2.Enabled = True
        resetBot()
    End Sub

    Private Sub Button35_Click_1(sender As Object, e As EventArgs) Handles Button35.Click
        Reset()
        cmbMoreVehicles.SelectedIndex = 0
        theurl = ""
        NICount = 0
        rolltheclipThread("C:/Soundboard/Cheryl/WRAPUP/have a great day.mp3")

        cmbDispo.Text = "No English"
        CurrentQ = 31
        Timer2.Enabled = True
        resetBot()
    End Sub

    Dim spot1 As Integer
    Dim spot2 As Integer

    Private Sub Button84_Click_1(sender As Object, e As EventArgs) Handles Button84.Click
        rolltheclipThread("C:/SOUNDBOARD/CHERYL/REBUTTALS/JANUARY FEB MARCH APRIL.mp3")
    End Sub

    Private Sub Button38_Click_1(sender As Object, e As EventArgs) Handles Button38.Click
        rolltheclipThread("C:\SoundBoard\Cheryl\TIE INS\Great What's Your Best Guess.mp3")
    End Sub

    Private Sub Button4_Click_2(sender As Object, e As EventArgs) Handles Button4.Click

        rolltheclipThread("C:\SoundBoard\Cheryl\REBUTTALS\THIS WILL BE REAL QUICK.mp3")

    End Sub

    Private Sub Button77_Click_1(sender As Object, e As EventArgs) Handles Button77.Click
        rolltheclipThread("C:\Users\Insurance Express\Desktop\Cheryl MP3\Old Sounds\JUSTABOUTDONE.mp3")
    End Sub

    Private Sub Button59_Click_1(sender As Object, e As EventArgs) Handles Button59.Click
        rolltheclipThread("C:\SoundBoard\Cheryl\PERSONAL INFO\I would just need an email address that you have access to.mp3")
    End Sub

    Private Sub Button19_Click_2(sender As Object, e As EventArgs) Handles Button19.Click

        rolltheclipThread("C:\SoundBoard\Cheryl\REBUTTALS\HappyWithInsurance.mp3")


    End Sub

    Private Sub Button32_Click_3(sender As Object, e As EventArgs) Handles Button32.Click

        rolltheclipThread("C:\SoundBoard\Cheryl\REBUTTALS\this info.mp3")



    End Sub

    Private Sub Button69_Click_1(sender As Object, e As EventArgs) Handles Button69.Click

        rolltheclipThread("C:\SoundBoard\Cheryl\REBUTTALS\I actually have this information.mp3")

    End Sub

    Private Sub Button26_Click_2(sender As Object, e As EventArgs) Handles Button26.Click

        rolltheclipThread("C:\SoundBoard\Cheryl\REBUTTALS\I Already Have Insurance rebuttal.mp3")


    End Sub

    Private Sub Button25_Click_2(sender As Object, e As EventArgs) Handles Button25.Click

        rolltheclipThread("C:\SoundBoard\Cheryl\REBUTTALS\nothing to be interested in.mp3")


    End Sub

    Private Sub Button49_Click_2(sender As Object, e As EventArgs) Handles Button49.Click

        rolltheclipThread("C:\SoundBoard\Cheryl\REBUTTALS\EMAIL REBUTTAL.mp3")

    End Sub

    Private Sub Button58_Click_2(sender As Object, e As EventArgs) Handles Button58.Click

        rolltheclipThread("C:\SoundBoard\Cheryl\REBUTTALS\NEW ADDRESS REBUTTAL.mp3")

    End Sub

    Private Sub Button56_Click_2(sender As Object, e As EventArgs) Handles Button56.Click

        rolltheclipThread("C:\SoundBoard\Cheryl\REBUTTALS\P.O BOX REBUTTAL.mp3")

    End Sub

    Private Sub Button31_Click_3(sender As Object, e As EventArgs) Handles Button31.Click

        rolltheclipThread("C:\SoundBoard\Cheryl\REBUTTALS\ADDRESS REBUTTAL.mp3")

    End Sub

    Private Sub Button5_Click_2(sender As Object, e As EventArgs) Handles Button5.Click

        rolltheclipThread("C:\SoundBoard\Cheryl\REBUTTALS\My spouse takes care of that.mp3")


    End Sub

    Private Sub Button68_Click_2(sender As Object, e As EventArgs) Handles Button68.Click
        rolltheclipThread("C:\SoundBoard\Cheryl\REBUTTALS\Is this the spouse.mp3")
    End Sub

    Private Sub Button52_Click_2(sender As Object, e As EventArgs) Handles Button52.Click
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\I HAVE A FEW HERE.mp3")
    End Sub

    Private Sub Button48_Click_1(sender As Object, e As EventArgs) Handles Button48.Click
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\ONE AT A TIME PLEASE.mp3")
    End Sub

    Private Sub Make_Click_1(sender As Object, e As EventArgs) Handles Make.Click
        rolltheclipThread("c:\soundboard\cheryl\PUSHONS\chevyfordgmc.mp3")
    End Sub

    Private Sub insurance_Click_1(sender As Object, e As EventArgs) Handles insurance.Click
        rolltheclipThread("c:\soundboard\cheryl\PUSHONS\allstategeicostatefarm.mp3")

    End Sub

    Private Sub Button50_Click_1(sender As Object, e As EventArgs) Handles Button50.Click
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\CAN YOU JUST VERIFY THE MONTH.mp3")
    End Sub

    Private Sub Button55_Click_3(sender As Object, e As EventArgs) Handles Button55.Click
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\CAN YOU JUST VERIFY THE DAY.mp3")
    End Sub

    Private Sub Button54_Click(sender As Object, e As EventArgs) Handles Button54.Click
        rolltheclipThread("c:\soundboard\cheryl\REBUTTALS\CAN YOU JUST VERIFY THE YEAR.mp3")
    End Sub

    Private Sub Button14_Click_1(sender As Object, e As EventArgs) Handles Button14.Click
        rolltheclipThread("C:\SoundBoard\Cheryl\PERSONAL INFO\email.mp3")
    End Sub

    Private Sub Button72_Click_2(sender As Object, e As EventArgs) Handles Button72.Click
        rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\zip.mp3")
    End Sub

    Private Sub Button29_Click_4(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\VEHICLE INFO\YEAR OF THE FIRST VEHICLE.mp3")
    End Sub

    Private Sub Button30_Click_4(sender As Object, e As EventArgs)
        rolltheclipThread("c:\soundboard\cheryl\VEHICLE INFO\MAKE OF THE FIRST VEHICLE.mp3")
    End Sub

    Private Sub Button45_Click_2(sender As Object, e As EventArgs) Handles Button45.Click
        rolltheclipThread("c:\soundboard\cheryl\VEHICLE INFO\MODEL OF THE FIRST VEHICLE.mp3")
    End Sub

    Private Sub Button60_Click_1(sender As Object, e As EventArgs) Handles Button60.Click

        rolltheclipThread("C:\SoundBoard\Cheryl\REBUTTALS\REBUTTAL3.mp3")

    End Sub

    Private Sub Button29_Click_5(sender As Object, e As EventArgs) Handles Button29.Click
        If My.Computer.Keyboard.CtrlKeyDown = False Then
            Playlist(0) = ("C:\SoundBoard\Cheryl\REACTIONS\BEST NI REBUTTALS ZIP\BEST NI REBUTTALS\Im sure what.mp3")

        Else
            rolltheclipThread("C:\SoundBoard\Cheryl\REACTIONS\BEST NI REBUTTALS ZIP\BEST NI REBUTTALS\Im sure what.mp3")
        End If


    End Sub
    Dim NumClicks As Integer = 0
    Private Sub Button39_Click_2(sender As Object, e As EventArgs) Handles Button39.Click
        If NumClicks = 0 Then
            Reset()
            cmbMoreVehicles.SelectedIndex = 0
            theurl = ""
            NICount = 0
            cmbDispo.Text = "Entering Lead/Low"
            resetBot()
            rolltheclipThread("C:/Soundboard/Cheryl/WRAPUP/ENDCALL.mp3")
            NumClicks += 1
            CurrentQ = 31
            Timer2.Enabled = True
        Else
            Disposition = Net.WebRequest.Create("http://loudcloud9.ytel.com/x5/api/agent.php?source=test&user=101&pass=API101IEpost&agent_user=" & txtVerifierNum.Text & "&function=external_status&value=" & "1Auto")
            NumClicks = 0
        End If

    End Sub

    Private Sub Button44_Click_2(sender As Object, e As EventArgs) Handles Button44.Click
        If NumClicks = 0 Then
            Reset()
            cmbMoreVehicles.SelectedIndex = 0
            theurl = ""
            NICount = 0
            cmbDispo.Text = "Entering Lead/Low"
            rolltheclipThread("C:/Soundboard/Cheryl/WRAPUP/have a great day.mp3")
            CurrentQ = 31
            resetBot()
            Timer2.Enabled = True
        Else
            Disposition = Net.WebRequest.Create("http://loudcloud9.ytel.com/x5/api/agent.php?source=test&user=101&pass=API101IEpost&agent_user=" & txtVerifierNum.Text & "&function=external_status&value=" & "1Auto")
        End If

    End Sub

    Private Sub cmbYear_LocationChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button18_Click_5(sender As Object, e As EventArgs) Handles Button18.Click
        rolltheclipThread("C:\SoundBoard\Cheryl\REBUTTALS\Sorry to hear that 2.mp3")
    End Sub

    Private Sub Button27_Click_2(sender As Object, e As EventArgs) Handles Button27.Click
        If clipnum(9) = 0 Then
            rolltheclipThread("C:\Soundboard\Cheryl\REACTIONS\Loud-laugh.mp3")
            clipnum(9) += 1
        Else
            rolltheclipThread("C:\Soundboard\Cheryl\REACTIONS\softer-Laugh.mp3")
            clipnum(9) = 0

        End If

    End Sub

    Private Sub cmbYear_Disposed(sender As Object, e As EventArgs)

    End Sub

    Private Sub cmbYear_SelectionChangeCommitted(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button40_Click_1(sender As Object, e As EventArgs) Handles Button40.Click

        rolltheclipThread("C:/Soundboard/Cheryl/Rebuttals/NotThatCheap.mp3")


    End Sub
    Dim errorText As String = ""
    Private Sub txtVerifierNum_TextChanged(sender As Object, e As EventArgs) Handles txtVerifierNum.TextChanged

    End Sub
    Dim UserList(4) As Integer
    Public Sub ErrorOutput()
        Console.WriteLine(errorText)
    End Sub
    Public Sub errorHandler(ByVal sender As Object, ByVal e As SpeechErrorEventArgs) Handles m.OnConversationError

        errorText = e.SpeechErrorText
        BeginInvoke(New Action(AddressOf ErrorOutput))

    End Sub

    Public Structure Agent

    End Structure

    Public Function GenerateStats() As String
        req = Net.WebRequest.Create("http://loudcloud9.ytel.com/x5/api/non_agent.php?source=test&user=101&pass=API101IEpost&function=agent_status&agent_user=" & txtVerifierNum.Text & "&stage=csv&header=YES")
        Dim webResp As Net.WebResponse = req.GetResponse
        Dim webReader As New IO.StreamReader(webResp.GetResponseStream)
        Dim stats As String = webReader.ReadToEnd()
        webResp.Close()
        Return stats
    End Function



    Private Sub txtVerifierNum_Click(sender As Object, e As EventArgs) Handles txtVerifierNum.Click
        Dim dir As String = AppDomain.CurrentDomain.BaseDirectory

        txtVerifierNum.Text = InputBox("enter agent #: ")
        If txtVerifierNum.Text.ToLower() = "moo" Then
            Form3.Show()
            rolltheclipThread(dir + "cow-moo.mp3")
        ElseIf txtVerifierNum.Text = "philip j fry" Then
            rolltheclipThread(dir + "goodnewseveryone.mp3")
            tmrAgentStatus.Enabled = True
        Else
            Dim opt As New ChromeOptions
            opt.AddArguments("disable-popup-blocking")
            local_browser = New ChromeDriver("C:\Users\Insurance Express\Downloads\chromedriver_win32", opt)  ' fun fact, you can just pass Nothing as the profile and it'll work fine(:
            local_browser.Manage.Timeouts.ImplicitlyWait(TimeSpan.FromSeconds(10))
            local_browser.Navigate.GoToUrl("https://loudcloud9.ytel.com")
            local_browser.SwitchTo().Frame("top")
            Thread.Sleep(550)
            local_browser.FindElementById("login-agent").Click()
            local_browser.FindElementById("agent-login").SendKeys(txtVerifierNum.Text)
            Thread.Sleep(500)
            local_browser.FindElementById("agent-password").SendKeys("y" & txtVerifierNum.Text & "IE")
            Thread.Sleep(500)
            local_browser.FindElementById("btn-get-campaign").Click()
            Thread.Sleep(500)
            local_browser.FindElementById("select-campaign").Click()
            local_browser.FindElementById("select-campaign").FindElements(By.TagName("option")).Last.Click()
            Thread.Sleep(250)
            local_browser.FindElementById("btn-submit").Click()
            tmrAgentStatus.Enabled = True
        End If

    End Sub
    Dim req As Net.WebRequest
    Dim Hangup As Net.WebRequest
    Dim Disposition As Net.WebRequest

    'CustName(0) = local_browser.FindElementById("frmFirstName").GetAttribute("value")
    '                CustName(1) = local_browser.FindElementById("frmLastName").GetAttribute("value")
    '                btnTheirName.Text = CustName(0)
    '                globalFile = "C:\Soundboard\Cheryl\Names\" & CustName(0) & " 1.mp3"
    '                globalFile2 = "C:\Soundboard\Cheryl\Names\" & CustName(0) & " 3.mp3"
    '                globalfile3 = "C:\Soundboard\Cheryl\Names\" & CustName(0) & " 2.mp3"
    Dim alreadyLoaded As Boolean = False
    Public Sub getLeadWindow()
        If alreadyLoaded = False Then
            If Not local_browser.Url.Contains("forms.lead.co") Then
                If local_browser.WindowHandles.Count > 1 Then
                    Try
                        local_browser.SwitchTo().Window(local_browser.WindowHandles.Last)
                    Catch ex As Exception
                        Console.WriteLine("uncaught exception")
                    End Try
                Else
                    Exit Sub
                End If
            End If
            Try
                If local_browser.PageSource.Contains("Please respectfully") Then
                    cmbDispo.Text = "Not Interested"
                    DispositionCall()
                End If
                If local_browser.PageSource.Contains("not found") Then
                    cmbDispo.Text = "Not Available"
                    DispositionCall()
                End If
                If local_browser.PageSource.Contains("Lead(s)") Then
                    Exit Sub
                End If
                Dim name As String = local_browser.FindElementById("frmFirstName").GetAttribute("value")
                btnTheirName.Text = name
                CustName(0) = name
                CustName(1) = name
                Try
                    If My.Computer.FileSystem.FileExists("C:\Soundboard\Cheryl\Names\" & CustName(0) & " 1.mp3") Then
                        globalFile = "C:\Soundboard\Cheryl\Names\" & CustName(0) & " 1.mp3"
                        btnTheirName.BackgroundImage = Nothing
                    Else

                        btnTheirName.BackgroundImage = System.Drawing.Image.FromFile("C:/NoSoundClip.jpg")

                    End If
                Catch ex As Exception
                    Console.WriteLine(ex.InnerException)
                    Console.WriteLine(ex.StackTrace)
                End Try
                If My.Computer.FileSystem.FileExists("C:\Soundboard\Cheryl\Names\" & CustName(0) & " 3.mp3") Then
                    globalFile2 = "C:\Soundboard\Cheryl\Names\" & CustName(0) & " 3.mp3"
                    Button15.BackgroundImage = Nothing
                Else
                    Button15.BackgroundImage = System.Drawing.Image.FromFile("C:/NoSoundClip.jpg")
                    Button17.BackgroundImageLayout = 2

                End If
                If My.Computer.FileSystem.FileExists("C:\Soundboard\Cheryl\Names\" & CustName(0) & " 2.mp3") Then
                    globalfile3 = "C:\Soundboard\Cheryl\Names\" & CustName(0) & " 2.mp3"
                    Button17.BackgroundImage = Nothing
                Else
                    Button17.BackgroundImage = System.Drawing.Image.FromFile("C:/NoSoundClip.jpg")
                    Button17.BackgroundImageLayout = 2



                End If

                alreadyLoaded = True

            Catch ex As Exception
                Console.WriteLine(ex.InnerException)
            End Try
        End If
    End Sub

    Private Sub tmrAgentStatus_Tick(sender As Object, e As EventArgs) Handles tmrAgentStatus.Tick
        getLeadWindow()
        Label3.Text = CurrentQ
        lblQuestion.Text = CURRENTQUESTION(CurrentQ)
        If txtVerifierNum.Text <> "" Then
            Dim STATS As String = GenerateStats()
            tempStr = STATS.Split(",")
            If STATS.Contains("INCALL") Then
                If newcall = True Then
                    newcall = False
                    alreadyLoaded = False
                    lblStatus.Text = "STATUS: " & "INCALL"
                    Me.BackColor = Color.Green
                    introHello = True
                    calltimer.Enabled = True
                End If
            End If
            If STATS.Contains("DISPO") Then
                introHello = False
                alreadyLoaded = False
                lblStatus.Text = "STATUS: " & "DISPO"
            ElseIf STATS.Contains("READY") Then
                lblStatus.Text = "STATUS: " & "READY"
                Me.BackColor = Color.Yellow
                newcall = True
                alreadyLoaded = False
                btnPause.Text = "Pause"
                btnPause.BackColor = Color.Red
            ElseIf STATS.Contains("PAUSED") Then
                btnPause.Text = "Resume"
                btnPause.BackColor = Color.Green
                alreadyLoaded = False
                lblStatus.Text = "STATUS: " & "PAUSED"
                isecond += 0.75
                Me.BackColor = Color.Red

            ElseIf STATS.Contains("DEAD") Then
                introHello = False
                StopThatClip()
                alreadyLoaded = False
                CurrentQ = 31
                Timer2.Enabled = True
            End If
            If tempStr.Length > 12 Then
                lblName.Text = tempStr(12)
                lblName2.Text = tempStr(12)
                If CInt(lblCalls.Text) <> tempStr(11) Then
                    lblCalls.Text = tempStr(11)
                    lblCalls2.Text = tempStr(11)
                End If
            End If
        End If
    End Sub 'Sends API Call to get agent report


    Private Sub WebBrowser1_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs)
        moo()
    End Sub 'WebBrowser Object reserved for Dispositioning calls



    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Unregister()
        If local_browser IsNot Nothing Then
            local_browser.Dispose()
        End If
    End Sub

    Private Sub Button1_Click_4(sender As Object, e As EventArgs) Handles btnPause.Click
        Dim Pause As Net.WebRequest
        Dim resp As Net.WebResponse
        If btnPause.Text = "Pause" Then
            Pause = Net.WebRequest.Create("http://loudcloud9.ytel.com/x5/api/agent.php?source=test&user=101&pass=API101IEpost&agent_user=" & txtVerifierNum.Text & "&function=external_pause&value=" & "PAUSE")
            resp = Pause.GetResponse
            resp.Close()
            btnPause.Text = "Resume"
        Else
            Pause = Net.WebRequest.Create("http://loudcloud9.ytel.com/x5/api/agent.php?source=test&user=101&pass=API101IEpost&agent_user=" & txtVerifierNum.Text & "&function=external_pause&value=" & "RESUME")
            resp = Pause.GetResponse
            resp.Close()
            btnPause.Text = "Pause"
        End If
    End Sub 'pause button

    Private Sub moo()
        Console.WriteLine("moo")
    End Sub

    Private Sub Button1_Click_5(sender As Object, e As EventArgs)

    End Sub
    Dim SilenceReps As Integer = 0
    Dim introHello As Boolean = True

    Dim micStatus As Boolean
    Dim isQuestion As Boolean = True

    Public Sub isStopped(sender As Object, e As NAudio.Wave.StoppedEventArgs) Handles waveOut.PlaybackStopped
        inBetween = True
        Console.WriteLine("CHERYLBOT IS DONE SPEAKING...")
        newobjection = True
        Select Case clipType
            Case "Question"
                introHello = False
                m.StartMicAndRecognition()
            Case "Objection"
                Currently_Rebuttaling = False
                Timer2.Enabled = True
        End Select
    End Sub 'checks to see if clip is stopped 
    Private Sub cmbMoreVehicles_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbMoreVehicles.SelectedIndexChanged
        VehicleNum = cmbMoreVehicles.Text
    End Sub

    Private Sub txtVerifierNum_MouseCaptureChanged(sender As Object, e As EventArgs) Handles txtVerifierNum.MouseCaptureChanged

    End Sub



    Private Sub tbIntro_Click(sender As Object, e As EventArgs) Handles tbIntro.Click

    End Sub
    'Dim cds As ChromeDriverService = New ChromeDriverService()
    Private Sub testpagebutton_Click(sender As Object, e As EventArgs) Handles testpagebutton.Click
        newcall = False
        local_browser = New ChromeDriver("C:\Users\Insurance Express\Downloads\chromedriver_win32")
        local_browser.Navigate.GoToUrl("https://forms.lead.co/auto/?key=e2869270-7c7a-11e1-b0c4-0800200c9a66")
    End Sub

    Private Sub tbQuestions_Click(sender As Object, e As EventArgs) Handles tbQuestions.Click

    End Sub

    Dim calltime As Integer = 0
    Private Sub calltimer_Tick(sender As Object, e As EventArgs) Handles calltimer.Tick
        Dim STATS As String = GenerateStats()
        If STATS.Contains("INCALL") Then
            calltime += 1
        End If
        If STATS.Contains("DEAD") Then
            If calltime < 10 Then
                Console.WriteLine("not available")
                Console.WriteLine(calltime)
                cmbDispo.Text = "Not Available"
                DispositionCall()
            ElseIf calltime > 10 Or callPos = Intro Then
                cmbDispo.Text = "Not Interested"
                Console.WriteLine("not interested")
                Console.WriteLine(calltime)
                DispositionCall()
            End If
        End If
    End Sub
    Dim theSilence As Integer = 0
    Dim SilenceCap As Double = 3
    Dim CustomSilence(5) As Integer

    Private Sub tmrSilence_Tick(sender As Object, e As EventArgs) Handles tmrSilence.Tick
        If waveOut.PlaybackState = 0 Then
            Dim temp As Integer = 0
            theSilence += 100
            Console.WriteLine("*******************")
            Console.WriteLine("Customer has gone " & theSilence / 1000 & " second(s) without responding.")
            Console.WriteLine("Silence Buffer is currently " & SilenceCap & " seconds.")
            Console.WriteLine("*******************")
            If (theSilence / 1000) > SilenceCap Then
                HandleSilence()
            End If
        End If
    End Sub
    Public Sub HandleSilence()

        theSilence = 0
        handleResponse()
    End Sub
    Dim inBetween As Integer = 0

End Class

Public Class responser
    Public Property response As String
    Public Property model As String
End Class