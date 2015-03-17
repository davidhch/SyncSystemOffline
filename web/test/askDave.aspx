<%@ Page Language="VB" %>
<script Language="VB" Option="Explicit" runat="server">

	Sub Page_Load(Src as object, E as EventArgs)
		' Init default bounds the first time the page is loaded
		If Not Page.IsPostBack Then
			txtLowerBound.Text = 0
			txtUpperBound.Text = 100
		End If
	End Sub

	Sub btnSubmit_Click(Src as object, E as EventArgs)
		' Check validation status - the validation controls
		' contain all the logic needed.  Basically we check
		' to be sure we've got integers and that lower <= upper
		If Page.IsValid Then
			Dim intLowerBound, intUpperBound As Integer

			' Pull in values from the text boxes
			intLowerBound = CInt(txtLowerBound.Text)
			intUpperBound = CInt(txtUpperBound.Text)

			' Set the labels for output text to the values we just read in
			lblLowerBound.Text = intLowerBound
			lblUpperBound.Text = intUpperBound

			' Get the random number and display it in lblRandomNumber
			Dim nRNDValue
			nRNDValue = GetRandomNumberInRange(intLowerBound, intUpperBound)
			
			IF  txtQuestion.Text = "" Then
				lblRandomNumber.Text = "<br><img src=""dhstay.gif""><br>David Says....'GOWAN ASK AWAY'"
			Else
			If nRNDValue > 0 Then
				lblRandomNumber.Text = "<br><img src=""dhgo.gif""><br>NOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO!!!!!!  YOU LOSE - DAVIDS GONE!"
			End If
			If nRNDValue > 20 Then
				lblRandomNumber.Text = "<br><img src=""dhstay.gif""><br>David Says....'THATS A FIRST LINE SUPPORT LEVEL QUESTION - FECK OFF'"
			End If
			If nRNDValue > 30 Then
				lblRandomNumber.Text = "<br><img src=""dhstay.gif""><br>David Says....'PUT IT ON AN EMAIL SO I CAN LOOK AT IT'"
			End If
			If nRNDValue > 40 Then
				lblRandomNumber.Text = "<br><img src=""dhstay.gif""><br>David Says....'JUSTIN'S THE MAN TO TALK TO ABOUT THAT'"
			End If
			If nRNDValue > 70 Then
				lblRandomNumber.Text = "<br><img src=""dhstay.gif""><br>David Says....'THAT QUESTION SHOULD GO TO DARREN'"
			End If
			End If
			'lblRandomNumber.Text = nRNDValue
			txtQuestion.Text = ""
		End If
	End Sub

	Function GetRandomNumberInRange(intLowerBound As Integer, intUpperBound As Integer)
		Dim RandomGenerator As Random
		Dim intRandomNumber As Integer
		
		' Create and init the randon number generator
		RandomGenerator = New Random()
		'RandomGenerator = New Random(DateTime.Now.Millisecond)

		' Get the next random number
        intRandomNumber = RandomGenerator.Next(intLowerBound, intUpperBound + 1)

		' Return the random # as the function's return value
		GetRandomNumberInRange = intRandomNumber
	End Function

</script>

<html>
<head>
 <title>SHOULD I STAY OR SHOULD I GO</title>
</head>
<body>


<asp:Label id="lblLowerBound"  style="display:none" runat="server" />
<asp:Label id="lblUpperBound"  style="display:none" runat="server" />
<strong><asp:Label id="lblRandomNumber" runat="server" /></strong><br />

<form runat="server">
<asp:TextBox id="txtQuestion" size="50" maxlength="50" runat="server" />
<asp:TextBox id="txtLowerBound" size="5" maxlength="5" style="display:none" runat="server" />
<asp:TextBox id="txtUpperBound" size="5" maxlength="5" style="display:none" runat="server" />

<asp:Button id="btnSubmit" text="ASK DAVID A QUESTION QUESTION" onClick="btnSubmit_Click" runat="server" />

<asp:RangeValidator id="validtxtLowerBound" 
	ControlToValidate="txtLowerBound" 
	Display="Dynamic"
	MinimumValue="-9999"
	MaximumValue="10000" 
	Type="Integer" 
	Text="<p>Please enter a number between -9999 and 10000 for the lower bound.</p>"
	ForeColor="#FF0000" 
	runat="server" />
<asp:RangeValidator id="validtxtUpperBound" 
	ControlToValidate="txtUpperBound" 
	Display="Dynamic"
	MinimumValue="-9999"
	MaximumValue="10000" 
	Type="Integer" 
	Text="<p>Please enter a number between -9999 and 10000 for the upper bound.</p>"
	ForeColor="#FF0000" 
	runat="server" />
<asp:CompareValidator id="validtxtLowerUpperBound"
	ControlToValidate="txtUpperBound"
	ControlToCompare="txtLowerBound"
	Type="Integer"
	Operator="GreaterThanEqual"
	Text="<p>Please enter an upper bound that is larger then the lower bound.</p>"
	ForeColor="#FF0000"
	runat="server" />
</form>

</body>
</html>
