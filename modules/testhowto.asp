<!-- module 18 -->
<%
txtPageHeader = "Take the Stresscom"
%>
<!-- #Include file="../inc/mod-top.asp" -->
			<p align="justify"><strong>Directions for Taking Stresscom</strong></p>
			<p align="justify">Stresscom presents a series of statements that measure your ability to deal with stress over the long term.  Stresscom attempts to measure your stress in relation to your life, work, style, and personality.</p>
			<p align="justify">Please complete the test as truthfully as possible and reflecting your general feelings.  There are no right or wrong answers.  There is no time limit.</p>
			<p align="justify">Statements presented allow you to express your present opinion as follows:</p>
			<p align="justify"><strong>Example</strong></p>
			<div align="center">
			<table border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td class="TestStatementDirections">&nbsp;</td>
					<td class="TestHead" width="60">Strongly<br>Disagree</td>
					<td class="TestHead" width="60">Disagree</td>
					<td class="TestHead" width="60">Undecided</td>
					<td class="TestHead" width="60">Agree</td>
					<td class="TestHead" width="60">Strongly<br>Agree</td>
				</tr>				
				<tr>
					<td class="TestStatementDirections" valign="top">I enjoy changes in my life</td>
					<td class="TestHead"><input class="TestRadio" type="radio" value="1" name="st1"></td>
					<td class="TestHead"><input class="TestRadio" type="radio" value="2" name="st2"></td>
					<td class="TestHead"><input class="TestRadio" type="radio" value="3" name="st3"></td>
					<td class="TestHead"><input class="TestRadio" type="radio" value="4" name="st4"></td>
					<td class="TestHead"><input class="TestRadio" type="radio" value="5" name="st5"></td>
				</tr>
			</table>
			</div>
			<p align="justify">If you agree with the above statement, you would choose <strong>Agree</strong> or <strong>Strongly Agree</strong>.  If not, select <strong>Disagree</strong> or <strong>Strongly Disagree</strong>.  If you can't decide, select <strong>Undecided</strong></p>
			<p align="right"><button onClick="window.location='modules/ssfi.asp?tid=<%= Request.QueryString("tid") %>'">Continue the Stresscom Page 1 of 3</button></p>
			<p align="justify">&nbsp;</p>
		<!-- #Include file="../inc/mod-bot.asp" -->