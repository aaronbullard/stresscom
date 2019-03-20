<!-- Module 12 -->  <!-- score.asp -->
<!-- #Include file="../inc/connect.inc" -->
<%
txtPageHeader = "Your Stresscom Results - Controlling the Emotions of Stress"
%>
<!-- #Include file="../inc/mod-top.asp" -->
<%
Select Case Request.QueryString("act")
	Case Else
		intStressScheduleID = Request.QueryString("tid")
%>
			<table>
				<tr>
					<td class="bodycopy" valign="top">
<p><strong>Controlling the Emotions of Stress</strong></p>
<p>Generally, the major issue connected with stress that leads to all kinds of other consequences is the feeling of loss of control. With loss of control, there are generally three major emotional results:  the hurry sickness, perfectibility syndrome, and anger in.</p>
<p><strong>Hurry Sickness</strong></p>
<p>Hurry sickness is something that results in exstress. We find ourselves caught up in a destructive pattern that causes us to always hurry. When this happens, we end up pressured, rushed, and feeling out of control. Unfortunately, hurrying through life usually has an unsuccessful result, causing more feelings of frustration, and we lose our ability to even relax.</p>
<p>We get addicted to the adrenaline rush that accompanies the hurry sickness. We like others to see us stressed because that makes us seem important in their eyes.  Many in this category have often grown up being constantly reminded and admonished for their lateness. They fear being late, and become nervous, stressed, and out of control. They often plan more events in one day than they can possibly complete, and end up rushed and out of time. Over-planning my also give them sense of self-importance.  At the end of some days, they are not where they would like to be and may feel physically sick.</p>
<p><strong>Perfectibility</strong></p>
<p>When you were younger, did you need to tear up one copy after another of your drawings because they were not perfect? Or, did you have a sense of not ever really being satisfied with how you did something? Well, if you did, you probably had a sense of perfectibility about yourself. In this crazy and ever-changing world, you need to recognize that perfection is hard to achieve.</p>
<p>On the other hand, we live in a world where we are bombarded with images of being perfect, to have the perfect hair, teeth, clothes, and weight. Every magazine encourages us to create the perfect home environment or meal. There are hundreds of tips on how to find the perfect mate, the perfect job, and the perfect vacation. Our society has set some high expectations of what the perfect life should look like, and how it can be obtained. Somehow when they are not perfect, we become very disappointed and angry. The belief that your best efforts were not ever good enough becomes a feeling of loss of control.</p>
<p><strong>Anger In</strong></p>
<p>The third area of exstress emotion is called "anger in." It is related to loss of control.  One university research center for behavioral medicine reports that people who are chronically angry have a four to seven times greater risk of heart attacks and cancer than those who are not anger prone, and that those who hold their anger inside and simmer in silence are at even greater risk. Anger causes a constriction of your cardiovascular system, and thus your blood pressure is increased, resulting in hypertension. "Anger in"— that is, hostility unexpressed— was found to be a major player in stress.</p>
<p>Anger is a very common emotion that we all experience. At times, some anger can motivate us to change things for the better, take up a just cause, defend our beliefs, or simply make us feel better because we let certain emotions out.</p>
<p>Anger that is seething just below the surface, however, can be unhealthy, and has three ways of expressing itself. Anger can be (1) "held in" and unexpressed; (2) passive-aggressive anger, which allows it to be expressed in a totally unrelated area of concern; or (3) expressed to others in a controlled communication.</p>
<p>Some of us lash out at others unexpectedly, or behave in destructive ways that isolate us from friends and family. Unexpressed anger can create frustration that can block our abilities to reason, or to find solutions for what is causing us to be angry. There are reasons that we get into unhealthy patterns of dealing with our anger, such as broken relationships, family crisis, financial or health issues, or any situation where resolution may be overwhelming. Most of all, anger takes away our happiness and peace of mind.</p>
<hr width="50%">
<hr>
<hr width="50%">
<p><button onClick="window.location='?mod=20&tid=<%= intStressScheduleID %>'">Continue to Techniques of Stress Management</button>
					</td>
				</tr>
			</table>
<%
End Select
%>
<!-- #Include file="../inc/mod-bot.asp" -->