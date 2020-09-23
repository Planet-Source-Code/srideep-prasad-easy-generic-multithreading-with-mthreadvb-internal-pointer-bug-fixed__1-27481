﻿<div align="center">

## Easy Generic Multithreading with MThreadVB \(Internal Pointer bug fixed\)


</div>

### Description

MThreadVB is a generic multithreader, allowing you to multithread any function or sub. To find out more, read on !
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Srideep Prasad](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/srideep-prasad.md)
**Level**          |Intermediate
**User Rating**    |4.9 (34 globes from 7 users)
**Compatibility**  |VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/srideep-prasad-easy-generic-multithreading-with-mthreadvb-internal-pointer-bug-fixed__1-27481/archive/master.zip)





### Source Code

<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Important Notice</title>
</head>
<body>
<p><b>MThreadVB - The easy way to multithread !</b></p>
<p>MThreadVB is a generic multithreader for VB - to which I have been making a
few changes here and there.... But it seems that in one of my updates I had
inadvertently referenced an independent DLL called VBConsole.Dll and had
forgotten to remove it.&nbsp; (This was done for testing and experimentation
purposes)...I had also forgotten to remove an invalid object variable
reference.... As a result, the update may not have worked.... I apologize for
any inconvenience and those of you who had downloaded the buggy code can
download the updated version now !&nbsp; Plus this new update has quite a few
more features (and took quite some time to add too !)</p>
<p><b><font color="#000080"><u>Fixes / Enhancements</u></font></b></p>
<p>1&gt;The VBConsole.dll reference problem has been fixed....</p>
<p>2&gt;Now defines a new property ObjectInThreadContext, that returns the
reference to the parent object containing the multithreaded sub in context to
the new thread</p>
<p>3&gt;With this, you can now implement File I/O and show forms (though I do
not very much recommend showing forms from multithreaded procedures), from multithreaded subs
(The Form show bug was reported by Robin Lobel - Special thanks to him for doing
so !)</p>
<p>4&gt;Some users it seems are having problems showing forms within
multithreaded procedures. Therefore I have updated the code to actually
demonstrate how to actually show forms from multithreaded procedures....</p>
<p>5&gt;A serious pointer dereferencing bug was causing problems when the
multithreaded sub had a relatively big name. This has now been fixed !</p>
<p>Here is the link to the bug fixed code -</p>
<p><a href="http://planet-source-code.com/vb/default.asp?lngCId=26900&amp;lngWId=1">http://planet-source-code.com/vb/default.asp?lngCId=26900&amp;lngWId=1</a>&nbsp;
</p>
<p>Do not hesitate to mail be if you notice some bug or problem....</p>
<p>Please remember that many of the enhancements were made possible due to
feedback from people at PSC.... Please continue to give your feedback regarding
any problems in the functioning of MThreadVB</p>
<p><b><u>Note: If you consider this code worth your vote, please vote for the main
page by clicking the link above... Or register it by clicking <a href="http://www.planet-source-code.com/vb/scripts/voting/VoteOnCodeRating.asp?lngWId=1&amp;txtCodeId=26900&amp;optCodeRatingValue=5">here
!</a></u></b></p>
</body>
</html>

